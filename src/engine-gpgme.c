/* engine-gpgme.c - Crypto engine with GPGME
 *	Copyright (C) 2005 g10 Code GmbH
 *
 * This file is part of GPGol.
 *
 * GPGol is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public License
 * as published by the Free Software Foundation; either version 2.1 
 * of the License, or (at your option) any later version.
 *  
 * GPGol is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
 * General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public License
 * along with GPGol; if not, write to the Free Software Foundation, 
 * Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA 
 */

/* Please note that we assume UTF-8 strings everywhere except when
   noted. */
   

#include <config.h>

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <time.h>
#include <errno.h>

#define COBJMACROS
#include <windows.h>
#include <objidl.h> /* For IStream. */

#include "gpgme.h"
#include "keycache.h"
#include "intern.h"
#include "passcache.h"
#include "engine.h"

#define TRACEPOINT() do { log_debug ("%s:%s:%d: tracepoint\n", \
                                       __FILE__, __func__, __LINE__); \
                        } while (0)


static char *debug_file = NULL;
static int init_done = 0;


static void add_verify_attestation (gpgme_data_t at, 
                                    gpgme_ctx_t ctx, 
                                    gpgme_verify_result_t res,
                                    const char *filename);



static void
cleanup (void)
{
  if (debug_file)
    {
      xfree (debug_file);
      debug_file = NULL;
    }
}


/* Enable or disable GPGME debug mode. */
void
op_set_debug_mode (int val, const char *file)
{
  const char *s= "GPGME_DEBUG";

  cleanup ();
  if (val > 0) 
    {
      debug_file = (char *)xcalloc (1, strlen (file) + strlen (s) + 2);
      sprintf (debug_file, "%s=%d;%s", s, val, file);
      putenv (debug_file);
    }
  else
    putenv ("GPGME_DEBUG=");
}


/* Cleanup static resources. */
void
op_deinit (void)
{
  cleanup ();
  cleanup_keycache_objects ();
}


/* Initialize the operation system. */
int
op_init (void)
{
  gpgme_error_t err;

  if (init_done == 1)
    return 0;

  if (!gpgme_check_version (NEED_GPGME_VERSION)) 
    {
      log_debug ("gpgme is too old (need %s, have %s)\n",
                 NEED_GPGME_VERSION, gpgme_check_version (NULL) );
      return -1;
    }

  err = gpgme_engine_check_version (GPGME_PROTOCOL_OpenPGP);
  if (err)
    {
      log_debug ("gpgme can't find a suitable OpenPGP backend: %s\n",
                 gpgme_strerror (err));
      return err;
    }
  
  /*init_keycache_objects ();*/
  init_done = 1;
  return 0;
}


/* The read callback used by GPGME to read data from an IStream object. */
static ssize_t
stream_read_cb (void *handle, void *buffer, size_t size)
{
  LPSTREAM stream = handle;
  HRESULT hr;
  ULONG nread;

  /* For EOF detection we assume that Read returns no error and thus
     nread will be 0.  The specs say that "Depending on the
     implementation, either S_FALSE or an error code could be returned
     when reading past the end of the stream"; thus we are not really
     sure whether our assumption is correct.  OTOH, at another place
     the docuemntation says that the implementation used by
     ISequentialStream exhibits the same EOF behaviour has found on
     the MSDOS FAT file system.  So we seem to have good karma. */
  hr = IStream_Read (stream, buffer, size, &nread);
  if (hr != S_OK)
    {
      log_debug ("%s:%s: Read failed: hr=%#lx", __FILE__, __func__, hr);
      errno = EIO;
      return -1;
    }
  return nread;
}

/* The write callback used by GPGME to write data to an IStream object. */
static ssize_t
stream_write_cb (void *handle, const void *buffer, size_t size)
{
  LPSTREAM stream = handle;
  HRESULT hr;
  ULONG nwritten;

  hr = IStream_Write (stream, buffer, size, &nwritten);
  if (hr != S_OK)
    {
      log_debug ("%s:%s: Write failed: hr=%#lx", __FILE__, __func__, hr);
      errno = EIO;
      return -1;
    }
  return nwritten;
}


/* This routine should be called immediately after an operation to
   make sure that the passphrase cache gets updated. ERR is expected
   to be the error code from the gpgme operation and PASS_CB_VALUE the
   context used by the passphrase callback.  

   On any error we flush a possible passphrase for the used keyID from
   the cache.  On success we store the passphrase into the cache.  The
   cache will take care of the supplied TTL and for example actually
   delete it if the TTL is 0 or an empty value is used. We also wipe
   the passphrase from the context here. */
static void
update_passphrase_cache (int err, struct decrypt_key_s *pass_cb_value)
{
  if (*pass_cb_value->keyid)
    {
      if (err)
        passcache_put (pass_cb_value->keyid, NULL, 0);
      else
        passcache_put (pass_cb_value->keyid, pass_cb_value->pass,
                       pass_cb_value->ttl);
    }
  if (pass_cb_value->pass)
    {
      wipestring (pass_cb_value->pass);
      xfree (pass_cb_value->pass);
      pass_cb_value->pass = NULL;
    }
}




/* Try to figure out why the encryption failed and provide a more
   suitable error code than the one returned by the encryption
   routine. */
static gpgme_error_t
check_encrypt_result (gpgme_ctx_t ctx, gpgme_error_t err)
{
  gpgme_encrypt_result_t res;

  res = gpgme_op_encrypt_result (ctx);
  if (!res)
    return err;
  if (res->invalid_recipients != NULL)
    return gpg_error (GPG_ERR_UNUSABLE_PUBKEY);
  /* XXX: we need to do more here! */
  return err;
}


/* Encrypt the data in INBUF into a newly malloced buffer stored on
   success at OUTBUF. The recipients are expected in the NULL
   terminated array KEYS. If SIGN_KEY is not NULl, the data will also
   be signed using this key.  TTL is the time the passphrase should be
   cached. */
int
op_encrypt (const char *inbuf, char **outbuf, gpgme_key_t *keys,
            gpgme_key_t sign_key, int ttl)
{
  struct decrypt_key_s dk;
  gpgme_data_t in = NULL;
  gpgme_data_t out = NULL;
  gpgme_error_t err;
  gpgme_ctx_t ctx = NULL;
    
  memset (&dk, 0, sizeof dk);
  dk.ttl = ttl;
  dk.flags = 0x01; /* FIXME: what is that? */

  *outbuf = NULL;

  op_init ();
  err = gpgme_new (&ctx);
  if (err)
    goto leave;

  err = gpgme_data_new_from_mem (&in, inbuf, strlen (inbuf), 1);
  if (err)
    goto leave;

  err = gpgme_data_new (&out);
  if (err)
    goto leave;

  gpgme_set_textmode (ctx, 1);
  gpgme_set_armor (ctx, 1);
  if (sign_key)
    {
      gpgme_set_passphrase_cb (ctx, passphrase_callback_box, &dk);
      dk.ctx = ctx;
      err = gpgme_signers_add (ctx, sign_key);
      if (!err)
        err = gpgme_op_encrypt_sign (ctx, keys, GPGME_ENCRYPT_ALWAYS_TRUST,
                                     in, out);
      dk.ctx = NULL;
      update_passphrase_cache (err, &dk);
    }
  else
    err = gpgme_op_encrypt (ctx, keys, GPGME_ENCRYPT_ALWAYS_TRUST, in, out);
  if (err)
    err = check_encrypt_result (ctx, err);
  else
    {
      /* Return the buffer but first make sure it is a string. */
      if (gpgme_data_write (out, "", 1) == 1)
        {
          *outbuf = gpgme_data_release_and_get_mem (out, NULL);
          out = NULL; 
        }
    }


 leave:
  if (ctx)
    gpgme_release (ctx);
  if (in)
    gpgme_data_release (in);
  if (out)
    gpgme_data_release (out);
  return err;
}



/* Encrypt the stream INSTREAM to the OUTSTREAM for all recpients
   given in the NULL terminated array KEYS.  If SIGN_KEY is not NULL
   the message will also be signed. */
int
op_encrypt_stream (LPSTREAM instream, LPSTREAM outstream, gpgme_key_t *keys,
                   gpgme_key_t sign_key, int ttl)
{
  struct decrypt_key_s dk;
  struct gpgme_data_cbs cbs;
  gpgme_data_t in = NULL;
  gpgme_data_t out = NULL;
  gpgme_ctx_t ctx = NULL;
  gpgme_error_t err;

  memset (&cbs, 0, sizeof cbs);
  cbs.read = stream_read_cb;
  cbs.write = stream_write_cb;

  memset (&dk, 0, sizeof dk);
  dk.ttl = ttl;
  dk.flags = 1;

  err = gpgme_data_new_from_cbs (&in, &cbs, instream);
  if (err)
    goto fail;

  err = gpgme_data_new_from_cbs (&out, &cbs, outstream);
  if (err)
    goto fail;

  err = gpgme_new (&ctx);
  if (err)
    goto fail;

  gpgme_set_armor (ctx, 1);
  /* FIXME:  We should not hardcode always trust. */
  if (sign_key)
    {
      gpgme_set_passphrase_cb (ctx, passphrase_callback_box, &dk);
      dk.ctx = ctx;
      err = gpgme_signers_add (ctx, sign_key);
      if (!err)
        err = gpgme_op_encrypt_sign (ctx, keys, GPGME_ENCRYPT_ALWAYS_TRUST,
                                     in, out);
      dk.ctx = NULL;
      update_passphrase_cache (err, &dk);
    }
  else
    err = gpgme_op_encrypt (ctx, keys, GPGME_ENCRYPT_ALWAYS_TRUST, in, out);
  if (err)
    err = check_encrypt_result (ctx, err);

 fail:
  if (ctx)
    gpgme_release (ctx);
  if (in)
    gpgme_data_release (in);
  if (out)
    gpgme_data_release (out);
  return err;
}


/* Sign and encrypt the data in INBUF into a newly allocated buffer at
   OUTBUF. */
int
op_sign (const char *inbuf, char **outbuf, int mode,
         gpgme_key_t sign_key, int ttl)
{
  struct decrypt_key_s dk;
  gpgme_error_t err;
  gpgme_data_t in = NULL;
  gpgme_data_t out = NULL;
  gpgme_ctx_t ctx = NULL;

  memset (&dk, 0, sizeof dk);
  dk.ttl = ttl;
  dk.flags = 1;

  *outbuf = NULL;
  op_init ();
  
  err = gpgme_new (&ctx);
  if (err) 
    goto leave;

  err = gpgme_data_new_from_mem (&in, inbuf, strlen (inbuf), 1);
  if (err)
    goto leave;

  err = gpgme_data_new (&out);
  if (err)
    goto leave;

  if (sign_key)
    gpgme_signers_add (ctx, sign_key);

  if (mode == GPGME_SIG_MODE_CLEAR)
    gpgme_set_textmode (ctx, 1);
  gpgme_set_armor (ctx, 1);

  gpgme_set_passphrase_cb (ctx, passphrase_callback_box, &dk);
  dk.ctx = ctx;
  err = gpgme_op_sign (ctx, in, out, mode);
  dk.ctx = NULL;
  update_passphrase_cache (err, &dk);

  if (!err)
    {
      /* Return the buffer but first make sure it is a string. */
      if (gpgme_data_write (out, "", 1) == 1)
        {
          *outbuf = gpgme_data_release_and_get_mem (out, NULL);
          out = NULL; 
        }
    }

 leave:
  if (ctx)
    gpgme_release (ctx);
  if (in)
    gpgme_data_release (in);
  if (out)
    gpgme_data_release (out);
  return err;
}


/* Create a signature from INSTREAM and write it to OUTSTREAM.  Use
   signature mode MODE and a passphrase caching time of TTL. */
int
op_sign_stream (LPSTREAM instream, LPSTREAM outstream, int mode,
                gpgme_key_t sign_key, int ttl)
{
  struct gpgme_data_cbs cbs;
  struct decrypt_key_s dk;
  gpgme_data_t in = NULL;
  gpgme_data_t out = NULL;
  gpgme_ctx_t ctx = NULL;
  gpgme_error_t err;
  
  memset (&cbs, 0, sizeof cbs);
  cbs.read = stream_read_cb;
  cbs.write = stream_write_cb;

  memset (&dk, 0, sizeof dk);
  dk.ttl = ttl;
  dk.flags = 0x01; /* fixme: Use a macro for documentation reasons. */

  err = gpgme_data_new_from_cbs (&in, &cbs, instream);
  if (err)
    goto fail;

  err = gpgme_data_new_from_cbs (&out, &cbs, outstream);
  if (err)
    goto fail;

  err = gpgme_new (&ctx);
  if (err)
    goto fail;
  
  if (sign_key)
    gpgme_signers_add (ctx, sign_key);

  if (mode == GPGME_SIG_MODE_CLEAR)
    gpgme_set_textmode (ctx, 1);
  gpgme_set_armor (ctx, 1);

  gpgme_set_passphrase_cb (ctx, passphrase_callback_box, &dk);
  dk.ctx = ctx;
  err = gpgme_op_sign (ctx, in, out, mode);
  dk.ctx = NULL;
  update_passphrase_cache (err, &dk);
  
 fail:
  if (in)
    gpgme_data_release (in);
  if (out)
    gpgme_data_release (out);
  if (ctx)
    gpgme_release (ctx);    
  return err;
}



/* Run the decryption.  Decrypts INBUF to OUTBUF, caller must xfree
   the result at OUTBUF.  TTL is the time in seconds to cache a
   passphrase.  If FILENAME is not NULL it will be displayed along
   with status outputs. If ATTESTATION is not NULL a text with the
   result of the signature verification will get printed to it. */
int 
op_decrypt (const char *inbuf, char **outbuf, int ttl, const char *filename,
            gpgme_data_t attestation)
{
  struct decrypt_key_s dk;
  gpgme_data_t in = NULL;
  gpgme_data_t out = NULL;
  gpgme_ctx_t ctx;
  gpgme_error_t err;
  
  *outbuf = NULL;
  op_init ();

  memset (&dk, 0, sizeof dk);
  dk.ttl = ttl;

  err = gpgme_new (&ctx);
  if (err)
    return err;

  err = gpgme_data_new_from_mem (&in, inbuf, strlen (inbuf), 1);
  if (err)
    goto leave;
  err = gpgme_data_new (&out);
  if (err)
    goto leave;

  gpgme_set_passphrase_cb (ctx, passphrase_callback_box, &dk);
  dk.ctx = ctx;
  err = gpgme_op_decrypt_verify (ctx, in, out);
  dk.ctx = NULL;
  update_passphrase_cache (err, &dk);

  /* Act upon the result of the decryption operation. */
  if (!err) 
    {
      /* Decryption succeeded.  Store the result at OUTBUF. */
      gpgme_verify_result_t res;

      /* Return the buffer but first make sure it is a string. */
      if (gpgme_data_write (out, "", 1) == 1)
        {
          *outbuf = gpgme_data_release_and_get_mem (out, NULL);
          out = NULL; 
        }

      /* Now check the state of any signature. */
      res = gpgme_op_verify_result (ctx);
      if (res && res->signatures)
        verify_dialog_box (res, filename);
      if (res && res->signatures && attestation)
        add_verify_attestation (attestation, ctx, res, filename);
    }
  else if (gpgme_err_code (err) == GPG_ERR_DECRYPT_FAILED)
    {
      /* The decryption failed.  See whether we can determine the real
         problem. */
      gpgme_decrypt_result_t res;
      res = gpgme_op_decrypt_result (ctx);
      if (res != NULL && res->recipients != NULL &&
          gpgme_err_code (res->recipients->status) == GPG_ERR_NO_SECKEY)
        err = GPG_ERR_NO_SECKEY;
      /* XXX: return the keyids */
    }
  else
    {
      /* Decryption failed for other reasons. */
    }


  /* If the callback indicated a cancel operation, set the error
     accordingly. */
  if (err && (dk.opts & OPT_FLAG_CANCEL))
    err = gpg_error (GPG_ERR_CANCELED);
  
leave:    
  if (ctx)
    gpgme_release (ctx);
  if (in)
    gpgme_data_release (in);
  if (out)
    gpgme_data_release (out);
  return err;
}


/* Decrypt the GPGME data object IN into the data object OUT.  Returns
   0 on success or an gpgme error code on failure.  If FILENAME is not
   NULL it will be displayed along with status outputs. If ATTESTATION
   is not NULL a text with the result of the signature verification
   will get printed to it. */
static int
decrypt_stream (gpgme_data_t in, gpgme_data_t out, int ttl,
                const char *filename, gpgme_data_t attestation)
{    
  struct decrypt_key_s dk;
  gpgme_ctx_t ctx = NULL;
  gpgme_error_t err;
  
  memset (&dk, 0, sizeof dk);
  dk.ttl = ttl;

  err = gpgme_new (&ctx);
  if (err)
    goto fail;

  gpgme_set_passphrase_cb (ctx, passphrase_callback_box, &dk);
  dk.ctx = ctx;
  err = gpgme_op_decrypt_verify (ctx, in, out);
  dk.ctx = NULL;
  update_passphrase_cache (err, &dk);
  /* Act upon the result of the decryption operation. */
  if (!err) 
    {
      gpgme_verify_result_t res;

      /* Decryption succeeded.  Now check the state of the signatures. */
      res = gpgme_op_verify_result (ctx);
      if (res && res->signatures)
        verify_dialog_box (res, filename);
      if (res && res->signatures && attestation)
        add_verify_attestation (attestation, ctx, res, filename);
    }
  else if (gpgme_err_code (err) == GPG_ERR_DECRYPT_FAILED)
    {
      /* The decryption failed.  See whether we can determine the real
         problem. */
      gpgme_decrypt_result_t res;
      res = gpgme_op_decrypt_result (ctx);
      if (res != NULL && res->recipients != NULL &&
          gpgme_err_code (res->recipients->status) == GPG_ERR_NO_SECKEY)
        err = GPG_ERR_NO_SECKEY;
      /* XXX: return the keyids */
    }
  else
    {
      /* Decryption failed for other reasons. */
    }


  /* If the callback indicated a cancel operation, set the error
     accordingly. */
  if (err && (dk.opts & OPT_FLAG_CANCEL))
    err = gpg_error (GPG_ERR_CANCELED);

 fail:
  if (ctx)
    gpgme_release (ctx);
  return err;
}

/* Decrypt the stream INSTREAM directly to the stream OUTSTREAM.
   Returns 0 on success or an gpgme error code on failure.  If
   FILENAME is not NULL it will be displayed along with status
   outputs. */
int
op_decrypt_stream (LPSTREAM instream, LPSTREAM outstream, int ttl,
                   const char *filename, gpgme_data_t attestation)
{
  struct gpgme_data_cbs cbs;
  gpgme_data_t in = NULL;
  gpgme_data_t out = NULL;    
  gpgme_error_t err;
  
  memset (&cbs, 0, sizeof cbs);
  cbs.read = stream_read_cb;
  cbs.write = stream_write_cb;

  err = gpgme_data_new_from_cbs (&in, &cbs, instream);
  if (!err)
    err = gpgme_data_new_from_cbs (&out, &cbs, outstream);
  if (!err)
    err = decrypt_stream (in, out, ttl, filename, attestation);

  if (in)
    gpgme_data_release (in);
  if (out)
    gpgme_data_release (out);
  return err;
}


/* Decrypt the stream INSTREAM directly to the newly allocated buffer OUTBUF.
   Returns 0 on success or an gpgme error code on failure.  If
   FILENAME is not NULL it will be displayed along with status
   outputs. */
int
op_decrypt_stream_to_buffer (LPSTREAM instream, char **outbuf, int ttl,
                             const char *filename, gpgme_data_t attestation)
{    
  struct gpgme_data_cbs cbs;
  gpgme_data_t in = NULL;
  gpgme_data_t out = NULL;    
  gpgme_error_t err;
  
  *outbuf = NULL;

  memset (&cbs, 0, sizeof cbs);
  cbs.read = stream_read_cb;

  err = gpgme_data_new_from_cbs (&in, &cbs, instream);
  if (!err)
    err = gpgme_data_new (&out);
  if (!err)
    err = decrypt_stream (in, out, ttl, filename, attestation);
  if (!err)
    {
      /* Return the buffer but first make sure it is a string. */
      if (gpgme_data_write (out, "", 1) == 1)
        {
          *outbuf = gpgme_data_release_and_get_mem (out, NULL);
          out = NULL; 
        }
    }

  if (in)
    gpgme_data_release (in);
  if (out)
    gpgme_data_release (out);
  return err;
}


/* Decrypt the stream INSTREAM directly to the GPGME data object OUT.
   Returns 0 on success or an gpgme error code on failure.  If
   FILENAME is not NULL it will be displayed along with status
   outputs. */
int
op_decrypt_stream_to_gpgme (LPSTREAM instream, gpgme_data_t out, int ttl,
                            const char *filename, gpgme_data_t attestation)
{
  struct gpgme_data_cbs cbs;
  gpgme_data_t in = NULL;
  gpgme_error_t err;
  
  memset (&cbs, 0, sizeof cbs);
  cbs.read = stream_read_cb;

  err = gpgme_data_new_from_cbs (&in, &cbs, instream);
  if (!err)
    err = decrypt_stream (in, out, ttl, filename, attestation);

  if (in)
    gpgme_data_release (in);
  return err;
}



/* Verify a message in INBUF and return the new message (i.e. the one
   with stripped off dash escaping) in a newly allocated buffer
   OUTBUF. If OUTBUF is NULL only the verification result will be
   displayed (this is suitable for PGP/MIME messages).  A dialog box
   will show the result of the verification.  If FILENAME is not NULL
   it will be displayed along with status outputs.  If ATTESTATION is
   not NULL a text with the result of the signature verification will
   get printed to it. */
int
op_verify (const char *inbuf, char **outbuf, const char *filename,
           gpgme_data_t attestation)
{
  gpgme_data_t in = NULL;
  gpgme_data_t out = NULL;
  gpgme_ctx_t ctx = NULL;
  gpgme_error_t err;
  gpgme_verify_result_t res = NULL;

  if (outbuf)
    *outbuf = NULL;

  op_init ();

  err = gpgme_new (&ctx);
  if (err)
    goto leave;

  err = gpgme_data_new_from_mem (&in, inbuf, strlen (inbuf), 1);
  if (err)
    goto leave;

  err = gpgme_data_new (&out);
  if (err)
    goto leave;

  err = gpgme_op_verify (ctx, in, NULL, out);
  if (!err)
    {
      if (outbuf) 
        {
          /* Return the buffer but first make sure it is a string. */
          if (gpgme_data_write (out, "", 1) == 1)
            {
              *outbuf = gpgme_data_release_and_get_mem (out, NULL);
              out = NULL; 
            }
	}
      res = gpgme_op_verify_result (ctx);
    }
  if (res) 
    verify_dialog_box (res, filename);
  if (res && attestation)
    add_verify_attestation (attestation, ctx, res, filename);

 leave:
  if (out)
    gpgme_data_release (out);
  if (in)
    gpgme_data_release (in);
  if (ctx)
    gpgme_release (ctx);
  return err;
}

/* Verify a detached message where the data is to be read from the
   DATA_STREAM and the signature itself is expected to be the string
   SIG_STRING.  FILENAME will be shown by the verification status
   dialog box.  If ATTESTATION is not NULL a text with the result of
   the signature verification will get printed to it.  */ 
int
op_verify_detached_sig (LPSTREAM data_stream,
                        const char *sig_string, const char *filename,
                        gpgme_data_t attestation)
{
  struct gpgme_data_cbs cbs;
  gpgme_data_t data = NULL;
  gpgme_data_t sig = NULL;
  gpgme_ctx_t ctx = NULL;
  gpgme_error_t err;
  gpgme_verify_result_t res = NULL;

  memset (&cbs, 0, sizeof cbs);
  cbs.read = stream_read_cb;
  cbs.write = stream_write_cb;

  op_init ();

  err = gpgme_new (&ctx);
  if (err)
    goto leave;

  err = gpgme_data_new_from_cbs (&data, &cbs, data_stream);
  if (err)
    goto leave;

  err = gpgme_data_new_from_mem (&sig, sig_string, strlen (sig_string), 0);
  if (err)
    goto leave;

  err = gpgme_op_verify (ctx, sig, data, NULL);
  if (!err)
    {
      res = gpgme_op_verify_result (ctx);
      if (res) 
        verify_dialog_box (res, filename);
      if (res && attestation)
        add_verify_attestation (attestation, ctx, res, filename);
    }

 leave:
  if (data)
    gpgme_data_release (data);
  if (sig)
    gpgme_data_release (sig);
  if (ctx)
    gpgme_release (ctx);
  return err;
}



static void
at_puts (gpgme_data_t a, const char *s)
{
  gpgme_data_write (a, s, strlen (s));
}

static void 
at_print_time (gpgme_data_t a, time_t t)
{
  char buf[200];

  strftime (buf, sizeof (buf)-1, "%c", localtime (&t));
  at_puts (a, buf);
}

static void 
at_fingerprint (gpgme_data_t a, gpgme_key_t key)
{
  const char *s;
  int i, is_pgp;
  char *buf, *p;
  const char *prefix = _("Fingerprint: ");

  if (!key)
    return;
  s = key->subkeys ? key->subkeys->fpr : NULL;
  if (!s)
    return;
  is_pgp = (key->protocol == GPGME_PROTOCOL_OpenPGP);

  buf = xmalloc ( strlen (prefix) + strlen(s) * 4 + 2 );
  p = stpcpy (buf, prefix);
  if (is_pgp && strlen (s) == 40)
    { 
      /* v4 style formatted. */
      for (i=0; *s && s[1] && s[2] && s[3] && s[4]; s += 4, i++)
        {
          *p++ = s[0];
          *p++ = s[1];
          *p++ = s[2];
          *p++ = s[3];
          *p++ = ' ';
          if (i == 4)
            *p++ = ' ';
        }
    }
  else
    { 
      /* v3 style or X.509 formatted. */
      for (i=0; *s && s[1] && s[2]; s += 2, i++)
        {
          *p++ = s[0];
          *p++ = s[1];
          *p++ = is_pgp? ' ':':';
          if (is_pgp && i == 7)
            *p++ = ' ';
        }
    }

  /* Just in case print remaining odd digits */
  for (; *s; s++)
    *p++ = *s;
  *p++ = '\n';
  *p = 0;
  at_puts (a, buf);
  xfree (buf);
}


/* Print common attributes of the signature summary SUM.  Returns
   trues if a severe warning has been encountered. */
static int 
at_sig_summary (gpgme_data_t a,  
                unsigned long sum, gpgme_signature_t sig, gpgme_key_t key)
{
  int severe = 0;

  if ((sum & GPGME_SIGSUM_VALID))
    at_puts (a, _("This signature is valid\n"));
  if ((sum & GPGME_SIGSUM_GREEN))
    at_puts (a, _("signature state is \"green\"\n"));
  if ((sum & GPGME_SIGSUM_RED))
    at_puts (a, _("signature state is \"red\"\n"));

  if ((sum & GPGME_SIGSUM_KEY_REVOKED))
    {
      at_puts (a, _("Warning: One of the keys has been revoked\n"));
      severe = 1;
    }
  
  if ((sum & GPGME_SIGSUM_KEY_EXPIRED))
    {
      time_t t = key->subkeys->expires ? key->subkeys->expires : 0;

      if (t)
        {
          at_puts (a, _("Warning: The key used to create the "
                        "signature expired at: "));
          at_print_time (a, t);
          at_puts (a, "\n");
        }
      else
        at_puts (a, _("Warning: At least one certification key "
                      "has expired\n"));
    }

  if ((sum & GPGME_SIGSUM_SIG_EXPIRED))
    {
      at_puts (a, _("Warning: The signature expired at: "));
      at_print_time (a, sig ? sig->exp_timestamp : 0);
      at_puts (a, "\n");
    }

  if ((sum & GPGME_SIGSUM_KEY_MISSING))
    at_puts (a, _("Can't verify due to a missing key or certificate\n"));

  if ((sum & GPGME_SIGSUM_CRL_MISSING))
    {
      at_puts (a, _("The CRL is not available\n"));
      severe = 1;
    }

  if ((sum & GPGME_SIGSUM_CRL_TOO_OLD))
    {
      at_puts (a, _("Available CRL is too old\n"));
      severe = 1;
    }

  if ((sum & GPGME_SIGSUM_BAD_POLICY))
    at_puts (a, _("A policy requirement was not met\n"));

  if ((sum & GPGME_SIGSUM_SYS_ERROR))
    {
      const char *t0 = NULL, *t1 = NULL;

      at_puts (a, _("A system error occured"));

      /* Try to figure out some more detailed system error information. */
      if (sig)
	{
	  t0 = "";
	  t1 = sig->wrong_key_usage ? "Wrong_Key_Usage" : "";
	}

      if (t0 || t1)
        {
          at_puts (a, ": ");
          if (t0)
            at_puts (a, t0);
          if (t1 && !(t0 && !strcmp (t0, t1)))
            {
              if (t0)
                at_puts (a, ",");
              at_puts (a, t1);
            }
        }
      at_puts (a, "\n");
    }

  return severe;
}


/* Print the validity of a key used for one signature. */
static void 
at_sig_validity (gpgme_data_t a, gpgme_signature_t sig)
{
  const char *txt = NULL;

  switch (sig ? sig->validity : 0)
    {
    case GPGME_VALIDITY_UNKNOWN:
      txt = _("WARNING: We have NO indication whether "
              "the key belongs to the person named "
              "as shown above\n");
      break;
    case GPGME_VALIDITY_UNDEFINED:
      break;
    case GPGME_VALIDITY_NEVER:
      txt = _("WARNING: The key does NOT BELONG to "
              "the person named as shown above\n");
      break;
    case GPGME_VALIDITY_MARGINAL:
      txt = _("WARNING: It is NOT certain that the key "
              "belongs to the person named as shown above\n");
      break;
    case GPGME_VALIDITY_FULL:
    case GPGME_VALIDITY_ULTIMATE:
      txt = NULL;
      break;
    }

  if (txt)
    at_puts (a, txt);
}


/* Print a text with the attestation of the signature verification
   (which is in RES) to A.  FILENAME may also be used in the
   attestation. */
static void
add_verify_attestation (gpgme_data_t a, gpgme_ctx_t ctx,
                        gpgme_verify_result_t res, const char *filename)
{
  time_t created;
  const char *fpr, *uid;
  gpgme_key_t key = NULL;
  int i, anybad = 0, anywarn = 0;
  unsigned int sum;
  gpgme_user_id_t uids = NULL;
  gpgme_signature_t sig;
  gpgme_error_t err;

  if (!gpgme_data_seek (a, 0, SEEK_CUR))
    {
      /* Nothing yet written to the stream.  Insert the curretn time. */
      at_puts (a, _("Verification started at: "));
      at_print_time (a, time (NULL));
      at_puts (a, "\n\n");
    }

  at_puts (a, _("Verification result for: "));
  at_puts (a, filename ? filename : _("[unnamed part]"));
  at_puts (a, "\n");
  if (res)
    {
      for (sig = res->signatures; sig; sig = sig->next)
        {
          created = sig->timestamp;
          fpr = sig->fpr;
          sum = sig->summary;

          if (gpg_err_code (sig->status) != GPG_ERR_NO_ERROR)
            anybad = 1;

          err = gpgme_get_key (ctx, fpr, &key, 0);
          uid = !err && key->uids && key->uids->uid ? key->uids->uid : "[?]";

          if ((sum & GPGME_SIGSUM_GREEN))
            {
              at_puts (a, _("Good signature from: "));
              at_puts (a, uid);
              at_puts (a, "\n");
              for (i = 1, uids = key->uids; uids; i++, uids = uids->next)
                {
                  if (uids->revoked)
                    continue;
                  at_puts (a, _("                aka: "));
                  at_puts (a, uids->uid);
                  at_puts (a, "\n");
                }
              at_puts (a, _("            created: "));
              at_print_time (a, created);
              at_puts (a, "\n");
              if (at_sig_summary (a, sum, sig, key))
                anywarn = 1;
              at_sig_validity (a, sig);
            }
          else if ((sum & GPGME_SIGSUM_RED))
            {
              at_puts (a, _("*BAD* signature claimed to be from: "));
              at_puts (a, uid);
              at_puts (a, "\n");
              at_sig_summary (a, sum, sig, key);
            }
          else if (!anybad && key && (key->protocol == GPGME_PROTOCOL_OpenPGP))
            { /* We can't decide (yellow) but this is a PGP key with a
                 good signature, so we display what a PGP user
                 expects: The name, fingerprint and the key validity
                 (which is neither fully or ultimate). */
              at_puts (a, _("Good signature from: "));
              at_puts (a, uid);
              at_puts (a, "\n");
              at_puts (a, _("            created: "));
              at_print_time (a, created);
              at_puts (a, "\n");
              at_sig_validity (a, sig);
              at_fingerprint (a, key);
              if (at_sig_summary (a, sum, sig, key))
                anywarn = 1;
            }
          else /* can't decide (yellow) */
            {
              at_puts (a, _("Error checking signature"));
              at_puts (a, "\n");
              at_sig_summary (a, sum, sig, key);
            }
          
          gpgme_key_release (key);
        }

      if (!anybad )
        {
          gpgme_sig_notation_t notation;
          
          for (sig = res->signatures; sig; sig = sig->next)
            {
              if (!sig->notations)
                continue;
              at_puts (a, _("*** Begin Notation (signature by: "));
              at_puts (a, sig->fpr);
              at_puts (a, ") ***\n");
              for (notation = sig->notations; notation;
                   notation = notation->next)
                {
                  if (notation->name)
                    {
                      at_puts (a, notation->name);
                      at_puts (a, "=");
                    }
                  if (notation->value)
                    {
                      at_puts (a, notation->value);
                      if (!(*notation->value
                            && (notation->value[strlen (notation->value)-1]
                                =='\n')))
                        at_puts (a, "\n");
                    }
                }
              at_puts (a, _("*** End Notation ***\n"));
            }
	}
    }
  at_puts (a, "\n");
}




/* Try to find a key for each item in array NAMES. If one ore more
   items were not found, they are stored as malloced strings to the
   newly allocated array UNKNOWN at the corresponding position.  Found
   keys are stored in the newly allocated array KEYS. If N is not NULL
   the total number of items will be stored at that address.  Note,
   that both UNKNOWN may have NULL entries inbetween. The fucntion
   returns the nuber of keys not found. Caller needs to releade KEYS
   and UNKNOWN. 

   FIXME: The calling convetion is far to complicated.  Needs to be revised.

*/
int 
op_lookup_keys (char **names, gpgme_key_t **keys, char ***unknown, size_t *n)
{
    int i, pos=0;
    gpgme_key_t k;

    for (i=0; names[i]; i++)
	;
    if (n)
	*n = i;
    *unknown = (char **)xcalloc (i+1, sizeof (char*));
    *keys = (gpgme_key_t *)xcalloc (i+1, sizeof (gpgme_key_t));
    for (i=0; names[i]; i++) {
	/*k = find_gpg_email(id[i]);*/
	k = get_gpg_key (names[i]);
	if (!k)
	    (*unknown)[pos++] = xstrdup (names[i]);
	else
	    (*keys)[i] = k;
    }
    return (i-pos);
}



/* Copy the data from the GPGME object DAT to a newly created file
   with name OUTFILE.  Returns 0 on success. */
static gpgme_error_t
data_to_file (gpgme_data_t *dat, const char *outfile)
{
  FILE *out;
  char *buf;
  size_t n=0;

  out = fopen (outfile, "wb");
  if (!out)
    return GPG_ERR_UNKNOWN_ERRNO; /* FIXME: We need to check why we
                                     can't use errno here. */
  /* FIXME: Why at all are we using an in memory object wqhen we are
     later going to write to a file anyway. */
  buf = gpgme_data_release_and_get_mem (*dat, &n);
  *dat = NULL;
  if (!n)
    {
      fclose (out);
      return GPG_ERR_EOF; /* FIXME:  wrap this into a gpgme_error() */
    }
  fwrite (buf, 1, n, out);
  fclose (out);
  /* FIXME: We have no error checking above. */
  xfree (buf);
  return 0;
}


int
op_export_keys (const char *pattern[], const char *outfile)
{      
    /* @untested@ */
    gpgme_ctx_t ctx=NULL;
    gpgme_data_t  out=NULL;    
    gpgme_error_t err;

    err = gpgme_new (&ctx);
    if (err)
	return err;
    err = gpgme_data_new (&out);
    if (err) {
	gpgme_release (ctx);
	return err;
    }

    gpgme_set_armor (ctx, 1);
    err = gpgme_op_export_ext (ctx, pattern, 0, out);
    if (!err)
	data_to_file (&out, outfile);

    gpgme_data_release (out);  
    gpgme_release (ctx);  
    return err;
}


const char *
userid_from_key (gpgme_key_t k)
{
  if (k && k->uids && k->uids->uid)
    return k->uids->uid;
  else
    return "?";
}

const char *
keyid_from_key (gpgme_key_t k)
{
  
  if (k && k->subkeys && k->subkeys->keyid)
    return k->subkeys->keyid;
  else
    return "????????";
}


const char*
op_strerror (int err)
{
    return gpgme_strerror (err);
}


