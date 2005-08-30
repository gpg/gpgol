/* engine-gpgme.c - Crypto engine with GPGME
 *	Copyright (C) 2005 g10 Code GmbH
 *
 * This file is part of GPGME Dialogs.
 *
 * GPGME Dialogs is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public License
 * as published by the Free Software Foundation; either version 2.1 
 * of the License, or (at your option) any later version.
 *  
 * GPGME Dialogs is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
 * General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public License
 * along with GPGME Dialogs; if not, write to the Free Software Foundation, 
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
      size_t n = 0;	
      *outbuf = gpgme_data_release_and_get_mem (out, &n);
      (*outbuf)[n] = 0;
      out = NULL;
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
      size_t n = 0;
      *outbuf = gpgme_data_release_and_get_mem (out, &n);
      (*outbuf)[n] = 0;
      out = NULL;
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
   with status outputs. */
int 
op_decrypt (const char *inbuf, char **outbuf, int ttl, const char *filename)
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
      size_t n = 0;
      gpgme_verify_result_t res;

      *outbuf = gpgme_data_release_and_get_mem (out, &n);
      (*outbuf)[n] = 0; /* Make sure it is really a string. */
      out = NULL; /* (That GPGME object is no any longer valid.) */

      /* Now check the state of any signature. */
      res = gpgme_op_verify_result (ctx);
      if (res && res->signatures)
        verify_dialog_box (res, filename);
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

/* Decrypt the stream INSTREAM directly to the stream OUTSTREAM.
   Returns 0 on success or an gpgme error code on failure.  If
   FILENAME is not NULL it will be displayed along with status
   outputs. */
int
op_decrypt_stream (LPSTREAM instream, LPSTREAM outstream, int ttl,
                   const char *filename)
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

  err = gpgme_data_new_from_cbs (&in, &cbs, instream);
  if (err)
    goto fail;

  err = gpgme_new (&ctx);
  if (err)
    goto fail;

  err = gpgme_data_new_from_cbs (&out, &cbs, outstream);
  if (err)
    goto fail;

  gpgme_set_passphrase_cb (ctx, passphrase_callback_box, &dk);
  dk.ctx = ctx;
  err = gpgme_op_decrypt (ctx, in, out);
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
  if (in)
    gpgme_data_release (in);
  if (out)
    gpgme_data_release (out);
  if (ctx)
    gpgme_release (ctx);
  return err;
}

/* Decrypt the stream INSTREAM directly to the newly allocated buffer OUTBUF.
   Returns 0 on success or an gpgme error code on failure.  If
   FILENAME is not NULL it will be displayed along with status
   outputs. */
int
op_decrypt_stream_to_buffer (LPSTREAM instream, char **outbuf, int ttl,
                             const char *filename)
{    
  struct decrypt_key_s dk;
  struct gpgme_data_cbs cbs;
  gpgme_data_t in = NULL;
  gpgme_data_t out = NULL;    
  gpgme_ctx_t ctx = NULL;
  gpgme_error_t err;
  
  *outbuf = NULL;

  memset (&cbs, 0, sizeof cbs);
  cbs.read = stream_read_cb;

  memset (&dk, 0, sizeof dk);
  dk.ttl = ttl;

  err = gpgme_data_new_from_cbs (&in, &cbs, instream);
  if (err)
    goto fail;

  err = gpgme_new (&ctx);
  if (err)
    goto fail;

  err = gpgme_data_new (&out);
  if (err)
    goto fail;

  gpgme_set_passphrase_cb (ctx, passphrase_callback_box, &dk);
  dk.ctx = ctx;
  err = gpgme_op_decrypt (ctx, in, out);
  dk.ctx = NULL;
  update_passphrase_cache (err, &dk);
  /* Act upon the result of the decryption operation. */
  if (!err) 
    {
      /* Decryption succeeded.  Store the result at OUTBUF. */
      size_t n = 0;
      gpgme_verify_result_t res;

      *outbuf = gpgme_data_release_and_get_mem (out, &n);
      (*outbuf)[n] = 0; /* Make sure it is really a string. */
      out = NULL; /* (That GPGME object is no any longer valid.) */

      /* Now check the state of the signatures. */
      res = gpgme_op_verify_result (ctx);
      if (res && res->signatures)
        verify_dialog_box (res, filename);
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
  if (in)
    gpgme_data_release (in);
  if (out)
    gpgme_data_release (out);
  if (ctx)
    gpgme_release (ctx);
  return err;
}



/* Verify a message in INBUF and return the new message (i.e. the one
   with stripped off dash escaping) in a newly allocated buffer
   OUTBUF. If OUTBUF is NULL only the verification result will be
   displayed (this is suitable for PGP/MIME messages).  A dialog box
   will show the result of the verification.  If FILENAME is not NULL
   it will be displayed along with status outputs. */
int
op_verify (const char *inbuf, char **outbuf, const char *filename)
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
      size_t n=0;
      if (outbuf) 
        {
          *outbuf = gpgme_data_release_and_get_mem (out, &n);
          (*outbuf)[n] = 0;
          out = NULL;
	}
      res = gpgme_op_verify_result (ctx);
    }
  if (res) 
    verify_dialog_box (res, filename);

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
   dialog box. */ 
int
op_verify_detached_sig (LPSTREAM data_stream,
                        const char *sig_string, const char *filename)
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


