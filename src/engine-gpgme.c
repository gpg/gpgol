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

static char *debug_file = NULL;
static int init_done = 0;

/* dummy */
struct _gpgme_engine_info  _gpgme_engine_ops_gpgsm;

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
    if (val > 0) {
	debug_file = (char *)xcalloc (1, strlen (file) + strlen (s) + 2);
	sprintf (debug_file, "%s=%d:%s", s, val, file);
	putenv (debug_file);
    }
    else {
	putenv ("GPGME_DEBUG=");
    }
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

  if (!gpgme_check_version (NEED_GPGME_VERSION)) {
      log_debug ("gpgme is too old (need %s, have %s)\n",
                 NEED_GPGME_VERSION, gpgme_check_version (NULL) );
      return -1;
  }

  err = gpgme_engine_check_version (GPGME_PROTOCOL_OpenPGP);
  if (err) {
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


/* Release a GPGME key array KEYS. */
static void
free_recipients (gpgme_key_t *keys)
{
  int i;
  
  if (!keys)
    return;
  for (i=0; keys[i]; i++)
    gpgme_key_release (keys[i]);
  xfree (keys);
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


int
op_encrypt_start (const char *inbuf, char **outbuf)
{
    gpgme_key_t *keys = NULL;
    int opts = 0;
    int err;

    recipient_dialog_box (&keys, &opts);
    if (opts & OPT_FLAG_CANCEL)
	return 0;
    err = op_encrypt ((void *)keys, inbuf, outbuf);
    free_recipients (keys);
    return err;
}


long 
ftello (FILE *f)
{
    /* XXX: find out if this is really needed */
    printf ("fd %d pos %ld\n", fileno(f), (long)ftell(f));
    return ftell (f);
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


/* FIXME: Why is RSET a void* ???*/
int
op_sign_encrypt_file (void *rset, const char *infile, const char *outfile)
{
  gpgme_data_t in = NULL;
  gpgme_data_t out = NULL;
  gpgme_ctx_t ctx = NULL;
  gpgme_error_t err;
  gpgme_key_t *keys = (gpgme_key_t*)rset;
  struct decrypt_key_s *hd;
  
  hd = xcalloc (1, sizeof *hd);

  err = gpgme_data_new_from_file (&in, infile, 1);
  if (err)
    goto fail;
    
  err = gpgme_new (&ctx);
  if (err)
    goto fail;
        
  err = gpgme_data_new (&out);
  if (err)
    goto fail;

  gpgme_set_passphrase_cb (ctx, passphrase_callback_box, hd);
  hd->ctx = ctx;    
  err = gpgme_op_encrypt_sign (ctx, keys, GPGME_ENCRYPT_ALWAYS_TRUST, in, out);
  hd->ctx = NULL;    
  update_passphrase_cache (err, hd);

  if (!err)
    err = data_to_file (&out, outfile);
    
 fail:
  if (ctx)
    gpgme_release (ctx);
  if (in)
    gpgme_data_release (in);
  if (out)
    gpgme_data_release (out);
  xfree (hd);
  return err;
}


/* Sign the file with name INFILE and write the output to a new file
   OUTFILE. PASSCB is the passphrase callback to use and DK the value
   we will pass to it.  MODE is one of the GPGME signing modes. */
static int
do_sign_file (gpgme_passphrase_cb_t pass_cb,
              struct decrypt_key_s *dk,
              int mode, const char *infile, const char *outfile)
{
  gpgme_data_t in = NULL;
  gpgme_data_t out = NULL;
  gpgme_ctx_t ctx = NULL;
  gpgme_error_t err;

  err = gpgme_data_new_from_file (&in, infile, 1);
  if (err)
    goto fail;    

  err = gpgme_new (&ctx);
  if (err)
    goto fail;
  
  gpgme_set_armor (ctx, 1);
  
  err = gpgme_data_new (&out);
  if (err)
    goto fail;

  gpgme_set_passphrase_cb (ctx, pass_cb, dk);
  dk->ctx = ctx;
  err = gpgme_op_sign (ctx, in, out, mode);
  dk->ctx = NULL;
  update_passphrase_cache (err, dk);
  
  if (!err)
    err = data_to_file (&out, outfile);

 fail:
  if (in)
    gpgme_data_release (in);
  if (out)
    gpgme_data_release (out);
  if (ctx)
    gpgme_release (ctx);    
  return err;
}


int
op_sign_file (int mode, const char *infile, const char *outfile, int ttl)
{
  struct decrypt_key_s *hd;
  gpgme_error_t err;
  
  hd = xcalloc (1, sizeof *hd);
  hd->ttl = ttl;
  hd->flags = 0x01;

  err = do_sign_file (passphrase_callback_box, hd, mode, infile, outfile);

  xfree (hd);
  return err;
}


/* Create a signature from INSTREAM and write it to OUTSTREAM.  Use
   signature mode MODE and a passphrase caching time of TTL. */
int
op_sign_stream (LPSTREAM instream, LPSTREAM outstream, int mode, int ttl)
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
  dk.flags = 0x01; /* fixme: Use a more macro for documentation reasons. */

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


/* Decrypt file INFILE into file OUTFILE. */
int
op_decrypt_file (const char *infile, const char *outfile)
{    
  gpgme_data_t in = NULL;
  gpgme_data_t out = NULL;    
  gpgme_ctx_t ctx = NULL;
  gpgme_error_t err;
  struct decrypt_key_s *dk;
  
  dk = xcalloc (1, sizeof *dk);

  err = gpgme_data_new_from_file (&in, infile, 1);
  if (err)
    goto fail;

  err = gpgme_new (&ctx);
  if (err)
    goto fail;

  err = gpgme_data_new (&out);
  if (err)
    goto fail;

  gpgme_set_passphrase_cb (ctx, passphrase_callback_box, dk);
  dk->ctx = ctx;
  err = gpgme_op_decrypt (ctx, in, out);
  dk->ctx = NULL;
  update_passphrase_cache (err, dk);

  if (!err)
    err = data_to_file (&out, outfile);

 fail:
  if (in)
    gpgme_data_release (in);
  if (out)
    gpgme_data_release (out);
  if (ctx)
    gpgme_release (ctx);
  xfree (dk);
  return err;
}



/* Decrypt the stream INSTREAM directly to the stream OUTSTREAM.
   Returns 0 on success or an gpgme error code on failure. */
int
op_decrypt_stream (LPSTREAM instream, LPSTREAM outstream)
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

 fail:
  if (in)
    gpgme_data_release (in);
  if (out)
    gpgme_data_release (out);
  if (ctx)
    gpgme_release (ctx);
  return err;
}


/* Try to figure out why the encryption failed and provide a suitable
   error code. */
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



/* Decrypt the data in file INFILE and write the ciphertext to the new
   file OUTFILE. */
/*FIXME: Why is RSET a void*???.  Why do we have this fucntion when
  there is already a version yusing streams? */
int
op_encrypt_file (void *rset, const char *infile, const char *outfile)
{
  gpgme_data_t in = NULL;
  gpgme_data_t out = NULL;
  gpgme_error_t err;
  gpgme_ctx_t ctx = NULL;

  err = gpgme_data_new_from_file (&in, infile, 1);
  if (err)
    goto fail;

  err = gpgme_new (&ctx);
  if (err)
    goto fail;

  err = gpgme_data_new (&out);
  if (err)
    goto fail;

  gpgme_set_armor (ctx, 1);
  err = gpgme_op_encrypt (ctx, (gpgme_key_t*)rset,
                          GPGME_ENCRYPT_ALWAYS_TRUST, in, out);
  if (!err)
    err = data_to_file (&out, outfile);
  else
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

/* Encrypt data from INFILE into the newly created file OUTFILE.  This
   version of the function internally works on streams to avoid
   copying data into memory buffers. */
/*FIXME: Why is RSET a void*??? */
int
op_encrypt_file_io (void *rset, const char *infile, const char *outfile)
{
  FILE *fin = NULL;
  FILE *fout = NULL;
  gpgme_data_t in = NULL;
  gpgme_data_t out = NULL;
  gpgme_error_t err;
  gpgme_ctx_t ctx = NULL;

  fin = fopen (infile, "rb");
  if (!fin)
    {
      err = GPG_ERR_UNKNOWN_ERRNO; /* FIXME: use errno */
      goto fail;
    }

  err = gpgme_data_new_from_stream (&in, fin);
  if (err)
    goto fail;

  fout = fopen (outfile, "wb");
  if (!fout)
    {
      err = GPG_ERR_UNKNOWN_ERRNO;
      goto fail;
    }

  err = gpgme_data_new_from_stream (&out, fout);
  if (err)
    goto fail;

  err = gpgme_new (&ctx);
  if (err)
    goto fail;
  
  gpgme_set_armor (ctx, 1);
  err = gpgme_op_encrypt (ctx, (gpgme_key_t*)rset,
                          GPGME_ENCRYPT_ALWAYS_TRUST, in, out);

 fail:
  if (fin)
    fclose (fin);
  /* FIXME: Shouldn't we remove the file on error here? */
  if (fout)
    fclose (fout);
  if (out)
    gpgme_data_release (out);
  if (in)
    gpgme_data_release (in);
  if (ctx)
    gpgme_release (ctx);
  return err;
}

/* Yet another encrypt versions; this time frommone memory buffer to a
   newly allocated one. */
/*FIXME: Why is RSET a void*??? */
int
op_encrypt (void *rset, const char *inbuf, char **outbuf)
{
  gpgme_key_t *keys = (gpgme_key_t *)rset;
  gpgme_data_t in = NULL;
  gpgme_data_t out = NULL;
  gpgme_error_t err;
  gpgme_ctx_t ctx = NULL;
    
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
  err = gpgme_op_encrypt (ctx, keys, GPGME_ENCRYPT_ALWAYS_TRUST, in, out);
  if (!err)
    {
      size_t n = 0;	
      *outbuf = gpgme_data_release_and_get_mem (out, &n);
      (*outbuf)[n] = 0;
      out = NULL;
    }
  else
    err = check_encrypt_result (ctx, err);

 leave:
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
/*FIXME: Why is RSET a void*??? */
int
op_sign_encrypt (void *rset, void *locusr, const char *inbuf, char **outbuf)
{
  gpgme_key_t *keys = (gpgme_key_t*)rset;
  gpgme_ctx_t ctx = NULL;
  gpgme_data_t in = NULL;
  gpgme_data_t out = NULL;
  gpgme_error_t err;
  struct decrypt_key_s *dk = NULL;

  op_init ();
  
  *outbuf = NULL;

  err = gpgme_new (&ctx);
  if (err)
    goto leave;

  dk = (struct decrypt_key_s *)xcalloc (1, sizeof *dk);
  dk->flags = 0x01;
  
  err = gpgme_data_new_from_mem (&in, inbuf, strlen (inbuf), 1);
  if (err)
    goto leave;

  err = gpgme_data_new (&out);
  if (err)
    goto leave;
  
  gpgme_set_passphrase_cb (ctx, passphrase_callback_box, dk);
  gpgme_set_armor (ctx, 1);
  gpgme_signers_add (ctx, (gpgme_key_t)locusr);
  err = gpgme_op_encrypt_sign (ctx, keys, GPGME_ENCRYPT_ALWAYS_TRUST, in, out);
  if (!err) 
    {
      size_t n = 0;
      *outbuf = gpgme_data_release_and_get_mem (out, &n);
      (*outbuf)[n] = 0;
      out = NULL;
    }
  else
    err = check_encrypt_result (ctx, err);

 leave:
  free_decrypt_key (dk);
  if (ctx)
    gpgme_release (ctx);
  if (out)
    gpgme_data_release (out);
  if (in)
    gpgme_data_release (in);
  return err;
}


static int
do_sign (gpgme_key_t locusr, const char *inbuf, char **outbuf)
{
  gpgme_error_t err;
  gpgme_data_t in = NULL;
  gpgme_data_t out = NULL;
  gpgme_ctx_t ctx = NULL;
  struct decrypt_key_s *dk = NULL;

  log_debug ("engine-gpgme.do_sign: enter\n");

  *outbuf = NULL;
  op_init ();
  
  err = gpgme_new (&ctx);
  if (err) 
    goto leave;

  dk = (struct decrypt_key_s *)xcalloc (1, sizeof *dk);
  dk->flags = 0x01;

  err = gpgme_data_new_from_mem (&in, inbuf, strlen (inbuf), 1);
  if (err)
    goto leave;

  err = gpgme_data_new (&out);
  if (err)
    goto leave;
    
  gpgme_set_passphrase_cb (ctx, passphrase_callback_box, dk);
  gpgme_signers_add (ctx, locusr);
  gpgme_set_textmode (ctx, 1);
  gpgme_set_armor (ctx, 1);
  err = gpgme_op_sign (ctx, in, out, GPGME_SIG_MODE_CLEAR);
  if (!err)
    {
      size_t n = 0;
      *outbuf = gpgme_data_release_and_get_mem (out, &n);
      (*outbuf)[n] = 0;
      out = NULL;
    }

 leave:
  free_decrypt_key (dk);
  if (ctx)
    gpgme_release (ctx);
  if (in)
    gpgme_data_release (in);
  if (out)
    gpgme_data_release (out);
  log_debug ("%s:%s: leave (rc=%d (%s))\n",
             __FILE__, __func__, err, gpgme_strerror (err));
  return err;
}



int
op_sign_start (const char *inbuf, char **outbuf)
{
  gpgme_key_t locusr = NULL;
  int err;

  log_debug ("engine-gpgme.op_sign_start: enter\n");
  err = signer_dialog_box (&locusr, NULL);
  if (err == -1)
    { /* Cancel */
      log_debug ("engine-gpgme.op_sign_start: leave (canceled)\n");
      return 0;
    }
  err = do_sign (locusr, inbuf, outbuf);
  log_debug ("engine-gpgme.op_sign_start: leave (rc=%d (%s))\n",
             err, gpg_strerror (err));
  return err;
}



/* Worker function for the decryption.  PASS_CB is the passphrase
   callback and PASS_CB_VALUE is passed as opaque argument to the
   callback.  INBUF is the string with ciphertext.  On success a newly
   allocated string with the plaintext will be stored at the address
   of OUTBUF and 0 returned.  On failure, NULL will be stored at
   OUTBUF and a gpgme error code is returned. */
static int
do_decrypt (gpgme_passphrase_cb_t pass_cb,
            struct decrypt_key_s *pass_cb_value,
	    const char *inbuf, char **outbuf)
{
  gpgme_data_t in = NULL;
  gpgme_data_t out = NULL;
  gpgme_ctx_t ctx;
  gpgme_error_t err;
  
  *outbuf = NULL;
  op_init ();

  err = gpgme_new (&ctx);
  if (err)
    return err;

  err = gpgme_data_new_from_mem (&in, inbuf, strlen (inbuf), 1);
  if (err)
    goto leave;
  err = gpgme_data_new (&out);
  if (err)
    goto leave;

  /* Set the callback and run the actual decryption routine.  We need
     to access the gpgme context in the callback, thus store it in the
     arg for the time of the call. */
  gpgme_set_passphrase_cb (ctx, pass_cb, pass_cb_value);
  pass_cb_value->ctx = ctx;    
  err = gpgme_op_decrypt_verify (ctx, in, out);
  pass_cb_value->ctx = NULL;    
  update_passphrase_cache (err, pass_cb_value);

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
        verify_dialog_box (res);
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

      /* XXX: automatically spawn verify_start in case the text was signed? */
    }


    /* If the callback indicated a cancel operation, clear the error. */
  if (pass_cb_value->opts & OPT_FLAG_CANCEL)
    err = 0;
  

leave:    
  gpgme_release (ctx);
  gpgme_data_release (in);
  gpgme_data_release (out);
  return err;
}


/* Run the decryption.  Decrypts INBUF to OUTBUF, caller must xfree
   the result at OUTBUF.  TTL is the time in seconds to cache a
   passphrase. */
int
op_decrypt_start (const char *inbuf, char **outbuf, int ttl)
{
  int err;
  struct decrypt_key_s *hd;
    
  hd = xcalloc (1, sizeof *hd);
  hd->ttl = ttl;
  err = do_decrypt (passphrase_callback_box, hd, inbuf, outbuf);

  free_decrypt_key (hd);
  return err;
}


/* Verify a message in INBUF and return the new message (i.e. the one
   with stripped off dash escaping) in a newly allocated buffer
   OUTBUF. IF OUTBUF is NULL only the verification result will be
   displayed (this is suitable for PGP/MIME messages).  A dialog box
   will show the result of the verification. */
int
op_verify_start (const char *inbuf, char **outbuf)
{
  gpgme_data_t in = NULL;
  gpgme_data_t out = NULL;
  gpgme_ctx_t ctx = NULL;
  gpgme_error_t err;
  gpgme_verify_result_t res = NULL;

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
    verify_dialog_box (res);

 leave:
  if (out)
    gpgme_data_release (out);
  if (in)
    gpgme_data_release (in);
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


const char*
op_strerror (int err)
{
    return gpgme_strerror (err);
}
