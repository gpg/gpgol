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
#include <windows.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <time.h>

#include "gpgme.h"
#include "intern.h"
#include "engine.h"
#include "keycache.h"

static char *debug_file = NULL;
static int init_done = 0;

/* dummy */
struct _gpgme_engine_info  _gpgme_engine_ops_gpgsm;

static void
cleanup (void)
{
    if (debug_file) {
	free (debug_file);
	debug_file = NULL;
    }
}


/* enable or disable GPGME debug mode. */
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
  gpgme_check_version (NULL);
  err = gpgme_engine_check_version (GPGME_PROTOCOL_OpenPGP);
  if (err)
      return err;
  /*init_keycache_objects ();*/
  init_done = 1;
  return 0;
}



int
op_encrypt_start (const char *inbuf, char **outbuf)
{
    gpgme_key_t *keys = NULL;
    int opts = 0;
    int err;

    recipient_dialog_box (&keys, &opts);
    err = op_encrypt ((void *)keys, inbuf, outbuf);
    free (keys);
    return err;
}

long 
ftello (FILE *f)
{
    /* XXX: find out if this is really needed */
    printf ("fd %d pos %d\n", fileno(f), ftell(f));
    return ftell (f);
}


gpgme_error_t
data_to_file (gpgme_data_t *dat, const char *outfile)
{
    FILE *out;
    char *buf;
    size_t n=0;

    out = fopen (outfile, "wb");
    if (!out)
	return GPG_ERR_UNKNOWN_ERRNO;
    buf = gpgme_data_release_and_get_mem (*dat, &n);
    *dat = NULL;
    if (n == 0) {
	fclose (out);
	return GPG_ERR_EOF;
    }
    fwrite (buf, 1, n, out);
    fclose (out);
    xfree (buf);
    return 0;
}


int
op_sign_encrypt_file (void *rset, const char *infile, const char *outfile)
{
    /* XXX: fill it with life */
    return 0;
}


int
op_sign_file (int mode, const char *infile, const char *outfile)
{
    gpgme_data_t in=NULL;
    gpgme_data_t out=NULL;
    gpgme_ctx_t ctx=NULL;
    gpgme_error_t err;
    struct decrypt_key_s *hd;

    err = gpgme_data_new_from_file (&in, infile, 1);
    if (err)
	return err;

    hd = xcalloc (1, sizeof *hd);
    err = gpgme_new (&ctx);
    if (err)
	goto fail;

    gpgme_set_passphrase_cb (ctx, passphrase_callback_box, hd);

    err = gpgme_data_new (&out);
    if (err)
	goto fail;

    err = gpgme_op_sign (ctx, in, out, mode);
    if (!err)
	err = data_to_file (&out, outfile);

fail:
    if (in != NULL)
	gpgme_data_release (in);
    if (out != NULL)
	gpgme_data_release (out);
    if (ctx != NULL)
	gpgme_release (ctx);
    xfree (hd);
    return err;
}


int
op_decrypt_file (const char *infile, const char *outfile)
{
    gpgme_data_t in = NULL;
    gpgme_data_t out = NULL;
    gpgme_error_t err;
    gpgme_ctx_t ctx = NULL;
    struct decrypt_key_s *hd;

    err = gpgme_data_new_from_file (&in, infile, 1);
    if (err)
	return err;

    hd = xcalloc (1, sizeof *hd);
    err = gpgme_new (&ctx);
    if (err)
	goto fail;
    gpgme_set_passphrase_cb (ctx, passphrase_callback_box, hd);

    err = gpgme_data_new (&out);
    if (err)
	goto fail;

    err = gpgme_op_decrypt (ctx, in, out);
    if (!err)
	err = data_to_file (&out, outfile);

fail:
    if (in != NULL)
	gpgme_data_release (in);
    if (out != NULL)
	gpgme_data_release (out);
    if (ctx != NULL)
	gpgme_release (ctx);
    xfree (hd);
    return err;
}


int
op_encrypt_file (void *rset, const char *infile, const char *outfile)
{
    gpgme_data_t in=NULL;
    gpgme_data_t out=NULL;
    gpgme_error_t err;
    gpgme_ctx_t ctx=NULL;

    err = gpgme_data_new_from_file (&in, infile, 1);
    if (err)
	return err;

    err = gpgme_new (&ctx);
    if (err)
	goto fail;

    err = gpgme_data_new (&out);
    if (err)
	goto fail;

    gpgme_set_armor (ctx, 1);
    err = gpgme_op_encrypt (ctx, (gpgme_key_t*)rset, GPGME_ENCRYPT_ALWAYS_TRUST, in, out);
    if (!err)
	err = data_to_file (&out, outfile);

fail:
    if (ctx != NULL)
	gpgme_release (ctx);
    if (in != NULL)
	gpgme_data_release (in);
    if (out != NULL)
	gpgme_data_release (out);
    return err;
}


int
op_encrypt_file_io (void *rset, const char *infile, const char *outfile)
{
    FILE *fin = NULL, *fout=NULL;
    gpgme_data_t in = NULL;
    gpgme_data_t out = NULL;
    gpgme_error_t err;
    gpgme_ctx_t ctx = NULL;

    fin = fopen (infile, "rb");
    if (fin == NULL)
	return GPG_ERR_UNKNOWN_ERRNO;
    err = gpgme_data_new_from_stream (&in, fin);
    if (err)
	goto fail;

    fout = fopen (outfile, "wb");
    if (fout == NULL) {
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

    err = gpgme_op_encrypt (ctx, (gpgme_key_t*)rset, GPGME_ENCRYPT_ALWAYS_TRUST, in, out);

fail:
    if (fin != NULL)
	fclose (fin);
    if (fout != NULL)
	fclose (fout);
    if (out != NULL)
	gpgme_data_release (out);
    if (in != NULL)
	gpgme_data_release (in);
    if (ctx != NULL)
	gpgme_release (ctx);
    return err;
}

int
op_encrypt (void *rset, const char *inbuf, char **outbuf)
{
    gpgme_key_t *keys = (gpgme_key_t *)rset;
    gpgme_data_t in=NULL, out=NULL;
    gpgme_error_t err;
    gpgme_ctx_t ctx;
    
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

    gpgme_set_armor (ctx, 1);
    err = gpgme_op_encrypt (ctx, keys, GPGME_ENCRYPT_ALWAYS_TRUST, in, out);
    if (!err) {
	size_t n = 0;
	*outbuf = gpgme_data_release_and_get_mem (out, &n);
	(*outbuf)[n] = 0;
	out = NULL;
    }

leave:
    gpgme_release (ctx);
    gpgme_data_release (in);
    gpgme_data_release (out);
    return 0;
}


int
op_sign_encrypt_start (const char *inbuf, char **outbuf)
{
    gpgme_key_t *keys = NULL;
    gpgme_key_t locusr=NULL;
    int opts = 0;
    int err;

    recipient_dialog_box (&keys, &opts);
    signer_dialog_box (&locusr, NULL);
    
    err = op_sign_encrypt ((void*)keys, (void*)locusr, inbuf, outbuf);
    free (keys);

    return err;
}

int
op_sign_encrypt (void *rset, void *locusr, const char *inbuf, char **outbuf)
{
    gpgme_ctx_t ctx;
    gpgme_key_t *keys = (gpgme_key_t*)rset;
    gpgme_data_t in=NULL, out=NULL;
    gpgme_error_t err;
    struct decrypt_key_s *hd;

    op_init ();
    *outbuf = NULL;

    err = gpgme_new (&ctx);
    if (err)
	return err;

    hd = (struct decrypt_key_s *)xcalloc (1, sizeof *hd);

    err = gpgme_data_new_from_mem (&in, inbuf, strlen (inbuf), 1);
    if (err)
	goto leave;
    err = gpgme_data_new (&out);
    if (err)
	goto leave;

    gpgme_set_passphrase_cb (ctx, passphrase_callback_box, hd);
    gpgme_set_armor (ctx, 1);
    gpgme_signers_add (ctx, (gpgme_key_t)locusr);
    err = gpgme_op_encrypt_sign (ctx, keys, GPGME_ENCRYPT_ALWAYS_TRUST, in, out);
    if (!err) {
	size_t n =0 ;
	*outbuf = gpgme_data_release_and_get_mem (out, &n);
	(*outbuf)[n] = 0;
	out = NULL;
    }

leave:
    free_decrypt_key (hd);
    gpgme_release (ctx);
    gpgme_data_release (out);
    gpgme_data_release (in);
    return err;
}

int
op_sign_start (const char *inbuf, char **outbuf)
{
    gpgme_key_t locusr = NULL;

    signer_dialog_box (&locusr, NULL);
    return op_sign ((void*)locusr, inbuf, outbuf);
}


int
op_sign (void *locusr, const char *inbuf, char **outbuf)
{
    gpgme_error_t err;
    gpgme_data_t in=NULL, out=NULL;
    gpgme_ctx_t ctx;
    struct decrypt_key_s *hd;

    *outbuf = NULL;
    op_init ();
    err = gpgme_new (&ctx);
    if (err)
	return err;

    hd = (struct decrypt_key_s *)xcalloc (1, sizeof *hd);
    hd->flags = 0x01;

    err = gpgme_data_new_from_mem (&in, inbuf, strlen (inbuf), 1);
    if (err)
	goto leave;
    err = gpgme_data_new (&out);
    if (err)
	goto leave;
    
    gpgme_set_passphrase_cb (ctx, passphrase_callback_box, hd);
    gpgme_signers_add (ctx, (gpgme_key_t)locusr);
    gpgme_set_textmode (ctx, 1);
    gpgme_set_armor (ctx, 1);
    err = gpgme_op_sign (ctx, in, out, GPGME_SIG_MODE_CLEAR);
    if (!err) {
	size_t n = 0;
	*outbuf = gpgme_data_release_and_get_mem (out, &n);
	(*outbuf)[n] = 0;
	out = NULL;
    }

leave:
    free_decrypt_key (hd);
    gpgme_release (ctx);
    gpgme_data_release (in);
    gpgme_data_release (out);
    return 0;
}


int
op_decrypt (int (*pass_cb)(void *, const char*, const char*, int, int),
	    void *pass_cb_value,
	    const char *inbuf, char **outbuf)
{
    gpgme_data_t in=NULL, out=NULL;
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

    gpgme_set_passphrase_cb (ctx, pass_cb, pass_cb_value);
    err = gpgme_op_decrypt_verify (ctx, in, out);
    if (!err) {
	size_t n = 0;
	*outbuf = gpgme_data_release_and_get_mem (out, &n);
	(*outbuf)[n] = 0;
	out = NULL;
    }

    /* XXX: automatically spawn verify_start in case the text was signed? */
    {
	gpgme_verify_result_t res;
	res = gpgme_op_verify_result (ctx);
	if (res != NULL && res->signatures != NULL)
	    verify_dialog_box (res);
    }

leave:    
    gpgme_release (ctx);
    gpgme_data_release (in);
    gpgme_data_release (out);
    return err;
}


int
op_decrypt_start_ext (const char *inbuf, char **outbuf, cache_item_t *ret_itm)
{
    struct decrypt_key_s *hd;
    cache_item_t itm;
    int err;

    hd = (struct decrypt_key_s *)xcalloc (1, sizeof *hd);
    err = op_decrypt (passphrase_callback_box, hd, inbuf, outbuf);

    if (ret_itm != NULL) {
	itm = cache_item_new ();
	itm->pass = hd->pass; 
	hd->pass = NULL;
	memcpy (itm->keyid, hd->keyid, sizeof (hd->keyid));
	*ret_itm = itm;
    }
    free_decrypt_key (hd);
    return err;
}

int
op_decrypt_start (const char *inbuf, char **outbuf)
{
    struct decrypt_key_s *hd;
    int err;

    hd = (struct decrypt_key_s *)xcalloc (1, sizeof *hd);
    err = op_decrypt (passphrase_callback_box, hd, inbuf, outbuf);

    free_decrypt_key (hd);
    return err;
}

int
op_decrypt_next (int (*pass_cb)(void *, const char*, const char*, int, int),
		 void *pass_cb_value,
		 const char *inbuf, char **outbuf)
{
    return op_decrypt (pass_cb, pass_cb_value, inbuf, outbuf);
}


int
op_verify_start (const char *inbuf, char **outbuf)
{
    gpgme_data_t in=NULL, out=NULL;
    gpgme_ctx_t ctx;
    gpgme_error_t err;
    gpgme_verify_result_t res=NULL;

    op_init ();
    *outbuf = NULL;

    err = gpgme_new (&ctx);
    if (err)
	return err;

    err = gpgme_data_new_from_mem (&in, inbuf, strlen (inbuf), 1);
    if (err)
	goto leave;
    err = gpgme_data_new (&out);
    if (err)
	goto leave;

    err = gpgme_op_verify (ctx, in, NULL, out);
    if (!err) {
	size_t n=0;
	if (outbuf != NULL) {
	    *outbuf = gpgme_data_release_and_get_mem (out, &n);
	    (*outbuf)[n] = 0;
	    out = NULL;
	}
	res = gpgme_op_verify_result (ctx);
    }
    if (res != NULL) 
	verify_dialog_box (res);

leave:
    gpgme_data_release (out);
    gpgme_data_release (in);
    gpgme_release (ctx);
    return err;
}


/* Try to find a key for each item in @id. If one ore more items were
   not found, it is added to @unknown at the same position.
   @n is the total amount of items to find. */
int 
op_lookup_keys (char **id, gpgme_key_t **keys, char ***unknown, size_t *n)
{
    int i, pos=0;
    gpgme_key_t k;

    for (i=0; id[i] != NULL; i++)
	;
    if (n)
	*n = i;
    *unknown = (char**)xcalloc (i+1, sizeof (char*));
    *keys = (gpgme_key_t*)xcalloc (i+1, sizeof (gpgme_key_t));
    for (i=0; id[i] != NULL; i++) {
	/*k = find_gpg_email(id[i]);*/
	k = get_gpg_key (id[i]);
	log_debug ("c:\\mapigpgme.log", "%s: %p\r\n", id[i], k);
	if (!k)	    
	    (*unknown)[pos++] = xstrdup (id[i]);
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
