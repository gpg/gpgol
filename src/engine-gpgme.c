/* engine-gpgme.c - Crypto engine with GPGME
 *	Copyright (C) 2005, 2006, 2007 g10 Code GmbH
 *
 * This file is part of GpgOL.
 *
 * GpgOL is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public License
 * as published by the Free Software Foundation; either version 2.1 
 * of the License, or (at your option) any later version.
 *  
 * GpgOL is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public License
 * along with this program; if not, see <http://www.gnu.org/licenses/>.
 */


#include <config.h>

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <time.h>
#include <errno.h>
#include <assert.h>
#define WIN32_LEAN_AND_MEAN 
#include <windows.h>

#include "common.h"
#include "passcache.h"
#include "engine.h"
#include "engine-gpgme.h"

#define TRACEPOINT() do { log_debug ("%s:%s:%d: tracepoint\n", \
                                       SRCNAME, __func__, __LINE__); \
                        } while (0)


/* Because we are using asynchronous gpgme commands, we need to have a
   closure to cleanup allocated resources and run the code required
   adfter gpgme finished the command (e.g. getting the signature
   verification result.  Thus all functions need to implement a
   closure function and pass it using a closure_data_t object via the
   gpgme_progress_cb hack.  */
struct closure_data_s;
typedef struct closure_data_s *closure_data_t;
struct closure_data_s
{
  void (*closure)(closure_data_t, gpgme_ctx_t, gpg_error_t);
  engine_filter_t filter;
  struct passphrase_cb_s pw_cb;  /* Passphrase callback info.  */
  int with_verify;
  gpgme_data_t sigobj;
};


static int basic_init_done = 0;
static int init_done = 0;


static DWORD WINAPI waiter_thread (void *dummy);
static CRITICAL_SECTION waiter_thread_lock;
static void update_passphrase_cache (int err, 
                                     struct passphrase_cb_s *pass_cb_value);
/* static void add_verify_attestation (gpgme_data_t at,  */
/*                                     gpgme_ctx_t ctx,  */
/*                                     gpgme_verify_result_t res, */
/*                                     const char *filename); */



static void
cleanup (void)
{
  /* Fixme: We should stop the thread.  */
}


/* Cleanup static resources. */
void
op_gpgme_deinit (void)
{
  cleanup ();
}


/* First part of the gpgme initialization.  This is sufficient if we
   only use the gpgme_data_t stuff.  */
int
op_gpgme_basic_init (void)
{
  if (basic_init_done)
    return 0;

  if (!gpgme_check_version (NEED_GPGME_VERSION)) 
    {
      log_debug ("gpgme is too old (need %s, have %s)\n",
                 NEED_GPGME_VERSION, gpgme_check_version (NULL) );
      return gpg_error (GPG_ERR_GENERAL);
    }

  basic_init_done = 1;
  return 0;
}


int
op_gpgme_init (void)
{
  gpgme_error_t err;

  if (init_done)
    return 0;

  err = op_gpgme_basic_init ();
  if (err)
    return err;

  err = gpgme_engine_check_version (GPGME_PROTOCOL_OpenPGP);
  if (err)
    {
      log_debug ("gpgme can't find a suitable OpenPGP backend: %s\n",
                 gpgme_strerror (err));
      return err;
    }

  err = gpgme_engine_check_version (GPGME_PROTOCOL_CMS);
  if (err)
    {
      log_debug ("gpgme can't find a suitable CMS backend: %s\n",
                 gpgme_strerror (err));
      return err;
    }

  {
    HANDLE th;
    DWORD tid;

    InitializeCriticalSection (&waiter_thread_lock);
    th = CreateThread (NULL, 128*1024, waiter_thread, NULL, 0, &tid);
    if (th == INVALID_HANDLE_VALUE)
      log_error ("failed to start the gpgme waiter thread\n");
    else
      CloseHandle (th);
  }

  init_done = 1;
  return 0;
}


/* The worker for the asynchronous commands.  */
static DWORD WINAPI
waiter_thread (void *dummy)
{
  gpgme_ctx_t ctx;
  gpg_error_t err;
  void *a_voidptr;
  closure_data_t closure_data;

  (void)dummy;

  for (;;)
    {
      /*  Note: We don't use hang because this will end up in a tight
          loop and does not do a voluntary context switch.  Thus we do
          this by ourself.  Actually it would be better to start
          gpgme_wait only if we really have something to do but that
          is a bit more complicated.  */
      EnterCriticalSection (&waiter_thread_lock);
      ctx = gpgme_wait (NULL, &err, 0);
      LeaveCriticalSection (&waiter_thread_lock);
      if (ctx)
        {
          gpgme_get_progress_cb (ctx, NULL, &a_voidptr);
          closure_data = a_voidptr;
          assert (closure_data);
          assert (closure_data->closure);
          closure_data->closure (closure_data, ctx, err);
          xfree (closure_data);
          gpgme_release (ctx);
        }
      else if (err)
        log_debug ("%s:%s: gpgme_wait failed: %s\n", SRCNAME, __func__, 
                   gpg_strerror (err));
      else
        Sleep (50);
    }
}


void
engine_gpgme_cancel (void *cancel_data)
{
  gpg_error_t err;
  gpgme_ctx_t ctx = cancel_data;

  if (ctx)
    {
      EnterCriticalSection (&waiter_thread_lock);
      err = gpgme_cancel (ctx);
      LeaveCriticalSection (&waiter_thread_lock);
      if (err)
        log_debug ("%s:%s: gpgme_cancel failed: %s\n", 
                   SRCNAME, __func__,  gpg_strerror (err));
    }
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
update_passphrase_cache (int err, struct passphrase_cb_s *pass_cb_value)
{
  if (!pass_cb_value)
    return;
  if (pass_cb_value->keyid && *pass_cb_value->keyid)
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
  if (res->invalid_recipients)
    return gpg_error (GPG_ERR_UNUSABLE_PUBKEY);
  /* XXX: we need to do more here! */
  return err;
}


/* Release an array of GPGME keys. */
static void 
release_key_array (gpgme_key_t *keys)
{
  int i;

  if (keys)
    {
      for (i = 0; keys[i]; i++) 
	gpgme_key_release (keys[i]);
      xfree (keys);
    }
}

/* Return the number of strings in the array STRINGS. */
static size_t
count_strings (char **strings)
{
  size_t i;

  for (i=0; strings[i]; i++)
    ;
  return i;
}

/* Return the number of keys in the gpgme_key_t array KEYS.  */
static size_t
count_keys (gpgme_key_t *keys)
{
  size_t i;
  
  for (i=0; keys[i]; i++)
    ;
  return i;
}


/* Return an array of gpgme key objects derived from thye list of
   strings in RECPIENTS. */
static gpg_error_t
prepare_recipient_keys (gpgme_key_t **r_keys, char **recipients, HWND hwnd)
{
  gpg_error_t err;
  gpgme_key_t *keys = NULL;
  char **unknown = NULL;
  size_t n_keys, n_unknown, n_recp;
  int i;

  *r_keys = NULL;
  if (op_lookup_keys (recipients, &keys, &unknown))
    {
      log_debug ("%s:%s: leave (lookup keys failed)\n", SRCNAME, __func__);
      return gpg_error (GPG_ERR_GENERAL);  
    }

  n_recp = count_strings (recipients);
  n_keys = count_keys (keys);
  n_unknown = count_strings (unknown);

  log_debug ("%s:%s: found %d recipients, need %d, unknown=%d\n",
             SRCNAME, __func__, (int)n_keys, (int)n_recp, (int)n_unknown);
  
  if (n_keys != n_recp)
    {
      unsigned int opts;
      gpgme_key_t *keys2;

      log_debug ("%s:%s: calling recipient_dialog_box2", SRCNAME, __func__);
      opts = recipient_dialog_box2 (keys, unknown, &keys2);
      release_key_array (keys);
      keys = keys2;
      if ( (opts & OPT_FLAG_CANCEL) ) 
        {
          err = gpg_error (GPG_ERR_CANCELED);
          goto leave;
	}
    }


  /* If a default key has been set, add it to the list of keys.  Check
     that the key is actually available. */
  if (opt.enable_default_key && opt.default_key && *opt.default_key)
    {
      gpgme_key_t defkey;
      
      defkey = op_get_one_key (opt.default_key);
      if (!defkey)
        {
          MessageBox (hwnd,
                 _("The configured default encryption certificate is not "
                   "available or does not unambigiously specify one. "
                   "Please fix this in the option dialog.\n\n"
                   "This message won't be be encrypted to this certificate!"),
                   _("Encryption"), MB_ICONWARNING|MB_OK);
        }
      else
        {
          gpgme_key_t *tmpkeys;

          n_keys = count_keys (keys) + 1;
          tmpkeys = xcalloc (n_keys+1, sizeof *tmpkeys);
          for (i = 0; keys[i]; i++) 
            {
              tmpkeys[i] = keys[i];
              gpgme_key_ref (tmpkeys[i]);
            }
          tmpkeys[i++] = defkey;
          tmpkeys[i] = NULL;
          release_key_array (keys);
          keys = tmpkeys;
        }
    }
  
  if (keys)
    {
      for (i=0; keys[i]; i++)
        log_debug ("%s:%s: recp.%d 0x%s %s\n", SRCNAME, __func__,
                   i, keyid_from_key (keys[i]), userid_from_key (keys[i]));
    }
  *r_keys = keys;
  keys = NULL;
  err = 0;

 leave:
  release_key_array (keys);
  return err;
}


/* Not that this closure is called in the context of the
   waiter_thread.  */
static void
encrypt_closure (closure_data_t cld, gpgme_ctx_t ctx, gpg_error_t err)
{
  if (cld->pw_cb.ctx)
    {
      /* Signing was also request; thus update the passphrase cache.  */ 
      update_passphrase_cache (err, &cld->pw_cb);
    }
  if (err)
    err = check_encrypt_result (ctx, err);
  engine_private_finished (cld->filter, err);
}


/* Encrypt the data from INDATA to the OUTDATA object for all
   recpients given in the NULL terminated array RECIPIENTS.  This
   function terminates with success and then expects the caller to
   wait for the result of the encryption using engine_wait.  FILTER is
   used for asynchronous commnication with the engine module.  HWND is
   the window handle of the current window and used to maintain the
   correct relationship between a popups and the active window.  */
int
op_gpgme_encrypt (protocol_t protocol, 
                  gpgme_data_t indata, gpgme_data_t outdata,
                  engine_filter_t filter, void *hwnd,
                  char **recipients)
{
  gpg_error_t err;
  closure_data_t cld;
  gpgme_ctx_t ctx = NULL;
  gpgme_key_t *keys = NULL;

  cld = xcalloc (1, sizeof *cld);
  cld->closure = encrypt_closure;
  cld->filter = filter;

  err = prepare_recipient_keys (&keys, recipients, NULL);
  if (err)
    goto leave;

  err = gpgme_new (&ctx);
  if (err)
    goto leave;
  gpgme_set_progress_cb (ctx, NULL, cld);
  switch (protocol)
    {
    case PROTOCOL_OPENPGP:  /* Gpgme's default.  */
      break;
    case PROTOCOL_SMIME:
      err = gpgme_set_protocol (ctx, GPGME_PROTOCOL_CMS);
      break;
    default:
      err = gpg_error (GPG_ERR_UNSUPPORTED_PROTOCOL);
      break;
    }
  if (err)
    goto leave;


  gpgme_set_armor (ctx, 1);
  /* FIXME:  We should not hardcode always trust. */
/*   if (sign_key) */
/*     { */
/*       gpgme_set_passphrase_cb (ctx, passphrase_callback_box, &cld->pw_cb); */
/*       cld->pw_cb.ctx = ctx; */
/*       cld->pw_cb.ttl = opt.passwd_ttl; */
/*       err = gpgme_signers_add (ctx, sign_key); */
/*       if (!err) */
/*         err = gpgme_op_encrypt_sign_start (ctx, keys, */
/*                                            GPGME_ENCRYPT_ALWAYS_TRUST, */
/*                                            indata, outdata); */
/*     } */
/*   else */
    err = gpgme_op_encrypt_start (ctx, keys, GPGME_ENCRYPT_ALWAYS_TRUST, 
                                  indata, outdata);

 leave:
  if (err)
    {
      xfree (cld);
      gpgme_release (ctx);
    }
  else
    engine_private_set_cancel (filter, ctx);
  release_key_array (keys);
  return err;
}




/* Not that this closure is called in the context of the
   waiter_thread.  */
static void
sign_closure (closure_data_t cld, gpgme_ctx_t ctx, gpg_error_t err)
{
  update_passphrase_cache (err, &cld->pw_cb);
  engine_private_finished (cld->filter, err);
}


/* Created a detached signature for INDATA and write it to OUTDATA.
   On termination of the signing command engine_private_finished() is
   called with FILTER as the first argument.  */
int
op_gpgme_sign (protocol_t protocol, 
               gpgme_data_t indata, gpgme_data_t outdata,
               engine_filter_t filter, void *hwnd)
{
  gpg_error_t err;
  closure_data_t cld;
  gpgme_ctx_t ctx = NULL;
  gpgme_key_t sign_key = NULL;

  if (signer_dialog_box (&sign_key, NULL, 0) == -1)
    {
      log_debug ("%s:%s: leave (dialog failed)\n", SRCNAME, __func__);
      return gpg_error (GPG_ERR_CANCELED);  
    }

  cld = xcalloc (1, sizeof *cld);
  cld->closure = sign_closure;
  cld->filter = filter;

  err = gpgme_new (&ctx);
  if (err)
    goto leave;
  gpgme_set_progress_cb (ctx, NULL, cld);
  switch (protocol)
    {
    case PROTOCOL_OPENPGP:  /* Gpgme's default.  */
      break;
    case PROTOCOL_SMIME:
      err = gpgme_set_protocol (ctx, GPGME_PROTOCOL_CMS);
      break;
    default:
      err = gpg_error (GPG_ERR_UNSUPPORTED_PROTOCOL);
      break;
    }
  if (err)
    goto leave;

  gpgme_set_armor (ctx, 1);
  gpgme_set_passphrase_cb (ctx, passphrase_callback_box, &cld->pw_cb);
  cld->pw_cb.ctx = ctx;
  cld->pw_cb.ttl = opt.passwd_ttl;
  err = gpgme_signers_add (ctx, sign_key);
  if (!err)
    err = gpgme_op_sign_start (ctx, indata, outdata, GPGME_SIG_MODE_DETACH);

 leave:
  if (err)
    {
      xfree (cld);
      gpgme_release (ctx);
    }
  else
    engine_private_set_cancel (filter, ctx);
  gpgme_key_unref (sign_key);
  return err;
}



/* Not that this closure is called in the context of the
   waiter_thread.  */
static void
decrypt_closure (closure_data_t cld, gpgme_ctx_t ctx, gpg_error_t err)
{
  update_passphrase_cache (err, &cld->pw_cb);

  if (!err && !cld->with_verify) 
    ;
  else if (!err) 
    {
      gpgme_verify_result_t res;

      /* Decryption succeeded.  Now check the state of the signatures. */
      res = gpgme_op_verify_result (ctx);
      if (res && res->signatures)
        verify_dialog_box (gpgme_get_protocol (ctx), res, NULL);
    }
  else if (gpg_err_code (err) == GPG_ERR_DECRYPT_FAILED)
    {
      /* The decryption failed.  See whether we can figure out a more
         suitable error code.  */
      gpgme_decrypt_result_t res;

      res = gpgme_op_decrypt_result (ctx);
      if (res && res->recipients 
          && gpgme_err_code (res->recipients->status) == GPG_ERR_NO_SECKEY)
        err = gpg_error (GPG_ERR_NO_SECKEY);
      /* Fixme: return the keyids */
    }
  else
    {
      /* Decryption failed for other reasons. */
    }

  /* If the passphrase callback indicated a cancel operation, change
     the the error code accordingly. */
  if (err && (cld->pw_cb.opts & OPT_FLAG_CANCEL))
    err = gpg_error (GPG_ERR_CANCELED);

  engine_private_finished (cld->filter, err);
}


/* Decrypt data from INDATA to OUTDATE.  If WITH_VERIFY is set, a
   signature of PGP/MIME combined message is also verified the same
   way as with op_gpgme_verify.  */
int
op_gpgme_decrypt (protocol_t protocol,
                  gpgme_data_t indata, gpgme_data_t outdata, 
                  engine_filter_t filter, void *hwnd,
                  int with_verify)
{
  gpgme_error_t err;
  closure_data_t cld;
  gpgme_ctx_t ctx = NULL;
  
  cld = xcalloc (1, sizeof *cld);
  cld->closure = decrypt_closure;
  cld->filter = filter;
  cld->with_verify = with_verify;

  err = gpgme_new (&ctx);
  if (err)
    goto leave;
  gpgme_set_progress_cb (ctx, NULL, cld);
  switch (protocol)
    {
    case PROTOCOL_OPENPGP:  /* Gpgme's default.  */
      break;
    case PROTOCOL_SMIME:
      err = gpgme_set_protocol (ctx, GPGME_PROTOCOL_CMS);
      break;
    default:
      err = gpg_error (GPG_ERR_UNSUPPORTED_PROTOCOL);
      break;
    }
  if (err)
    goto leave;

  /* Note: We do no error checking for the next call because some
     backends may not implement a command hanler at all.  */
  gpgme_set_passphrase_cb (ctx, passphrase_callback_box, &cld->pw_cb);
  cld->pw_cb.ctx = ctx;

  if (with_verify) 
    err = gpgme_op_decrypt_verify_start (ctx, indata, outdata);
  else
    err = gpgme_op_decrypt_start (ctx, indata, outdata);


 leave:
  if (err)
    {
      xfree (cld);
      gpgme_release (ctx);
    }
  else
    engine_private_set_cancel (filter, ctx);
  return err;
}




/* Not that this closure is called in the context of the
   waiter_thread.  */
static void
verify_closure (closure_data_t cld, gpgme_ctx_t ctx, gpg_error_t err)
{
  if (!err)
    {
      gpgme_verify_result_t res;

      res = gpgme_op_verify_result (ctx);
      if (res) 
        verify_dialog_box (gpgme_get_protocol (ctx), res, NULL);
    }
  gpgme_data_release (cld->sigobj);
  engine_private_finished (cld->filter, err);
}


/* Verify a detached message where the data is in the gpgme object
   DATA and the signature given as the string SIGNATUEE. */
int
op_gpgme_verify (gpgme_protocol_t protocol, 
                 gpgme_data_t data, const char *signature,
                 engine_filter_t filter, void *hwnd)
{
  gpgme_error_t err;
  closure_data_t cld;
  gpgme_ctx_t ctx = NULL;
  gpgme_data_t sigobj = NULL;

  cld = xcalloc (1, sizeof *cld);
  cld->closure = verify_closure;
  cld->filter = filter;

  err = gpgme_new (&ctx);
  if (err)
    goto leave;

  gpgme_set_progress_cb (ctx, NULL, cld);
  switch (protocol)
    {
    case PROTOCOL_OPENPGP:  /* Gpgme's default.  */
      break;
    case PROTOCOL_SMIME:
      err = gpgme_set_protocol (ctx, GPGME_PROTOCOL_CMS);
      break;
    default:
      err = gpg_error (GPG_ERR_UNSUPPORTED_PROTOCOL);
      break;
    }
  if (err)
    goto leave;

  err = gpgme_data_new_from_mem (&sigobj, signature, strlen (signature), 0);
  if (err)
    goto leave;
  cld->sigobj = sigobj;

  err = gpgme_op_verify_start (ctx, sigobj, data, NULL);

 leave:
  if (err)
    {
      gpgme_data_release (sigobj);
      xfree (cld);
      gpgme_release (ctx);
    }
  else
    engine_private_set_cancel (filter, ctx);
  return err;
}






#if 0
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
   true if a severe warning has been encountered. */
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
      at_puts (a, _("Warning: One of the certificates has been revoked\n"));
      severe = 1;
    }
  
  if ((sum & GPGME_SIGSUM_KEY_EXPIRED))
    {
      time_t t = key->subkeys->expires ? key->subkeys->expires : 0;

      if (t)
        {
          at_puts (a, _("Warning: The certificate used to create the "
                        "signature expired at: "));
          at_print_time (a, t);
          at_puts (a, "\n");
        }
      else
        at_puts (a, _("Warning: At least one certification certificate "
                      "has expired\n"));
    }

  if ((sum & GPGME_SIGSUM_SIG_EXPIRED))
    {
      at_puts (a, _("Warning: The signature expired at: "));
      at_print_time (a, sig ? sig->exp_timestamp : 0);
      at_puts (a, "\n");
    }

  if ((sum & GPGME_SIGSUM_KEY_MISSING))
    at_puts (a, _("Can't verify due to a missing certificate\n"));

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
              "this certificate belongs to the person named "
              "as shown above\n");
      break;
    case GPGME_VALIDITY_UNDEFINED:
      break;
    case GPGME_VALIDITY_NEVER:
      txt = _("WARNING: The certificate does NOT BELONG to "
              "the person named as shown above\n");
      break;
    case GPGME_VALIDITY_MARGINAL:
      txt = _("WARNING: It is NOT certain that the certificate "
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
      /* Nothing yet written to the stream.  Insert the current time. */
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
#endif



/* Try to find a key for each item in array NAMES. Items not found are
   stored as malloced strings in the newly allocated array UNKNOWN.
   Found keys are stored in the newly allocated array KEYS.  Both
   arrays are terminated by a NULL entry.  Caller needs to release
   KEYS and UNKNOWN.

   Returns: 0 on success. However success may also be that one or all
   keys are unknown.
*/
int 
op_lookup_keys (char **names, gpgme_key_t **keys, char ***unknown)
{
  gpgme_error_t err;
  gpgme_ctx_t ctx;
  size_t n;
  int i, kpos, upos;
  gpgme_key_t k, k2;

  *keys = NULL;
  *unknown = NULL;

  err = gpgme_new (&ctx);
  if (err)
    return -1; /* Error. */

  for (n=0; names[n]; n++)
    ;

  *keys =  xcalloc (n+1, sizeof *keys);
  *unknown = xcalloc (n+1, sizeof *unknown);

  for (i=kpos=upos=0; names[i]; i++)
    {
      k = NULL;
      err = gpgme_op_keylist_start (ctx, names[i], 0);
      if (!err)
        {
          err = gpgme_op_keylist_next (ctx, &k);
          if (!err && !gpgme_op_keylist_next (ctx, &k2))
            {
              /* More than one matching key available.  Take this one
                 as unknown. */
              gpgme_key_release (k);
              gpgme_key_release (k2);
              k = k2 = NULL;
            }
        }
      gpgme_op_keylist_end (ctx);

      
      /* only useable keys will be added otherwise they will be stored
         in unknown (marked with their status). */
      if (k && !k->revoked && !k->disabled && !k->expired)
        (*keys)[kpos++] = k;
      else if (k)
	{
	  char *p, *fmt = "%s (%s)";
	  char *warn = k->revoked? "revoked" : k->expired? "expired" : "disabled";
	  
	  p = xcalloc (1, strlen (names[i]) + strlen (warn) + strlen (fmt) +1);
	  sprintf (p, fmt, names[i], warn);
	  (*unknown)[upos++] = p;
	  gpgme_key_release (k);
	}
      else if (!k)
        (*unknown)[upos++] = xstrdup (names[i]);
    }

  gpgme_release (ctx);
  return 0;
}


/* Return a GPGME key object matching PATTERN.  If no key matches or
   the match is ambiguous, return NULL. */
gpgme_key_t 
op_get_one_key (char *pattern)
{
  gpgme_error_t err;
  gpgme_ctx_t ctx;
  gpgme_key_t k, k2;

  err = gpgme_new (&ctx);
  if (err)
    return NULL; /* Error. */
  err = gpgme_op_keylist_start (ctx, pattern, 0);
  if (!err)
    {
      err = gpgme_op_keylist_next (ctx, &k);
      if (!err && !gpgme_op_keylist_next (ctx, &k2))
        {
          /* More than one matching key available.  Return an error
             instead. */
          gpgme_key_release (k);
          gpgme_key_release (k2);
          k = k2 = NULL;
        }
    }
  gpgme_op_keylist_end (ctx);
  gpgme_release (ctx);
  return k;
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
  gpgme_free (buf);
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

