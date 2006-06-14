/* passphrase-dialog.c
 *	Copyright (C) 2004 Timo Schulz
 *	Copyright (C) 2005, 2006 g10 Code GmbH
 *
 * This file is part of GPGol.
 * 
 * GPGol is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2 of the License, or (at your option) any later version.
 * 
 * GPGol is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Lesser General Public License for more details.
 * 
 * You should have received a copy of the GNU Lesser General Public
 * License along with this program; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA
 * 02110-1301, USA.
 */

#include <config.h>

#include <windows.h>
#include <time.h>
#include <assert.h>
#include <gpgme.h>

#include "gpgol-ids.h"
#include "passcache.h"
#include "intern.h"

#define TRACEPOINT() do { log_debug ("%s:%s:%d: tracepoint\n", \
                                     SRCNAME, __func__, __LINE__); \
                        } while (0)

/* Object to maintai8n state in the dialogs. */
struct dialog_context_s
{
  struct decrypt_key_s *dec; /* The decryption info. */

  gpgme_key_t *keyarray;     /* NULL or an array of keys. */

  int hide_state;            /* Flag indicating that some stuff is hidden. */

  unsigned int use_as_cb;    /* This is used by the passphrase callback. */

  int no_encrypt_warning;    /* Print a warning after cancel. */
};


static char const allhexdigits[] = "1234567890ABCDEFabcdef";


static void
set_key_hint (struct decrypt_key_s *dec, HWND dlg, int ctrlid)
{
  const char *s = dec->user_id;
  char *key_hint;
  
  if (s && dec->keyid) 
    {
      char stop_char;
      size_t i = 0;

      key_hint = xmalloc (17 + strlen (s) + strlen (dec->keyid) + 32);
      if (strchr (s, '<') && strchr (s, '>'))
        stop_char = '<';
      else if (strchr (s, '(') && strchr (s, ')'))
        stop_char = '(';
      else
        stop_char = 0;
      while (*s != stop_char)
        key_hint[i++] = *s++;
      key_hint[i++] = ' ';
      if (dec->keyid && strlen (dec->keyid) > 8)
        sprintf (key_hint+i, "(0x%s)", dec->keyid+8);
      else
        key_hint[i] = 0;
    }
  else
    key_hint = xstrdup (_("No key hint given."));
  SendDlgItemMessage (dlg, ctrlid, CB_ADDSTRING, 0, 
                      (LPARAM)(const char *)key_hint);
  SendDlgItemMessage (dlg, ctrlid, CB_SETCURSEL, 0, 0);
  xfree (key_hint);
}

/* Release the key array ARRAY as well as all COUNT keys. */
static void
release_keyarray (gpgme_key_t *array)
{
  size_t n;

  if (array)
    {
      for (n=0; array[n]; n++)
        gpgme_key_release (array[n]);
      xfree (array);
    }
}

/* Return the number of keys in the key array KEYS. */
static size_t
count_keys (gpgme_key_t *keys)
{
  size_t n = 0;

  if (keys)
    for (; *keys; keys++)
      n++;
  return n;
}


static void
load_recipbox (HWND dlg, int ctlid, gpgme_ctx_t ctx)
{
  gpgme_decrypt_result_t res;
  gpgme_recipient_t rset, r;
  gpgme_ctx_t keyctx = NULL;
  gpgme_key_t key;
  gpgme_error_t err;
  char *buffer, *p;
  size_t n;
  
  if (!ctx)
    return;

  /* Lump together all recipients of the message. */
  res = gpgme_op_decrypt_result (ctx);
  if (!res || !res->recipients)
    return;
  rset = res->recipients;
  for (n=0, r = rset; r; r = r->next)
    if (r->keyid)
      n += strlen (r->keyid) + 1;    
  buffer = xmalloc (n + 1);
  for (p=buffer, r = rset; r; r = r->next)
    if (r->keyid)
      p = stpcpy (stpcpy (p, r->keyid), " ");

  /* Run a key list on all of them to fill up the list box. */
  err = gpgme_new (&keyctx);
  if (err)
    goto fail;
  err = gpgme_op_keylist_start (keyctx, buffer, 0);
  if (err)
    goto fail;

  while (!gpgme_op_keylist_next (keyctx, &key))
    {
      if (key && key->uids && key->uids->uid)
	{
	  char *utf8_uid = utf8_to_native (key->uids->uid);
	  SendDlgItemMessage (dlg, ctlid, LB_ADDSTRING, 0,
			      (LPARAM)(const char *)utf8_uid);
	  xfree (utf8_uid);
	}
      if (key)
	gpgme_key_release (key);
    }

 fail:
  if (keyctx)
    gpgme_release (keyctx);
  xfree (buffer);
}



/* Return a string with the short description of the algorithm.  This
   function is guaranteed to not return NULL or a string longer that 3
   bytes. */
const char*
get_pubkey_algo_str (gpgme_pubkey_algo_t alg)
{
  switch (alg)
    {
    case GPGME_PK_RSA:
    case GPGME_PK_RSA_E:
    case GPGME_PK_RSA_S:
      return "RSA";
      
    case GPGME_PK_ELG_E:
    case GPGME_PK_ELG:
      return "ELG";

    case GPGME_PK_DSA:
      return "DSA";

    default:
      break;
    }
  
  return "???";
}


/* Fill a combo box with all keys and return an error with those
   keys. */
static gpgme_key_t *
load_secbox (HWND dlg, int ctlid)
{
  gpg_error_t err;
  gpgme_ctx_t ctx;
  gpgme_key_t key;
  int pos;
  gpgme_key_t *keyarray;
  size_t keyarray_size;

  err = gpgme_new (&ctx);
  if (err)
    return NULL;

  err = gpgme_op_keylist_start (ctx, NULL, 1);
  if (err)
    {
      log_error ("failed to initiate key listing: %s\n", gpg_strerror (err));
      gpgme_release (ctx);
      return NULL;
    }

  keyarray_size = 20; 
  keyarray = xcalloc (keyarray_size+1, sizeof *keyarray);
  pos = 0;

  while (!gpgme_op_keylist_next (ctx, &key)) 
    {
      const char *email, *keyid, *algo;
      char *p, *name;
      long idx;
      
      if (key->revoked || key->expired || key->disabled || key->invalid)
        {
          gpgme_key_release (key);
          continue;
        }
      if (!key->uids || !key->subkeys)
        {
          gpgme_key_release (key);
          continue;
        }
        
      if (!key->uids->name)
        name = strdup ("");
      else
	name = utf8_to_native (key->uids->name);
      email = key->uids->email;
      if (!email)
	email = "";
      keyid = key->subkeys->keyid;
      if (!keyid || strlen (keyid) < 8)
        {
	  xfree (name);
          gpgme_key_release (key);
          continue;
        }

      algo = get_pubkey_algo_str (key->subkeys->pubkey_algo);
      p = xmalloc (strlen (name) + strlen (email) + strlen (keyid+8) + 3 + 20);
      if (email && *email)
	sprintf (p, "%s <%s> (0x%s, %s)", name, email, keyid+8, algo);
      else
	sprintf (p, "%s (0x%s, %s)", name, keyid+8, algo);
      idx = SendDlgItemMessage (dlg, ctlid, CB_ADDSTRING, 0, 
				(LPARAM)(const char *)p);
      xfree (p);
      xfree (name);
      if (idx < 0) /* Error. */
        {
          gpgme_key_release (key);
          continue;
        }
      
      SendDlgItemMessage (dlg, ctlid, CB_SETITEMDATA, idx, (LPARAM)pos);

      if (pos >= keyarray_size)
        {
          gpgme_key_t *tmparr;
          size_t i;

          keyarray_size += 10;
          tmparr = xcalloc (keyarray_size, sizeof *tmparr);
          for (i=0; i < pos; i++)
            tmparr[i] = keyarray[i];
          xfree (keyarray);
          keyarray = tmparr;
        }
      keyarray[pos++] = key;
    }
  SendDlgItemMessage (dlg, ctlid, CB_SETCURSEL, 0, 0);

  gpgme_op_keylist_end (ctx);
  gpgme_release (ctx);
  return keyarray;
}


static BOOL CALLBACK
decrypt_key_dlg_proc (HWND dlg, UINT msg, WPARAM wparam, LPARAM lparam)
{
  /* Fixme: We should not use a static here but keep it in an array
     index by DLG. */
  static struct dialog_context_s *context; 
  struct decrypt_key_s *dec;
  size_t n;
  const char *warn;

  if (msg == WM_INITDIALOG)
    {
      context = (struct dialog_context_s *)lparam;
      context->hide_state = 1;
      dec = context->dec;
      if (dec && context->use_as_cb) 
        {
          dec->opts = 0;
          dec->pass = NULL;
          set_key_hint (dec, dlg, IDC_DEC_KEYLIST);
          EnableWindow (GetDlgItem (dlg, IDC_DEC_KEYLIST), FALSE);
	}
      if (dec && dec->last_was_bad)
        SetDlgItemText (dlg, IDC_DEC_HINT, 
                        (dec && dec->last_was_bad)?
                        _("Invalid passphrase; please try again..."):"");

      if (dec && !context->use_as_cb)
	context->keyarray = load_secbox (dlg, IDC_DEC_KEYLIST);

      CheckDlgButton (dlg, IDC_DEC_HIDE, BST_CHECKED);
      center_window (dlg, NULL);
      if (dec && dec->hide_pwd) 
        {
          ShowWindow (GetDlgItem (dlg, IDC_DEC_HIDE), SW_HIDE);
          ShowWindow (GetDlgItem (dlg, IDC_DEC_PASS), SW_HIDE);
          ShowWindow (GetDlgItem (dlg, IDC_DEC_PASSINF), SW_HIDE);
	}
      else
        SetFocus (GetDlgItem (dlg, IDC_DEC_PASS));

      if (!context->use_as_cb)
	SetWindowText (dlg, _("Select Signing Key"));
      SetForegroundWindow (dlg);
      return FALSE;
    }

  if (!context)
    return FALSE;

  dec = context->dec;

  switch (msg) 
    {
    case WM_DESTROY:
      context->hide_state = 1;
      break;

    case WM_SYSCOMMAND:
      if (wparam == SC_CLOSE)
        EndDialog (dlg, TRUE);
      break;

    case WM_COMMAND:
      switch (HIWORD (wparam)) 
        {
	case BN_CLICKED:
          if ((int)LOWORD (wparam) == IDC_DEC_HIDE)
            {
              HWND hwnd;

              context->hide_state ^= 1;
              hwnd = GetDlgItem (dlg, IDC_DEC_PASS);
              SendMessage (hwnd, EM_SETPASSWORDCHAR,
                           context->hide_state? '*' : 0, 0);
              SetFocus (hwnd);
	    }
          break;
	}

      switch (LOWORD (wparam)) 
        {
	case IDOK:
          n = SendDlgItemMessage (dlg, IDC_DEC_PASS, WM_GETTEXTLENGTH, 0, 0);
          if (n && dec) 
            {
              dec->pass = xmalloc (n + 2);
              GetDlgItemText (dlg, IDC_DEC_PASS, dec->pass, n+1);
	    }
          if (dec && !context->use_as_cb) 
            {
              int idx, pos;

              idx = SendDlgItemMessage (dlg, IDC_DEC_KEYLIST,
                                        CB_GETCURSEL, 0, 0);
              pos = SendDlgItemMessage (dlg, IDC_DEC_KEYLIST,
                                        CB_GETITEMDATA, idx, 0);
              if (pos >= 0 && pos < count_keys (context->keyarray))
                {
                  dec->signer = context->keyarray[pos];
                  gpgme_key_ref (dec->signer);
                }
	    }
          EndDialog (dlg, TRUE);
          break;
          
	case IDCANCEL:
          if (context->no_encrypt_warning)
            {
              warn = _("If you cancel this dialog, the message will be sent"
                       " in cleartext!\n\n"
                       "Do you really want to cancel?");
            }
          else if (dec && context->use_as_cb && (dec->flags & 0x01)) 
            {
              warn = _("If you cancel this dialog, the message"
                       " will be sent without signing.\n\n"
                       "Do you really want to cancel?");
            }
          else
            warn = NULL;

          if (warn)
            {
              n = MessageBox (dlg, warn, _("Secret Key Dialog"),
                              MB_ICONWARNING|MB_YESNO);
              if (n == IDNO)
                return FALSE;
	    }
          if (dec)
            {
              dec->opts = OPT_FLAG_CANCEL;
              dec->pass = NULL;
            }
          EndDialog (dlg, FALSE);
          break;
	}
      break; /*WM_COMMAND*/
    }

  return FALSE;
}


static BOOL CALLBACK
decrypt_key_ext_dlg_proc (HWND dlg, UINT msg, WPARAM wparam, LPARAM lparam)
{
  /* Fixme: We should not use a static here but keep it in an array
     index by DLG. */
  static struct dialog_context_s *context; 
  struct decrypt_key_s * dec;
  size_t n;
  const char *warn;

  if (msg == WM_INITDIALOG)
    {
      context = (struct dialog_context_s *)lparam;
      context->hide_state = 1;
      dec = context->dec;
      if (dec) 
        {
          dec->opts = 0;
          dec->pass = NULL;
          set_key_hint (dec, dlg, IDC_DECEXT_KEYLIST);
          EnableWindow (GetDlgItem (dlg, IDC_DECEXT_KEYLIST), FALSE);
	}

      SetDlgItemText (dlg, IDC_DECEXT_HINT, 
                      (dec && dec->last_was_bad)?
                      _("Invalid passphrase; please try again..."):"");
      if (dec)
        load_recipbox (dlg, IDC_DECEXT_RSET, (gpgme_ctx_t)dec->ctx);

      CheckDlgButton (dlg, IDC_DECEXT_HIDE, BST_CHECKED);
      center_window (dlg, NULL);
      SetFocus (GetDlgItem (dlg, IDC_DECEXT_PASS));
      SetForegroundWindow (dlg);
      return FALSE;
    }

  if (!context)
    return FALSE;

  dec = context->dec;

  switch (msg) 
    {
    case WM_DESTROY:
      context->hide_state = 1;
      break;

    case WM_SYSCOMMAND:
      if (wparam == SC_CLOSE)
        EndDialog (dlg, TRUE);
      break;

    case WM_COMMAND:
      switch (HIWORD (wparam)) 
        {
	case BN_CLICKED:
          if ((int)LOWORD (wparam) == IDC_DECEXT_HIDE) 
            {
              HWND hwnd;
              
              context->hide_state ^= 1;
              hwnd = GetDlgItem (dlg, IDC_DECEXT_PASS);
              SendMessage (hwnd, EM_SETPASSWORDCHAR,
                           context->hide_state? '*' : 0, 0);
              SetFocus (hwnd);
	    }
          break;
	}

      switch (LOWORD (wparam)) 
        {
	case IDOK:
          n = SendDlgItemMessage (dlg, IDC_DECEXT_PASS, WM_GETTEXTLENGTH,0,0);
          if (n && dec) 
            {
              dec->pass = xmalloc ( n + 2 );
              GetDlgItemText (dlg, IDC_DECEXT_PASS, dec->pass, n+1 );
	    }
          EndDialog (dlg, TRUE);
          break;

	case IDCANCEL:
          if (context->no_encrypt_warning)
            {
              warn = _("If you cancel this dialog, the message will be sent"
                       " in cleartext!\n\n"
                       "Do you really want to cancel?");
            }
          else if (dec && context->use_as_cb && (dec->flags & 0x01)) 
            {
              warn = _("If you cancel this dialog, the message"
                       " will be sent without signing.\n"
                       "Do you really want to cancel?");
            }
          else
            warn = NULL;

          if (warn)
            {
              n = MessageBox (dlg, warn, _("Secret Key Dialog"),
                              MB_ICONWARNING|MB_YESNO);
              if (n == IDNO)
                return FALSE;
	    }
          if (dec)
            {
              dec->opts = OPT_FLAG_CANCEL;
              dec->pass = NULL;
            }
          EndDialog (dlg, FALSE);
          break;
	}
      break; /*WM_COMMAND*/
    }

  return FALSE;
}

/* Display a signer dialog which contains all secret keys, useable for
   signing data.  The key is returned in R_KEY.  The passprase in
   r_passwd.  If Encrypting is true, the message will get encrypted
   later. */
int 
signer_dialog_box (gpgme_key_t *r_key, char **r_passwd, int encrypting)
{
  struct dialog_context_s context; 
  struct decrypt_key_s dec;
  int resid;

  memset (&context, 0, sizeof context);
  memset (&dec, 0, sizeof dec);
  dec.hide_pwd = 1;
  context.dec = &dec;
  context.no_encrypt_warning = encrypting;

  if (!strncmp (gettext_localename (), "de", 2))
    resid = IDD_DEC_DE;
  else
    resid = IDD_DEC;
  DialogBoxParam (glob_hinst, (LPCTSTR)resid, GetDesktopWindow (),
                  decrypt_key_dlg_proc, (LPARAM)&context);

  if (dec.signer) 
    {
      if (r_passwd)
        *r_passwd = dec.pass;
      else 
        {	    
          if (dec.pass)
            wipestring (dec.pass);
          xfree (dec.pass);
        }
      dec.pass = NULL;
      *r_key = dec.signer;
      dec.signer = NULL;
    }
  if (dec.pass)
    wipestring (dec.pass);
  xfree (dec.pass);
  if (dec.signer)
    gpgme_key_release (dec.signer);
  release_keyarray (context.keyarray);
  return (dec.opts & OPT_FLAG_CANCEL)? -1 : 0;
}


/* GPGME passphrase callback function. It starts the decryption dialog
   to request the passphrase from the user.  See the GPGME manual for
   a description of the arguments. */
gpgme_error_t
passphrase_callback_box (void *opaque, const char *uid_hint, 
			 const char *pass_info,
			 int prev_was_bad, int fd)
{
  struct decrypt_key_s *dec = opaque;
  DWORD nwritten = 0;
  char keyidstr[16+1];
  int resid;

  log_debug ("passphrase_callback_box: enter (uh=`%s',pi=`%s')\n", 
             uid_hint?uid_hint:"(null)", pass_info?pass_info:"(null)");

  *keyidstr = 0; 

  /* First remove a possible passphrase from the return structure. */
  if (dec->pass)
    wipestring (dec->pass);
  xfree (dec->pass);
  dec->pass = NULL;

  /* For some reasons the cancel flag has been set - write an empty
     passphrase and close the handle to indicate the cancel state to
     the backend. */
  if (dec->opts & OPT_FLAG_CANCEL)
    {
      /* Casting the FD to a handle is okay as gpgme uses OS handles. */
      WriteFile ((HANDLE)fd, "\n", 1, &nwritten, NULL);
      CloseHandle ((HANDLE)fd);
      log_debug ("passphrase_callback_box: leave (due to cancel flag)\n");
      return -1;
    }

  /* Parse the information to get the keyid we use for caching of the
     passphrase. If we got suitable information, we will have a proper
     16 character string in KEYIDSTR; if not KEYIDSTR has been set to
     empty. */
  if (pass_info)
    {
      /* As of now (gpg 1.4.2) these information are possible:

         1. Standard passphrase requested:
            "<long main keyid> <long keyid> <keytype> <keylength>"
         2. Passphrase for symmetric key requested:
            "<cipher_algo> <s2k_mode> <s2k_hash>"
         3. PIN for a card requested.
            "<card_type> <chvno>"
      
       For caching we need to use the long keyid from case 1; the main
       keyid can't be used because a key may have different
       passphrases on the subkeys.  Caching for symmetrical keys is
       not possible because there is no information on what
       key(i.e. passphrase) to use.  Caching of PINs is not yet
       possible because we don't have information on the card's serial
       number yet; that must be solved by gpgme. 

       To detect case 1 we simply check whether the first token
       consists in its entire of 16 hex digits.
      */
      const char *s;
      int i;

      for (s=pass_info, i=0; *s && strchr (allhexdigits, *s) ; s++, i++)
        ;
      if (i == 16)
        {
          while (*s == ' ')
            s++;
          for (i=0; *s && strchr (allhexdigits, *s) && i < 16; s++, i++)
            keyidstr[i] = *s;
          keyidstr[i] = 0;
          if (*s)
            s++;
          if (i != 16 || !strchr (allhexdigits, *s))
            {
              log_debug ("%s: oops: does not look like pass_info\n", __func__);
              *keyidstr = 0;
            }
        }
    } 
  
  log_debug ("%s: using keyid 0x%s\n", __func__, keyidstr);

  /* Now check how to proceed. */
  if (prev_was_bad) 
    {
      log_debug ("%s: last passphrase was bad\n", __func__);
      /* Flush a possible cache entry for that keyID. */
      if (*keyidstr)
        passcache_put (keyidstr, NULL, 0);
    }
  else if (*keyidstr)
    {
      dec->pass = passcache_get (keyidstr);
      log_debug ("%s: getting passphrase for 0x%s from cache: %s\n",
                 __func__, keyidstr, dec->pass? "hit":"miss");
    }

  /* Copy the keyID into the context. */
  assert (strlen (keyidstr) < sizeof dec->keyid);
  strcpy (dec->keyid, keyidstr);

  /* If we have no cached passphrase, popup the passphrase dialog. */
  if (!dec->pass)
    {
      int rc;
      const char *s;
      struct dialog_context_s context;

      memset (&context, 0, sizeof context);
      context.dec = dec;
      context.use_as_cb = 1;

      /* Construct the user ID. */
      if (uid_hint)
        {
          /* Skip the first token (keyID). */
          for (s=uid_hint; *s && *s != ' '; s++)
             ;
          while (*s == ' ')
            s++;
        }
      else
        s = "[no user Id]";
      xfree (dec->user_id);
      dec->user_id = utf8_to_native (s);
      dec->last_was_bad = prev_was_bad;
      if (dec->flags & 0x01)
        {
          if (!strncmp (gettext_localename (), "de", 2))
            resid = IDD_DEC_DE;
          else
            resid = IDD_DEC;
          rc = DialogBoxParam (glob_hinst, (LPCSTR)resid,
                               GetDesktopWindow (),
                               decrypt_key_dlg_proc, (LPARAM)&context);
        }
      else
        {
          if (!strncmp (gettext_localename (), "de", 2))
            resid = IDD_DEC_EXT_DE;
          else
            resid = IDD_DEC_EXT;
          rc = DialogBoxParam (glob_hinst, (LPCTSTR)resid,
                               GetDesktopWindow (),
                               decrypt_key_ext_dlg_proc, (LPARAM)&context);
        }
      if (rc <= 0) 
        log_debug_w32 (-1, "%s: dialog failed (rc=%d)", __func__, rc);
      release_keyarray (context.keyarray);
    }
  else 
    log_debug ("%s:%s: dec=%p dec->pass=`[censored]'\n",
               SRCNAME, __func__, dec);

  /* If we got a passphrase, send it to the FD. */
  if (dec->pass) 
    {
      log_debug ("passphrase_callback_box: sending passphrase ...\n");
      WriteFile ((HANDLE)fd, dec->pass, strlen (dec->pass), &nwritten, NULL);
    }

  WriteFile((HANDLE)fd, "\n", 1, &nwritten, NULL);

  log_debug ("passphrase_callback_box: leave\n");
  return 0;
}


/* Release the context which was used in the passphrase callback. */
void
free_decrypt_key (struct decrypt_key_s * ctx)
{
    if (!ctx)
	return;
    if (ctx->pass) {
        wipestring (ctx->pass);
	xfree (ctx->pass);
    }
    xfree (ctx->user_id);
    xfree(ctx);
}
