/* passphrase-dialog.c
 *	Copyright (C) 2004 Timo Schulz
 *	Copyright (C) 2005 g10 Code GmbH
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
#include "keycache.h"
#include "passcache.h"
#include "intern.h"
#include "usermap.h"


static char const allhexdigits[] = "1234567890ABCDEFabcdef";


#if 0 /* XXX: left over from old code?? */
static void
add_string_list (HWND hbox, const char **list, int start_idx)
{
    const char * s;
    int i;

    for (i=0; (s=list[i]); i++)
	SendMessage (hbox, CB_ADDSTRING, 0, (LPARAM)(const char *)s);
    SendMessage (hbox, CB_SETCURSEL, (WPARAM) start_idx, 0);
}
#endif


static void
set_key_hint (struct decrypt_key_s * dec, HWND dlg, int ctrlid)
{
    const char *s = dec->user_id;
    char *key_hint;
    char stop_char=0;
    size_t i=0;

    if (dec->user_id != NULL) {
	key_hint = (char *)xmalloc (17 + strlen (dec->user_id) + 32);
	if (strchr (s, '<') && strchr (s, '>'))
	    stop_char = '<';
	else if (strchr (s, '(') && strchr (s, ')'))
	    stop_char = '(';
	while (s && *s != stop_char)
	    key_hint[i++] = *s++;
	key_hint[i++] = ' ';
	sprintf (key_hint+i, "(0x%s)", dec->keyid+8);
    }
    else
	key_hint = xstrdup ("No key hint given.");
    SendDlgItemMessage (dlg, ctrlid, CB_ADDSTRING, 0, 
			(LPARAM)(const char *)key_hint);
    SendDlgItemMessage (dlg, ctrlid, CB_SETCURSEL, 0, 0);
    xfree (key_hint);
}


static void
load_recipbox (HWND dlg, int ctlid, gpgme_ctx_t ctx)
{
  gpgme_decrypt_result_t res;
  gpgme_recipient_t rset, r;
  gpgme_ctx_t keyctx=NULL;
  gpgme_key_t key;
  gpgme_error_t err;
  char *p;
  int c=0;
  
  if (ctx == NULL)
    return;
  res = gpgme_op_decrypt_result (ctx);
  if (res == NULL || res->recipients == NULL)
    return;
  rset = res->recipients;
  for (r = rset; r; r = r->next)
    c++;    
  p = xcalloc (1, c*(17+2));
  for (r = rset; r; r = r->next) {
    strcat (p, r->keyid);
    strcat (p, " ");
  }

  err = gpgme_new (&keyctx);
  if (err)
    goto fail;
  err = gpgme_op_keylist_start (keyctx, p, 0);
  if (err)
    goto fail;

  r = rset;
  for (;;) 
    {
      const char *uid;
      err = gpgme_op_keylist_next (keyctx, &key);
      if (err)
	break;
      uid = gpgme_key_get_string_attr (key, GPGME_ATTR_USERID, NULL, 0);
      if (uid != NULL) 
	SendDlgItemMessage (dlg, ctlid, LB_ADDSTRING, 0,
			    (LPARAM)(const char *)uid);
	gpgme_key_release (key);
	key = NULL;
	r = r->next;
    }

fail:
    if (keyctx != NULL)
	gpgme_release (keyctx);
    xfree (p);
}


static void
load_secbox (HWND dlg, int ctlid)
{
    gpgme_key_t sk;
    size_t n=0, doloop=1;
    void *ctx=NULL;

    enum_gpg_seckeys (NULL, &ctx);
    while (doloop) {
	const char *name, *email, *keyid, *algo;
	char *p;

	if (enum_gpg_seckeys (&sk, &ctx))
	    doloop = 0;

	if (gpgme_key_get_ulong_attr (sk, GPGME_ATTR_KEY_REVOKED, NULL, 0) ||
	    gpgme_key_get_ulong_attr (sk, GPGME_ATTR_KEY_EXPIRED, NULL, 0) ||
	    gpgme_key_get_ulong_attr (sk, GPGME_ATTR_KEY_INVALID, NULL, 0))
	    continue;
	
	name = gpgme_key_get_string_attr (sk, GPGME_ATTR_NAME, NULL, 0);
	email = gpgme_key_get_string_attr (sk, GPGME_ATTR_EMAIL, NULL, 0);
	keyid = gpgme_key_get_string_attr (sk, GPGME_ATTR_KEYID, NULL, 0);
	algo = gpgme_key_get_string_attr (sk, GPGME_ATTR_ALGO, NULL, 0);
	if (!email)
	    email = "";
	p = (char *)xcalloc (1, strlen (name) + strlen (email) + 17 + 32);
	if (email && strlen (email))
	    sprintf (p, "%s <%s> (0x%s, %s)", name, email, keyid+8, algo);
	else
	    sprintf (p, "%s (0x%s, %s)", name, keyid+8, algo);
	SendDlgItemMessage (dlg, ctlid, CB_ADDSTRING, 0, 
			    (LPARAM)(const char *) p);
	xfree (p);
    }
    
    ctx = NULL;
    reset_gpg_seckeys (&ctx);
    doloop = 1;
    n = 0;
    while (doloop) {
	if (enum_gpg_seckeys (&sk, &ctx))
	    doloop = 0;
	if (gpgme_key_get_ulong_attr (sk, GPGME_ATTR_KEY_REVOKED, NULL, 0) ||
	    gpgme_key_get_ulong_attr (sk, GPGME_ATTR_KEY_EXPIRED, NULL, 0) ||
	    gpgme_key_get_ulong_attr (sk, GPGME_ATTR_KEY_INVALID, NULL, 0))
	    continue;
	SendDlgItemMessage (dlg, ctlid, CB_SETITEMDATA, n, (LPARAM)(DWORD)sk);
	n++;
    }
    SendDlgItemMessage (dlg, ctlid, CB_SETCURSEL, 0, 0);
    reset_gpg_seckeys (&ctx);
}


static BOOL CALLBACK
decrypt_key_dlg_proc (HWND dlg, UINT msg, WPARAM wparam, LPARAM lparam)
{
    static struct decrypt_key_s * dec;
    static int hide_state = 1;
    size_t n;

    switch (msg) {
    case WM_INITDIALOG:
        log_debug ("decrypt_key_dlg_proc: WM_INITDIALOG\n");
	dec = (struct decrypt_key_s *)lparam;
	if (dec && dec->use_as_cb) {
	    dec->opts = 0;
	    dec->pass = NULL;
	    set_key_hint (dec, dlg, IDC_DEC_KEYLIST);
	    EnableWindow (GetDlgItem (dlg, IDC_DEC_KEYLIST), FALSE);
	}
	if (dec && dec->last_was_bad)
	    SetDlgItemText (dlg, IDC_DEC_HINT, "Invalid passphrase; please try again...");
	else
	    SetDlgItemText (dlg, IDC_DEC_HINT, "");
	if (dec && !dec->use_as_cb)
	    load_secbox (dlg, IDC_DEC_KEYLIST);
	CheckDlgButton (dlg, IDC_DEC_HIDE, BST_CHECKED);
	center_window (dlg, NULL);
	if (dec->hide_pwd) {
	    ShowWindow (GetDlgItem (dlg, IDC_DEC_HIDE), SW_HIDE);
	    ShowWindow (GetDlgItem (dlg, IDC_DEC_PASS), SW_HIDE);
	    ShowWindow (GetDlgItem (dlg, IDC_DEC_PASSINF), SW_HIDE);
	    /* XXX: make the dialog window smaller */
	}
	else
	    SetFocus (GetDlgItem (dlg, IDC_DEC_PASS));
	SetForegroundWindow (dlg);
	return FALSE;

    case WM_DESTROY:
        log_debug ("decrypt_key_dlg_proc: WM_DESTROY\n");
	hide_state = 1;
	break;

    case WM_SYSCOMMAND:
        log_debug ("decrypt_key_dlg_proc: WM_SYSCOMMAND\n");
	if (wparam == SC_CLOSE)
	    EndDialog (dlg, TRUE);
	break;

    case WM_COMMAND:
        log_debug ("decrypt_key_dlg_proc: WM_COMMAND\n");
	switch (HIWORD (wparam)) {
	case BN_CLICKED:
	    if ((int)LOWORD (wparam) == IDC_DEC_HIDE) {
		HWND hwnd;

		hide_state ^= 1;
		hwnd = GetDlgItem (dlg, IDC_DEC_PASS);
		SendMessage (hwnd, EM_SETPASSWORDCHAR, hide_state? '*' : 0, 0);
		SetFocus (hwnd);
	    }
	    break;
	}
	switch (LOWORD (wparam)) {
	case IDOK:
	    n = SendDlgItemMessage (dlg, IDC_DEC_PASS, WM_GETTEXTLENGTH, 0, 0);
	    if (n) {
		dec->pass = (char *)xcalloc (1, n+2);
		GetDlgItemText (dlg, IDC_DEC_PASS, dec->pass, n+1);
	    }
	    if (!dec->use_as_cb) {
		int idx = SendDlgItemMessage (dlg, IDC_DEC_KEYLIST, 
					      CB_GETCURSEL, 0, 0);
		dec->signer = (gpgme_key_t)SendDlgItemMessage (dlg, IDC_DEC_KEYLIST,
							       CB_GETITEMDATA, idx, 0);
		gpgme_key_ref (dec->signer);
	    }
	    EndDialog (dlg, TRUE);
	    break;

	case IDCANCEL:
	    if (dec && dec->use_as_cb && (dec->flags & 0x01)) {
		const char *warn = "If you cancel this dialog, the message will be sent without signing.\n\n"
				   "Do you really want to cancel?";
		n = MessageBox (dlg, warn, "Secret Key Dialog", MB_ICONWARNING|MB_YESNO);
		if (n == IDNO)
		    return FALSE;
	    }
	    dec->opts = OPT_FLAG_CANCEL;
	    dec->pass = NULL;
	    EndDialog (dlg, FALSE);
	    break;
	}
	break;
    }
    return FALSE;
}


static BOOL CALLBACK
decrypt_key_ext_dlg_proc (HWND dlg, UINT msg, WPARAM wparam, LPARAM lparam)
{
    static struct decrypt_key_s * dec;
    static int hide_state = 1;
    size_t n;

    switch (msg) {
    case WM_INITDIALOG:
	dec = (struct decrypt_key_s *)lparam;
	if (dec != NULL) {
	    dec->opts = 0;
	    dec->pass = NULL;
	    set_key_hint (dec, dlg, IDC_DECEXT_KEYLIST);
	    EnableWindow (GetDlgItem (dlg, IDC_DECEXT_KEYLIST), FALSE);
	}
	if (dec && dec->last_was_bad)
	    SetDlgItemText (dlg, IDC_DECEXT_HINT, "Invalid passphrase; please try again...");
	else
	    SetDlgItemText (dlg, IDC_DECEXT_HINT, "");
	if (dec != NULL)
	    load_recipbox (dlg, IDC_DECEXT_RSET, (gpgme_ctx_t)dec->ctx);
	CheckDlgButton (dlg, IDC_DECEXT_HIDE, BST_CHECKED);
	center_window (dlg, NULL);
	SetFocus (GetDlgItem (dlg, IDC_DECEXT_PASS));
	SetForegroundWindow (dlg);
	return FALSE;

    case WM_DESTROY:
	hide_state = 1;
	break;

    case WM_SYSCOMMAND:
	if (wparam == SC_CLOSE)
	    EndDialog (dlg, TRUE);
	break;

    case WM_COMMAND:
	switch (HIWORD (wparam)) {
	case BN_CLICKED:
	    if ((int)LOWORD (wparam) == IDC_DECEXT_HIDE) {
		HWND hwnd;

		hide_state ^= 1;
		hwnd = GetDlgItem (dlg, IDC_DECEXT_PASS);
		SendMessage (hwnd, EM_SETPASSWORDCHAR, hide_state? '*' : 0, 0);
		SetFocus (hwnd);
	    }
	    break;
	}
	switch (LOWORD (wparam)) {
	case IDOK:
	    n = SendDlgItemMessage (dlg, IDC_DECEXT_PASS, WM_GETTEXTLENGTH, 0, 0);
	    if (n) {
		dec->pass = (char *)xcalloc( 1, n+2 );
		GetDlgItemText( dlg, IDC_DECEXT_PASS, dec->pass, n+1 );
	    }
	    EndDialog (dlg, TRUE);
	    break;

	case IDCANCEL:
	    if (dec && dec->use_as_cb && (dec->flags & 0x01)) {
		const char *warn = "If you cancel this dialog, the message will be sent without signing.\n"
				   "Do you really want to cancel?";
		n = MessageBox (dlg, warn, "Secret Key Dialog", MB_ICONWARNING|MB_YESNO);
		if (n == IDNO)
		    return FALSE;
	    }
	    dec->opts = OPT_FLAG_CANCEL;
	    dec->pass = NULL;
	    EndDialog (dlg, FALSE);
	    break;
	}
	break;
    }
    return FALSE;
}

/* Display a signer dialog which contains all secret keys, useable
   for signing data. The key is returned in r_key. The password in
   r_passwd. */
int 
signer_dialog_box (gpgme_key_t *r_key, char **r_passwd)
{
    struct decrypt_key_s hd;
    int rc = 0;

    memset(&hd, 0, sizeof (hd));
    hd.hide_pwd = 1;
    DialogBoxParam (glob_hinst, (LPCTSTR)IDD_DEC, GetDesktopWindow (),
		    decrypt_key_dlg_proc, (LPARAM)&hd);
    if (hd.signer) {
	if (r_passwd)
	    *r_passwd = xstrdup (hd.pass);
	else {	    
	    xfree (hd.pass);
	    hd.pass = NULL;
	}
	*r_key = hd.signer;
    }
    if (hd.opts & OPT_FLAG_CANCEL)
	rc = -1;
    memset (&hd, 0, sizeof (hd));    
    return rc;
}


/* GPGME passphrase callback function. It starts the decryption dialog
   to request the passphrase from the user.  See the GPGME manual for
   a description of the arguments. */
gpgme_error_t
passphrase_callback_box (void *opaque, const char *uid_hint, 
			 const char *pass_info,
			 int prev_was_bad, int fd)
{
  struct decrypt_key_s *hd = opaque;
  DWORD nwritten = 0;
  char keyidstr[16+1];

  log_debug ("passphrase_callback_box: enter (uh=`%s',pi=`%s')\n", 
             uid_hint?uid_hint:"(null)", pass_info?pass_info:"(null)");

  *keyidstr = 0; 

  /* First remove a possible passphrase from the return structure. */
  if (hd->pass)
    wipestring (hd->pass);
  xfree (hd->pass);
  hd->pass = NULL;

  /* For some reasons the cancel flag has been set - write an empty
     passphrase and close the handle to indicate the cancel state to
     the backend. */
  if (hd->opts & OPT_FLAG_CANCEL)
    {
      /* FIXME: Casting an FD to a handles is very questionable.  We
         need to check this. */
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
       not possible becuase there is no information on what
       key(i.e. passphrase) to use.  Caching of of PINs is not yet
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
      hd->pass = passcache_get (keyidstr);
      log_debug ("%s: getting passphrase for 0x%s from cache: %s\n",
                 __func__, keyidstr, hd->pass? "hit":"miss");
    }

  /* Copy the keyID into the context. */
  assert (strlen (keyidstr) < sizeof hd->keyid);
  strcpy (hd->keyid, keyidstr);

  /* If we have no cached pssphrase, popup the passphrase dialog. */
  if (!hd->pass)
    {
      int rc;
      const char *s;

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
      xfree (hd->user_id);
      hd->user_id = xstrdup (s);
          
      hd->last_was_bad = prev_was_bad;
      hd->use_as_cb = 1; /* FIXME: For what is this used? */
      if (hd->flags & 0x01)
        rc = DialogBoxParam (glob_hinst, (LPCSTR)IDD_DEC,
                             GetDesktopWindow (),
                             decrypt_key_dlg_proc, (LPARAM)hd);
      else
        rc = DialogBoxParam (glob_hinst, (LPCTSTR)IDD_DEC_EXT,
                             GetDesktopWindow (),
                             decrypt_key_ext_dlg_proc, (LPARAM)hd);
      if (rc <= 0) 
        log_debug_w32 (-1, "%s: dialog failed (rc=%d)", __func__, rc);
    }
  else 
    log_debug ("%s:%s: hd=%p hd->pass=`[censored]'\n",
               __FILE__, __func__, hd);

  /* If we got a passphrase, send it to the FD. */
  if (hd->pass) 
    {
      log_debug ("passphrase_callback_box: sending passphrase ...\n");
      WriteFile ((HANDLE)fd, hd->pass, strlen (hd->pass), &nwritten, NULL);
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
