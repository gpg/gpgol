/* verify-dialog.c
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
#include <gpgme.h>

#include "gpgol-ids.h"
#include "intern.h"
#include "util.h"

struct dialog_context
{
  gpgme_verify_result_t res;
  const char *filename;
};


static char*
get_timestamp (time_t l)
{
    static char buf[64];
    struct tm * tm;

    if (l == 0) {
	sprintf (buf, "????" "-??" "-?? ??" ":??" ":??");
	return buf;
    }
	
    tm = localtime (&l);
    sprintf (buf, "%04d-%02d-%02d %02d:%02d:%02d",
	     tm->tm_year+1900, tm->tm_mon+1, tm->tm_mday,
	     tm->tm_hour, tm->tm_min, tm->tm_sec);
    return buf;
}


static int
load_akalist (HWND dlg, gpgme_key_t key)
{
  gpgme_user_id_t u;
  char *uid;
  int n = 0;
  
  u = key->uids;
  if (!u->next)
    return n;
  for (u=u->next; u; u=u->next) 
    {
      uid = utf8_to_native (u->uid);
      SendDlgItemMessage (dlg, IDC_VRY_AKALIST, LB_ADDSTRING,
			  0, (LPARAM)(const char*)uid);
      xfree (uid);
      n++;
    }
  return n;
}


static void 
load_sigbox (HWND dlg, gpgme_verify_result_t ctx)
{
  gpgme_error_t err;
  gpgme_key_t key;
  char buf[2+16+1];
  char *p;
  const char *s;
  int stat;
  int valid, no_key = 0, n = 0;
  
  s = get_timestamp (ctx->signatures->timestamp);
  SetDlgItemText (dlg, IDC_VRY_TIME, s);
  
  s = ctx->signatures->fpr;
  if (strlen (s) == 40)
    strncpy (buf+2, s+40-8, 8);
  else if (strlen (s) == 32) /* MD5:RSAv3 */
    strncpy (buf+2, s+32-8, 8);
  else
    strncpy (buf+2, s+8, 8);
  buf[10] = 0;
  buf[0] = '0'; 
  buf[1] = 'x';
  SetDlgItemText (dlg, IDC_VRY_KEYID, buf);

  {
    gpgme_ctx_t gctx;

    key = NULL;
    if (!gpgme_new (&gctx))
      {
        err = gpgme_get_key (gctx, buf+2, &key, 0);
        if (err)
          {
            log_debug ("getting key `%s' failed: %s",
                       buf+2, gpg_strerror (err));
            key = NULL;
          }
        gpgme_release (gctx);
      }
  }

  stat = ctx->signatures->summary;
  if (stat & GPGME_SIGSUM_RED)
    s = _("BAD signature!");
  else if (!stat || (stat & GPGME_SIGSUM_GREEN))
    s = _("Good signature");
  else if (stat & GPGME_SIGSUM_KEY_REVOKED)
    s = _("Good signature from revoked key");
  else if (stat & GPGME_SIGSUM_KEY_EXPIRED)
    s = _("Good signature from expired key");
  else if (stat & GPGME_SIGSUM_SIG_EXPIRED)
    s = _("Good expired signature");
  else if (stat & GPGME_SIGSUM_KEY_MISSING) 
    {
      s = _("Could not check signature: missing key");
      no_key = 1;
    }
  else
    s = _("Verification error");
  /* XXX: if we have a key we do _NOT_ trust, stat is 'wrong' */
  SetDlgItemText (dlg, IDC_VRY_STATUS, s);
  
  if (key && key->uids) 
    {
      p = utf8_to_native (key->uids->uid);
      SetDlgItemText (dlg, IDC_VRY_ISSUER, p);
      xfree (p);
      
      n = load_akalist (dlg, key);
      gpgme_key_release (key);
      if (n == 0)
	EnableWindow (GetDlgItem (dlg, IDC_VRY_AKALIST), FALSE);
    }
  else 
    {
      s = _("User-ID not found");
      SetDlgItemText (dlg, IDC_VRY_ISSUER, s);
    }
  
  s = get_pubkey_algo_str (ctx->signatures->pubkey_algo);
  SetDlgItemText (dlg, IDC_VRY_PKALGO, s);
  
  valid = ctx->signatures->validity;
  if (stat & GPGME_SIGSUM_RED)
    {
      /* This is a BAD signature; give a hint to the user. */
      SetDlgItemText (dlg, IDC_VRY_HINT, 
               _("This may be due to a wrong option setting"));
    }
  else if (stat & GPGME_SIGSUM_SIG_EXPIRED) 
    {
      const char *fmt;
    
      fmt = _("Signature expired on %s");
      s = get_timestamp (ctx->signatures->exp_timestamp);
      p = xmalloc (strlen (s)+1+strlen (fmt)+2);
      sprintf (p, fmt, s);
      SetDlgItemText (dlg, IDC_VRY_HINT, s);
      xfree (p);
    }
  else if (valid < GPGME_VALIDITY_MARGINAL) 
    {
      switch (valid) 
	{
	case GPGME_VALIDITY_NEVER:
	  s = _("Signature issued by a key we do NOT trust.");
	  break;
	  
	default:
	  if (no_key)
	    s = "";
	  else
	    s = _("Signature issued by a non-valid key.");
	  break;
	}
      SetDlgItemText (dlg, IDC_VRY_HINT, s);
    }
}


static BOOL CALLBACK
verify_dlg_proc (HWND dlg, UINT msg, WPARAM wparam, LPARAM lparam)
{
    static struct dialog_context *ctx;

    switch (msg) {
    case WM_INITDIALOG:
      
	ctx = (struct dialog_context *)lparam;
	load_sigbox (dlg, ctx->res);
	center_window (dlg, NULL);
	SetForegroundWindow (dlg);
        if (ctx->filename)
          {
            const char *s = _("Verification Result");
            char *tmp = xmalloc (strlen (ctx->filename) 
                                 + strlen (s) + 100);
            strcpy (stpcpy (stpcpy (stpcpy (tmp, s),
                                    " ("), ctx->filename), ")");
            SetWindowText (dlg, tmp);
            xfree (tmp);
          }
	break;

    case WM_COMMAND:
	switch (LOWORD(wparam)) {
	case IDOK:
	    EndDialog (dlg, TRUE);
	    break;
	}
	break;
    }
    return FALSE;
}


/* Display the verify dialog based on the gpgme result in
   RES. FILENAME is used to modify the caption of the dialog; it may
   be NULL. */
int
verify_dialog_box (gpgme_verify_result_t res, const char *filename)
{
  struct dialog_context ctx;
  int resid;

  memset (&ctx,0, sizeof ctx);
  ctx.res = res;
  ctx.filename = filename;

  if (!strncmp (gettext_localename (), "de", 2))
    resid = IDD_VRY_DE;
  else
    resid = IDD_VRY;
  DialogBoxParam (glob_hinst, (LPCTSTR)resid, GetDesktopWindow (),
                  verify_dlg_proc, (LPARAM)&ctx);
  return res->signatures->summary == GPGME_SIGSUM_GREEN? 0 : -1;
}
