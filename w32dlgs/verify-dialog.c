/* verify-dialog.c
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
#include <time.h>

#include "resource.h"
#include "gpgme.h"
#include "keycache.h"
#include "intern.h"

static char*
get_timestamp (time_t l)
{
    static char buf[64];
    struct tm * tm;

    if (l == 0) {
	sprintf (buf, "????-??-?? ??:??:??");
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
    int n = 0;

    u = key->uids;
    if (!u->next)
	return n;
    for (u=u->next; u; u=u->next) {
	SendDlgItemMessage (dlg, IDC_VRY_AKALIST, LB_ADDSTRING,
			    0, (LPARAM)(const char*)u->uid);
	n++;
    }
    return n;
}


static void 
load_sigbox (HWND dlg, gpgme_verify_result_t ctx)
{
    gpgme_key_t key;
    char *s, buf[2+16+1];
    char *p;
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
    /*key = find_gpg_key (buf+2, 0);*/
    key = get_gpg_key (buf+2);
    
    stat = ctx->signatures->summary;
    if (stat & GPGME_SIGSUM_GREEN)
	s = "Good signature";
    else if (stat & GPGME_SIGSUM_RED)
	s = "BAD signature!";
    else if (stat & GPGME_SIGSUM_KEY_REVOKED)
	s = "Good signature from revoked key";
    else if (stat & GPGME_SIGSUM_KEY_EXPIRED)
	s = "Good signature from expired key";
    else if (stat & GPGME_SIGSUM_SIG_EXPIRED)
	s = "Good expired signature";
    else if (stat & GPGME_SIGSUM_KEY_MISSING) {
	s = "Could not check signature: missing key";
	no_key = 1;
    }
    else
	s = "Verification error";
    /* XXX: if we have a key we do _NOT_ trust, stat is 'wrong' */
    SetDlgItemText (dlg, IDC_VRY_STATUS, s);
    
    if (key) {
	s = (char*)gpgme_key_get_string_attr (key, GPGME_ATTR_USERID, NULL, 0);
	SetDlgItemText (dlg, IDC_VRY_ISSUER, s);

	n = load_akalist (dlg, key);
	gpgme_key_release (key);
	if (n == 0)
	    EnableWindow (GetDlgItem (dlg, IDC_VRY_AKALIST), FALSE);
    }
    else {
	s = "User-ID not found";
	SetDlgItemText (dlg, IDC_VRY_ISSUER, s);
    }

    s = "???";
    switch (ctx->signatures->pubkey_algo) {
    case 1: s = "RSA"; break;
    case 17:s = "DSA"; break;
    }
    SetDlgItemText (dlg, IDC_VRY_PKALGO, s);

    valid = ctx->signatures->validity;
    if (stat & GPGME_SIGSUM_SIG_EXPIRED) {
	char *fmt;

	fmt = "Signature expired on %s";
	s = get_timestamp (ctx->signatures->exp_timestamp);
	p = xmalloc (strlen (s)+1+strlen (fmt)+2);
	sprintf (p, fmt, s);
	SetDlgItemText (dlg, IDC_VRY_HINT, s);
	xfree (p);
    }
    else if (valid < GPGME_VALIDITY_MARGINAL) {
	switch (valid) {
	case GPGME_VALIDITY_NEVER:
	    s = "Signature issued by a key we do NOT trust.";
	    break;

	default:
	    if (no_key)
		s = "";
	    else
		s = "Signature issued by a non-valid key.";
	    break;
	}
	SetDlgItemText (dlg, IDC_VRY_HINT, s);
    }
}


static BOOL CALLBACK
verify_dlg_proc (HWND dlg, UINT msg, WPARAM wparam, LPARAM lparam)
{
    static gpgme_verify_result_t ctx;

    switch (msg) {
    case WM_INITDIALOG:
	ctx = (gpgme_verify_result_t)lparam;
	load_sigbox (dlg, ctx);
	center_window (dlg, NULL);
	SetForegroundWindow (dlg);
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


/* Display the verify dialog based on the gpgme result in @res. */
int
verify_dialog_box (gpgme_verify_result_t res)
{
    DialogBoxParam (glob_hinst, (LPCTSTR)IDD_VRY, GetDesktopWindow (),
		    verify_dlg_proc, (LPARAM)res);
    return res->signatures->summary == GPGME_SIGSUM_GREEN? 0 : -1;
}
