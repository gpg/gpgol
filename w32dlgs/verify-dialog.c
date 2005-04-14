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
#include "intern.h"
#include "keycache.h"

static char*
get_timestamp (time_t l)
{
    static char buf[64];
    struct tm * tm;

    tm = localtime (&l);
    sprintf (buf, "%04d-%02d-%02d %02d:%02d:%02d",
	     tm->tm_year+1900, tm->tm_mon+1, tm->tm_mday,
	     tm->tm_hour, tm->tm_min, tm->tm_sec);
    return buf;
}


static void
load_akalist (HWND dlg, gpgme_key_t key)
{
    gpgme_user_id_t u;

    u=key->uids;
    if (!u->next)
	return;
    for (u=u->next; u; u=u->next) {
	SendDlgItemMessage (dlg, IDC_VRY_AKALIST, LB_ADDSTRING,
			    0, (LPARAM)(const char*)u->uid);
    }
}

static void 
load_sigbox (HWND dlg, gpgme_verify_result_t ctx)
{
    gpgme_key_t key;
    char *s, buf[2+16+1];
    int stat;

    s = get_timestamp (ctx->signatures->timestamp);
    SetDlgItemText (dlg, IDC_VRY_TIME, s);

    s = ctx->signatures->fpr;
    if (strlen (s) == 40)
	strncpy (buf+2, s+40-8, 8);
    else
	strncpy (buf+2, s+8, 8);
    buf[10] = 0;
    buf[0] = '0'; buf[1] = 'x';
    SetDlgItemText (dlg, IDC_VRY_KEYID, buf);
    key = find_gpg_key (buf+2, 0);

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
    else if (stat & GPGME_SIGSUM_KEY_MISSING)
	s = "Could not check signature: missing key";
    SetDlgItemText (dlg, IDC_VRY_STATUS, s);
    
    if (key) {
	s = (char*)gpgme_key_get_string_attr (key, GPGME_ATTR_USERID, NULL, 0);
	SetDlgItemText (dlg, IDC_VRY_ISSUER, s);

	s = (char*)gpgme_key_get_string_attr (key, GPGME_ATTR_ALGO, NULL, 0);
	SetDlgItemText (dlg, IDC_VRY_PKALGO, s);

	load_akalist (dlg, key);
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


int
verify_dialog_box (gpgme_verify_result_t res)
{
    DialogBoxParam (glob_hinst, (LPCTSTR)IDD_VRY, GetDesktopWindow (),
		    verify_dlg_proc, (LPARAM)res);
    return res->signatures->summary == GPGME_SIGSUM_GREEN? 0 : -1;
}
