/* passphrase-dialog.c
 *	Copyright (C) 2004 Timo Schulz
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


static void
add_string_list( HWND hbox, const char ** list, int start_idx )
{
    const char * s;
    int i;

    for( i=0; (s=list[i]); i++ )
	SendMessage( hbox, CB_ADDSTRING, 0, (LPARAM)(const char *)s );
    SendMessage( hbox, CB_SETCURSEL, (WPARAM) start_idx, 0 );
}


static void
set_key_hint( struct decrypt_key_s * dec, HWND dlg, int ctrlid )
{
    const char * s = dec->user_id;
    char * key_hint;
    char stop_char=0;
    size_t i=0;

    if( dec->user_id ) {
	key_hint = (char *)malloc( 17 + strlen( dec->user_id ) + 32 );
	if( strchr( s, '<' ) && strchr( s, '>' ) )
	    stop_char = '<';
	else if( strchr( s, '(' ) && strchr( s, ')' ) )
	    stop_char = '(';
	while( s && *s != stop_char )
	    key_hint[i++] = *s++;
	key_hint[i++] = ' ';
	sprintf( key_hint+i, "(0x%s)", dec->keyid+8 );	
    }
    else
	key_hint = strdup( "Symmetrical Decryption" );
    SendDlgItemMessage( dlg, ctrlid, CB_ADDSTRING, 0, (LPARAM)(const char *)key_hint );	    
    SendDlgItemMessage( dlg, ctrlid, CB_SETCURSEL, 0, 0 );
    free( key_hint );
}


static void
load_secbox( HWND dlg )
{
    gpgme_key_t sk;
    size_t n=0;
    void *ctx=NULL;

    enum_gpg_seckeys( NULL, &ctx);
    while( !enum_gpg_seckeys( &sk, &ctx ) ) {
	const char * name, * email, * keyid, * algo;
	char * p;

	if( gpgme_key_get_ulong_attr( sk, GPGME_ATTR_KEY_REVOKED, NULL, 0 )
	  ||gpgme_key_get_ulong_attr( sk, GPGME_ATTR_KEY_EXPIRED, NULL, 0 )
	  ||gpgme_key_get_ulong_attr( sk, GPGME_ATTR_KEY_INVALID, NULL, 0 ) )
	    continue;
	
	name = gpgme_key_get_string_attr( sk, GPGME_ATTR_NAME, NULL, 0 );
	email = gpgme_key_get_string_attr( sk, GPGME_ATTR_EMAIL, NULL, 0 );
	keyid = gpgme_key_get_string_attr( sk, GPGME_ATTR_KEYID, NULL, 0 );
	algo = gpgme_key_get_string_attr( sk, GPGME_ATTR_ALGO, NULL, 0 );
	if( !email )
	    email = "";
	p = (char *)calloc( 1, strlen( name ) + strlen( email ) + 17 + 32 );
	if( email && strlen( email ) )
	    sprintf( p, "%s <%s> (0x%s, %s)", name, email, keyid+8, algo );
	else
	    sprintf( p, "%s (0x%s, %s)", name, keyid+8, algo );
	SendDlgItemMessage( dlg, IDC_DEC_KEYLIST, CB_ADDSTRING, 0, (LPARAM)(const char *) p );
	free(p);
    }
    
    ctx=NULL;
    reset_gpg_seckeys(&ctx);    
    while( !enum_gpg_seckeys( &sk, &ctx ) ) {
	SendDlgItemMessage( dlg, IDC_DEC_KEYLIST, CB_SETITEMDATA, n, (LPARAM)(DWORD)sk );
	n++;
    }
    SendDlgItemMessage( dlg, IDC_DEC_KEYLIST, CB_SETCURSEL, 0, 0 );
}


static BOOL CALLBACK
decrypt_key_dlg_proc( HWND dlg, UINT msg, WPARAM wparam, LPARAM lparam )
{
    static struct decrypt_key_s * dec;
    static int hide_state = 1;
    size_t n;

    switch( msg ) {
    case WM_INITDIALOG:
	dec = (struct decrypt_key_s *)lparam;
	if (dec && dec->use_as_cb) {
	    dec->opts = 0;
	    dec->pass = NULL;
	    set_key_hint (dec, dlg, IDC_DEC_KEYLIST);
	    EnableWindow( GetDlgItem( dlg, IDC_DEC_KEYLIST ), FALSE );
	}
	if (dec && !dec->use_as_cb)
	    load_secbox(dlg);
	CheckDlgButton( dlg, IDC_DEC_HIDE, BST_CHECKED );
	center_window (dlg, NULL);
	SetFocus (GetDlgItem (dlg, IDC_DEC_PASS));
	SetForegroundWindow (dlg);
	return FALSE;

    case WM_COMMAND:
	switch( HIWORD( wparam ) ) {
	case BN_CLICKED:
	    if( (int)LOWORD( wparam ) == IDC_DEC_HIDE ) {
		HWND hwnd;

		hide_state ^= 1;
		hwnd = GetDlgItem( dlg, IDC_DEC_PASS );
		SendMessage( hwnd, EM_SETPASSWORDCHAR, hide_state? '*' : 0, 0 );
		SetFocus( hwnd );
	    }
	    break;
	}
	switch( LOWORD(wparam) ) {
	case IDOK:
	    n = SendDlgItemMessage( dlg, IDC_DEC_PASS, WM_GETTEXTLENGTH, 0, 0 );
	    if (n) {
		dec->pass = (char *)calloc( 1, n+2 );
		GetDlgItemText( dlg, IDC_DEC_PASS, dec->pass, n+1 );
	    }
	    if (!dec->use_as_cb) {
		int idx = SendDlgItemMessage(dlg, IDC_DEC_KEYLIST, CB_GETCURSEL, 0, 0);
		dec->signer = (gpgme_key_t)SendDlgItemMessage(dlg, IDC_DEC_KEYLIST, 
							      CB_GETITEMDATA, idx, 0);
	    }
	    EndDialog( dlg, TRUE );
	    break;

	case IDCANCEL:
	    dec->opts = OPT_FLAG_CANCEL;
	    dec->pass = NULL;
	    EndDialog( dlg, FALSE );
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
signer_dialog_box(gpgme_key_t *r_key, char **r_passwd)
{
    struct decrypt_key_s hd;
    memset(&hd, 0, sizeof hd);
    DialogBoxParam(glob_hinst, (LPCTSTR)IDD_DEC, GetDesktopWindow(),
		    decrypt_key_dlg_proc, (LPARAM)&hd );
    if (hd.signer) {
	if (r_passwd)
	    *r_passwd = strdup(hd.pass);
	*r_key = hd.signer;
    }
    memset (&hd, 0, sizeof hd);
    return 0;
}


/* GPGME passphrase callback function. It starts the decryption dialog
   to request the passphrase from the user. */
const char * 
passphrase_callback_box(void *opaque, const char *uid_hint, const char *pass_info,
			int prev_was_bad, int fd)
{
    struct decrypt_key_s * hd = (struct decrypt_key_s *)opaque;

    if (prev_was_bad) {
	free (hd->pass);
	hd->pass = NULL;
    }

    if (hd && uid_hint && !hd->pass) {
	const char * s = uid_hint;
	size_t i=0;


	while( s && *s != '\n' )
	    s++;
	s++;
	while( s && *s != ' ' )
	    hd->keyid[i++] = *s++;
	hd->keyid[i] = '\0'; s++;
	if( hd->user_id )
	    free( hd->user_id );
	hd->user_id = (char *)calloc( 1, strlen( s ) + 2 );
	strcpy( hd->user_id, s );
	
	if (hd->opts & OPT_FLAG_CANCEL)
	    return "";
	hd->use_as_cb = 1;
	DialogBoxParam( glob_hinst, (LPCTSTR)IDD_DEC, GetDesktopWindow(),
			decrypt_key_dlg_proc, (LPARAM)hd );
    }
    return hd? hd->pass : "";
}


/* Release the context which was used in the passphrase callback. */
void
free_decrypt_key(struct decrypt_key_s * ctx)
{
    if (!ctx)
	return;
    if (ctx->pass) {
	free(ctx->pass);
	ctx->pass = NULL;
    }
    if (ctx->user_id) {
	free(ctx->user_id);
	ctx->user_id = NULL;
    }
    free(ctx);
}
