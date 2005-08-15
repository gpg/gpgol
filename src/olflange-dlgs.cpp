/* olflange-dlgs.cpp - New dialogs for Outlook.
 *	Copyright (C) 2001 G Data Software AG, http://www.gdata.de
 *	Copyright (C) 2004, 2005 g10 Code GmbH
 * 
 * This file is part of OutlGPG.
 * 
 * OutlGPG is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2 of the License, or (at your option) any later version.
 * 
 * OutlGPG is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Lesser General Public License for more details.
 * 
 * You should have received a copy of the GNU Lesser General Public
 * License along with this program; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA
 * 02110-1301, USA.
 */

#ifdef HAVE_CONFIG_H
#include <config.h>
#endif

#include <windows.h>
#include <prsht.h>

#include "mymapi.h"
#include "mymapitags.h"
#include "MapiGPGME.h"

#include "olflange-def.h"
#include "olflange-ids.h"


/* GPGOptionsDlgProc -
   Handles the notifications sent for managing the options property page. */
bool GPGOptionsDlgProc (HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    BOOL bMsgResult;    
    static LPNMHDR pnmhdr;
    static HWND hWndPage;
    static int enable = 1;
    
    switch (uMsg) {
    case WM_INITDIALOG:
	const char *s;
	s = NULL;
	s = m_gpg->getDefaultKey ();
	enable = s && *s? 1 : 0;
	EnableWindow (GetDlgItem (hDlg, IDC_ENCRYPT_TO), enable==0? FALSE: TRUE);
	if (enable == 1)
	    CheckDlgButton (hDlg, IDC_ENCRYPT_WITH_STANDARD_KEY, BST_CHECKED);
	SetDlgItemText (hDlg, IDC_VERSION_INFO, 
		        "Version "VERSION " ("__DATE__")");
	return TRUE;
		
    case WM_COMMAND:
	if (HIWORD (wParam) == BN_CLICKED &&
	    LOWORD (wParam) == IDC_ENCRYPT_WITH_STANDARD_KEY) {
	    enable ^= 1;
	    EnableWindow (GetDlgItem (hDlg, IDC_ENCRYPT_TO), enable==0? FALSE: TRUE);
	}
	if (LOWORD(wParam) == IDC_GPG_OPTIONS)
	    m_gpg->startConfigDialog (hDlg);
	break;
	
    case WM_NOTIFY:
	pnmhdr = ((LPNMHDR) lParam);
	switch (pnmhdr->code) {
	case PSN_KILLACTIVE:
	    bMsgResult = FALSE;  /* allow this page to receive PSN_APPLY */
	    break;

	case PSN_SETACTIVE: {
	    TCHAR s[20];
	    const char * f;
	    
	    if (m_gpg->getDefaultKey () != NULL)		
		SetDlgItemText (hDlg, IDC_ENCRYPT_TO, m_gpg->getDefaultKey ());
	    wsprintf(s, "%d", m_gpg->getStorePasswdTime ());
	    SendDlgItemMessage(hDlg, IDC_TIME_PHRASES, WM_SETTEXT, 0, (LPARAM) s);
	    f = m_gpg->getLogFile ();
	    SendDlgItemMessage (hDlg, IDC_DEBUG_LOGFILE, WM_SETTEXT, 0, (LPARAM)f);
	    hWndPage = pnmhdr->hwndFrom;   // to be used in WM_COMMAND
	    SendDlgItemMessage (hDlg, IDC_ENCRYPT_DEFAULT, BM_SETCHECK, 
				m_gpg->getEncryptDefault () ? 1 : 0, 0L);
	    SendDlgItemMessage (hDlg, IDC_SIGN_DEFAULT, BM_SETCHECK, 
			        m_gpg->getSignDefault () ? 1 : 0, 0L);
	    SendDlgItemMessage (hDlg, IDC_ENCRYPT_WITH_STANDARD_KEY, BM_SETCHECK, 
			        m_gpg->getEncryptWithDefaultKey () && enable? 1 : 0, 0L);
	    SendDlgItemMessage (hDlg, IDC_SAVE_DECRYPTED, BM_SETCHECK, 
				m_gpg->getSaveDecryptedAttachments () ? 1 : 0, 0L);
	    SendDlgItemMessage (hDlg, IDC_SIGN_ATTACHMENTS, BM_SETCHECK,
				m_gpg->getSignAttachments() ? 1 : 0, 0L);
	    bMsgResult = FALSE;  /* accepts activation */
	    break; }
		
	case PSN_APPLY:	{
	    TCHAR s[201];

	    GetDlgItemText (hDlg, IDC_ENCRYPT_TO, s, 200);
	    if (strlen (s) > 0 && strchr (s, ' ')) {
		MessageBox (hDlg, "The default key cannot contain any spaces.",
			    "Outlook GnuPG-Plugin", MB_ICONERROR|MB_OK);
		bMsgResult = PSNRET_INVALID_NOCHANGEPAGE ;
		break;
	    }
	    if (strlen (s) == 0)
		m_gpg->setEncryptWithDefaultKey (false);
	    else
		m_gpg->setEncryptWithDefaultKey (SendDlgItemMessage(hDlg, IDC_ENCRYPT_WITH_STANDARD_KEY, BM_GETCHECK, 0, 0L));
	    SendDlgItemMessage (hDlg, IDC_TIME_PHRASES, WM_GETTEXT, 20, (LPARAM) s);		
	    m_gpg->setStorePasswdTime (atol (s));
	    SendDlgItemMessage (hDlg, IDC_DEBUG_LOGFILE, WM_GETTEXT, 200, (LPARAM)s);
	    m_gpg->setLogFile (s);
	    SendDlgItemMessage (hDlg, IDC_ENCRYPT_TO, WM_GETTEXT, 200, (LPARAM)s);
	    m_gpg->setDefaultKey (s);
		
	    m_gpg->setEncryptDefault (SendDlgItemMessage(hDlg, IDC_ENCRYPT_DEFAULT, BM_GETCHECK, 0, 0L));
	    m_gpg->setSignDefault (SendDlgItemMessage(hDlg, IDC_SIGN_DEFAULT, BM_GETCHECK, 0, 0L));
	    m_gpg->setSaveDecryptedAttachments (SendDlgItemMessage(hDlg, IDC_SAVE_DECRYPTED, BM_GETCHECK, 0, 0L));
	    m_gpg->setSignAttachments (SendDlgItemMessage (hDlg, IDC_SIGN_ATTACHMENTS,  BM_GETCHECK, 0, 0L));
	    m_gpg->writeOptions ();
	    bMsgResult = PSNRET_NOERROR;
	    break; }
		
	case PSN_HELP:
	    MessageBox (pnmhdr->hwndFrom,
    "This is OutlGPG version " PACKAGE_VERSION "\n"
    "Copyright (C) 2005 g10 Code GmbH\n"
    "\n"
    "OutlGPG is a plugin for Outlook to allow encryption and\n"
    "signing of messages using the OpenPGP standard. It makes\n"
    "use of the GnuPG software (http://www.gnupg.org).\n"
    "\n"
    "OutlGPG is free software; you can redistribute it and/or\n"
    "modify it under the terms of the GNU Lesser General Public\n"
    "License as published by the Free Software Foundation; either\n"
    "version 2.1 of the License, or (at your option) any later version.\n"
    "\n"
    "OutlGPG is distributed in the hope that it will be useful,\n"
    "but WITHOUT ANY WARRANTY; without even the implied warranty of\n"
    "MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU\n"
    "Lesser General Public License for more details.\n"
    "\n"
    "You should have received a copy of the GNU Lesser General Public\n"
    "License along with this library; if not, write to the Free Software\n"
    "Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA\n"
    "02110-1301, USA.\n",
                  "GnuPG Plugin", MB_OK);
	    bMsgResult = TRUE;
	    break;

	default:
	    bMsgResult = FALSE;
	    break;	    
	}
	SetWindowLong (hDlg, DWL_MSGRESULT, bMsgResult);
	break;

    default:
	bMsgResult = FALSE;
	break;		
    }
    return bMsgResult;
}
