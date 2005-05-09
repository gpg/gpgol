/* GPGOptionsDlg.cpp - gpg options dialog procedure
 *	Copyright (C) 2001 G Data Software AG, http://www.gdata.de
 *	Copyright (C) 2004, 2005 g10 Code GmbH
 * 
 * This file is part of the G DATA Outlook Plugin for GnuPG.
 * 
 * This plugin is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2 of the License, or (at your option) any later version.
 * 
 * This plugin is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Lesser General Public License for more details.
 * 
 * You should have received a copy of the GNU Lesser General
 * Public License along with this plugin; if not, write to the 
 * Free Software Foundation, Inc., 59 Temple Place - Suite 330, 
 * Boston, MA 02111-1307, USA.
 */

#include "stdafx.h"
#include "GPGExchange.h"
#include "resource.h"
#include "MapiGPGME.h"


/* GPGOptionsDlgProc -
   Handles the notifications sent for managing the options property page. */
BOOL CALLBACK GPGOptionsDlgProc (HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    USES_CONVERSION;
    BOOL bMsgResult;    
    static LPNMHDR pnmhdr;
    static HWND hWndPage;
    
    switch (uMsg) {
    case WM_INITDIALOG:
	SetDlgItemText (hDlg, IDC_VERSION_INFO, "Version "VERSION " ("__DATE__")");
	return TRUE;
		
    case WM_COMMAND:
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

	    wsprintf(s, "%d", m_gpg->getStorePasswdTime ());
	    SendDlgItemMessage(hDlg, IDC_TIME_PHRASES, WM_SETTEXT, 0, (LPARAM) s);
	    f = m_gpg->getLogFile ();
	    SendDlgItemMessage (hDlg, IDC_DEBUG_LOGFILE, WM_SETTEXT, 0, (LPARAM)f);
	    hWndPage = pnmhdr->hwndFrom;   /* to be used in WM_COMMAND */
	    SendDlgItemMessage (hDlg, IDC_ENCRYPT_DEFAULT, BM_SETCHECK, 
				m_gpg->getEncryptDefault () ? 1 : 0, 0L);
	    SendDlgItemMessage (hDlg, IDC_SIGN_DEFAULT, BM_SETCHECK, 
			        m_gpg->getSignDefault () ? 1 : 0, 0L);
	    SendDlgItemMessage (hDlg, IDC_ENCRYPT_WITH_STANDARD_KEY, BM_SETCHECK, 
			        m_gpg->getEncryptWithDefaultKey () ? 1 : 0, 0L);
	    SendDlgItemMessage (hDlg, IDC_SAVE_DECRYPTED, BM_SETCHECK, 
				m_gpg->getSaveDecryptedAttachments () ? 1 : 0, 0L);
	    bMsgResult = FALSE;  /* accepts activation */
	    break; }
		
	case PSN_APPLY:	{
	    TCHAR s[200];
	    SendDlgItemMessage (hDlg, IDC_TIME_PHRASES, WM_GETTEXT, 20, (LPARAM) s);		
	    m_gpg->setStorePasswdTime (atol (s));
	    SendDlgItemMessage (hDlg, IDC_DEBUG_LOGFILE, WM_GETTEXT, 200, (LPARAM)s);
	    m_gpg->setLogFile (s);
		
	    m_gpg->setEncryptDefault (SendDlgItemMessage(hDlg, IDC_ENCRYPT_DEFAULT, BM_GETCHECK, 0, 0L));
	    m_gpg->setSignDefault (SendDlgItemMessage(hDlg, IDC_SIGN_DEFAULT, BM_GETCHECK, 0, 0L));		
	    m_gpg->setEncryptWithDefaultKey (SendDlgItemMessage(hDlg, IDC_ENCRYPT_WITH_STANDARD_KEY, BM_GETCHECK, 0, 0L));
	    m_gpg->setSaveDecryptedAttachments (SendDlgItemMessage(hDlg, IDC_SAVE_DECRYPTED, BM_GETCHECK, 0, 0L));
	    m_gpg->writeOptions ();
	    bMsgResult = PSNRET_NOERROR;
	    break; }
		
	case PSN_HELP:
	    MessageBox (pnmhdr->hwndFrom,
			"This plugin was orginally created by G-DATA and called \"GPGExch\" http://www.gdata.de/gpg\n\n"
			"Later versions of the plugin were created by g10 Code GmbH (http://www.g10code.com)\n",
			"Outlook GnuPG-Plugin", MB_OK);
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
