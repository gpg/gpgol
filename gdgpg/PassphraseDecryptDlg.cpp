/* PasspharseDecryptDlg.cpp - implementation of the dialog class to
 *                  input the passpharse for decryption
 *	Copyright (C) 2001 G Data Software AG, http://www.gdata.de
 *	Copyright (C) 2004 g10 Code GmbH
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

#include <atlframe.h>
#include <atlctrls.h>
#include <atldlgs.h>
#include <atlctrlw.h>

#include "resource.h"

#include "PassphraseDecryptDlg.h"




LRESULT 
CPassphraseDecryptDlg::OnInitDialog (UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
    m_ctlPassphrase = GetDlgItem (IDC_PASSPHRASE);
    m_ctlKeyList = GetDlgItem (IDC_KEY_LIST);

    string sKeys = m_sKeys + "\n";
    while (sKeys.find ('\n') != -1)
    {
	string sKey = sKeys.substr (0, sKeys.find ('\n'));
	sKeys = sKeys.substr (sKeys.find ('\n')+1);
	m_ctlKeyList.AddString (sKey.c_str ());
    }
    if (m_nUnknownKeys > 0)
    {
	TCHAR s[100], s1[100];
	LoadString(_Module.m_hInstResource, IDS_UNKNOWN_KEYS, s1, sizeof(s1));
	wsprintf(s, s1, m_nUnknownKeys);
	m_ctlKeyList.AddString(s);	
    }
    
    CenterWindow ();
    return TRUE;
}


LRESULT
CPassphraseDecryptDlg::OnOk(WORD /*wNotifyCode*/, WORD wID, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
    TCHAR s[256];
    m_ctlPassphrase.GetWindowText (s, 255);
    m_sPassphrase = s;
    EndDialog(wID);
    return 0;
}


LRESULT 
CPassphraseDecryptDlg::OnCancel(WORD /*wNotifyCode*/, WORD wID, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
    EndDialog (wID);
    return 0;
}

string 
CPassphraseDecryptDlg::GetPassphrase (void)
{
	return m_sPassphrase;
}

void 
CPassphraseDecryptDlg::SetKeys (string sKeys, int nUnknownKeys)
{
    m_sKeys = sKeys;
    m_nUnknownKeys = nUnknownKeys;
}
