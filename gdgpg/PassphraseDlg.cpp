/* PassphraseDlg.cpp - implementation of the passphrase dialog class
 *	Copyright (C) 2001 G Data Software AG, http://www.gdata.de
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

#include "PassphraseDlg.h"

LRESULT 
CPassphraseDlg::OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
    // connect controls
    m_ctlPassphrase = GetDlgItem(IDC_PASSPHRASE);
    m_ctlMessage = GetDlgItem(IDC_MESSAGE);
    m_ctlMessage.SetWindowText(m_sMessage.c_str());
    // center the dialog on the screen
    CenterWindow();
    return TRUE;
}

LRESULT 
CPassphraseDlg::OnOk(WORD /*wNotifyCode*/, WORD wID, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{

    TCHAR s[256];
    m_ctlPassphrase.GetWindowText(s, 255);
    m_sPassphrase = s;
    EndDialog(wID);
    return 0;
}

LRESULT 
CPassphraseDlg::OnCancel(WORD /*wNotifyCode*/, WORD wID, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
    EndDialog(wID);
    return 0;
}

string 
CPassphraseDlg::GetPassphrase()
{
    return m_sPassphrase;
}

void 
CPassphraseDlg::SetMessage(string sMessage)
{
    m_sMessage = sMessage;
}

