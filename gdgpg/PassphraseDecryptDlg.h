/* PasspharseDecryptDlg.h - declaration of the dialog class to
 *                  input the passpharse for decryption
 * Copyright (C) 2001 G Data Software AG, http://www.gdata.de
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

#if !defined(AFX_PASSPHRASEDECRYPTDLG_H__195E40C5_2C38_4DE0_9E2C_C2D11B1F9FEA__INCLUDED_)
#define AFX_PASSPHRASEDECRYPTDLG_H__195E40C5_2C38_4DE0_9E2C_C2D11B1F9FEA__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CPassphraseDecryptDlg : public CDialogImpl<CPassphraseDecryptDlg>  
{
public:
	enum { IDD = IDD_PASSPHRASE_DECRYPT };

	string GetPassphrase();
	void SetKeys(string sKeys, int nUnknownKeys);

protected:
	CEdit    m_ctlPassphrase;
	CListBox m_ctlKeyList;
	string   m_sPassphrase;
	string   m_sKeys;
	int      m_nUnknownKeys;

protected:
	BEGIN_MSG_MAP_EX(CPassphraseDecryptDlg)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		COMMAND_ID_HANDLER(IDOK, OnOk)
		COMMAND_ID_HANDLER(IDCANCEL, OnCancel)
	END_MSG_MAP()

	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnOk(WORD /*wNotifyCode*/, WORD wID, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnCancel(WORD /*wNotifyCode*/, WORD wID, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
};

#endif // !defined(AFX_PASSPHRASEDECRYPTDLG_H__195E40C5_2C38_4DE0_9E2C_C2D11B1F9FEA__INCLUDED_)
