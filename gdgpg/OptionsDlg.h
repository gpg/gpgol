/* OptionsDlg.h - declaration of the options dialog class
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

#if !defined(AFX_OPTIONSDLG_H__B64CE15A_A281_4E34_8758_84EE5E63AB30__INCLUDED_)
#define AFX_OPTIONSDLG_H__B64CE15A_A281_4E34_8758_84EE5E63AB30__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class COptionsDlg : public CDialogImpl<COptionsDlg>  
{
public:
	enum { IDD = IDD_OPTIONS };

	string m_sGPG;
	string m_sKeyManager;

protected:
	CEdit   m_ctlGPG;
	CEdit   m_ctlKeyManager;

protected:
	BEGIN_MSG_MAP_EX(COptionsDlg)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		COMMAND_ID_HANDLER(IDOK, OnOk)
		COMMAND_ID_HANDLER(IDCANCEL, OnCancel)
		COMMAND_ID_HANDLER(IDC_SELECT_GPG, OnSelectGPG)
		COMMAND_ID_HANDLER(IDC_SELECT_KEY_MANAGER, OnSelectKeyManager)
	END_MSG_MAP()

	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnOk(WORD /*wNotifyCode*/, WORD wID, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnCancel(WORD /*wNotifyCode*/, WORD wID, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnSelectGPG(WORD /*wNotifyCode*/, WORD wID, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnSelectKeyManager(WORD /*wNotifyCode*/, WORD wID, HWND /*hWndCtl*/, BOOL& /*bHandled*/);

	BOOL SelectFile(string &sFilename, LPCSTR pszTitle, LPSTR pszFilter);
};

#endif // !defined(AFX_OPTIONSDLG_H__B64CE15A_A281_4E34_8758_84EE5E63AB30__INCLUDED_)
