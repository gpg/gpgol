/* SelectKeyDlg.h - declaration of the select key dialog class
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

#if !defined(AFX_SELECTKEYDLG_H__59DD6C72_37A1_4E79_98D9_2C3C99BFDF8A__INCLUDED_)
#define AFX_SELECTKEYDLG_H__59DD6C72_37A1_4E79_98D9_2C3C99BFDF8A__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "key.h"

class CSelectKeyDlg : public CDialogImpl<CSelectKeyDlg>  
{
public:
	enum { IDD = IDD_SELECT_KEY };

	CSelectKeyDlg();
	void SetKeyList(vector<CKey>* pKeyList);
	void SetExceptionKeyID(string sID);

protected:
	CButton	      m_ctlOKButton;
	CButton	      m_ctlAddButton;
	CButton	      m_ctlDeleteButton;
	CListViewCtrl m_ctlKeyList;
	CListViewCtrl m_ctlRecipientList;
	string        m_sExceptionKeyID;

	vector<CKey>* m_pKeyList; 

protected:
	BEGIN_MSG_MAP_EX(CSelectKeyDlg)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		COMMAND_ID_HANDLER_EX(IDC_ADD, OnAdd)
		COMMAND_ID_HANDLER_EX(IDC_DELETE, OnDelete)
		COMMAND_ID_HANDLER(IDOK, OnOk)
		COMMAND_ID_HANDLER(IDCANCEL, OnCancel)
		NOTIFY_HANDLER(IDC_KEY_LIST, NM_DBLCLK, OnKeyListDoubleClick)
		NOTIFY_HANDLER(IDC_KEY_LIST, LVN_ITEMCHANGED, OnKeyListSelChanged)
		NOTIFY_HANDLER(IDC_RECIPIENT_LIST, NM_DBLCLK, OnRecipientListDoubleClick)
		NOTIFY_HANDLER(IDC_RECIPIENT_LIST, LVN_ITEMCHANGED, OnRecipientListSelChanged)
	END_MSG_MAP()

	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnOk(WORD /*wNotifyCode*/, WORD wID, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnCancel(WORD /*wNotifyCode*/, WORD wID, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnKeyListDoubleClick(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);
	LRESULT OnKeyListSelChanged(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);
	LRESULT OnRecipientListDoubleClick(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);
	LRESULT OnRecipientListSelChanged(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);

	void OnAdd(UINT, int, HWND);
	void OnDelete(UINT, int, HWND);
	void FillLists();
	void AddRecipients();
	void DeleteRecipients();
	void InsertKeyInList(CListViewCtrl* pListCtrl, CKey* pKey);

};

#endif // !defined(AFX_SELECTKEYDLG_H__59DD6C72_37A1_4E79_98D9_2C3C99BFDF8A__INCLUDED_)
