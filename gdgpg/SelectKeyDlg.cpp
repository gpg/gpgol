/* SelectKeyDlg.cpp - implementation of the select key dialog class
 *	Copyright (C) 2001 G Data Software AG, http://www.gdata.de
 *	Copyright (C) 2003 Timo Schulz
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

#include "SelectKeyDlg.h"


CSelectKeyDlg::CSelectKeyDlg()
{	
    m_pKeyList = NULL;
}


LRESULT 
CSelectKeyDlg::OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, 
			    LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
    TCHAR s[50];

    m_ctlOKButton = GetDlgItem(IDOK);
    m_ctlAddButton = GetDlgItem(IDC_ADD);
    m_ctlDeleteButton = GetDlgItem(IDC_DELETE);
    m_ctlKeyList = GetDlgItem(IDC_KEY_LIST);
    m_ctlRecipientList = GetDlgItem(IDC_RECIPIENT_LIST);

    ::SendMessage (m_ctlKeyList, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, 
		    (LPARAM) ::SendMessage(m_ctlKeyList, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
		    | LVS_EX_FULLROWSELECT);
    ::SendMessage (m_ctlRecipientList, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, 
		    (LPARAM) ::SendMessage(m_ctlRecipientList, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0) 
		    | LVS_EX_FULLROWSELECT);

    CenterWindow ();

    LoadString (_Module.m_hInstResource, IDS_KEYLIST_NAME, s, sizeof(s));
    m_ctlKeyList.InsertColumn (0, s, LVCFMT_LEFT, 180, -1);
    m_ctlRecipientList.InsertColumn (0, s, LVCFMT_LEFT, 180, -1);

    LoadString (_Module.m_hInstResource, IDS_KEYLIST_TRUST, s, sizeof(s));
    m_ctlKeyList.InsertColumn (1, s, LVCFMT_LEFT, 80, -1);
    m_ctlRecipientList.InsertColumn (1, s, LVCFMT_LEFT, 80, -1);

    LoadString(_Module.m_hInstResource, IDS_KEYLIST_ID, s, sizeof(s));
    m_ctlKeyList.InsertColumn (2, s, LVCFMT_LEFT, 80, -1);
    m_ctlRecipientList.InsertColumn (2, s, LVCFMT_LEFT, 80, -1);

    LoadString (_Module.m_hInstResource, IDS_KEYLIST_VALID, s, sizeof(s));
    m_ctlKeyList.InsertColumn (3, s, LVCFMT_LEFT, 100, -1);
    m_ctlRecipientList.InsertColumn (3, s, LVCFMT_LEFT, 100, -1);

    FillLists ();
    return TRUE;
}


void 
CSelectKeyDlg::OnAdd(UINT, int, HWND)
{
    AddRecipients ();
}

void 
CSelectKeyDlg::OnDelete(UINT, int, HWND)
{
    DeleteRecipients ();
}


LRESULT 
CSelectKeyDlg::OnOk(WORD /*wNotifyCode*/, WORD wID, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
    EndDialog (wID);
    return 0;
}


LRESULT 
CSelectKeyDlg::OnCancel(WORD /*wNotifyCode*/, WORD wID, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
    EndDialog (wID);
    return 0;
}


void 
CSelectKeyDlg::SetKeyList(vector<CKey>* pKeyList)
{
    m_pKeyList = pKeyList;
}


void 
CSelectKeyDlg::FillLists ()
{
    m_ctlKeyList.DeleteAllItems ();
    if (m_pKeyList == NULL)
	return;
    for (vector<CKey>::iterator i = m_pKeyList->begin(); 
	 i != m_pKeyList->end(); i++)
    {
	CListViewCtrl* pListCtrl = i->m_bSelected ? &m_ctlRecipientList : &m_ctlKeyList;
	if (!i->IsValidKey() || !i->CanDo (CAN_ENCRYPT))
	    continue;
	if (i->GetKeyID (0) != m_sExceptionKeyID)
	    InsertKeyInList(pListCtrl, i);
    }
    m_ctlOKButton.EnableWindow(m_ctlRecipientList.GetItemCount() > 0);
    m_ctlAddButton.EnableWindow(m_ctlKeyList.GetSelectedCount() > 0);
    m_ctlDeleteButton.EnableWindow(m_ctlRecipientList.GetSelectedCount() > 0);
}

LRESULT 
CSelectKeyDlg::OnKeyListDoubleClick(int idCtrl, LPNMHDR pnmh, BOOL& bHandled)
{
    AddRecipients ();
    return 0;
}


LRESULT 
CSelectKeyDlg::OnKeyListSelChanged (int idCtrl, LPNMHDR pnmh, BOOL& bHandled)
{
    m_ctlAddButton.EnableWindow (m_ctlKeyList.GetSelectedCount() > 0);
    return 0;
}


LRESULT 
CSelectKeyDlg::OnRecipientListDoubleClick (int idCtrl, LPNMHDR pnmh, BOOL& bHandled)
{
    DeleteRecipients ();
    return 0;
}


LRESULT 
CSelectKeyDlg::OnRecipientListSelChanged (int idCtrl, LPNMHDR pnmh, BOOL& bHandled)
{
    m_ctlDeleteButton.EnableWindow (m_ctlRecipientList.GetSelectedCount() > 0);
    return 0;
}


void 
CSelectKeyDlg::AddRecipients ()
{
    while (m_ctlKeyList.GetSelectedCount() > 0)
    {
	int nSel = m_ctlKeyList.GetNextItem (-1, LVNI_SELECTED);
	CKey* pKey = (CKey*) m_ctlKeyList.GetItemData (nSel);
	pKey->m_bSelected = TRUE;
	m_ctlKeyList.DeleteItem (nSel);
	InsertKeyInList (&m_ctlRecipientList, pKey);	
    }
    m_ctlOKButton.EnableWindow (m_ctlRecipientList.GetItemCount() > 0);
    m_ctlAddButton.EnableWindow (m_ctlKeyList.GetSelectedCount() > 0);
    m_ctlDeleteButton.EnableWindow (m_ctlRecipientList.GetSelectedCount() > 0);
}

void 
CSelectKeyDlg::DeleteRecipients (void)
{
    while (m_ctlRecipientList.GetSelectedCount () > 0)
    {
	int nSel = m_ctlRecipientList.GetNextItem (-1, LVNI_SELECTED);
	CKey* pKey = (CKey*) m_ctlRecipientList.GetItemData(nSel);
	pKey->m_bSelected = FALSE;
	m_ctlRecipientList.DeleteItem(nSel);
	InsertKeyInList(&m_ctlKeyList, pKey);	
    }
    m_ctlOKButton.EnableWindow(m_ctlRecipientList.GetItemCount() > 0);
    m_ctlAddButton.EnableWindow(m_ctlKeyList.GetSelectedCount() > 0);
    m_ctlDeleteButton.EnableWindow(m_ctlRecipientList.GetSelectedCount() > 0);
}


void 
CSelectKeyDlg::InsertKeyInList(CListViewCtrl* pListCtrl, CKey* pKey)
{
    char keyid[32];
    const char * p;
    int nItem;

    p = pKey->GetKeyID (-1).c_str();
    _snprintf (keyid, sizeof keyid-1, "0x%s", p+8);
    nItem = pListCtrl->InsertItem(0, pKey->m_sUser.c_str());
    pListCtrl->SetItemText(nItem, 1, pKey->GetTrustString().c_str());
    pListCtrl->SetItemText(nItem, 2, keyid);
    pListCtrl->SetItemText(nItem, 3, pKey->m_sValid.c_str());
    pListCtrl->SetItemData(nItem, (DWORD) pKey);
}

void 
CSelectKeyDlg::SetExceptionKeyID(string sID)
{
    m_sExceptionKeyID = sID;
}
