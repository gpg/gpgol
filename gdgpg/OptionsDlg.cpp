/* OptionsDlg.cpp - implementation of the options dialog class
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

#include "OptionsDlg.h"

LRESULT 
COptionsDlg::OnInitDialog (UINT /*uMsg*/, WPARAM /*wParam*/, 
			   LPARAM /*lParam*/, BOOL& /*bHandled*/)
{

    m_ctlGPG = GetDlgItem (IDC_GPG);
    m_ctlKeyManager = GetDlgItem (IDC_KEY_MANAGER);

    m_ctlGPG.SetWindowText (m_sGPG.c_str());
    m_ctlKeyManager.SetWindowText (m_sKeyManager.c_str());
   
    CenterWindow ();

    return TRUE;
}


LRESULT 
COptionsDlg::OnOk (WORD /*wNotifyCode*/, WORD wID, 
		   HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{	
    TCHAR s[MAX_PATH];

    m_ctlGPG.GetWindowText(s, sizeof(s));
    m_sGPG = s;
    m_ctlKeyManager.GetWindowText(s, sizeof(s));
    m_sKeyManager = s;
    EndDialog(wID);
    return 0;
}


LRESULT 
COptionsDlg::OnCancel(WORD /*wNotifyCode*/, WORD wID, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
    EndDialog(wID);
    return 0;
}


LRESULT
COptionsDlg::OnSelectGPG(WORD /*wNotifyCode*/, WORD wID, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
    TCHAR szTitle[256] = "";
    TCHAR szFilter[256] = "";

    LoadString (_Module.m_hInstResource, IDS_SELECT_PROGRAM, szTitle, sizeof(szTitle));
    LoadString (_Module.m_hInstResource, IDS_SELECT_PROGRAM_FILTER, szFilter, sizeof(szFilter));
    if (SelectFile (m_sGPG, szTitle, szFilter))
	m_ctlGPG.SetWindowText (m_sGPG.c_str());
    return 0;
}


LRESULT
COptionsDlg::OnSelectKeyManager (WORD /*wNotifyCode*/, WORD wID, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
    TCHAR szTitle[256] = "";
    TCHAR szFilter[256] = "";
    const char * s;

    LoadString (_Module.m_hInstResource, IDS_SELECT_PROGRAM, szTitle, sizeof(szTitle));
    LoadString (_Module.m_hInstResource, IDS_SELECT_PROGRAM_FILTER, szFilter, sizeof(szFilter));
    if (SelectFile(m_sKeyManager, szTitle, szFilter))
    {
	s = m_sKeyManager.c_str ();
	if ((strstr (s, "winpt.exe") || strstr (s, "WinPT.exe") || strstr (s, "WINPT.EXE"))
	    && !strstr (s, "--keymanager"))
	    m_sKeyManager += " --keymanager";
	m_ctlKeyManager.SetWindowText(m_sKeyManager.c_str());
    }
    return 0;
}


BOOL 
COptionsDlg::SelectFile (string &sFilename, LPCSTR pszTitle, LPSTR pszFilter)
{
    OPENFILENAME ofn;	
    TCHAR szFile[MAX_PATH];
    int nLength = strlen(pszFilter);

    for (int i=0; i<nLength; i++)
    {
	if (pszFilter[i] == '|')
	    pszFilter[i] = '\0';
    }

    if (strlen (sFilename.c_str ()) > MAX_PATH)
	return FALSE; /* avoid buffer overflow */
    strcpy (szFile, sFilename.c_str());
    memset (&ofn, 0, sizeof (ofn));
    ofn.lStructSize = sizeof (ofn);
    ofn.hwndOwner = m_hWnd;
    ofn.lpstrFile = szFile;
    ofn.nMaxFile = MAX_PATH;
    //	ofn.lpstrDefExt = "*.exe";
    ofn.lpstrFileTitle = NULL;
    ofn.nMaxFileTitle = 0;
    ofn.Flags |= OFN_HIDEREADONLY | OFN_FILEMUSTEXIST;
    ofn.lpstrTitle = pszTitle;
    ofn.lpstrFilter = pszFilter;
    if (GetOpenFileName (&ofn))
    {
	sFilename = szFile;
	return TRUE;	
    }
    return FALSE;
}
