/* config-dialog.c
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
#include <shellapi.h>
#include <shlobj.h>
#include <time.h>
#include <gpgme.h>
#include "resource.h"
#include "intern.h"

static char*
get_open_file_name (const char *dir)
{
    static char fname[MAX_PATH+1];
    OPENFILENAME ofn;

    memset (&ofn, 0, sizeof ofn);
    memset (fname, 0, sizeof fname);
    ofn.hwndOwner = GetDesktopWindow();
    ofn.hInstance = glob_hinst;
    ofn.lpstrTitle = "Select GnuPG binary";
    ofn.lStructSize = sizeof (ofn);
    ofn.lpstrInitialDir = dir;
    ofn.lpstrFilter = "*.EXE";
    ofn.lpstrFile = fname;
    ofn.nMaxFile = sizeof(fname)-1;
    if (GetOpenFileName(&ofn) == FALSE)
	return NULL;
    return fname;
}

static char*
get_folder (void)
{
    static char fname[MAX_PATH+1];
    BROWSEINFO bi;
    ITEMIDLIST * il;

    memset (&bi, 0, sizeof (bi));
    memset (fname, 0, sizeof (fname));
    bi.hwndOwner = GetDesktopWindow();
    bi.lpszTitle = "Select GnuPG home directory";
    il = SHBrowseForFolder (&bi);
    if (il) {
	SHGetPathFromIDList (il, fname);
	return fname;
    }
    return NULL;
}


static int
load_config_value_ext (char **val)
{
    static char buf[MAX_PATH+64];
    BOOL ec;

    memset (buf, 0, sizeof (buf));
    /* MSDN: This buffer must be at least MAX_PATH characters in size. */
    ec = SHGetSpecialFolderPath (HWND_DESKTOP, buf, CSIDL_APPDATA, TRUE);
    if (ec != 1)
	return -1;
    strcat (buf, "\\gnupg");
    if (GetFileAttributes (buf) == 0xFFFFFFFF)
	return -1;
    *val = buf;
    return 0;
}


static int
load_config_value (const char *key, char **val)
{
    HKEY h;
    DWORD size=0, type;
    int ec;

    ec = RegOpenKeyEx (HKEY_CURRENT_USER, "Software\\GNU\\GnuPG", 0, KEY_READ, &h);
    if (ec != ERROR_SUCCESS)
	return load_config_value_ext (val);

    ec = RegQueryValueEx(h, key, NULL, &type, NULL, &size);
    if (ec != ERROR_SUCCESS) {
	RegCloseKey(h);
	return -1;
    }
    *val = calloc(1, size+1);
    if (!*val)
	abort ();
    ec = RegQueryValueEx (h, key, NULL, &type, (BYTE*)*val, &size);
    if (ec != ERROR_SUCCESS) {
	free (*val); *val = NULL;
	RegCloseKey (h);
	return -1;
    }
    return 0;
}


static int
store_config_value(const char *key, const char *val)
{
    HKEY h;
    int ec;

    ec = RegOpenKeyEx(HKEY_CURRENT_USER, "Software\\GNU\\GnuPG", 0, KEY_ALL_ACCESS, &h);
    if (ec != ERROR_SUCCESS)
	return -1;
    ec = RegSetValueEx (h, key, 0, REG_SZ, (const BYTE*)val, strlen(val));
    if (ec != ERROR_SUCCESS) {
	RegCloseKey(h);
	return -1;
    }
    RegCloseKey(h);
    return 0;
}


static BOOL CALLBACK
config_dlg_proc(HWND dlg, UINT msg, WPARAM wparam, LPARAM lparam)
{
    char *buf;
    char name[MAX_PATH];
    int n;

    switch (msg) {
    case WM_INITDIALOG:
	center_window(dlg, 0);
	if (!load_config_value("gpgProgram", &buf)) {
	    SetDlgItemText(dlg, IDC_OPT_GPGPRG, buf);
	    free (buf);
	}
	if (!load_config_value("HomeDir", &buf)) {
	    SetDlgItemText(dlg, IDC_OPT_HOMEDIR, buf);
	    free (buf);
	}
	break;

    case WM_COMMAND:
	switch (LOWORD(wparam)) {
	case IDC_OPT_SELPRG:
	    buf = get_open_file_name(NULL);
	    if (buf && *buf)
		SetDlgItemText(dlg, IDC_OPT_GPGPRG, buf);
	    break;

	case IDC_OPT_SELHOMEDIR:
	    buf = get_folder();
	    if (buf && *buf)
		SetDlgItemText(dlg, IDC_OPT_HOMEDIR, buf);
	    break;

	case IDOK:
	    n = GetDlgItemText(dlg, IDC_OPT_GPGPRG, name, MAX_PATH-1);
	    if (n > 0)
		store_config_value("gpgProgram", name);
	    n = GetDlgItemText(dlg, IDC_OPT_HOMEDIR, name, MAX_PATH-1);
	    if (n > 0)
		store_config_value("HomeDir", name);
	    EndDialog(dlg, TRUE);
	    break;
	}
	break;
    }

    return FALSE;
}


void
config_dialog_box(HWND parent)
{
    if (parent == NULL)
	parent = GetDesktopWindow();
    DialogBoxParam(glob_hinst, (LPCTSTR)IDD_OPT, parent,
		    config_dlg_proc, 0);
}
