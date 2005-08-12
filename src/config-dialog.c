/* config-dialog.c
 *	Copyright (C) 2005 g10 Code GmbH
 *	Copyright (C) 2003 Timo Schulz
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
#include <sys/stat.h>
#include <gpgme.h>

#include "olgpgcoredlgs.h"
#include "keycache.h"
#include "intern.h"

#define REGPATH "Software\\GNU\\GnuPG"

static char*
get_open_file_name (const char *dir)
{
    static char fname[MAX_PATH+1];
    OPENFILENAME ofn;

    memset (&ofn, 0, sizeof (ofn));
    memset (fname, 0, sizeof (fname));
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


static void 
SHFree (void *p) 
{         
    IMalloc *pm;         
    SHGetMalloc (&pm);
    if (pm) {
	pm->lpVtbl->Free(pm,p);
	pm->lpVtbl->Release(pm);         
    } 
} 


/* Open the common dialog to select a folder. Caller has to free the string. */
static char*
get_folder (void)
{
    char fname[MAX_PATH+1];
    BROWSEINFO bi;
    ITEMIDLIST * il;
    char *path = NULL;

    memset (&bi, 0, sizeof (bi));
    memset (fname, 0, sizeof (fname));
    bi.hwndOwner = GetDesktopWindow ();
    bi.lpszTitle = "Select GnuPG home directory";
    il = SHBrowseForFolder (&bi);
    if (il != NULL) {
	SHGetPathFromIDList (il, fname);
	path = xstrdup (fname);
	SHFree (il);
    }
    return path;
}


static int
load_config_value_ext (char **val)
{
  static char buf[MAX_PATH+64];
  
  /* MSDN: This buffer must be at least MAX_PATH characters in size. */
  memset (buf, 0, sizeof (buf));
  if (w32_shgetfolderpath (NULL, CSIDL_APPDATA|CSIDL_FLAG_CREATE, 
                           NULL, 0, buf) < 0)
    return -1;
  strcat (buf, "\\gnupg");
  if (GetFileAttributes (buf) == 0xFFFFFFFF)
    return -1;
  *val = buf;
  return 0;
}


static char*
expand_path (const char *path)
{
    DWORD len;
    char *p;

    len = ExpandEnvironmentStrings (path, NULL, 0);
    if (!len)
	return NULL;
    len += 1;
    p = xcalloc (1, len+1);
    if (!p)
	return NULL;
    len = ExpandEnvironmentStrings (path, p, len);
    if (!len) {
	xfree (p);
	return NULL;
    }
    return p; 
}

static int
load_config_value (HKEY hk, const char *path, const char *key, char **val)
{
    HKEY h;
    DWORD size=0, type;
    int ec;

    if (hk == NULL)
	hk = HKEY_CURRENT_USER;
    ec = RegOpenKeyEx (hk, path, 0, KEY_READ, &h);
    if (ec != ERROR_SUCCESS)
	return load_config_value_ext (val);

    ec = RegQueryValueEx(h, key, NULL, &type, NULL, &size);
    if (ec != ERROR_SUCCESS) {
	RegCloseKey (h);
	return -1;
    }
    if (type == REG_EXPAND_SZ) {
	char tmp[256]; /* XXX: do not use a static buf */
	RegQueryValueEx (h, key, NULL, NULL, (BYTE*)tmp, &size);
	*val = expand_path (tmp);
    }
    else {
	*val = xcalloc(1, size+1);
	ec = RegQueryValueEx (h, key, NULL, &type, (BYTE*)*val, &size);
	if (ec != ERROR_SUCCESS) {
	    xfree (*val);
	    *val = NULL;
	    RegCloseKey (h);
	    return -1;
	}
    }
    RegCloseKey (h);
    return 0;
}



static int
store_config_value (HKEY hk, const char *path, const char *key, const char *val)
{
    HKEY h;
    int type = REG_SZ;
    int ec;

    if (hk == NULL)
	hk = HKEY_CURRENT_USER;
    ec = RegOpenKeyEx (hk, path, 0, KEY_ALL_ACCESS, &h);
    if (ec != ERROR_SUCCESS)
	return -1;
    if (strchr (val, '%'))
	type = REG_EXPAND_SZ;
    ec = RegSetValueEx (h, key, 0, type, (const BYTE*)val, strlen(val));
    if (ec != ERROR_SUCCESS) {
	RegCloseKey(h);
	return -1;
    }
    RegCloseKey(h);
    return 0;
}


static int
does_folder_exist (const char *path)
{
    int attrs = GetFileAttributes (path);
    int err = 0;

    if (attrs == 0xFFFFFFFF)
	err = -1;
    else if (!(attrs & FILE_ATTRIBUTE_DIRECTORY))
	err = -1;
    if (err != 0) {
	const char *fmt = "\"%s\" either does not exist or is not a directory";
	char *p = xmalloc (strlen (fmt) + strlen (path) + 2 + 2);
	sprintf (p, fmt, path);
	MessageBox (NULL, p, "Config Error", MB_ICONERROR|MB_OK);
	xfree (p);
    }
    return err;
}


static int
does_file_exist (const char *name, int is_file)
{
    struct stat st;
    const char *s;
    char *p, *name2;
    int err = 0;

    /* check WinPT specific flags */
    if ((p=strstr (name, "--keymanager"))) {
	name2 = xcalloc (1, (p-name)+2);
	strncpy (name2, name, (p-name)-1);
    }
    else
	name2 = xstrdup (name);

    if (stat (name2, &st) == -1) {
	s = "\"%s\" does not exist.";
	p = xmalloc (strlen (s) + strlen (name2) + 2);
	sprintf (p, s, name2);
	MessageBox (NULL, p, "Config Error", MB_ICONERROR|MB_OK);
	err = -1;
    }
    else if (is_file && !(st.st_mode & _S_IFREG)) {
	s = "\"%s\" is not a regular file.";
	p = xmalloc (strlen (s) + strlen (name2) + 2);
	sprintf (p, s, name2);
	MessageBox (NULL, p, "Config Error", MB_ICONERROR|MB_OK);
	err = -1;
    }
    xfree (name2);
    xfree (p);
    return err;
}


static void
error_box (const char *title)
{	
    TCHAR buf[256];
    DWORD last_err;

    last_err = GetLastError ();
    FormatMessage (FORMAT_MESSAGE_FROM_SYSTEM, NULL, last_err, 
		   MAKELANGID (LANG_NEUTRAL, SUBLANG_DEFAULT), 
		   buf, sizeof (buf)-1, NULL);
    MessageBox (NULL, buf, title, MB_OK);
}


static BOOL CALLBACK
config_dlg_proc (HWND dlg, UINT msg, WPARAM wparam, LPARAM lparam)
{
    char *buf = NULL;
    char name[MAX_PATH+1];
    int n;

    switch (msg) {
    case WM_INITDIALOG:
	center_window (dlg, 0);
	if (!load_config_value (NULL, REGPATH, "gpgProgram", &buf)) {
	    SetDlgItemText (dlg, IDC_OPT_GPGPRG, buf);
	    xfree (buf); 
	    buf=NULL;
	}
	if (!load_config_value (NULL, REGPATH, "HomeDir", &buf)) {
	    SetDlgItemText (dlg, IDC_OPT_HOMEDIR, buf);
	    xfree (buf); 
	    buf=NULL;
	}
	if (!load_config_value (NULL, REGPATH, "keyManager", &buf)) {
	    SetDlgItemText (dlg, IDC_OPT_KEYMAN, buf);
	    xfree (buf);
	    buf=NULL;
	}
	break;

    case WM_COMMAND:
	switch (LOWORD(wparam)) {
	case IDC_OPT_SELPRG:
	    buf = get_open_file_name (NULL);
	    if (buf && *buf)
		SetDlgItemText(dlg, IDC_OPT_GPGPRG, buf);
	    break;

	case IDC_OPT_SELHOMEDIR:
	    buf = get_folder ();
	    if (buf && *buf)
		SetDlgItemText(dlg, IDC_OPT_HOMEDIR, buf);
	    xfree (buf);
	    break;

	case IDC_OPT_SELKEYMAN:
	    buf = get_open_file_name (NULL);
	    if (buf && *buf)
		SetDlgItemText (dlg, IDC_OPT_KEYMAN, buf);
	    break;

	case IDOK:
	    n = GetDlgItemText (dlg, IDC_OPT_GPGPRG, name, MAX_PATH-1);
	    if (n > 0) {
		if (does_file_exist (name, 1))
		    return FALSE;
		if (store_config_value (NULL, REGPATH, "gpgProgram", name))
		    error_box ("GPG Config");
	    }
	    n = GetDlgItemText (dlg, IDC_OPT_KEYMAN, name, MAX_PATH-1);
	    if (n > 0) {
		if (does_file_exist (name, 1))
		    return FALSE;
		if (store_config_value (NULL, REGPATH, "keyManager", name))
		    error_box ("GPG Config");
	    }
	    n = GetDlgItemText (dlg, IDC_OPT_HOMEDIR, name, MAX_PATH-1);
	    if (n > 0) {
		if (does_folder_exist (name))
		    return FALSE;
		if (store_config_value (NULL, REGPATH, "HomeDir", name))
		    error_box ("GPG Config");
	    }
	    EndDialog (dlg, TRUE);
	    break;
	}
	break;
    }

    return FALSE;
}


/* Display GPG configuration dialog. */
void
config_dialog_box (HWND parent)
{
    int resid=0;

    switch (GetUserDefaultLangID ()) {
    case 0x0407:    resid = IDD_OPT_DE;break;
    default:	    resid = IDD_OPT; break;
    }

    if (parent == NULL)
	parent = GetDesktopWindow ();
    DialogBoxParam (glob_hinst, (LPCTSTR)resid, parent,
		    config_dlg_proc, 0);
}


/* Start the key manager specified by the registry entry 'keyManager'. */
int
start_key_manager (void)
{
    PROCESS_INFORMATION pi;
    STARTUPINFO si;
    char *keyman = NULL;
    
    if (load_config_value (NULL, REGPATH, "keyManager", &keyman))
	return -1;

    /* create startup info for the gpg process */
    memset (&si, 0, sizeof (si));
    si.cb = sizeof (STARTUPINFO);
    si.dwFlags = STARTF_USESHOWWINDOW;
    si.wShowWindow = SW_SHOW;    

    if (CreateProcess (NULL, keyman,
			NULL, NULL, TRUE, CREATE_DEFAULT_ERROR_MODE,
			NULL, NULL, &si, &pi) == TRUE) {
	CloseHandle(pi.hProcess);
	CloseHandle(pi.hThread);	
    }

    xfree (keyman);
    return 0;
}


/* Store a key in the registry with the key given by @key and the 
   value @value. */
int
store_extension_value (const char *key, const char *val)
{
    return store_config_value (HKEY_LOCAL_MACHINE, 
	"Software\\Microsoft\\Exchange\\Client\\Extensions\\OutlGPG", 
	key, val);
}

/* Load a key from the registry with the key given by @key. The value is
   returned in @val and needs to freed by the caller. */
int
load_extension_value (const char *key, char **val)
{
    return load_config_value (HKEY_LOCAL_MACHINE, 
	"Software\\Microsoft\\Exchange\\Client\\Extensions\\OutlGPG", 
	key, val);
}
