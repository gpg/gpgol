/* config-dialog.c
 *	Copyright (C) 2005, 2008 g10 Code GmbH
 *	Copyright (C) 2003 Timo Schulz
 *
 * This file is part of GpgOL.
 * 
 * GpgOL is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 * 
 * GpgOL is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Lesser General Public License for more details.
 * 
 * You should have received a copy of the GNU Lesser General Public License
 * along with this program; if not, see <http://www.gnu.org/licenses/>.
 */

#include <config.h>

#include <windows.h>
#include <shellapi.h>
#include <shlobj.h>
#include <time.h>
#include <sys/stat.h>
#include <unistd.h>

#include "common.h"
#include "gpgol-ids.h"
#include "dialogs.h"

/* Registry path to GnuPG */
#define REGPATH "Software\\GNU\\GnuPG"

/* Registry path to store plugin settings */
#define GPGOL_REGPATH "Software\\GNU\\GpgOL"


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

    *val = NULL;
    if (hk == NULL)
	hk = HKEY_CURRENT_USER;
    ec = RegOpenKeyEx (hk, path, 0, KEY_READ, &h);
    if (ec != ERROR_SUCCESS)
	return -1;

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
  int type;
  int ec;
  
  if (hk == NULL)
    hk = HKEY_CURRENT_USER;
  ec = RegCreateKeyEx (hk, path, 0, NULL, REG_OPTION_NON_VOLATILE,
                       KEY_ALL_ACCESS, NULL, &h, NULL);
  if (ec != ERROR_SUCCESS)
    {
      log_debug_w32 (ec, "creating/opening registry key `%s' failed", path);
      return -1;
    }
  type = strchr (val, '%')? REG_EXPAND_SZ : REG_SZ;
  ec = RegSetValueEx (h, key, 0, type, (const BYTE*)val, strlen (val));
  if (ec != ERROR_SUCCESS)
    {
      log_debug_w32 (ec, "saving registry key `%s'->`%s' failed", path, key);
      RegCloseKey(h);
      return -1;
    }
  RegCloseKey(h);
  return 0;
}


/* To avoid writing a dialog template for each language we use gettext
   for the labels and hope that there is enough space in the dialog to
   fit teh longest translation.  */
static void
config_dlg_set_labels (HWND dlg)
{
  static struct { int itemid; const char *label; } labels[] = {
    { IDC_T_DEBUG_LOGFILE,  N_("Debug output (for analysing problems)")},
    { 0, NULL}
  };
  int i;

  for (i=0; labels[i].itemid; i++)
    SetDlgItemText (dlg, labels[i].itemid, _(labels[i].label));
  
}  

static BOOL CALLBACK
config_dlg_proc (HWND dlg, UINT msg, WPARAM wparam, LPARAM lparam)
{
  char name[MAX_PATH+1];
  int n;
  const char *s;

  (void)lparam;
  
  switch (msg) 
    {
    case WM_INITDIALOG:
      center_window (dlg, 0);
      s = get_log_file ();
      SetDlgItemText (dlg, IDC_DEBUG_LOGFILE, s? s:"");
      config_dlg_set_labels (dlg);
      break;
      
    case WM_COMMAND:
      switch (LOWORD (wparam)) 
        {
	case IDOK:
          n = GetDlgItemText (dlg, IDC_DEBUG_LOGFILE, name, MAX_PATH-1);
          set_log_file (n>0?name:NULL);
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
#ifndef WIN64
  int resid;

  resid = IDD_EXT_OPTIONS;

  if (!parent)
    parent = GetDesktopWindow ();
  DialogBoxParam (glob_hinst, (LPCTSTR)resid, parent, config_dlg_proc, 0);
#else
  (void)parent;
  (void)config_dlg_proc;
#endif
}



/* Store a key in the registry with the key given by @key and the 
   value @value. */
int
store_extension_value (const char *key, const char *val)
{
    return store_config_value (HKEY_CURRENT_USER, GPGOL_REGPATH, key, val);
}

/* Load a key from the registry with the key given by @key. The value is
   returned in @val and needs to freed by the caller. */
int
load_extension_value (const char *key, char **val)
{
    return load_config_value (HKEY_CURRENT_USER, GPGOL_REGPATH, key, val);
}
