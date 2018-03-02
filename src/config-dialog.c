/* config-dialog.c
 * Copyright (C) 2005, 2008 g10 Code GmbH
 * Copyright (C) 2003 Timo Schulz
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

#include "common.h"
#include "gpgol-ids.h"
#include "dialogs.h"

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
#ifndef _WIN64
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
