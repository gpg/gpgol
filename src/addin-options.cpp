/* addin-options.cpp - Options for the Ol >= 2010 Addin
 * Copyright (C) 2015 by Bundesamt für Sicherheit in der Informationstechnik
 * Software engineering by Intevation GmbH
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

#ifdef HAVE_CONFIG_H
#include "config.h"
#endif

#include <windows.h>
#include "dialogs.h"
#include "common.h"

#include <string>

#include <gpgme++/context.h>
#include <gpgme++/data.h>

/* To avoid writing a dialog template for each language we use gettext
   for the labels and hope that there is enough space in the dialog to
   fit the longest translation.. */
static void
set_labels (HWND dlg)
{
  static struct { int itemid; const char *label; } labels[] = {
    { IDC_G_GENERAL,        N_("General")},
    { IDC_ENABLE_SMIME,     N_("Enable the S/MIME support")},

    { IDC_G_SEND,           N_("Message sending")},
    { IDC_ENCRYPT_DEFAULT,  N_("&Encrypt new messages by default")},
    { IDC_SIGN_DEFAULT,     N_("&Sign new messages by default")},
    { IDC_INLINE_PGP,       N_("&Send OpenPGP mails without "
                               "attachments as PGP/Inline")},
    { IDC_REPLYCRYPT,       N_("S&elect crypto settings automatically "
                               "for reply and forward")},
    { IDC_AUTORRESOLVE,     N_("&Resolve recipient keys automatically")},


    { IDC_GPG_OPTIONS,      N_("Debug...")},
    { IDC_GPG_CONF,         N_("Configure GnuPG")},
    { IDC_VERSION_INFO,     N_("Version ")VERSION},
    { 0, NULL}
  };
  int i;

  for (i=0; labels[i].itemid; i++)
    SetDlgItemText (dlg, labels[i].itemid, _(labels[i].label));
}

static void
launch_kleo_config (HWND hDlg)
{
  char *uiserver = get_uiserver_name ();
  bool showError = false;
  if (uiserver)
    {
      std::string path (uiserver);
      xfree (uiserver);
      if (path.find("kleopatra.exe") != std::string::npos)
        {
        size_t dpos;
        if ((dpos = path.find(" --daemon")) != std::string::npos)
            {
              path.erase(dpos, strlen(" --daemon"));
            }
          auto ctx = GpgME::Context::createForEngine(GpgME::SpawnEngine);
          if (!ctx)
            {
              log_error ("%s:%s: No spawn engine.",
                         SRCNAME, __func__);
            }
            std::string parentWid = std::to_string ((int) (intptr_t) hDlg);
            const char *argv[] = {path.c_str(),
                                  "--config",
                                  "--parent-windowid",
                                  parentWid.c_str(),
                                  NULL };
            log_debug ("%s:%s: Starting %s %s %s",
                       SRCNAME, __func__, path.c_str(), argv[1], argv[2]);
            GpgME::Data d(GpgME::Data::null);
            ctx->spawnAsync(path.c_str(), argv, d, d,
                            d, (GpgME::Context::SpawnFlags) (
                                GpgME::Context::SpawnAllowSetFg |
                                GpgME::Context::SpawnShowWindow));
        }
      else
        {
          showError = true;
        }
    }
  else
    {
      showError = true;
    }

  if (showError)
    {
      MessageBox (NULL,
                  _("Could not find Kleopatra.\n"
                  "Please reinstall Gpg4win with the Kleopatra component enabled."),
                  _("GpgOL"),
                  MB_ICONINFORMATION|MB_OK);
    }

}

static INT_PTR CALLBACK
options_window_proc (HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  (void)lParam;
  switch (uMsg)
    {
      case WM_INITDIALOG:
        {
          SendDlgItemMessage (hDlg, IDC_ENABLE_SMIME, BM_SETCHECK,
                              !!opt.enable_smime, 0L);
          SendDlgItemMessage (hDlg, IDC_ENCRYPT_DEFAULT, BM_SETCHECK,
                              !!opt.encrypt_default, 0L);
          SendDlgItemMessage (hDlg, IDC_SIGN_DEFAULT, BM_SETCHECK,
                              !!opt.sign_default, 0L);
          SendDlgItemMessage (hDlg, IDC_INLINE_PGP, BM_SETCHECK,
                              !!opt.inline_pgp, 0L);
          SendDlgItemMessage (hDlg, IDC_REPLYCRYPT, BM_SETCHECK,
                              !!opt.reply_crypt, 0L);
          SendDlgItemMessage (hDlg, IDC_AUTORRESOLVE, BM_SETCHECK,
                              !!opt.autoresolve, 0L);
          set_labels (hDlg);
          ShowWindow (GetDlgItem (hDlg, IDC_GPG_OPTIONS),
                      opt.enable_debug ? SW_SHOW : SW_HIDE);
        }
      return 1;
      case WM_LBUTTONDOWN:
        {
          return 1;
        }
      case WM_COMMAND:
        switch (LOWORD (wParam))
          {
            case IDOK:
              {
                opt.enable_smime = !!SendDlgItemMessage
                  (hDlg, IDC_ENABLE_SMIME, BM_GETCHECK, 0, 0L);

                opt.encrypt_default = !!SendDlgItemMessage
                  (hDlg, IDC_ENCRYPT_DEFAULT, BM_GETCHECK, 0, 0L);
                opt.sign_default = !!SendDlgItemMessage
                  (hDlg, IDC_SIGN_DEFAULT, BM_GETCHECK, 0, 0L);
                opt.inline_pgp = !!SendDlgItemMessage
                  (hDlg, IDC_INLINE_PGP, BM_GETCHECK, 0, 0L);

                opt.reply_crypt = !!SendDlgItemMessage
                  (hDlg, IDC_REPLYCRYPT, BM_GETCHECK, 0, 0L);

                opt.autoresolve = !!SendDlgItemMessage
                  (hDlg, IDC_AUTORRESOLVE, BM_GETCHECK, 0, 0L);

                write_options ();
                EndDialog (hDlg, TRUE);
                break;
              }
            case IDC_GPG_CONF:
              launch_kleo_config (hDlg);
              break;
            case IDC_GPG_OPTIONS:
              config_dialog_box (hDlg);
              break;
          }
      case WM_SYSCOMMAND:
        switch (LOWORD (wParam))
          {
            case SC_CLOSE:
              EndDialog (hDlg, TRUE);
          }

      break;
    }
  return 0;
}

void
options_dialog_box (HWND parent)
{
  int resid;

  resid = IDD_ADDIN_OPTIONS;

  if (!parent)
    parent = GetDesktopWindow ();
  DialogBoxParam (glob_hinst, MAKEINTRESOURCE (resid), parent,
                  options_window_proc, 0);
}
