/* olflange-dlgs.cpp - New dialogs for Outlook.
 * Copyright (C) 2004, 2005, 2006, 2007, 2008 g10 Code GmbH
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
#include <config.h>
#endif

#include <windows.h>
#include <prsht.h>

#include "mymapi.h"
#include "mymapitags.h"
#include "myexchext.h"
#include "common.h"
#include "display.h"
#include "olflange-def.h"
#include "dialogs.h"
#include "engine.h"

/* To avoid writing a dialog template for each language we use gettext
   for the labels and hope that there is enough space in the dialog to
   fit teh longest translation.  */
static void
set_labels (HWND dlg)
{
  static struct { int itemid; const char *label; } labels[] = {
    { IDC_G_GENERAL,        N_("General")},
    { IDC_ENABLE_SMIME,     N_("Enable the S/MIME support")},

    { IDC_G_SEND,           N_("Message sending")},
    { IDC_ENCRYPT_DEFAULT,  N_("&Encrypt new messages by default")},
    { IDC_SIGN_DEFAULT,     N_("&Sign new messages by default")},

    { IDC_G_RECV,           N_("Message receiving")},
//     { IDC_PREVIEW_DECRYPT,  N_("Also decrypt in preview window")},
    { IDC_PREFER_HTML,      N_("Show HTML view if possible")},
    { IDC_BODY_AS_ATTACHMENT, N_("Present encrypted message as attachment")},

    { IDC_GPG_OPTIONS,      "Debug..."},
    { IDC_GPG_CONF,         N_("Crypto Engine")},
    { IDC_VERSION_INFO,  "Version " VERSION},
    { 0, NULL}
  };
  int i;

  for (i=0; labels[i].itemid; i++)
    SetDlgItemText (dlg, labels[i].itemid, _(labels[i].label));

}


/* GPGOptionsDlgProc -
   Handles the notifications sent for managing the options property page. */
bool
GPGOptionsDlgProc (HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  BOOL bMsgResult = FALSE;
  static LPNMHDR pnmhdr;
//   static BOOL openpgp_state = FALSE;
//   static BOOL smime_state = FALSE;

  switch (uMsg)
    {
    case WM_INITDIALOG:
      {
//         openpgp_state = (opt.default_protocol == PROTOCOL_OPENPGP);
//         smime_state = (opt.default_protocol == PROTOCOL_SMIME);

        EnableWindow (GetDlgItem (hDlg, IDC_SMIME_DEFAULT),
                      !!opt.enable_smime);
        set_labels (hDlg);
        ShowWindow (GetDlgItem (hDlg, IDC_GPG_OPTIONS),
                    opt.enable_debug? SW_SHOW : SW_HIDE);
      }
      return TRUE;

    case WM_LBUTTONDOWN:
      {
        int x = LOWORD (lParam);
        int y = HIWORD (lParam);
        RECT rect_banner = {0,0,0,0};
        RECT rect_dlg = {0,0,0,0};
        HWND bitmap;

        GetWindowRect (hDlg, &rect_dlg);
        bitmap = GetDlgItem (hDlg, IDC_G10CODE_STRING);
        if (bitmap)
          GetWindowRect (bitmap, &rect_banner);

        rect_banner.left   -= rect_dlg.left;
        rect_banner.right  -= rect_dlg.left;
        rect_banner.top    -= rect_dlg.top;
        rect_banner.bottom -= rect_dlg.top;

        if (x >= rect_banner.left && x <= rect_banner.right
            && y >= rect_banner.top && y <= rect_banner.bottom)
          {
            ShellExecute (NULL, "open",
                          "http://www.g10code.com/p-gpgol.html",
                          NULL, NULL, SW_SHOWNORMAL);
          }
      }
      break;

    case WM_COMMAND:
      if (HIWORD (wParam) == BN_CLICKED)
	{
	  /* If dialog state has been changed, activate the confirm button. */
	  switch (wParam)
	    {
	    case IDC_ENABLE_SMIME:
	    case IDC_ENCRYPT_DEFAULT:
	    case IDC_SIGN_DEFAULT:
// 	    case IDC_OPENPGP_DEFAULT:
// 	    case IDC_SMIME_DEFAULT:
// 	    case IDC_PREVIEW_DECRYPT:
	      SendMessage (GetParent (hDlg), PSM_CHANGED, (WPARAM)hDlg, 0L);
	      break;
	    }
	}
      if (HIWORD (wParam) == BN_CLICKED &&
	  LOWORD (wParam) == IDC_ENABLE_SMIME)
	{
	  opt.enable_smime = !opt.enable_smime;
	  EnableWindow (GetDlgItem (hDlg, IDC_SMIME_DEFAULT),
                        opt.enable_smime);
	}
//       if (HIWORD (wParam) == BN_CLICKED &&
// 	  LOWORD (wParam) == IDC_OPENPGP_DEFAULT)
// 	{
// 	  openpgp_state = !openpgp_state;
//           if (openpgp_state)
//             {
//               smime_state = 0;
//               SendDlgItemMessage (hDlg, IDC_SMIME_DEFAULT, BM_SETCHECK,0,0L);
//             }
// 	}
//       if (HIWORD (wParam) == BN_CLICKED &&
// 	  LOWORD (wParam) == IDC_SMIME_DEFAULT)
// 	{
// 	  smime_state = !smime_state;
//           if (smime_state)
//             {
//               openpgp_state = 0;
//               SendDlgItemMessage (hDlg, IDC_OPENPGP_DEFAULT, BM_SETCHECK,0,0L);
//             }
// 	}
      if (opt.enable_debug && LOWORD (wParam) == IDC_GPG_OPTIONS)
	config_dialog_box (hDlg);
      else if (LOWORD (wParam) == IDC_GPG_CONF)
        engine_start_confdialog (hDlg);
      break;

    case WM_NOTIFY:
      pnmhdr = ((LPNMHDR) lParam);
      switch (pnmhdr->code)
        {
	case PSN_KILLACTIVE:
          bMsgResult = FALSE;  /*(Allow this page to receive PSN_APPLY. */
          break;

	case PSN_SETACTIVE:
          SendDlgItemMessage (hDlg, IDC_ENABLE_SMIME, BM_SETCHECK,
                              !!opt.enable_smime, 0L);

          SendDlgItemMessage (hDlg, IDC_ENCRYPT_DEFAULT, BM_SETCHECK,
                              !!opt.encrypt_default, 0L);
          SendDlgItemMessage (hDlg, IDC_SIGN_DEFAULT, BM_SETCHECK,
                              !!opt.sign_default, 0L);
//           SendDlgItemMessage (hDlg, IDC_OPENPGP_DEFAULT, BM_SETCHECK,
//                                 openpgp_state, 0L);
//           SendDlgItemMessage (hDlg, IDC_SMIME_DEFAULT, BM_SETCHECK,
//                               smime_state, 0L);

//           SendDlgItemMessage (hDlg, IDC_PREVIEW_DECRYPT, BM_SETCHECK,
//                               !!opt.preview_decrypt, 0L);
          SendDlgItemMessage (hDlg, IDC_PREFER_HTML, BM_SETCHECK,
				!!opt.prefer_html, 0L);
          SendDlgItemMessage (hDlg, IDC_BODY_AS_ATTACHMENT, BM_SETCHECK,
				!!opt.body_as_attachment, 0L);
          bMsgResult = FALSE;  /* Accepts activation. */
          break;

	case PSN_APPLY:
          opt.enable_smime = !!SendDlgItemMessage
            (hDlg, IDC_ENABLE_SMIME, BM_GETCHECK, 0, 0L);

          opt.encrypt_default = !!SendDlgItemMessage
            (hDlg, IDC_ENCRYPT_DEFAULT, BM_GETCHECK, 0, 0L);
          opt.sign_default = !!SendDlgItemMessage
            (hDlg, IDC_SIGN_DEFAULT, BM_GETCHECK, 0, 0L);

//           if (openpgp_state)
//             opt.default_protocol = PROTOCOL_OPENPGP;
//           else if (smime_state && opt.enable_smime)
//             opt.default_protocol = PROTOCOL_SMIME;
//           else
            opt.default_protocol = PROTOCOL_UNKNOWN;

//           opt.preview_decrypt = !!SendDlgItemMessage
//             (hDlg, IDC_PREVIEW_DECRYPT, BM_GETCHECK, 0, 0L);
          opt.prefer_html = !!SendDlgItemMessage
            (hDlg, IDC_PREFER_HTML, BM_GETCHECK, 0, 0L);
          opt.body_as_attachment = !!SendDlgItemMessage
            (hDlg, IDC_BODY_AS_ATTACHMENT, BM_GETCHECK, 0, 0L);

          /* Make sure that no new-version-installed warning will pop
             up on the next start.  Not really needed as the warning
             dialog set this too, but it doesn't harm to do it again. */
          opt.git_commit = GIT_COMMIT;

          write_options ();
          bMsgResult = PSNRET_NOERROR;
          break;

	case PSN_HELP:
          {
            const char cpynotice[] = "Copyright (C) 2009 g10 Code GmbH";
            const char en_notice[] =
      "GpgOL is a plugin for Outlook to allow encryption and\n"
      "signing of messages using the OpenPGP and S/MIME standard.\n"
      "It uses the GnuPG software (http://www.gnupg.org). Latest\n"
      "release information are accessible by clicking on the logo.\n"
      "\n"
      "GpgOL is free software; you can redistribute it and/or\n"
      "modify it under the terms of the GNU Lesser General Public\n"
      "License as published by the Free Software Foundation; either\n"
      "version 2.1 of the License, or (at your option) any later version.\n"
      "\n"
      "GpgOL is distributed in the hope that it will be useful,\n"
      "but WITHOUT ANY WARRANTY; without even the implied warranty of\n"
      "MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the\n"
      "GNU Lesser General Public License for more details.\n"
      "\n"
      "You should have received a copy of the GNU Lesser General Public "
      "License\n"
      "along with this program; if not, see <http://www.gnu.org/licenses/>.";

            /* TRANSLATORS: See the source for the full english text.  */
            const char notice_key[] = N_("-#GpgOLFullHelpText#-");
            const char *notice;
            char header[300];
            char *buffer;
            size_t nbuffer;

            snprintf (header, sizeof header, _("This is GpgOL version %s"),
                      PACKAGE_VERSION);
            notice = _(notice_key);
            if (!strcmp (notice, notice_key))
              notice = en_notice;
            nbuffer = strlen (header)+strlen (cpynotice)+strlen (notice)+20;
            buffer = (char*)xmalloc (nbuffer);
            snprintf (buffer, nbuffer, "%s\n%s\n\n%s\n",
                      header, cpynotice, notice);
            MessageBox (pnmhdr->hwndFrom, buffer, "GpgOL", MB_OK);
            xfree (buffer);
          }
          bMsgResult = TRUE;
          break;

	default:
          bMsgResult = FALSE;
          break;
	}

#ifndef _WIN64
      /* SetWindowLong is not portable according to msdn
         it should be replaced by SetWindowLongPtr. But
         as this here is code for Outlook < 2010 we don't
         care as there is no 64bit version for that. */
      SetWindowLong (hDlg, DWL_MSGRESULT, bMsgResult);
#endif
      break;

    default:
      bMsgResult = FALSE;
      break;
    }

  return bMsgResult;
}
