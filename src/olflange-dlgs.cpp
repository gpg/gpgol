/* olflange-dlgs.cpp - New dialogs for Outlook.
 *	Copyright (C) 2001 G Data Software AG, http://www.gdata.de
 *	Copyright (C) 2004, 2005, 2006, 2007 g10 Code GmbH
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
#include "ol-ext-callback.h"
#include "olflange-def.h"
#include "dialogs.h"


/* To avoid writing a dialog template for each language we use gettext
   for the labels and hope that there is enough space in the dialog to
   fit teh longest translation.  */
static void
set_labels (HWND dlg)
{
  static struct { int itemid; const char *label; } labels[] = {
    { IDC_ENCRYPT_DEFAULT,  N_("&Encrypt new messages by default")},
    { IDC_SIGN_DEFAULT,     N_("&Sign new messages by default")},
    { IDC_OPENPGP_DEFAULT,  N_("Use OPENPGP by default")},
    { IDC_SMIME_DEFAULT,    N_("Use S/MIME by default")},
    { IDC_ENABLE_SMIME,     N_("Enable the S/MIME support")},
    { IDC_ENCRYPT_WITH_STANDARD_KEY, 
                     N_("Also encrypt message with the default certificate")},
    { IDC_PREVIEW_DECRYPT,  N_("Also decrypt in preview window")},
    { IDC_PREFER_HTML,      N_("Show HTML view if possible")},

    { IDC_G_PASSPHRASE,     N_("Passphrase")},
    { IDC_T_PASSPHRASE_TTL, N_("Cache &passphrase for")}, 
    { IDC_T_PASSPHRASE_MIN, N_("minutes")},

    { IDC_GPG_OPTIONS,      N_("Ad&vanced..")},
    { IDC_VERSION_INFO,  "Version "VERSION " ("__DATE__")"},
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
  static HWND hWndPage;
  static BOOL openpgp_state = FALSE;
  static BOOL smime_state = FALSE;
    
  switch (uMsg) 
    {
    case WM_INITDIALOG:
      {
        /* We don't use this anymore.  Actually I don't know why we
           used it at all.  Note that this unused code has been
           converted to use LoadImage instead of LoadBitmap. */
/*         HDC hdc = GetDC (hDlg); */
/*         if (hdc) */
/*           { */
/*             int bits_per_pixel = GetDeviceCaps (hdc, BITSPIXEL); */
/*             HANDLE bitmap; */
                
/*             if (bits_per_pixel > 15) */
/*               { */
/*                 bitmap = LoadImage (glob_hinst, MAKEINTRESOURCE(IDB_BANNER), */
/*                                     IMAGE_BITMAP, 0, 0, */
/*                                     LR_CREATEDIBSECTION | LR_LOADTRANSPARENT); */
/*                 if (bitmap) */
/*                   { */
/*                     HBITMAP old = (HBITMAP) SendDlgItemMessage */
/*                       (hDlg, IDC_BITMAP, STM_SETIMAGE, */
/*                        IMAGE_BITMAP, (LPARAM)bitmap); */
/*                     if (old) */
/*                       DeleteObject (old);	 */
/*                   }	 */
/*               }		 */
/*             ReleaseDC (hDlg, hdc);	 */
/*           } */
        
        openpgp_state = opt.default_protocol = PROTOCOL_OPENPGP;
        smime_state = opt.default_protocol = PROTOCOL_SMIME;

	EnableWindow (GetDlgItem (hDlg, IDC_ENCRYPT_TO),
                      !!opt.enable_default_key);
        EnableWindow (GetDlgItem (hDlg, IDC_SMIME_DEFAULT), 
                      !!opt.enable_smime);
	if (opt.enable_default_key)
          CheckDlgButton (hDlg, IDC_ENCRYPT_WITH_STANDARD_KEY, BST_CHECKED);
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
        bitmap = GetDlgItem (hDlg, IDC_BITMAP);
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
	    case IDC_ENCRYPT_WITH_STANDARD_KEY:
	    case IDC_PREFER_HTML:
	    case IDC_SIGN_DEFAULT:
	    case IDC_OPENPGP_DEFAULT:
	    case IDC_SMIME_DEFAULT:
	    case IDC_PREVIEW_DECRYPT:
	    case IDC_ENABLE_SMIME:
	      SendMessage (GetParent (hDlg), PSM_CHANGED, (WPARAM)hDlg, 0L);
	      break;
	    }
	}
      if (HIWORD (wParam) == BN_CLICKED &&
	  LOWORD (wParam) == IDC_ENCRYPT_WITH_STANDARD_KEY) 
	{
	  opt.enable_default_key = !opt.enable_default_key;
	  EnableWindow (GetDlgItem (hDlg, IDC_ENCRYPT_TO), 
			!!opt.enable_default_key);
	}
      if (HIWORD (wParam) == BN_CLICKED &&
	  LOWORD (wParam) == IDC_ENABLE_SMIME) 
	{
	  opt.enable_smime = !opt.enable_smime;
	  EnableWindow (GetDlgItem (hDlg, IDC_SMIME_DEFAULT), 
                        opt.enable_smime);
	}
      if (HIWORD (wParam) == BN_CLICKED &&
	  LOWORD (wParam) == IDC_OPENPGP_DEFAULT) 
	{
	  openpgp_state = !openpgp_state;
          if (openpgp_state)
            {
              smime_state = 0;
              SendDlgItemMessage (hDlg, IDC_SMIME_DEFAULT, BM_SETCHECK,0,0L);
            }
	}
      if (HIWORD (wParam) == BN_CLICKED &&
	  LOWORD (wParam) == IDC_SMIME_DEFAULT) 
	{
	  smime_state = !smime_state;
          if (smime_state)
            {
              openpgp_state = 0;
              SendDlgItemMessage (hDlg, IDC_OPENPGP_DEFAULT, BM_SETCHECK,0,0L);
            }
	}
      if (opt.enable_debug && LOWORD (wParam) == IDC_GPG_OPTIONS)
	config_dialog_box (hDlg);
      break;
	
    case WM_NOTIFY:
	pnmhdr = ((LPNMHDR) lParam);
	switch (pnmhdr->code) {
	case PSN_KILLACTIVE:
	    bMsgResult = FALSE;  /* allow this page to receive PSN_APPLY */
	    break;

	case PSN_SETACTIVE: {
	    TCHAR s[30];
	    
	    if (opt.default_key && *opt.default_key)		
                SetDlgItemText (hDlg, IDC_ENCRYPT_TO, opt.default_key);
            else
		SetDlgItemText (hDlg, IDC_ENCRYPT_TO, "");
	    wsprintf (s, "%d", opt.passwd_ttl/60);
	    SendDlgItemMessage (hDlg, IDC_TIME_PHRASES, WM_SETTEXT,
                               0, (LPARAM) s);
	    hWndPage = pnmhdr->hwndFrom;   // to be used in WM_COMMAND
	    SendDlgItemMessage (hDlg, IDC_ENCRYPT_DEFAULT, BM_SETCHECK, 
				!!opt.encrypt_default, 0L);
	    SendDlgItemMessage (hDlg, IDC_SIGN_DEFAULT, BM_SETCHECK, 
			        !!opt.sign_default, 0L);
	    SendDlgItemMessage (hDlg, IDC_ENCRYPT_WITH_STANDARD_KEY,
                                BM_SETCHECK, opt.enable_default_key, 0L);
            SendDlgItemMessage (hDlg, IDC_OPENPGP_DEFAULT, BM_SETCHECK, 
                                openpgp_state, 0L);
            SendDlgItemMessage (hDlg, IDC_SMIME_DEFAULT, BM_SETCHECK, 
                                smime_state, 0L);
	    SendDlgItemMessage (hDlg, IDC_ENABLE_SMIME, BM_SETCHECK,
				!!opt.enable_smime, 0L);
	    SendDlgItemMessage (hDlg, IDC_PREVIEW_DECRYPT, BM_SETCHECK,
				!!opt.preview_decrypt, 0L);
	    SendDlgItemMessage (hDlg, IDC_PREFER_HTML, BM_SETCHECK,
				!!opt.prefer_html, 0L);
	    bMsgResult = FALSE;  /* accepts activation */
	    break; }
		
	case PSN_APPLY:	{
	    TCHAR s[201];
            
            opt.enable_default_key = !!SendDlgItemMessage
              (hDlg, IDC_ENCRYPT_WITH_STANDARD_KEY, BM_GETCHECK, 0, 0L);

            GetDlgItemText (hDlg, IDC_ENCRYPT_TO, s, 200);
            if (strlen (s) > 0 && strchr (s, ' ')) 
              {
                if (opt.enable_default_key)
                  {
                    MessageBox (hDlg,_("The default certificate may not"
                                       " contain any spaces."),
                                "GpgOL", MB_ICONERROR|MB_OK);
                    bMsgResult = PSNRET_INVALID_NOCHANGEPAGE;
                    break;
                  }
              }
            set_default_key (s);
 
	    SendDlgItemMessage (hDlg, IDC_TIME_PHRASES, WM_GETTEXT,
                                20, (LPARAM)s);		
	    opt.passwd_ttl = (int)atol (s)*60;
		
	    opt.encrypt_default = !!SendDlgItemMessage
              (hDlg, IDC_ENCRYPT_DEFAULT, BM_GETCHECK, 0, 0L);
	    opt.sign_default = !!SendDlgItemMessage 
              (hDlg, IDC_SIGN_DEFAULT, BM_GETCHECK, 0, 0L);
            opt.enable_smime = !!SendDlgItemMessage
              (hDlg, IDC_ENABLE_SMIME, BM_GETCHECK, 0, 0L);
            if (opt.enable_smime)
              {
                MessageBox (hDlg, 
          _("You have enabled GpgOL's support for the S/MIME protocol.\n\n"
            "New S/MIME messages are thus only viewable with GpgOL and "
            "not anymore with Outlook's internal S/MIME support.  Those "
            "message will even be unreadable by Outlook after GpgOL has "
            "been deinstalled.  A tool to mitigate this problem will be "
            "provided when GpgOL arrives at production quality status."),
                            "GpgOL", MB_ICONWARNING|MB_OK);
              }

	    if (openpgp_state)
              opt.default_protocol = PROTOCOL_OPENPGP;
	    else if (smime_state && opt.enable_smime)
              opt.default_protocol = PROTOCOL_SMIME;
            else
              opt.default_protocol = PROTOCOL_UNKNOWN;
            
            opt.preview_decrypt = !!SendDlgItemMessage
              (hDlg, IDC_PREVIEW_DECRYPT, BM_GETCHECK, 0, 0L);
            opt.prefer_html = !!SendDlgItemMessage
              (hDlg, IDC_PREFER_HTML, BM_GETCHECK, 0, 0L);

	    write_options ();
	    bMsgResult = PSNRET_NOERROR;
	    break; }
		
	case PSN_HELP: 
          {
            const char cpynotice[] = "Copyright (C) 2007 g10 Code GmbH";
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
	    bMsgResult = TRUE;
            break; }
          
	default:
	    bMsgResult = FALSE;
	    break;	    
	}
	SetWindowLong (hDlg, DWL_MSGRESULT, bMsgResult);
	break;

    default:
	bMsgResult = FALSE;
	break;		
    }
    return bMsgResult;
}
