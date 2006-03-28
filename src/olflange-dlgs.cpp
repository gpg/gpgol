/* olflange-dlgs.cpp - New dialogs for Outlook.
 *	Copyright (C) 2001 G Data Software AG, http://www.gdata.de
 *	Copyright (C) 2004, 2005, 2006 g10 Code GmbH
 * 
 * This file is part of GPGol.
 * 
 * GPGol is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2 of the License, or (at your option) any later version.
 * 
 * GPGol is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Lesser General Public License for more details.
 * 
 * You should have received a copy of the GNU Lesser General Public
 * License along with this program; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA
 * 02110-1301, USA.
 */

#ifdef HAVE_CONFIG_H
#include <config.h>
#endif

#include <windows.h>
#include <prsht.h>

#include "intern.h"
#include "mymapi.h"
#include "mymapitags.h"
#include "display.h"

#include "olflange-def.h"
#include "olflange-ids.h"


/* GPGOptionsDlgProc -
   Handles the notifications sent for managing the options property page. */
bool 
GPGOptionsDlgProc (HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  BOOL bMsgResult = FALSE;    
  static LPNMHDR pnmhdr;
  static HWND hWndPage;
    
  switch (uMsg) 
    {
    case WM_INITDIALOG:
      {
        HDC hdc = GetDC (hDlg);
        if (hdc)
          {
            int bits_per_pixel = GetDeviceCaps (hdc, BITSPIXEL);
            HBITMAP bitmap;
                
            ReleaseDC (hDlg, hdc);	
            if (bits_per_pixel > 15)
              {
                bitmap = LoadBitmap (glob_hinst, MAKEINTRESOURCE(IDB_BANNER));
                if (bitmap)
                  {
                    HBITMAP old = (HBITMAP) SendDlgItemMessage
                      (hDlg, IDC_BITMAP, STM_SETIMAGE,
                       IMAGE_BITMAP, (LPARAM)bitmap);
                    if (old)
                      DeleteObject (old);	
                  }	
              }		
          }
        
	EnableWindow (GetDlgItem (hDlg, IDC_ENCRYPT_TO),
                      !!opt.enable_default_key);
	if (opt.enable_default_key)
          CheckDlgButton (hDlg, IDC_ENCRYPT_WITH_STANDARD_KEY, BST_CHECKED);
	SetDlgItemText (hDlg, IDC_VERSION_INFO, 
		        "Version "VERSION " ("__DATE__")");
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
	if (HIWORD (wParam) == BN_CLICKED &&
	    LOWORD (wParam) == IDC_ENCRYPT_WITH_STANDARD_KEY) {
	    opt.enable_default_key = !opt.enable_default_key;
	    EnableWindow (GetDlgItem (hDlg, IDC_ENCRYPT_TO), 
			  !!opt.enable_default_key);
	}
	if (LOWORD(wParam) == IDC_GPG_OPTIONS)
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
	    wsprintf(s, "%d", opt.passwd_ttl);
	    SendDlgItemMessage(hDlg, IDC_TIME_PHRASES, WM_SETTEXT,
                               0, (LPARAM) s);
	    hWndPage = pnmhdr->hwndFrom;   // to be used in WM_COMMAND
	    SendDlgItemMessage (hDlg, IDC_ENCRYPT_DEFAULT, BM_SETCHECK, 
				!!opt.encrypt_default, 0L);
	    SendDlgItemMessage (hDlg, IDC_SIGN_DEFAULT, BM_SETCHECK, 
			        !!opt.sign_default, 0L);
	    SendDlgItemMessage (hDlg, IDC_ENCRYPT_WITH_STANDARD_KEY,
                                BM_SETCHECK, opt.enable_default_key, 0L);
	    SendDlgItemMessage (hDlg, IDC_SAVE_DECRYPTED, BM_SETCHECK, 
				!!opt.save_decrypted_attach, 0L);
	    SendDlgItemMessage (hDlg, IDC_SIGN_ATTACHMENTS, BM_SETCHECK,
				!!opt.auto_sign_attach, 0L);
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
                    MessageBox (hDlg,_("The default key may not"
                                       " contain any spaces."),
                                "GPGol", MB_ICONERROR|MB_OK);
                    bMsgResult = PSNRET_INVALID_NOCHANGEPAGE;
                    break;
                  }
              }
            set_default_key (s);
 
	    SendDlgItemMessage (hDlg, IDC_TIME_PHRASES, WM_GETTEXT,
                                20, (LPARAM)s);		
	    opt.passwd_ttl = (int)atol (s);
		
	    opt.encrypt_default = !!SendDlgItemMessage
              (hDlg, IDC_ENCRYPT_DEFAULT, BM_GETCHECK, 0, 0L);
	    opt.sign_default = !!SendDlgItemMessage 
              (hDlg, IDC_SIGN_DEFAULT, BM_GETCHECK, 0, 0L);
	    opt.save_decrypted_attach = !!SendDlgItemMessage
              (hDlg, IDC_SAVE_DECRYPTED, BM_GETCHECK, 0, 0L);
            opt.auto_sign_attach = !!SendDlgItemMessage
              (hDlg, IDC_SIGN_ATTACHMENTS, BM_GETCHECK, 0, 0L);
            opt.preview_decrypt = !!SendDlgItemMessage
              (hDlg, IDC_PREVIEW_DECRYPT, BM_GETCHECK, 0, 0L);
            opt.prefer_html = !!SendDlgItemMessage
              (hDlg, IDC_PREFER_HTML, BM_GETCHECK, 0, 0L);

	    write_options ();
	    bMsgResult = PSNRET_NOERROR;
	    break; }
		
	case PSN_HELP:
	    MessageBox (pnmhdr->hwndFrom,
    "This is GPGol version " PACKAGE_VERSION "\n"
    "Copyright (C) 2005 g10 Code GmbH\n"
    "\n"
    "GPGol is a plugin for Outlook to allow encryption and\n"
    "signing of messages using the OpenPGP standard. It makes\n"
    "use of the GnuPG software (http://www.gnupg.org). Latest\n"
    "release information are accessible by clicking on the logo.\n"
    "\n"
    "GPGol is free software; you can redistribute it and/or\n"
    "modify it under the terms of the GNU Lesser General Public\n"
    "License as published by the Free Software Foundation; either\n"
    "version 2.1 of the License, or (at your option) any later version.\n"
    "\n"
    "GPGol is distributed in the hope that it will be useful,\n"
    "but WITHOUT ANY WARRANTY; without even the implied warranty of\n"
    "MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the\n"
    "GNU Lesser General Public License for more details.\n"
    "\n"
    "You should have received a copy of the GNU Lesser General Public\n"
    "License along with this library; if not, write to the Free Software\n"
    "Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA\n"
    "02110-1301, USA.\n",
                  "GPGol", MB_OK);
	    bMsgResult = TRUE;
	    break;

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
