/* GPGOptionsDlg.cpp - gpg options dialog procedure
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
#include "GPGExchange.h"
#include "resource.h"
#include "gpg.h"


/* GPGOptionsDlgProc
 Handles the notifications sent for managing the options property page. */
BOOL CALLBACK GPGOptionsDlgProc (
	HWND hDlg,      // The handle to modeless dialog (the property page).
	UINT uMsg,      // The message.
	WPARAM wParam,  // wParam of wndproc
	LPARAM lParam)  // lParam of wndproc
{
    USES_CONVERSION;
    BOOL bMsgResult;
    static HBRUSH hBrush;
    static COLORREF GrayColor;
    static LPNMHDR pnmhdr;
    static HWND hWndPage;
    
    switch (uMsg)
    {
    case WM_INITDIALOG:
	{
	    LOGBRUSH lb;
	    GrayColor = (COLORREF)GetSysColor (COLOR_BTNFACE);
	    memset(&lb, 0, sizeof(LOGBRUSH));
	    lb.lbStyle = BS_SOLID;
	    lb.lbColor = GrayColor;
	    hBrush = CreateBrushIndirect (&lb);

	    int nBitsPerPixel = 0;
	    HDC hdc = GetDC (hDlg);
	    if (hdc != NULL)
	    {
		nBitsPerPixel = GetDeviceCaps (hdc, BITSPIXEL);
		ReleaseDC (hDlg, hdc);	
	    }
	    
	    if (nBitsPerPixel > 15)
	    {
		HBITMAP hbmp;
		hbmp = LoadBitmap (((CWinApp*) &theApp)->m_hInstance, MAKEINTRESOURCE(IDB_BANNER_HI));
		if (hbmp != NULL)
		{
		    HBITMAP hbmpOld;
		    hbmpOld = (HBITMAP) SendDlgItemMessage (hDlg, IDC_BITMAP, 
							    STM_SETIMAGE, IMAGE_BITMAP, 
							    (LPARAM) hbmp);
		    if (hbmpOld != NULL)
			DeleteObject (hbmpOld);	
		}	
	    }
	    SetDlgItemText (hDlg, IDC_VERSION_INFO, "Version "VERSION " ("__DATE__")");
	    return TRUE;
	
	}
	break;

	
    case WM_CTLCOLORDLG:
    case WM_CTLCOLORBTN:
    case WM_CTLCOLORSTATIC:
	if (hBrush != NULL)
	{
	    SetBkColor ((HDC)wParam, GrayColor);
	    return (BOOL)hBrush;
	}
	break;

		
    case WM_DESTROY:
	if (hBrush != NULL)
	    DeleteObject (hBrush);
	return TRUE;
  
		
    case WM_COMMAND:
	if (LOWORD(wParam) == IDC_GPG_OPTIONS)
	{
	    g_gpg.EditExtendedOptions (hDlg);
	    g_gpg.InvalidateKeyLists ();
	}
	break;

		
    case WM_LBUTTONDOWN:
	{
	    int xClick = LOWORD(lParam);
	    int yClick = HIWORD(lParam);
	    RECT rectBanner = {0,0,0,0};
	    RECT rectDlg = {0,0,0,0};
	    GetWindowRect(hDlg, &rectDlg);
	    HWND hwndBitmap = GetDlgItem(hDlg, IDC_BITMAP);
	    if (hwndBitmap != NULL)
		GetWindowRect(hwndBitmap, &rectBanner);

	    rectBanner.left -= rectDlg.left;
	    rectBanner.right -= rectDlg.left;
	    rectBanner.top -= rectDlg.top;
	    rectBanner.bottom -= rectDlg.top;

	    if ((xClick >= rectBanner.left) && (xClick <= rectBanner.right) &&
		(yClick >= rectBanner.top) && (yClick <= rectBanner.bottom))	
	    {
		// open link
		ShellExecute(NULL,"open","http://www.gdata.de/gpg",NULL,NULL,SW_SHOWNORMAL);	
	    }
	}
	break;

		
    case WM_NOTIFY:
	    pnmhdr = ((LPNMHDR) lParam);			
	    switch ( pnmhdr->code)
	    {
	    case PSN_KILLACTIVE:
		bMsgResult = FALSE;  // allow this page to receive PSN_APPLY
		break;

	    case PSN_SETACTIVE:
		{
		    TCHAR s[20];
		    const char * f;
		    wsprintf(s, "%d", g_gpg.GetStorePassPhraseTime());
		    SendDlgItemMessage(hDlg, IDC_TIME_PHRASES, WM_SETTEXT, 0, (LPARAM) s);
		    f = g_gpg.GetLogFile ();
		    SendDlgItemMessage (hDlg, IDC_DEBUG_LOGFILE, WM_SETTEXT, 0, (LPARAM)f);
		    hWndPage = pnmhdr->hwndFrom;   /* to be used in WM_COMMAND */
		    SendDlgItemMessage (hDlg, IDC_ENCRYPT_DEFAULT, BM_SETCHECK, 
					g_gpg.GetEncryptDefault () ? 1 : 0, 0L);
		    SendDlgItemMessage (hDlg, IDC_SIGN_DEFAULT, BM_SETCHECK, 
				        g_gpg.GetSignDefault () ? 1 : 0, 0L);
		    SendDlgItemMessage (hDlg, IDC_ENCRYPT_WITH_STANDARD_KEY, BM_SETCHECK, 
				        g_gpg.GetEncryptWithStandardKey () ? 1 : 0, 0L);
		    SendDlgItemMessage (hDlg, IDC_SAVE_DECRYPTED, BM_SETCHECK, 
					g_gpg.GetSaveDecrypted() ? 1 : 0, 0L);
		    bMsgResult = FALSE;  /* accepts activation */
		    break;
		
		}
		
	    case PSN_APPLY:	
		{
		    TCHAR s[200];
		    SendDlgItemMessage(hDlg, IDC_TIME_PHRASES, WM_GETTEXT, 20, (LPARAM) s);		
		    g_gpg.SetStorePassPhraseTime(atol(s));
		    SendDlgItemMessage (hDlg, IDC_DEBUG_LOGFILE, WM_GETTEXT, 200, (LPARAM)s);
		    g_gpg.SetLogFile (s);
		
		    g_gpg.SetEncryptDefault (SendDlgItemMessage(hDlg, IDC_ENCRYPT_DEFAULT, BM_GETCHECK, 0, 0L));		
		    g_gpg.SetSignDefault (SendDlgItemMessage(hDlg, IDC_SIGN_DEFAULT, BM_GETCHECK, 0, 0L));		
		    g_gpg.SetEncryptWithStandardKey (SendDlgItemMessage(hDlg, IDC_ENCRYPT_WITH_STANDARD_KEY, BM_GETCHECK, 0, 0L));
		    g_gpg.SetSaveDecrypted (SendDlgItemMessage(hDlg, IDC_SAVE_DECRYPTED, BM_GETCHECK, 0, 0L));
		    g_gpg.WriteGPGOptions ();
		    bMsgResult = PSNRET_NOERROR;
		    break;
		}
		
	    case PSN_HELP:                                              		
		MessageBox(pnmhdr->hwndFrom,
		    "For more information, please visit: http://www.gdata.de/gpg\n"
		    "Later versions of the plugin were created by g10 Code Gmbh (http://www.g10code.com)\n",
		    "G DATA GnuPG-Plugin", MB_OK);
		bMsgResult = TRUE;
		break;

	    default:
		bMsgResult = FALSE;
		break;
	    }
	    SetWindowLong( hDlg, DWL_MSGRESULT, bMsgResult);
	    break;

    default:
	bMsgResult = FALSE;
	break;		
    }
    return bMsgResult;
}
