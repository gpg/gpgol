/* GPGExchange.h - global declarations
 * Copyright (C) 2001 G Data Software AG, http://www.gdata.de
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

#ifndef INC_GPGEXCHANGE_H
#define INC_GPGEXCHANGE_H

class CGPGExchExtMessageEvents;
class CGPGExchExtCommands;
class CGPGExchExtPropertySheets;
class CGDGPGWrapper;
class CGPG;
class CGPGExchApp;

extern CGPGExchApp theApp;

extern CGPG g_gpg;

BOOL CALLBACK GPGOptionsDlgProc (HWND hDlg, UINT uMsg,
				 WPARAM wParam, LPARAM lParam);

#endif // INC_GPGEXCHANGE_H