/* olflange-def.h 
 * Copyright (C) 2005 g10 Code GmbH
 *
 * This file is part of GpgOL.
 *
 * GpgOL is free software; you can redistribute it and/or modify it
 * under the terms of the GNU Lesser General Public License as
 * published by the Free Software Foundation; either version 2.1 of
 * the License, or (at your option) any later version.
 *  
 * GpgOL is distributed in the hope that it will be useful, but
 * WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
 * General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public License
 * along with this program; if not, see <http://www.gnu.org/licenses/>.
 */

#ifndef OLFLANGE_DEF_H
#define OLFLANGE_DEF_H

class GpgolExtCommands;
class GpgolUserEvents;
class GpgolSessionEvents;
class GpgolMessageEvents;
class GpgolAttachedFileEvents;
class GpgolPropertySheets;
class GpgolItemEvents;

bool GPGOptionsDlgProc (HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam);

/* Retrieve the OOM object from the EECB.  */
LPDISPATCH get_eecb_object (LPEXCHEXTCALLBACK eecb);


#endif /*OLFLANGE_DEF_H*/

