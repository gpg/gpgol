/* revert.h - Declarations for revert.cpp.
 *	Copyright (C) 2008 g10 Code GmbH
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

#ifndef REVERT_H
#define REVERT_H

EXTERN_C LONG __stdcall gpgol_message_revert (LPMESSAGE message, 
                                              LONG do_save,
                                              ULONG save_flags);

EXTERN_C LONG __stdcall gpgol_mailitem_revert (LPDISPATCH mailitem);

EXTERN_C LONG __stdcall gpgol_folder_revert (LPDISPATCH mapifolderobj);


#endif /*REVERT_H*/
