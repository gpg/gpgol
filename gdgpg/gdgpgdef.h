/* GDGPGDEF.h - general definitions
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

#ifndef __GDGPGDEF_H_
#define __GDGPGDEF_H_

#define GDGPG_SUCCESS                    0
#define GDGPG_ERR_CANCEL                 1
#define GDGPG_ERR_GPG_FAILED             2
#define GDGPG_ERR_NO_DEST_FILE           3
#define GDGPG_ERR_NOT_CRYPTED_OR_SIGNED  4
#define GDGPG_ERR_CHECK_SIGNATURE        5
#define GPGPG_ERR_NO_STANDARD_KEY        6
#define GDGPG_ERR_KEY_MANAGER_NOT_EXISTS 7
#define GDGPG_ERR_NO_VALID_DECRYPT_KEY   8
#define GDGPG_ERR_NO_RECIPIENTS		 9
#define GDGPG_ERR_INV_RECIPIENTS	10

#endif //__GDGPGDEF_H_