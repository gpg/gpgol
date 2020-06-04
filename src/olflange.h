/* olflange.h - Flange between Outlook and the MapiGPGME class
 * Copyright (C) 2005, 2007 g10 Code GmbH
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

#ifndef OLFLANGE_H
#define OLFLANGE_H

#include "mymapi.h"
#include "mymapitags.h"
#include "mapihelp.h"

/* The GUID for this plugin.  */
#define CLSIDSTR_GPGOL   "{42d30988-1a3a-11da-c687-000d6080e735}"
DEFINE_GUID(CLSID_GPGOL, 0x42d30988, 0x1a3a, 0x11da,
            0xc6, 0x87, 0x00, 0x0d, 0x60, 0x80, 0xe7, 0x35);

/* For documentation: The GUID used for our custom properties:
   {31805ab8-3e92-11dc-879c-00061b031004}
 */

/* The ProgID used by us */
#define GPGOL_PROGID "GNU.GpgOL"
/* User friendly add in name */
#define GPGOL_PRETTY "GpgOL - The GnuPG Outlook Plugin"
/* Short description of the addin */
#define GPGOL_DESCRIPTION "Cryptography for Outlook"

EXTERN_C const char * __stdcall gpgol_check_version (const char *req_version);

EXTERN_C int get_ol_main_version (void);

void install_forms (void);
#endif /*OLFLANGE_H*/
