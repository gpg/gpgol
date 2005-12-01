/* pgpmime.h - PGP/MIME routines for Outlook
 *	Copyright (C) 2005 g10 Code GmbH
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
#ifndef PGPMIME_H
#define PGPMIME_H

#ifdef __cplusplus
extern "C" {
#if 0
}
#endif
#endif

int pgpmime_decrypt (LPSTREAM instream, int ttl, char **body,
                     gpgme_data_t attestation, HWND hwnd,
                     int preview_mode);



#ifdef __cplusplus
}
#endif
#endif /*PGPMIME_H*/
