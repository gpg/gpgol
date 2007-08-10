/* mimeparse.h - Multipart MIME parser.
 *	Copyright (C) 2007 g10 Code GmbH
 *
 * This file is part of GpgOL.
 * 
 * GpgOL is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2 of the License, or (at your option) any later version.
 * 
 * GpgOL is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Lesser General Public License for more details.
 * 
 * You should have received a copy of the GNU Lesser General Public
 * License along with this program; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA
 * 02110-1301, USA.
 */
#ifndef SMIME_H
#define SMIME_H

#ifdef __cplusplus
extern "C" {
#if 0
}
#endif
#endif


int mime_verify (const char *message, size_t messagelen, 
                 LPMESSAGE mapi_message, int is_smime,
                 int ttl, 
                 gpgme_data_t attestation, HWND hwnd, int preview_mode);
int mime_decrypt (LPSTREAM instream, LPMESSAGE mapi_message, int is_smime,
                  int ttl,
                  gpgme_data_t attestation, HWND hwnd, int preview_mode);


#ifdef __cplusplus
}
#endif
#endif /*SMIME_H*/
