/* mimeparser.h - MIME parser.
 *	Copyright (C) 2007 g10 Code GmbH
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

#ifndef MIMEPARSER_H
#define MIMEPARSER_H

#ifdef __cplusplus
extern "C" {
#if 0
}
#endif
#endif


int mime_verify (protocol_t protocol, const char *message, size_t messagelen, 
                 LPMESSAGE mapi_message, 
                 HWND hwnd, int preview_mode);
int mime_decrypt (protocol_t protocol, 
                  LPSTREAM instream, LPMESSAGE mapi_message,
                  HWND hwnd, int preview_mode);


#ifdef __cplusplus
}
#endif
#endif /*MIMEPARSER_H*/
