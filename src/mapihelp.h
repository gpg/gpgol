/* mapihelp.h - Helper functions for MAPI
 *	Copyright (C) 2005, 2007 g10 Code GmbH
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

#ifndef MAPIHELP_H
#define MAPIHELP_H

/* The list of message classes we support.  */
typedef enum 
  {
    MSGCLS_UNKNOWN = 0,
    MSGCLS_GPGSM,
    MSGCLS_GPGSM_MULTIPART_SIGNED
  }
msgclass_t;


void log_mapi_property (LPMESSAGE message, ULONG prop, const char *propname);
int mapi_change_message_class (LPMESSAGE message);
msgclass_t mapi_get_message_class (LPMESSAGE message);
int mapi_to_mime (LPMESSAGE message, const char *filename);



#endif /*MAPIHELP_H*/
