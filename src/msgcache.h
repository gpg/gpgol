/* msgcache.h - Interface to the message cache.
 *	Copyright (C) 2005 g10 Code GmbH
 *
 * This file is part of OutlGPG.
 * 
 * OutlGPG is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2 of the License, or (at your option) any later version.
 * 
 * OutlGPG is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Lesser General Public License for more details.
 * 
 * You should have received a copy of the GNU Lesser General Public
 * License along with this program; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA
 * 02110-1301, USA.
 */
#ifndef MSGCACHE_H
#define MSGCACHE_H

#ifdef __cplusplus
extern "C" {
#if 0
}
#endif
#endif

/* Initialize the message cache subsystem. */
int initialize_msgcache (void);

/* Put BODY into tye cace, derive the key from MESSAGE.  TRANSFER
   controls whether the cache will snatch ownership of body. */
void msgcache_put (char *body, int transfer, LPMESSAGE message);

/* Return the plaintext stored under a key derived from MESSAGE or
   NULL if none was found. */
const char *msgcache_get (LPMESSAGE message);


#ifdef __cplusplus
}
#endif
#endif /*MSGCACHE_H*/
