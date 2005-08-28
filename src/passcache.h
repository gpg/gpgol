/* passcache.h - Interface the passphrase cache for OutlGPG
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

#ifndef PASSCACHE_H
#define PASSCACHE_H


/* Initialize the passcache subsystem. */
int initialize_passcache (void);

/* Flush all entries. */
void passcache_flushall (void);

/* Store and retrieve a cached passphrase. */
void passcache_put (const char *key, const char *value, int ttl);
char *passcache_get (const char *key);

#endif /*PASSCACHE_H*/