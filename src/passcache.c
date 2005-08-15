/* passcache.c - passphrase cache for OutlGPG
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

/* We use a global passphrase cache.  The cache time is set at the
   time the passphrase gets stored.  */

#include <config.h>
#include <stdlib.h>
#include <string.h>
#include <windows.h> /* Argg: Only required for the locking. */

#include "util.h"
#include "passcache.h"


/* Mutex used to serialize access to the cache. */
static HANDLE cache_mutex;


/* Initialize this mode.  Called at a very early stage. Returns 0 on
   success. */
int
initialize_passcache (void)
{
  SECURITY_ATTRIBUTES sa;
  
  memset (&sa, 0, sizeof sa);
  sa.bInheritHandle = TRUE;
  sa.lpSecurityDescriptor = NULL;
  sa.nLength = sizeof sa;
  cache_mutex = CreateMutex (&sa, FALSE, NULL);
  return cache_mutex? 0 : -1;
}

/* Flush all entries from the cache. */
void
passcache_flushall (void)
{

}


/* Clear the passphrase stored under KEY.  Do nothing if there is no
   cached passphrase under this key. */
void 
passcache_flush (const char *key)
{


}



/* Store the passphrase in VALUE under KEY in out cache. Assign TTL
   seconds as maximum caching time.  If it already exists, merely
   updates the TTL.  */
void
passcache_put (const char *key, const char *value, int ttl)
{
  /* IF the TTL is 0 or value is NULL or empry, flush a possible
     entry. */

}


/* Return the passphrase stored under KEY as a newly malloced string.
   Caller must release that string using xfree.  Using this function
   won't update the TTL.  If no passphrase is available under this
   key, the function returns NULL. */
char *
passcache_get (const char *key)
{
  return NULL;
}
