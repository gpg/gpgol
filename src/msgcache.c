/* msgcache.cpp - Implementation of a message cache.
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

/* Due to some peculiarities of Outlook 2003 and possible also earlier
   versions we need fixup the text of a reply before editing starts.
   This is done using the Echache extension mechanism and this module
   provides the means of caching and locating messages.  To be exact,
   we don't cache entire messages but just the plaintext after
   decryption.

   What we do here is to save the plaintext in a list under a key
   taken from the PR_STORE_ENTRYID property.  It seems that this is a
   reliable key to match the message again after Reply has been
   called.  We can't use PR_ENTRYID because this one is different in
   the reply template message.

   To keep the memory size at bay we but a limit on the maximum cache
   size; thus depending on the total size of the messages the number
   of open inspectors with decrypted messages which can be matched
   against a reply template is limited. We try to make sure that there
   is at least one message; this makes sure that in the most common
   case the plaintext is always available.  We use a circular buffer
   so that the oldest messages are flushed from the cache first.  I
   don't think that it makes much sense to take the sieze of a message
   into account here.
*/
   

#ifdef HAVE_CONFIG_H
#include <config.h>
#endif

#include <windows.h>
#include <assert.h>
#include <string.h>

#include "mymapi.h"
#include "mymapitags.h"

#include "msgcache.h"
#include "util.h"


/* A Item to hold a cache message, i.e. the plaintext and a key. */
struct cache_item
{
  struct cache_item *next;  

  char *plaintext;  /* The malloced plaintext of the message.  This is
                       assumed to be a valid C String, UTF8
                       encoded. */
  size_t length;    /* The length of that plaintext used to compute
                       the total size of the cache. */

  /* The length of the key and the key itself.  The cache item is
     dynamically allocated to fit the size of the key.  Note, that
     the key is a binary blob. */
  size_t keylen;
  char key[1];
};
typedef struct cache_item *cache_item_t;


/* The actual cache is a simple list anchord at this global
   variable. */
static cache_item_t the_cache;

/* Mutex used to serialize access to the cache. */
static HANDLE cache_mutex;



/* Initialize this module.  Called at a very early stage during DLL
   loading.  Returns 0 on success. */
int
initialize_msgcache (void)
{
  SECURITY_ATTRIBUTES sa;
  
  memset (&sa, 0, sizeof sa);
  sa.bInheritHandle = TRUE;
  sa.lpSecurityDescriptor = NULL;
  sa.nLength = sizeof sa;
  cache_mutex = CreateMutex (&sa, FALSE, NULL);
  return cache_mutex? 0 : -1;
}


/* Put the BODY of a message into the cache.  BODY should be a
   malloced string, UTF8 encoded.  If TRANSFER is given as true, the
   ownership of the malloced memory for BODY is transferred to this
   module.  MESSAGE is the MAPI message object used to retrieve the
   storarge key for the BODY. */
void
msgcache_put (char *body, int transfer, LPMESSAGE message)
{
  cache_item_t item;

  item = xcalloc (1, sizeof *item);
  item->plaintext = xstrdup (body);
  item->length = strlen (body);

  the_cache = item;
}


/* Locate a plaintext stored under a key derived from the MAPI object
   MESSAGE and return it.  Returns NULL if no plaintext is available.
   FIXME: We need to make sure that the cache object is locked until
   it has been processed by the caller - required a
   msgcache_get_unlock fucntion or similar. */
const char *
msgcache_get (LPMESSAGE message)
{
  if (the_cache && the_cache->plaintext)
    {
      return the_cache->plaintext;
    }
  return NULL;
}

