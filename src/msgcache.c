/* msgcache.cpp - Implementation of a message cache.
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

   FUBOH: This damned Outlook think is fucked up beyond all hacking.
   After implementing this cace module here I realized that none of
   the properties are taken from the message to the reply template:
   Neither STORE_ENTRYID (which BTW is anyway constant within a
   folder), nor ENTRYID, nor SEARCHID nor RECORD_KEY.  The only way to
   solve this is by assuming that the last mail read will be the one
   which gets replied to.  To implement this, we mark a message as the
   active one through msgcache_set_active() called from the OnRead
   hook.  msgcache_get will then simply return the active message.
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

  int ref;          /* Reference counter, indicating how many callers
                       are currently accessing the plaintext of this
                       object. */

  int is_active;    /* Flag, see comment at top. */

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
  sa.bInheritHandle = FALSE;
  sa.lpSecurityDescriptor = NULL;
  sa.nLength = sizeof sa;
  cache_mutex = CreateMutex (&sa, FALSE, NULL);
  return cache_mutex? 0 : -1;
}


/* Acquire the mutex.  Returns 0 on success. */
static int 
lock_cache (void)
{
  int code = WaitForSingleObject (cache_mutex, INFINITE);
  if (code != WAIT_OBJECT_0)
    log_error ("%s:%s: waiting on mutex failed: code=%#x\n",
               __FILE__, __func__, code);
  return code != WAIT_OBJECT_0;
}

/* Release the mutex.  No error is returned because this is a fatal
   error anyway and there is no way to clean up. */
static void
unlock_cache (void)
{
  if (!ReleaseMutex (cache_mutex))
    log_error_w32 (-1, "%s:%s: ReleaseMutex failed", __FILE__, __func__);
}



/* Put the BODY of a message into the cache.  BODY should be a
   malloced string, UTF8 encoded.  If TRANSFER is given as true, the
   ownership of the malloced memory for BODY is transferred to this
   module.  MESSAGE is the MAPI message object used to retrieve the
   storarge key for the BODY. */
void
msgcache_put (char *body, int transfer, LPMESSAGE message)
{
  HRESULT hr;
  LPSPropValue lpspvFEID = NULL;
  cache_item_t item;
  size_t keylen;
  void *key;

  if (!message)
    return; /* No message: Nop. */

  hr = HrGetOneProp ((LPMAPIPROP)message, PR_SEARCH_KEY, &lpspvFEID);
  if (FAILED (hr))
    {
      log_error ("%s: HrGetOneProp failed: hr=%#lx\n", __func__, hr);
      return;
    }
    
  if ( PROP_TYPE (lpspvFEID->ulPropTag) != PT_BINARY )
    {
      log_error ("%s: HrGetOneProp returned unexpected property type\n",
                 __func__);
      MAPIFreeBuffer (lpspvFEID);
      return;
    }
  keylen = lpspvFEID->Value.bin.cb;
  key = lpspvFEID->Value.bin.lpb;

  if (!keylen || !key || keylen > 10000)
    {
      log_error ("%s: malformed PR_SEARCH_KEY\n", __func__);
      MAPIFreeBuffer (lpspvFEID);
      return;
    }

  item = xmalloc (sizeof *item + keylen - 1);
  item->next = NULL;
  item->ref = 0;
  item->is_active = 0;
  item->plaintext = transfer? body : xstrdup (body);
  item->length = strlen (body);
  item->keylen = keylen;
  memcpy (item->key, key, keylen);

  MAPIFreeBuffer (lpspvFEID);

  if (!lock_cache ())
    {
      /* FIXME: Decide whether to kick out some entries. */
      item->next = the_cache;
      the_cache = item;
      unlock_cache ();
    }
  msgcache_set_active (message);
}


/* If MESSAGE is in our cse set it's active flag and reset the active
   flag of all others.  */
void
msgcache_set_active (LPMESSAGE message)
{
  HRESULT hr;
  LPSPropValue lpspvFEID = NULL;
  cache_item_t item;
  size_t keylen = 0;
  void *key = NULL;
  int okay = 0;

  if (!message)
    return; /* No message: Nop. */
  if (!the_cache)
    return; /* No cache: avoid needless work. */
  
  hr = HrGetOneProp ((LPMAPIPROP)message, PR_SEARCH_KEY, &lpspvFEID);
  if (FAILED (hr))
    {
      log_error ("%s: HrGetOneProp failed: hr=%#lx\n", __func__, hr);
      goto leave;
    }
    
  if ( PROP_TYPE (lpspvFEID->ulPropTag) != PT_BINARY )
    {
      log_error ("%s: HrGetOneProp returned unexpected property type\n",
                 __func__);
      goto leave;
    }
  keylen = lpspvFEID->Value.bin.cb;
  key = lpspvFEID->Value.bin.lpb;

  if (!keylen || !key || keylen > 10000)
    {
      log_error ("%s: malformed PR_SEARCH_KEY\n", __func__);
      goto leave;
    }
  okay = 1;


 leave:
  if (!lock_cache ())
    {
      for (item = the_cache; item; item = item->next)
        item->is_active = 0;
      if (okay)
        {
          for (item = the_cache; item; item = item->next)
            {
              if (item->keylen == keylen 
                  && !memcmp (item->key, key, keylen))
                {
                  item->is_active = 1;
                  break;
                }
            }
        }
      unlock_cache ();
    }

  MAPIFreeBuffer (lpspvFEID);
}


/* Locate a plaintext stored under a key derived from the MAPI object
   MESSAGE and return it.  The user must provide the address of a void
   pointer variable which he later needs to pass to the
   msgcache_unref. Returns NULL if no plaintext is available;
   msgcache_unref is then not required but won't harm either. */
const char *
msgcache_get (LPMESSAGE message, void **refhandle)
{
  cache_item_t item;
  const char *result = NULL;

  *refhandle = NULL;

  if (!message)
    return NULL; /* No message: Nop. */
  if (!the_cache)
    return NULL; /* No cache: avoid needless work. */
  
#if 0 /* Old code, see comment at top of file. */
  HRESULT hr;
  LPSPropValue lpspvFEID = NULL;
  size_t keylen;
  void *key;

  hr = HrGetOneProp ((LPMAPIPROP)message, PR_SEARCH_KEY, &lpspvFEID);
  if (FAILED (hr))
    {
      log_error ("%s: HrGetOneProp failed: hr=%#lx\n", __func__, hr);
      return NULL;
    }
    
  if ( PROP_TYPE (lpspvFEID->ulPropTag) != PT_BINARY )
    {
      log_error ("%s: HrGetOneProp returned unexpected property type\n",
                 __func__);
      MAPIFreeBuffer (lpspvFEID);
      return NULL;
    }
  keylen = lpspvFEID->Value.bin.cb;
  key = lpspvFEID->Value.bin.lpb;

  if (!keylen || !key || keylen > 10000)
    {
      log_error ("%s: malformed PR_SEARCH_KEY\n", __func__);
      MAPIFreeBuffer (lpspvFEID);
      return NULL;
    }


  if (!lock_cache ())
    {
      for (item = the_cache; item; item = item->next)
        {
          if (item->keylen == keylen 
              && !memcmp (item->key, key, keylen))
            {
              item->ref++;
              result = item->plaintext; 
              *refhandle = item;
              break;
            }
        }
      unlock_cache ();
    }

  MAPIFreeBuffer (lpspvFEID);
#else /* New code. */
  if (!lock_cache ())
    {
      for (item = the_cache; item; item = item->next)
        {
          if (item->is_active)
            {
              item->ref++;
              result = item->plaintext; 
              *refhandle = item;
              break;
            }
        }
      unlock_cache ();
    }

#endif /* New code. */

  return result;
}


/* Release access to a value returned by msgcache_get.  REFHANDLE is
   the value as stored in the pointer variable by msgcache_get. */
void
msgcache_unref (void *refhandle)
{
  cache_item_t item;

  if (!refhandle)
    return;

  if (!lock_cache ())
    {
      for (item = the_cache; item; item = item->next)
        {
          if (item == refhandle)
            {
              if (item->ref < 1)
                log_error ("%s: zero reference count for item %p\n",
                           __func__, item);
              else
                item->ref--;
              /* Fixme: check whether this one has been scheduled for
                 removal. */
              break;
            }
        }
      unlock_cache ();
      if (!item)
        log_error ("%s: invalid reference handle %p detected\n",
                   __func__, refhandle);
    }
}
