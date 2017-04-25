/* msgcache.c - Implementation of a message cache.
 * Copyright (C) 2005 g10 Code GmbH
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

/* Due to some peculiarities of Outlook 2003 and possible also earlier
   versions we need fixup the text of a reply before editing starts.
   This is done using the Exchange extension mechanism and this module
   provides the means of caching and locating messages.  To be exact,
   we don't cache entire messages but just the plaintext after
   decryption.

   What we do here is to save the plaintext in a list under a key
   taken from the PR_CONVERSATION_INDEX property.  It seems that this
   is a reliable key to match the message again after Reply has been
   called.  The ConversationIndex as available as an OL item in the
   SENDNOTEMESSAGE hook is 6 bytes longer that the one we get from
   MAPI on the orginal message; thus we only compare up to the length
   we have stored when caching the message.  OL2003 using POP3 uses a
   ConversationIndex of 22/28 bytes.  Note that we could also read the
   ConversationIndex from the OL object but it is easier in
   msgcache_put to retrieve it direct from MAPI.

   The obvious problem with the conversation index is that it does
   only work with independent messages and fails badly when reading
   and replying to several mails from one message thread.

   The desitable soultion would be a way to know the rfc822 message-id
   of the orignal message when entering the reply form.  I can't find
   such a datum but it might be there anyway. I am here thinking of a
   MS coder who wants a reply which indicates the message id and the
   date of the message.

   To keep the memory size at bay we but a limit on the maximum cache
   size; thus depending on the total size of the messages the number
   of open inspectors with decrypted messages which can be matched
   against a reply template is limited. We try to make sure that there
   is at least one message; this makes sure that in the most common
   case the plaintext is always available.  We use a circular buffer
   so that the oldest messages are flushed from the cache first.  I
   don't think that it makes much sense to take the size of a message
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
#include "common.h"


/* We limit the size of the cache to this value. The cache might take
   up more space temporary if a new message is larger than this values
   or if several thereads are currently accessing the cache. */
#define MAX_CACHESIZE (512*1024)  /* 512k */


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
               SRCNAME, __func__, code);
  return code != WAIT_OBJECT_0;
}

/* Release the mutex.  No error is returned because this is a fatal
   error anyway and there is no way to clean up. */
static void
unlock_cache (void)
{
  if (!ReleaseMutex (cache_mutex))
    log_error_w32 (-1, "%s:%s: ReleaseMutex failed", SRCNAME, __func__);
}


/* Flush entries from the cache if it gets too large.  NOTE: This
   function needs to be called with a locked cache.  NEWSIZE is the
   size of the message we want to put in the cache later. */
static void
flush_if_needed (size_t newsize)
{
  cache_item_t item, prev;
  size_t total;
  
  for (total = newsize, item = the_cache; item; item = item->next)
    total += item->length;

  if (total <= MAX_CACHESIZE)
    return;

  /* Our algorithm to remove entries is pretty simple: We remove
     entries from the end until we are below the maximum size. */
 again:
  for (item = the_cache, prev = NULL; item; prev = item, item = item->next)
    if ( !item->next && !item->ref)
      {
        if (prev)
          prev->next = NULL;
        else
          the_cache = NULL;
        if (total > item->length)
          total -= item->length;
        else
          total = 0;
        xfree (item->plaintext);
        xfree (item);
        if (total > MAX_CACHESIZE)
          goto again;
        break;
      }
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
  void *refhandle;

  if (!message)
    return; /* No message: Nop. */

  hr = HrGetOneProp ((LPMAPIPROP)message, PR_CONVERSATION_INDEX, &lpspvFEID);
  if (FAILED (hr))
    {
      log_debug ("%s: HrGetOneProp failed: hr=%#lx\n", __func__, hr);
      return;
    }
    
  if ( PROP_TYPE (lpspvFEID->ulPropTag) != PT_BINARY )
    {
      log_debug ("%s: HrGetOneProp returned unexpected property type\n",
                 __func__);
      MAPIFreeBuffer (lpspvFEID);
      return;
    }
  keylen = lpspvFEID->Value.bin.cb;
  key = lpspvFEID->Value.bin.lpb;

  if (!keylen || !key || keylen > 100)
    {
      log_debug ("%s: malformed ConversationIndex\n", __func__);
      MAPIFreeBuffer (lpspvFEID);
      return;
    }

  if (msgcache_get (key, keylen, &refhandle) )
    {
      /* Update an existing cache item. We know that refhandle is
         actually the item, thus we can simply change it here. The
         reference counting makes sure that no other thread will steal
         it while we are updating.  But we still need to lock. */
      log_hexdump (key, keylen, "%s: updating key: ", __func__);
      item = refhandle;
      if (!lock_cache ())
        {
          xfree (item->plaintext);
          item->plaintext = transfer? body : xstrdup (body);
          item->length = strlen (body);
        }
      else /* Oops - locking failed:  Cleanup */
        {
          if (transfer)
            xfree (body);
        }

      msgcache_unref (refhandle);
      MAPIFreeBuffer (lpspvFEID);
    }
  else
    {
      /* Create a new cache entry. */
      item = xmalloc (sizeof *item + keylen - 1);
      item->next = NULL;
      item->ref = 0;
      item->plaintext = transfer? body : xstrdup (body);
      item->length = strlen (body);
      item->keylen = keylen;
      memcpy (item->key, key, keylen);
      log_hexdump (key, keylen, "%s: new cache key: ", __func__);
      
      MAPIFreeBuffer (lpspvFEID);
      
      if (!lock_cache ())
        {
          flush_if_needed (item->length);
          item->next = the_cache;
          the_cache = item;
          unlock_cache ();
        }
    }
}



/* Locate a plaintext stored under KEY of length KEYLEN and return it.
   The user must provide the address of a void pointer variable which
   he later needs to pass to the msgcache_unref. Returns NULL if no
   plaintext is available; msgcache_unref is then not required but
   won't harm either. */
const char *
msgcache_get (const void *key, size_t keylen, void **refhandle)
{
  cache_item_t item;
  const char *result = NULL;

  *refhandle = NULL;

  if (!key || !keylen)
    ; /* No key: Nop. */
  else if (!the_cache)
    ; /* No cache: avoid needless work. */
  else if (!lock_cache ())
    {
      for (item = the_cache; item; item = item->next)
        {
          if (keylen >= item->keylen 
              && !memcmp (key, item->key, item->keylen))
            {
              item->ref++;
              result = item->plaintext; 
              *refhandle = item;
              break;
            }
        }
      unlock_cache ();
    }

  log_hexdump (key, keylen, "%s: cache %s for key: ",
               __func__, result? "hit":"miss");
  return result;
}


/* Locate a plaintext stored for the mapi MESSSAGE and return it.  The
   user must provide the address of a void pointer variable which he
   later needs to pass to the msgcache_unref. Returns NULL if no
   plaintext is available; msgcache_unref is then not required but
   won't harm either. */
const char *
msgcache_get_from_mapi (LPMESSAGE message, void **refhandle)
{
  HRESULT hr;
  LPSPropValue lpspvFEID = NULL;
  const char *result = NULL;

  *refhandle = NULL;

  if (!message)
    return NULL; 

  hr = HrGetOneProp ((LPMAPIPROP)message, PR_CONVERSATION_INDEX, &lpspvFEID);
  if (FAILED (hr))
    {
      log_debug ("%s: HrGetOneProp failed: hr=%#lx\n", __func__, hr);
      return NULL;
    }
    
  if ( PROP_TYPE (lpspvFEID->ulPropTag) != PT_BINARY )
    {
      log_debug ("%s: HrGetOneProp returned unexpected property type\n",
                 __func__);
      MAPIFreeBuffer (lpspvFEID);
      return NULL;
    }
  result = msgcache_get (lpspvFEID->Value.bin.lpb, lpspvFEID->Value.bin.cb,
                         refhandle);
  MAPIFreeBuffer (lpspvFEID);
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
              /* We could here check whether this one has been
                 scheduled for removal.  However, I don't think this
                 is really required, we just wait for the next new
                 message which will then remove all pending ones. */
              break;
            }
        }
      unlock_cache ();
      if (!item)
        log_error ("%s: invalid reference handle %p detected\n",
                   __func__, refhandle);
    }
}
