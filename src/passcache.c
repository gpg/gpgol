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
#include <assert.h>
#include <time.h>
#include <windows.h> /* Argg: Only required for the locking. */

#include "util.h"
#include "passcache.h"


/* An item to hold a cached passphrase. */
struct cache_item
{
  /* We love linked lists; there are only a few passwords and access
     to them is not in any way time critical. */
  struct cache_item *next;  

  /* The Time to Live for this entry. */
  int ttl;

  /* The timestamp is updated with each access to the item and used
     along with TTL to exire this item. */
  time_t timestamp;

  /* The value of this item.  Malloced C String.  If this one is NULL
     this item may be deleted. */
  char *value;         

  /* The key for this item. C String. */
  char key[1];
};
typedef struct cache_item *cache_item_t;


/* The actual cache is a simple list anchord at this global
   variable. */
static cache_item_t the_cache;

/* Mutex used to serialize access to the cache. */
static HANDLE cache_mutex;


/* Initialize this mode.  Called at a very early stage. Returns 0 on
   success. */
int
initialize_passcache (void)
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


/* This is routine is used to remove all deleted entries from the
   linked list.  Deleted entries are marked by a value of NULL.  Note,
   that this routibne must be called in a locked state. */
static void
remove_deleted_items (void)
{
  cache_item_t item, prev;

 again:
  for (item = the_cache; item; item = item->next)
    if (!item->value)
      {
        if (item == the_cache)
          {
            the_cache = item->next;
            xfree (item);
          }
        else
          {
            for (prev=the_cache; prev->next; prev = prev->next)
              if (prev->next == item)
                {
                  prev->next = item->next;
                  xfree (item);
                  item = NULL;
                  break;
                }
            assert (!item);
          }
        goto again; /* Yes, we use this pretty dumb algorithm ;-) */
      }
}



/* Flush all entries from the cache. */
void
passcache_flushall (void)
{
  cache_item_t item;

  if (lock_cache ())
    return; /* FIXME: Should we pop up a message box? */ 

  for (item = the_cache; item; item = item->next)
    if (item->value)
      {
        wipestring (item->value);
        xfree (item->value);
        item->value = NULL;
      }
  remove_deleted_items ();

  unlock_cache ();
}


/* Store the passphrase in VALUE under KEY in out cache. Assign TTL
   seconds as maximum caching time.  If it already exists, merely
   updates the TTL. If the TTL is 0 or VALUE is NULL or empty, flush a
   possible entry. */
void
passcache_put (const char *key, const char *value, int ttl)
{
  cache_item_t item;

  if (!key || !*key)
    {
      log_error ("%s:%s: no key given", __FILE__, __func__);
      return;
    }

  if (lock_cache ())
    return; /* FIXME: Should we pop up a message box if a flush was
               requested? */
  
  for (item = the_cache; item; item = item->next)
    if (item->value && !strcmp (item->key, key))
      break;
  if (item && (!ttl || !value || !*value))
    {
      /* Delete this entry. */
      wipestring (item->value);
      xfree (item->value);
      item->value = NULL;
      /* Actual delete will happen before we allocate a new entry. */
    }
  else if (item)
    {
      /* Update this entry. */
      if (item->value)
        {
          wipestring (item->value);
          xfree (item->value);
        }
      item->value = xstrdup (value);
      item->ttl = ttl;
      item->timestamp = time (NULL);
    }
  else if (!ttl || !value || !*value)
    {
      log_debug ("%s:%s: ignoring attempt to add empty entry `%s'",
                 __FILE__, __func__, key);
    }
  else 
    {
      /* Create new cache entry. */
      remove_deleted_items ();
      item = xcalloc (1, sizeof *item + strlen (key));
      strcpy (item->key, key);
      item->ttl = ttl;
      item->value = xstrdup (value);
      item->timestamp = time (NULL);

      item->next = the_cache;
      the_cache = item;
    }

  unlock_cache ();
}


/* Return the passphrase stored under KEY as a newly malloced string.
   Caller must release that string using xfree.  Using this function
   won't update the TTL.  If no passphrase is available under this
   key, the function returns NULL.  Calling thsi function with KEY set
   to NULL will only expire old entries. */
char *
passcache_get (const char *key)
{
  cache_item_t item;
  char *result = NULL;
  time_t now = time (NULL);

  if (lock_cache ())
    return NULL;
  
  /* Expire entries. */
  for (item = the_cache; item; item = item->next)
    if (item->value && item->timestamp + item->ttl < now)
      {
        wipestring (item->value);
        xfree (item->value);
        item->value = NULL;
      }

  /* Look for the entry. */
  if (key && *key)
    {
      for (item = the_cache; item; item = item->next)
        if (item->value && !strcmp (item->key, key))
          {
            result = xstrdup (item->value);
            item->timestamp = time (NULL);
            break;
          }
    }

  unlock_cache ();

  return result;
}
