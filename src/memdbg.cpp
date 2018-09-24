/* @file memdbg.cpp
 * @brief Memory debugging helpers
 *
 * Copyright (C) 2018 Intevation GmbH
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

#include "config.h"

#include "memdbg.h"

#include "common_indep.h"

#include <gpg-error.h>

#include <unordered_map>
#include <string>

std::unordered_map <std::string, int> cppObjs;
std::unordered_map <void *, int> olObjs;
std::unordered_map <void *, std::string> olNames;
std::unordered_map <void *, std::string> allocs;

GPGRT_LOCK_DEFINE (memdbg_log);

#define DBGGUARD if (!(opt.enable_debug & DBG_OOM_EXTRA)) return

#ifdef HAVE_W32_SYSTEM
# include "oomhelp.h"
#endif

/* Returns true on a name change */
static bool
register_name (void *obj, const char *nameSuggestion)
{
#ifdef HAVE_W32_SYSTEM

  char *name = get_object_name ((LPUNKNOWN)obj);
  bool suggestionUsed = false;

  if (!name && nameSuggestion)
    {
      name = xstrdup (nameSuggestion);
      suggestionUsed = true;
    }
  if (!name)
    {
      auto it = olNames.find (obj);
      if (it != olNames.end())
        {
          if (it->second != "unknown")
            {
              log_debug ("%s:%s Ptr %p name change from %s to unknown",
                         SRCNAME, __func__, obj, it->second.c_str());
              it->second = "unknown";
              return true;
            }
        }
      return false;
    }

  std::string sName = name;
  xfree (name);

  auto it = olNames.find (obj);
  if (it != olNames.end())
    {
      if (it->second != sName)
        {
          log_debug ("%s:%s Ptr %p name change from %s to %s",
                     SRCNAME, __func__, obj, it->second.c_str(),
                     sName.c_str());
          it->second = sName;
          return !suggestionUsed;
        }
    }
  else
    {
      olNames.insert (std::make_pair (obj, sName));
    }
#else
  (void) obj;
#endif
  return false;
}

void
_memdbg_addRef (void *obj, const char *nameSuggestion)
{
  DBGGUARD;

  if (!obj)
    {
      return;
    }

  gpgrt_lock_lock (&memdbg_log);

  auto it = olObjs.find (obj);

  if (it == olObjs.end())
    {
      it = olObjs.insert (std::make_pair (obj, 0)).first;
    }
  if (register_name (obj, nameSuggestion) && it->second)
    {
      log_error ("%s:%s Name change without null ref on %p!",
                 SRCNAME, __func__, obj);
    }
  it->second++;

  gpgrt_lock_unlock (&memdbg_log);
}

void
memdbg_released (void *obj)
{
  DBGGUARD;

  if (!obj)
    {
      return;
    }

  gpgrt_lock_lock (&memdbg_log);

  auto it = olObjs.find (obj);

  if (it == olObjs.end())
    {
      log_error ("%s:%s Released %p without query if!!",
                 SRCNAME, __func__, obj);
      gpgrt_lock_unlock (&memdbg_log);
      return;
    }

  it->second--;

  if (it->second < 0)
    {
      log_error ("%s:%s Released %p below zero",
                 SRCNAME, __func__, obj);
    }
  gpgrt_lock_unlock (&memdbg_log);
}

void
memdbg_ctor (const char *objName)
{
  DBGGUARD;

  if (!objName)
    {
      TRACEPOINT;
      return;
    }

  gpgrt_lock_lock (&memdbg_log);

  const std::string nameStr (objName);

  auto it = cppObjs.find (nameStr);

  if (it == cppObjs.end())
    {
      it = cppObjs.insert (std::make_pair (nameStr, 0)).first;
    }
  it->second++;

  gpgrt_lock_unlock (&memdbg_log);
}

void
memdbg_dtor (const char *objName)
{
  DBGGUARD;

  if (!objName)
    {
      TRACEPOINT;
      return;
    }

  const std::string nameStr (objName);
  auto it = cppObjs.find (nameStr);

  if (it == cppObjs.end())
    {
      log_error ("%s:%s Dtor of %s before ctor",
                 SRCNAME, __func__, nameStr.c_str());
      gpgrt_lock_unlock (&memdbg_log);
      return;
    }

  it->second--;

  if (it->second < 0)
    {
      log_error ("%s:%s Dtor of %s more often then ctor",
                 SRCNAME, __func__, nameStr.c_str());
    }
  gpgrt_lock_unlock (&memdbg_log);
}


void
_memdbg_alloc (void *ptr, const char *srcname, const char *func, int line)
{
  DBGGUARD;

  if (!ptr)
    {
      TRACEPOINT;
      return;
    }

  gpgrt_lock_lock (&memdbg_log);

  const std::string identifier = std::string (srcname) + std::string (":") +
                                  std::string (func) + std::string (":") +
                                  std::to_string (line);

  auto it = allocs.find (ptr);

  if (it != allocs.end())
    {
      TRACEPOINT;
      gpgrt_lock_unlock (&memdbg_log);
      return;
    }
  allocs.insert (std::make_pair (ptr, identifier));

  gpgrt_lock_unlock (&memdbg_log);
}


int
memdbg_free (void *ptr)
{
  DBGGUARD false;

  if (!ptr)
    {
      TRACEPOINT;
      return false;
    }

  gpgrt_lock_lock (&memdbg_log);

  auto it = allocs.find (ptr);

  if (it == allocs.end())
    {
      log_error ("%s:%s Free unregistered: %p",
                 SRCNAME, __func__, ptr);
      gpgrt_lock_unlock (&memdbg_log);
      return false;
    }

  allocs.erase (it);

  gpgrt_lock_unlock (&memdbg_log);
  return true;
}

void
memdbg_dump ()
{
  DBGGUARD;
  gpgrt_lock_lock (&memdbg_log);
  log_debug(""
"------------------------------MEMORY DUMP----------------------------------");

  log_debug("-- C++ Objects --");
  for (const auto &pair: cppObjs)
    {
      log_debug("%s\t: %i", pair.first.c_str(), pair.second);
    }
  log_debug("-- C++ End --");
  log_debug("-- OL Objects --");
  for (const auto &pair: olObjs)
    {
      if (!pair.second)
        {
          continue;
        }
      const auto it = olNames.find (pair.first);
      if (it == olNames.end())
        {
          log_debug("%p\t: %i", pair.first, pair.second);
        }
      else
        {
          log_debug("%p:%s\t: %i", pair.first,
                    it->second.c_str (), pair.second);
        }
    }
  log_debug("-- OL End --");
  log_debug("-- Allocated Addresses --");
  for (const auto &pair: allocs)
    {
      log_debug ("%s: %p", pair.second.c_str(), pair.first);
    }
  log_debug("-- Allocated Addresses End --");

  log_debug(""
"------------------------------MEMORY END ----------------------------------");
  gpgrt_lock_unlock (&memdbg_log);
}
