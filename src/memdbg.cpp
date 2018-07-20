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

#include <map>
#include <string>

std::map <std::string, int> cppObjs;
std::map <void *, int> olObjs;
std::map <void *, std::string> olNames;

GPGRT_LOCK_DEFINE (memdbg_log);

#define DBGGUARD if (!(opt.enable_debug & DBG_OOM_EXTRA)) return

#ifdef HAVE_W32_SYSTEM
# include "oomhelp.h"
#endif

/* Returns true on a name change */
static bool
register_name (void *obj)
{
#ifdef HAVE_W32_SYSTEM

  char *name = get_object_name ((LPUNKNOWN)obj);

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
          return true;
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
memdbg_addRef (void *obj)
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
  if (register_name (obj) && it->second)
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
  (void)objName;
}

void
memdbg_dtor (const char *objName)
{
  (void)objName;

}

void
memdbg_dump ()
{
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

  log_debug(""
"------------------------------MEMORY END ----------------------------------");
  gpgrt_lock_unlock (&memdbg_log);
}