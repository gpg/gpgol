/* Copyright (C) 2018 Intevation GmbH
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

#include "dispcache.h"
#include "gpg-error.h"

#include "common.h"
#include "gpgoladdin.h"

#include <unordered_map>

GPGRT_LOCK_DEFINE (cache_lock);

class DispCache::Private
{
public:
  Private()
    {

    }

  void addDisp (int id, LPDISPATCH obj)
   {
     if (!id || !obj)
       {
         TRACEPOINT;
         return;
       }
     gpgrt_lock_lock (&cache_lock);
     auto it = m_cache.find (id);
     if (it != m_cache.end ())
       {
         log_debug ("%s:%s Item \"%i\" already cached. Replacing it.",
                    SRCNAME, __func__, id);
         gpgol_release (it->second);
         it->second = obj;
         gpgrt_lock_unlock (&cache_lock);
         return;
       }
     m_cache.insert (std::make_pair (id, obj));
     gpgrt_lock_unlock (&cache_lock);
     return;
   }

  LPDISPATCH getDisp (int id)
    {
      if (!id)
        {
          TRACEPOINT;
          return nullptr;
        }
      gpgrt_lock_lock (&cache_lock);

      const auto it = m_cache.find (id);
      if (it != m_cache.end())
        {
          LPDISPATCH ret = it->second;
          gpgrt_lock_unlock (&cache_lock);
          return ret;
        }
      gpgrt_lock_unlock (&cache_lock);
      return nullptr;
    }

  ~Private ()
    {
      gpgrt_lock_lock (&cache_lock);
      for (const auto it: m_cache)
        {
          gpgol_release (it.second);
        }
      gpgrt_lock_unlock (&cache_lock);
    }

private:
  std::unordered_map<int, LPDISPATCH> m_cache;
};

DispCache::DispCache (): d (new Private)
{
}

void
DispCache::addDisp (int id, LPDISPATCH obj)
{
  d->addDisp (id, obj);
}

LPDISPATCH
DispCache::getDisp (int id)
{
  return d->getDisp (id);
}

DispCache *
DispCache::instance ()
{
  return GpgolAddin::get_instance ()->get_dispcache ().get ();
}
