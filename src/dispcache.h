/* @file dispcache.h
 * @brief Cache for IDispatch objects that are reusable.
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
#ifndef DISPCACHE_H
#define DISPCACHE_H

#include "config.h"

#include <memory>

#include "oomhelp.h"

class GpgolAddin;

class DispCache
{
    friend class GpgolAddin;
protected:
    DispCache ();

public:
    /* Accessor. Returns the instance carried by
       gpgoladdin. */
    static DispCache *instance ();

    /* Add a IDispatch with the id id to the cache.

       The IDispatch is released on destruction of
       the cache.

       Id's are meant to be defined in dialogs.h
    */
    void addDisp (int id, LPDISPATCH obj);

    /* Get the according IDispatch object. */
    LPDISPATCH getDisp (int id);

private:
    class Private;
    std::shared_ptr<Private> d;
};

#endif
