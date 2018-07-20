#ifndef MEMDBG_H
#define MEMDBG_H

/* @file memdbg.h
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

#ifdef __cplusplus
extern "C" {
#if 0
}
#endif
#endif

#define memdbg_addRef(X) \
{ \
  if (X) \
    { \
      log_oom_extra ("%s:%s:%i AddRef on %p", \
                     SRCNAME, __func__, __LINE__, X); \
      _memdbg_addRef (X, __func__); \
    } \
}

void _memdbg_addRef (void *obj, const char *nameSuggestion);
void memdbg_released (void *obj);

void memdbg_ctor (const char *objName);
void memdbg_dtor (const char *objName);

void memdbg_dump(void);

#ifdef __cplusplus
}
#endif

#endif //MEMDBG_H
