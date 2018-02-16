/* @file cpphelp.h
 * @brief Common cpp helper stuff
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

#include <algorithm>
#include "cpphelp.h"

#include "common.h"

void
release_cArray (char **carray)
{
  if (carray)
    {
      for (int idx = 0; carray[idx]; idx++)
        xfree (carray[idx]);
      xfree (carray);
    }
}

void
rtrim(std::string &s) {
    s.erase(std::find_if(s.rbegin(), s.rend(), [](int ch) {
        return !std::isspace(ch);
    }).base(), s.end());
}

char **
vector_to_cArray(const std::vector<std::string> &vec)
{
  char ** ret = (char**) xmalloc (sizeof (char*) * (vec.size() + 1));
  for (size_t i = 0; i < vec.size(); i++)
    {
      ret[i] = strdup (vec[i].c_str());
    }
  ret[vec.size()] = NULL;
  return ret;
}
