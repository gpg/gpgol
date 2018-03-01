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

#include <gpgme++/context.h>
#include <gpgme++/error.h>
#include <gpgme++/configuration.h>

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

bool
in_de_vs_mode()
{
/* We cache the values only once. A change requires restart.
     This is because checking this is very expensive as gpgconf
     spawns each process to query the settings. */
  static bool checked;
  static bool vs_mode;

  if (checked)
    {
      return vs_mode;
    }
  GpgME::Error err;
  const auto components = GpgME::Configuration::Component::load (err);
  log_debug ("%s:%s: Checking for de-vs mode.",
             SRCNAME, __func__);
  if (err)
    {
      log_error ("%s:%s: Failed to get gpgconf components: %s",
                 SRCNAME, __func__, err.asString ());
      checked = true;
      vs_mode = false;
      return vs_mode;
    }
  for (const auto &component: components)
    {
      if (component.name () && !strcmp (component.name (), "gpg"))
        {
          for (const auto &option: component.options ())
            {
              if (option.name () && !strcmp (option.name (), "compliance") &&
                  option.currentValue ().stringValue () &&
                  !stricmp (option.currentValue ().stringValue (), "de-vs"))
                {
                  log_debug ("%s:%s: Detected de-vs mode",
                             SRCNAME, __func__);
                  checked = true;
                  vs_mode = true;
                  return vs_mode;
                }
            }
          checked = true;
          vs_mode = false;
          return vs_mode;
        }
    }
  checked = true;
  vs_mode = false;
  return false;
}
