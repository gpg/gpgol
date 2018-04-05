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

#include "cpphelp.h"

#include <algorithm>
#include <sstream>
#include <vector>
#include <iterator>

#include "common.h"

#include <gpgme++/context.h>
#include <gpgme++/error.h>
#include <gpgme++/configuration.h>

#include <windows.h>

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
rtrim(std::string &s)
{
  s.erase (std::find_if (s.rbegin(), s.rend(), [] (int ch) {
      return !std::isspace(ch);
  }).base(), s.end());
}

void
ltrim(std::string &s)
{
  s.erase (s.begin(), std::find_if (s.begin(), s.end(), [] (int ch) {
      return !std::isspace(ch);
  }));
}

void
trim(std::string &s)
{
  ltrim (s);
  rtrim (s);
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

std::vector <std::string>
cArray_to_vector(const char **cArray)
{
  std::vector<std::string> ret;

  if (!cArray)
    {
      return ret;
    }

  for (int i = 0; cArray[i]; i++)
    {
      ret.push_back (std::string (cArray[i]));
    }
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

std::map<std::string, std::string>
get_registry_subkeys (const char *path)
{
  HKEY theKey;
  std::map<std::string, std::string> ret;

  std::string regPath = GPGOL_REGPATH;
  regPath += "\\";
  regPath += path;

  if (RegOpenKeyEx (HKEY_CURRENT_USER,
                    regPath.c_str (),
                    0, KEY_ENUMERATE_SUB_KEYS | KEY_READ,
                    &theKey) != ERROR_SUCCESS)
    {
      TRACEPOINT;
      return ret;
    }

  DWORD values = 0,
        maxValueName = 0,
        maxValueLen = 0;

  DWORD err = RegQueryInfoKey (theKey,
                               nullptr,
                               nullptr,
                               nullptr,
                               nullptr,
                               nullptr,
                               nullptr,
                               &values,
                               &maxValueName,
                               &maxValueLen,
                               nullptr,
                               nullptr);

  if (err != ERROR_SUCCESS)
    {
      TRACEPOINT;
      RegCloseKey (theKey);
      return ret;
    }

  /* Add space for NULL */
  maxValueName++;
  maxValueLen++;

  char name[maxValueName + 1];
  char value[maxValueLen + 1];
  for (int i = 0; i < values; i++)
    {
      DWORD nameLen = maxValueName;
      err = RegEnumValue (theKey, i,
                          name,
                          &nameLen,
                          nullptr,
                          nullptr,
                          nullptr,
                          nullptr);

      if (err != ERROR_SUCCESS)
        {
          TRACEPOINT;
          continue;
        }

      DWORD type;
      DWORD valueLen = maxValueLen;
      err = RegQueryValueEx (theKey, name,
                             NULL, &type,
                             (BYTE*)value, &valueLen);

      if (err != ERROR_SUCCESS)
        {
          TRACEPOINT;
          continue;
        }
      if (type != REG_SZ)
        {
          TRACEPOINT;
          continue;
        }
      ret.insert (std::make_pair (std::string (name, nameLen),
                                  std::string (value, valueLen)));
    }
  RegCloseKey (theKey);
  return ret;
}

template<typename Out> void
internal_split (const std::string &s, char delim, Out result) {
  std::stringstream ss(s);
  std::string item;
  while (std::getline (ss, item, delim))
    {
      *(result++) = item;
    }
}

std::vector<std::string>
gpgol_split (const std::string &s, char delim)
{
  std::vector<std::string> elems;
  internal_split (s, delim, std::back_inserter (elems));
  return elems;
}
