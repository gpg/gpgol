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

#ifdef HAVE_CONFIG_H
# include "config.h"
#endif

#include "cpphelp.h"

#include <algorithm>
#include <sstream>
#include <vector>
#include <iterator>

#include "common_indep.h"

#ifdef HAVE_W32_SYSTEM
# include "common.h"
# include <windows.h>
#else
#include "common_indep.h"
#endif

void
release_cArray (char **carray)
{
  if (carray)
    {
      for (int idx = 0; carray[idx]; idx++)
        {
          xfree (carray[idx]);
        }
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

void
remove_whitespace (std::string &s)
{
  s.erase(remove_if(s.begin(), s.end(), isspace), s.end());
}

void
join(const std::vector<std::string>& v, const char *c, std::string& s)
{
  s.clear();
  for (auto p = v.begin(); p != v.end(); ++p)
    {
      s += *p;
      if (p != v.end() - 1)
        {
          s += c;
        }
    }
}

char **
vector_to_cArray(const std::vector<std::string> &vec)
{
  char ** ret = (char**) xmalloc (sizeof (char*) * (vec.size() + 1));
  for (size_t i = 0; i < vec.size(); i++)
    {
      ret[i] = xstrdup (vec[i].c_str());
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

#ifdef HAVE_W32_SYSTEM
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
#endif

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

std::string
string_to_hex(const std::string& input)
{
    static const char* const lut = "0123456789ABCDEF";
    size_t len = input.length();

    std::string output;
    output.reserve (3 * len + (len * 3 / 26));
    for (size_t i = 0; i < len; ++i)
    {
        const unsigned char c = input[i];
        output.push_back (lut[c >> 4]);
        output.push_back (lut[c & 15]);
        output.push_back (' ');
        if (i % 26 == 0)
          {
            output.push_back ('\n');
          }
    }
    return output;
}

bool
is_binary (const std::string &input)
{
  for (int i = 0; i < input.size() - 1; ++i)
    {
      const unsigned char c = input[i];
      if (c < 32 && c != 0x0d && c != 0x0a)
        {
          return true;
        }
    }
  return false;
}

const char *
to_cstr (const GpgME::Protocol &prot)
{
  return prot == GpgME::CMS ? "S/MIME" :
         prot == GpgME::OpenPGP ? "OpenPGP" :
         "Unknown Protocol";
}

void
find_and_replace(std::string& source, const std::string &find,
                 const std::string &replace)
{
  TSTART;
  for(std::string::size_type i = 0; (i = source.find(find, i)) != std::string::npos;)
    {
      source.replace(i, find.length(), replace);
      i += replace.length();
    }
  TRETURN;
}

bool starts_with(const std::string &s, const char *prefix)
{
  if (!prefix || s.empty ())
    {
      return false;
    }
  return s.substr(0, strlen (prefix)) == prefix;
}

bool starts_with(const std::string &s, const char prefix)
{
  return !s.empty() && s.front() == prefix;
}

static std::string
do_asprintf_s (const char *fmt, va_list a)
{
  char *buf;
  if (gpgrt_vasprintf (&buf, fmt, a) < 0)
    {
      STRANGEPOINT;
      return std::string ();
    }
  std::string ret (buf);
  free (buf);
  return ret;
}

std::string
asprintf_s (const char *fmt, ...)
{
  va_list a;
  va_start (a, fmt);
  std::string ret = do_asprintf_s (fmt, a);
  va_end (a);
  return ret;
}
