/* gpgolstr.cpp - String helper class for Outlook API.
 * Copyright (C) 2015 by Bundesamt für Sicherheit in der Informationstechnik
 * Software engineering by Intevation GmbH
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
#include "gpgolstr.h"
#include "common.h"

GpgOLStr::GpgOLStr(const char *str) :
  m_utf8str(NULL), m_widestr(NULL)
{
  if (!str)
    return;
  m_utf8str = strdup (str);
}

GpgOLStr::GpgOLStr(const wchar_t *str) :
  m_utf8str(NULL), m_widestr(NULL)
{
  if (!str)
    return;
  m_widestr = wcsdup (str);
}

GpgOLStr::~GpgOLStr()
{
  xfree (m_utf8str);
  xfree (m_widestr);
}

GpgOLStr::operator char*()
{
  if (!m_utf8str && m_widestr)
    {
      m_utf8str = wchar_to_utf8_2 (m_widestr, wcslen (m_widestr));
    }
  return m_utf8str;
}

GpgOLStr::operator wchar_t*()
{
  if (!m_widestr && m_utf8str)
    {
      m_widestr = utf8_to_wchar2 (m_utf8str, strlen (m_utf8str));
    }
  return m_widestr;
}
