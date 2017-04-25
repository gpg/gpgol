/* gpgolstr.h - String helper class for Outlook API.
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

#include <stddef.h>

/* Small string wrapper that handles wchar / utf8 conversion and
   can be used as a temporary object for an allocated string.
   Modifying the char or wchar_t pointer directly results in
   undefined behavior.
   They are intended to be used for API calls that expect a
   mutable string but are actually a constant.
   */
class GpgOLStr
{
public:
  GpgOLStr() : m_utf8str(NULL), m_widestr(NULL) {}

  GpgOLStr(const char *str);
  GpgOLStr(const wchar_t *widestr);
  ~GpgOLStr();

  operator char* ();
  operator wchar_t* ();

private:
  char *m_utf8str;
  wchar_t *m_widestr;
};
