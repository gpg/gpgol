/* attachment.cpp - Functions for attachment handling
 *    Copyright (C) 2005, 2007 g10 Code GmbH
 *    Copyright (C) 2015 Intevation GmbH
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
#include "common_indep.h"
#include "attachment.h"

#include <climits>

Attachment::Attachment()
{
}

Attachment::~Attachment()
{
  log_debug ("%s:%s", SRCNAME, __func__);
}

std::string
Attachment::get_display_name() const
{
  return m_utf8DisplayName;
}

void
Attachment::set_display_name(const char *name)
{
  m_utf8DisplayName = std::string(name);
}

void
Attachment::set_attach_type(attachtype_t type)
{
  m_type = type;
}

bool
Attachment::isSupported(GpgME::DataProvider::Operation op) const
{
  return op == GpgME::DataProvider::Read ||
         op == GpgME::DataProvider::Write ||
         op == GpgME::DataProvider::Seek ||
         op == GpgME::DataProvider::Release;
}

ssize_t
Attachment::read(void *buffer, size_t bufSize)
{
  return m_data.read (buffer, bufSize);
}

ssize_t
Attachment::write(const void *data, size_t size)
{
  return m_data.write (data, size);
}

off_t Attachment::seek(off_t offset, int whence)
{
  return m_data.seek (offset, whence);
}

void Attachment::release()
{
  /* No op. */
  log_debug ("%s:%s", SRCNAME, __func__);
}
