/* attachment.cpp - Functions for attachment handling
 * Copyright (C) 2005, 2007 g10 Code GmbH
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

#include "config.h"
#include "common_indep.h"
#include "attachment.h"

#include <climits>

Attachment::Attachment()
{
  memdbg_ctor ("Attachment");
}

Attachment::~Attachment()
{
  memdbg_dtor ("Attachment");
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
  if (!name)
    {
      log_error ("%s:%s: Display name set to null.",
                 SRCNAME, __func__);
      return;
    }
  m_utf8DisplayName = std::string(name);
}

void
Attachment::set_attach_type(attachtype_t type)
{
  m_type = type;
}

GpgME::Data &
Attachment::get_data()
{
  return m_data;
}

void
Attachment::set_content_id(const char *cid)
{
  m_cid = cid;
}

std::string
Attachment::get_content_id() const
{
  return m_cid;
}

void
Attachment::set_content_type(const char *ctype)
{
  m_ctype = ctype;
}

std::string
Attachment::get_content_type() const
{
  return m_ctype;
}
