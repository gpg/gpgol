/* @file recipient.cpp
 * @brief Information about a recipient.
 *
 * Copyright (C) 2020, g10 code GmbH
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
#include "recipient.h"
#include "debug.h"

#include <gpgme++/key.h>

Recipient::Recipient(const char *addr, int type)
{
  TSTART;
  if (addr)
    {
      m_mbox = GpgME::UserID::addrSpecFromString (addr);
    }
  setType (type);
  if (!m_mbox.size ())
    {
      log_error ("%s:%s: Recipient constructed without valid addr",
                 SRCNAME, __func__);
      m_type = invalidType;
    }
  TRETURN;
}

Recipient::Recipient() : m_type (invalidType)
{
}

void
Recipient::setType (int type)
{
  if (type > olBCC || type < olOriginator)
    {
      log_error ("%s:%s: Invalid recipient type %i",
                 SRCNAME, __func__, type);
      m_type = invalidType;
    }
  m_type = static_cast<recipientType> (type);
}

void
Recipient::setKeys (const std::vector<GpgME::Key> &keys)
{
  m_keys = keys;
}

std::string
Recipient::mbox () const
{
  return m_mbox;
}

Recipient::recipientType
Recipient::type () const
{
  return m_type;
}

std::vector<GpgME::Key>
Recipient::keys () const
{
  return m_keys;
}
