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
#include "cpphelp.h"

#include "mimemaker.h"

#include <gpgme++/key.h>

Recipient::Recipient(const char *addr,
                     const char *name, int type):
  m_index (-1)
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
  if (name && !strcmp (name, addr))
    {
      m_name = name;
    }

  TRETURN;
}

Recipient::Recipient(const char *addr, int type):
  Recipient(addr, nullptr, type)
{

}

Recipient::Recipient(const Recipient &other)
{
  m_type = other.type();
  m_mbox = other.mbox();
  m_keys = other.keys();
  m_index = other.index();
  m_name = other.name();
  m_addr = other.addr();
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

void
Recipient::setIndex (int i)
{
  m_index = i;
}

int
Recipient::index () const
{
  return m_index;
}

void
Recipient::dump (const std::vector<Recipient> &recps)
{
  log_data ("--- Begin recipient dump ---");
  if (recps.empty())
    {
      log_data ("Empty recipient list.");
    }
  for (const auto &recp: recps)
    {
      log_data ("Type: %i Mail: '%s'", recp.type (), recp.mbox ().c_str ());
      for (const auto &key: recp.keys ())
        {
          log_data ("Key: %s: %s", to_cstr (key.protocol ()),
                    key.primaryFingerprint ());
        }
      if (recp.keys().empty())
        {
          log_data ("unresolved");
        }
    }
  log_data ("--- End recipient dump ---");
}

std::string
Recipient::encodedDisplayName () const
{
  std::string ret;
  if (m_name.empty())
    {
      char *encoded = utf8_to_rfc2047b (m_addr.c_str ());
      if (encoded)
        {
          ret = encoded;
          xfree (encoded);
        }
      return ret;
    }
  std::string displayName = m_name + std::string (" <") +
                            m_addr + std::string (">");

  char *encDisp = utf8_to_rfc2047b (displayName.c_str ());
  if (encDisp)
    {
      ret = encDisp;
      xfree (encDisp);
    }
  return ret;
}

std::string
Recipient::name () const
{
  return m_name;
}

std::string
Recipient::addr () const
{
  return m_addr;
}
