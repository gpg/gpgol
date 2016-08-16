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
#include "common.h"
#include "attachment.h"
#include "mymapitags.h"
#include "mapihelp.h"
#include "gpgolstr.h"

#include <gpg-error.h>

Attachment::Attachment()
{
  HRESULT hr;
  hr = OpenStreamOnFile (MAPIAllocateBuffer, MAPIFreeBuffer,
                        (SOF_UNIQUEFILENAME | STGM_DELETEONRELEASE
                         | STGM_CREATE | STGM_READWRITE),
                        NULL, GpgOLStr("GPG"), &m_stream);
  if (FAILED (hr))
    {
      log_error ("%s:%s: can't create attachment: hr=%#lx\n",
                 SRCNAME, __func__, hr);
      m_stream = NULL;
    }
}

Attachment::~Attachment()
{
  log_debug ("%s:%s", SRCNAME, __func__);
  gpgol_release (m_stream);
}

LPSTREAM
Attachment::get_stream()
{
  return m_stream;
}

std::string
Attachment::get_display_name() const
{
  return m_utf8DisplayName;
}

std::string
Attachment::get_tmp_file_name() const
{
  return m_utf8FileName;
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

void
Attachment::set_hidden(bool value)
{
  m_hidden = value;
}

int
Attachment::write(const char *data, size_t size)
{
  if (!data || !size)
    {
      return 0;
    }
  if (!m_stream && m_stream->Write (data, size, NULL) != S_OK)
    {
      return 1;
    }
  if (m_stream->Commit (0) != S_OK)
    {
      return 1;
    }
  return 0;
}
