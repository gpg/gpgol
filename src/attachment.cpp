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

#include <climits>
#include <gpg-error.h>
#include <gpgme++/error.h>

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

Attachment::Attachment(LPSTREAM stream)
{
  if (stream)
    {
      stream->AddRef ();
    }
  m_stream = stream;
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
  if (!bufSize)
    {
      return 0;
    }
  if (!buffer || bufSize >= ULONG_MAX)
    {
      log_error ("%s:%s: Read invalid",
                 SRCNAME, __func__);
      GpgME::Error::setSystemError (GPG_ERR_EINVAL);
      return -1;
    }
  if (!m_stream)
    {
      log_error ("%s:%s: Read on null stream.",
                 SRCNAME, __func__);
      GpgME::Error::setSystemError (GPG_ERR_EIO);
      return -1;
    }

  ULONG cb = static_cast<size_t> (bufSize);
  ULONG bRead = 0;
  HRESULT hr = m_stream->Read (buffer, cb, &bRead);
  if (hr != S_OK && hr != S_FALSE)
    {
      log_error ("%s:%s: Read failed",
                 SRCNAME, __func__);
      GpgME::Error::setSystemError (GPG_ERR_EIO);
      return -1;
    }
  return static_cast<size_t>(bRead);
}

ssize_t
Attachment::write(const void *data, size_t size)
{
  if (!size)
    {
      return 0;
    }
  if (!data || size >= ULONG_MAX)
    {
      GpgME::Error::setSystemError (GPG_ERR_EINVAL);
      return -1;
    }
  if (!m_stream)
    {
      log_error ("%s:%s: Write on NULL stream. ",
                 SRCNAME, __func__);
      GpgME::Error::setSystemError (GPG_ERR_EIO);
      return -1;
    }
  ULONG written = 0;
  if (m_stream->Write (data, static_cast<ULONG>(size), &written) != S_OK)
    {
      GpgME::Error::setSystemError (GPG_ERR_EIO);
      log_error ("%s:%s: Write failed.",
                 SRCNAME, __func__);
      return -1;
    }
  if (m_stream->Commit (0) != S_OK)
    {
      log_error ("%s:%s: Commit failed. ",
                 SRCNAME, __func__);
      GpgME::Error::setSystemError (GPG_ERR_EIO);
      return -1;
    }
  return static_cast<ssize_t> (written);
}

off_t Attachment::seek(off_t offset, int whence)
{
  DWORD dwOrigin;
  switch (whence)
    {
      case SEEK_SET:
          dwOrigin = STREAM_SEEK_SET;
          break;
      case SEEK_CUR:
          dwOrigin = STREAM_SEEK_CUR;
          break;
      case SEEK_END:
          dwOrigin = STREAM_SEEK_END;
          break;
      default:
          GpgME::Error::setSystemError (GPG_ERR_EINVAL);
          return (off_t) - 1;
   }
  if (!m_stream)
    {
      log_error ("%s:%s: Seek on null stream.",
                 SRCNAME, __func__);
      GpgME::Error::setSystemError (GPG_ERR_EIO);
      return (off_t) - 1;
    }
  LARGE_INTEGER move = {0, 0};
  move.QuadPart = offset;
  ULARGE_INTEGER result;
  HRESULT hr = m_stream->Seek (move, dwOrigin, &result);

  if (hr != S_OK)
    {
      log_error ("%s:%s: Seek failed. ",
                 SRCNAME, __func__);
      GpgME::Error::setSystemError (GPG_ERR_EINVAL);
      return (off_t) - 1;
    }
  return result.QuadPart;
}

void Attachment::release()
{
  /* No op. */
  log_debug ("%s:%s", SRCNAME, __func__);
}
