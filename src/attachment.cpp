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
#include "mymapitags.h"
#include "mapihelp.h"
#include "mimedataprovider.h"

#include <climits>

#define COPYBUFSIZE (8 * 1024)

Attachment::Attachment()
{
  memdbg_ctor ("Attachment");
}

#ifndef BUILD_TESTS
Attachment::Attachment(LPDISPATCH attach)
{
  memdbg_ctor ("Attachment");
  if (!attach)
    {
      STRANGEPOINT;
      TRETURN;
    }
  int type = get_oom_int (attach, "Type");
  m_utf8DisplayName = get_oom_string_s (attach, "DisplayName");
  m_fileName = get_oom_string_s (attach, "FileName");
  char *cid = get_pa_string (attach, PR_ATTACH_CONTENT_ID_DASL);
  if (cid)
    {
      m_cid = cid;
      xfree (cid);
    }

  if (m_fileName.empty())
    {
      if (!m_utf8DisplayName.empty())
        {
          log_dbg ("No filename found. Using displayName instead.");
          m_fileName = m_utf8DisplayName;
        }
      else
        {
          log_dbg ("No filename found and displayName using unknow.");
          m_fileName = "unknown_attachment.dat";
          m_utf8DisplayName = m_fileName;
        }
    }

  log_dbg ("Creating attachment obj for type: %i for %s",
           type, anonstr (m_utf8DisplayName.c_str ()));
  if (type == olByValue)
    {
      log_dbg ("Copying attachment by value");
      LPATTACH mapi_attachment = nullptr;
      mapi_attachment = (LPATTACH) get_oom_iunknown (attach,
                                                     "MapiObject");
      if (!mapi_attachment)
        {
          log_debug ("%s:%s: Failed to get MapiObject of attachment: %p",
                     SRCNAME, __func__, attach);
          TRETURN;
        }
      LPSTREAM stream = nullptr;
      if (FAILED (gpgol_openProperty (mapi_attachment, PR_ATTACH_DATA_BIN,
                                      &IID_IStream, 0, MAPI_MODIFY,
                                      (LPUNKNOWN*) &stream)))
        {
          log_debug ("%s:%s: Failed to open stream for mapi_attachment: %p",
                     SRCNAME, __func__, mapi_attachment);
          gpgol_release (mapi_attachment);
          TRETURN;
        }
      gpgol_release (mapi_attachment);
      char buf[COPYBUFSIZE];
      HRESULT hr;
      ULONG bRead = 0;
      while ((hr = stream->Read (buf, COPYBUFSIZE, &bRead)) == S_OK ||
             hr == S_FALSE)
        {
          if (!bRead)
            {
              // EOF
              break;
            }
          m_data.write (buf, bRead);
        }
      gpgol_release (stream);
    }
  else
    {
      log_dbg ("Creating attachment for other type.");
      HANDLE hTmpFile = INVALID_HANDLE_VALUE;
      char *path = get_tmp_outfile_utf8 (m_fileName.c_str (), &hTmpFile);
      if (!path || hTmpFile == INVALID_HANDLE_VALUE)
        {
          STRANGEPOINT;
          TRETURN;
        }
      CloseHandle (hTmpFile);
      if (oom_save_as_file (attach, path))
        {
          log_err ("Failed to call SaveAsFile.");
          TRETURN;
        }
      wchar_t *wname = utf8_to_wchar (path);
      hTmpFile = CreateFileW (wname,
                              GENERIC_READ | GENERIC_WRITE,
                              FILE_SHARE_READ | FILE_SHARE_DELETE,
                              NULL,
                              OPEN_EXISTING,
                              FILE_FLAG_DELETE_ON_CLOSE,
                              NULL);
      xfree (wname);
      if (hTmpFile == INVALID_HANDLE_VALUE)
        {
          log_debug_w32 (-1, "%s:%s: Failed to open attachment file.",
                         SRCNAME, __func__);
        }

      /* Read the attachment file into data */
      readFullFile (hTmpFile, m_data);
      CloseHandle (hTmpFile);
    }
}
#endif

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

std::string
Attachment::get_file_name () const
{
  if (m_fileName.size ())
    {
      return m_fileName;
    }
  return m_utf8DisplayName;
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

void
Attachment::set_is_mime (const bool &val)
{
  m_is_mime = val;
}

bool
Attachment::is_mime () const
{
  return m_is_mime;
}

#ifdef _WIN32
int
Attachment::attach_to (LPDISPATCH mailitem, std::string &r_errStr,
                       int *r_errCode)
{
  TSTART;
  int err = 0;

  if (!mailitem)
    {
      STRANGEPOINT;
      TRETURN 1;
    }

  if (m_is_mime)
    {
      log_dbg ("Parsing mime data of embedded mail.");
      MimeDataProvider mimeinfo;
      char buf[4096];
      size_t nread;
      /* Rewind */
      m_data.rewind ();
      while ((nread = m_data.read (buf, 4096)) > 0)
        {
          mimeinfo.write (buf, nread);
        }
      m_data.rewind ();
      const auto subject = mimeinfo.get_header ("Subject");
      if (!subject.size ())
        {
          log_dbg ("Could not extract subject");
          m_utf8DisplayName = _("unknown embedded message");
        }
      else if (subject.size () >= MAX_PATH - 4)
        {
          log_dbg ("Subject too long %s", anonstr (subject.c_str ()));
          m_utf8DisplayName = subject.substr (0, MAX_PATH - 4);
        }
      else
        {
          log_dbg ("Using subject '%s' as display name",
                   anonstr (subject.c_str ()));
          m_utf8DisplayName = subject;
        }

      // We have ensured that subject is smaller then MAX_PATH - 4, even
      // though we are the unicode API to avoid ridicicoulous paths. Now
      // we can sanitize that display name to get a file name.
      // Check if the string contains a period
      m_fileName = sanitizeFileName (m_utf8DisplayName);
      if (m_fileName.find('.') == std::string::npos)
        {
          // If no period is found, append .eml
          m_fileName += ".eml";
        }
      log_data ("Displayname for encapsulated message: '%s' \n Filename: '%s'",
                m_utf8DisplayName.c_str (), m_fileName.c_str ());
    }

  /* Looking at this several years after the code was written I
     do not understand why we don't set a fallback name like
     unknown attachment in this case. But to keep bug compatibility
     I am not changing this.   aheinecke - 2024-08-22 */
  const auto dispName = get_display_name ();
  if (dispName.empty())
    {
      log_dbg ("%s:%s: Ignoring attachment without display name.",
                 SRCNAME, __func__);
      TRETURN 0;
    }

  wchar_t* wchar_name = utf8_to_wchar (m_fileName.empty() ? dispName.c_str ()
                                                          : m_fileName.c_str ());
  if (!wchar_name)
    {
      log_error ("%s:%s: Failed to convert '%s' to wchar.",
                 SRCNAME, __func__, anonstr (dispName.c_str()));
      TRETURN 0;
    }

  wchar_t* wchar_disp = utf8_to_wchar (dispName.size() > 60 ? dispName.substr(0,59).c_str() : dispName.c_str ());
  if (dispName.size() > 60 )
    {
      log_dbg ("%s:%s: Displayname > 60 chars trucating to ",
                 SRCNAME, __func__);
    }
  if (!wchar_disp)
    {
      log_error ("%s:%s: Failed to convert '%s' to wchar.",
                 SRCNAME, __func__, anonstr (dispName.c_str()));
      TRETURN 0;
    }

  HANDLE hFile;
  wchar_t* wchar_file = get_tmp_outfile (wchar_name,
                                         &hFile);
  if (!wchar_file)
    {
      log_error ("%s:%s: Failed to obtain a tmp filename for: %s",
                 SRCNAME, __func__, anonstr (dispName.c_str()));
      err = 1;
    }
  if (!err && copy_to (hFile))
    {
      log_error ("%s:%s: Failed to copy attachment %s to temp file",
                 SRCNAME, __func__, anonstr (dispName.c_str()));
      err = 1;
    }
  if (!err && add_oom_attachment (mailitem, wchar_file, wchar_disp,
                                  r_errStr, r_errCode))
    {
      log_error ("%s:%s: Failed to add attachment: %s",
                 SRCNAME, __func__, anonstr (dispName.c_str()));
      err = 1;
    }
  if (hFile && hFile != INVALID_HANDLE_VALUE)
    {
      CloseHandle (hFile);
    }
  if (wchar_file && !DeleteFileW (wchar_file))
    {
      log_error ("%s:%s: Failed to delete tmp attachment for: %s",
                 SRCNAME, __func__, anonstr (dispName.c_str()));
      err = 1;
    }
  xfree (wchar_file);
  xfree (wchar_name);
  xfree (wchar_disp);
  TRETURN err;
}

int
Attachment::copy_to (HANDLE hFile)
{
  TSTART;
  char copybuf[COPYBUFSIZE];
  size_t nread;

  /* Security considerations: Writing the data to a temporary
     file is necessary as neither MAPI manipulation works in the
     read event to transmit the data nor Property Accessor
     works (see above). From a security standpoint there is a
     short time where the temporary files are on disk. Tempdir
     should be protected so that only the user can read it. Thus
     we have a local attack that could also take the data out
     of Outlook. FILE_SHARE_READ is necessary so that outlook
     can read the file.

     A bigger concern is that the file is manipulated
     by another software to fake the signature state. So
     we keep the write exlusive to us.

     We delete the file before closing the write file handle.
  */

  /* Make sure we start at the beginning */
  m_data.seek (0, SEEK_SET);
  while ((nread = m_data.read (copybuf, COPYBUFSIZE)))
    {
      DWORD nwritten;
      if (!WriteFile (hFile, copybuf, nread, &nwritten, NULL))
        {
          log_error ("%s:%s: Failed to write in tmp attachment.",
                     SRCNAME, __func__);
          TRETURN 1;
        }
      if (nread != nwritten)
        {
          log_error ("%s:%s: Write truncated.",
                     SRCNAME, __func__);
          TRETURN 1;
        }
    }
  TRETURN 0;
}
#endif
