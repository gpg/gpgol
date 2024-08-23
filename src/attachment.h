/* attachment.h - Wrapper class for attachments
 * Copyright (C) 2005, 2007 g10 Code GmbH
 * Copyright (C) 2015, 2016 by Bundesamt für Sicherheit in der Informationstechnik
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
#ifndef ATTACHMENT_H
#define ATTACHMENT_H

#include <string>

#include <gpgme++/data.h>

#ifdef _WIN32
# include "oomhelp.h"
#endif
/** Helper class for attachment actions. */
class Attachment
{
public:
  enum olAttachmentType
    {
      olByValue = 1,
      olByReference = 4,
      olEmbeddeditem = 5,
      olOLE = 6,
    };
  /** Creates and opens a new in memory attachment. */
  Attachment();
  ~Attachment();

  /** Set the display name */
  void set_display_name(const char *name);
  std::string get_display_name() const;

  /** Get file name, usually the same as display name */
  std::string get_file_name() const;

  void set_attach_type(attachtype_t type);

  /* Content id */
  void set_content_id (const char *cid);
  std::string get_content_id() const;

  /* Content type */
  void set_content_type (const char *type);
  std::string get_content_type () const;

  /* Is the data itself in mime encoding */
  void set_is_mime (const bool &val);
  bool is_mime () const;

  /* get the underlying data structure */
  GpgME::Data& get_data();

#ifdef _WIN32
  /** Create a data struct from OOM attachment */
  Attachment (LPDISPATCH attach);

  /** Copy the attachment to a file handle */
  int copy_to (HANDLE hFile);

  /** Add the attachment to a Mailitem. Returns the error
   * information of the AddAttachment call in r_errStr and
   * r_errCode. */
  int attach_to (LPDISPATCH mailitem, std::string &r_errStr, int *r_errCode);
#endif
private:
  GpgME::Data m_data;
  std::string m_utf8DisplayName;
  std::string m_fileName;
  attachtype_t m_type;
  std::string m_cid;
  std::string m_ctype;
  bool m_is_mime = false;
};

#endif // ATTACHMENT_H
