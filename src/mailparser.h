/* @file mailparser.h
 * @brief Parse a mail and decrypt / verify accordingly
 *
 *    Copyright (C) 2016 Intevation GmbH
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

#ifndef MAILPARSER_H
#define MAILPARSER_H

#include "oomhelp.h"
#include "mapihelp.h"
#include "gpgme.h"

#include <string>
#include <vector>
#include <memory>

class Attachment;

class MailParser
{
public:
  /** Construct a new MailParser for the stream instream.
    instream is expected to point to a mime mail.
    Adds a reference to the stream and releases it on
    destruction. */
  MailParser(LPSTREAM instream, msgtype_t type);
  ~MailParser();

  /** Construct a new MailParser for an inline message where
    the content is pointet to by body.
  MailParser(const char *body, msgtype_t type);
  */
  /** Main entry point. Parses the Mail returns an
    * empty string on success or an error message on failure. */
  std::string parse();

  /** Get the Body converted to utf8. Call parse first. */
  std::shared_ptr<std::string> get_utf8_text_body();

  /** Get an alternative? HTML Body converted to utf8. Call parse first. */
  std::shared_ptr<std::string> get_utf8_html_body();

  /** Get the decrypted / verified attachments. Call parse first.
  */
  std::vector<std::shared_ptr<Attachment> > get_attachments();
private:
  std::vector<std::shared_ptr<Attachment> > m_attachments;
  std::shared_ptr<std::string> m_body;
  std::shared_ptr<std::string> m_htmlbody;

  /* State variables */
  LPSTREAM m_stream;
  msgtype_t m_type;
  bool m_error;
  bool m_in_data;
  gpgme_data_t signed_data;/* NULL or the data object used to collect
                              the signed data. It would be better to
                              just hash it but there is no support in
                              gpgme for this yet. */
  gpgme_data_t sig_data;  /* NULL or data object to collect the
                             signature attachment which should be a
                             signature then.  */
};

#endif /* MAILPARSER_H */
