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

#include <string>
#include <vector>
#include <memory>

#include <gpgme++/decryptionresult.h>
#include <gpgme++/verificationresult.h>
#include <gpgme++/data.h>

class Attachment;
class MimeDataProvider;

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

  /** Get the Body. Call parse first. */
  const std::string get_body() const;

  /** Get an alternative? HTML Body. Call parse first. */
  const std::string get_html_body() const;

  /** Get the decrypted / verified attachments. Call parse first.
  */
  std::vector<std::shared_ptr<Attachment> > get_attachments() const;
private:
  /* State variables */
  MimeDataProvider *m_inputprovider;
  MimeDataProvider *m_outputprovider;
  msgtype_t m_type;
  bool m_error;
  GpgME::DecryptionResult m_decrypt_result;
  GpgME::VerificationResult m_verify_result;
};

#endif /* MAILPARSER_H */
