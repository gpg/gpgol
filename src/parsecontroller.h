/* @file parsecontroller.h
 * @brief Controll the parsing and decrypt / verify of a mail
 *
 * Copyright (C) 2016 by Bundesamt für Sicherheit in der Informationstechnik
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

#ifndef PARSECONTROLLER_H
#define PARSECONTROLLER_H

#include <string>
#include <vector>
#include <memory>

#ifdef HAVE_CONFIG_H
#include <config.h>
#endif

#include "common_indep.h"

#include <gpgme++/decryptionresult.h>
#include <gpgme++/verificationresult.h>
#include <gpgme++/data.h>

class Attachment;
class MimeDataProvider;

#ifdef HAVE_W32_SYSTEM
#include "oomhelp.h"
#endif

/* A template for decryption errors / status message. */
extern const char decrypt_template[];
extern const char decrypt_template_html[];

class ParseController
{
public:
#ifdef HAVE_W32_SYSTEM
  /** Construct a new ParseController for the stream instream.
    instream is expected to point to a mime mail.
    Adds a reference to the stream and releases it on
    destruction. */
  ParseController(LPSTREAM instream, msgtype_t type);
#endif
  ParseController(FILE *instream, msgtype_t type);

  ~ParseController();

  /** Main entry point. After execution getters will become
  valid. */
  void parse();

  /** Get the Body. Call parse first. */
  const std::string get_body () const;

  /** Get the charset of the body. Call parse first.
    *
    * That is a bit of a clunky API to make testing
    * without outlook easier as we use mlang for Charset
    * conversion which is not available on GNU/Linux.
  */
  const std::string get_body_charset() const;
  const std::string get_html_charset() const;

  /** Get an alternative? HTML Body. Call parse first. */
  const std::string get_html_body () const;

  /** Get the decrypted / verified attachments. Call parse first.
  */
  std::vector<std::shared_ptr<Attachment> > get_attachments() const;

  const GpgME::DecryptionResult decrypt_result() const
  { return m_decrypt_result; }
  const GpgME::VerificationResult verify_result() const
  { return m_verify_result; }

  const std::string get_formatted_error() const
  { return m_error; }

  void setSender(const std::string &sender);

  bool shouldBlockHtml() const
  { return m_block_html; }

private:
  /* State variables */
  MimeDataProvider *m_inputprovider;
  MimeDataProvider *m_outputprovider;
  msgtype_t m_type;
  std::string m_error;
  GpgME::DecryptionResult m_decrypt_result;
  GpgME::VerificationResult m_verify_result;
  std::string m_sender;
  bool m_block_html;
};

#endif /* PARSECONTROLLER_H */
