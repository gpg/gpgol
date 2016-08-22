/* @file mailparser.cpp
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
#include "config.h"
#include "common.h"

#include "mailparser.h"
#include "attachment.h"
#include "mimedataprovider.h"

#include <gpgme++/context.h>
#include <gpgme++/decryptionresult.h>

#include <sstream>

using namespace GpgME;

MailParser::MailParser(LPSTREAM instream, msgtype_t type):
    m_body (std::shared_ptr<std::string>(new std::string())),
    m_htmlbody (std::shared_ptr<std::string>(new std::string())),
    m_input (Data(new MimeDataProvider(instream))),
    m_type (type),
    m_error (false)
{
  log_mime_parser ("%s:%s: Creating parser for stream: %p",
                   SRCNAME, __func__, instream);
}

MailParser::~MailParser()
{
  log_debug ("%s:%s", SRCNAME, __func__);
}

static void
operation_for_type(msgtype_t type, bool *decrypt,
                   bool *verify, Protocol *protocol)
{
  *decrypt = false;
  *verify = false;
  *protocol = Protocol::UnknownProtocol;
  switch (type)
    {
      case MSGTYPE_GPGOL_MULTIPART_ENCRYPTED:
      case MSGTYPE_GPGOL_PGP_MESSAGE:
        *decrypt = true;
        *protocol = Protocol::OpenPGP;
        break;
      case MSGTYPE_GPGOL_MULTIPART_SIGNED:
      case MSGTYPE_GPGOL_CLEAR_SIGNED:
        *verify = true;
        *protocol = Protocol::OpenPGP;
        break;
      case MSGTYPE_GPGOL_OPAQUE_SIGNED:
        *protocol = Protocol::CMS;
        *verify = true;
        break;
      case MSGTYPE_GPGOL_OPAQUE_ENCRYPTED:
        *protocol = Protocol::CMS;
        *decrypt = true;
        break;
      default:
        log_error ("%s:%s: Unknown data type: %i",
                   SRCNAME, __func__, type);
    }
}

std::string
MailParser::parse()
{
  // Wrap the input stream in an attachment / GpgME Data
  Protocol protocol;
  bool decrypt, verify;
  operation_for_type (m_type, &decrypt, &verify, &protocol);
  auto ctx = Context::createForProtocol (protocol);
  ctx->setArmor(true);

  if (decrypt)
    {
      Data output;
      log_debug ("%s:%s: Decrypting with protocol: %s",
                 SRCNAME, __func__,
                 protocol == OpenPGP ? "OpenPGP" :
                 protocol == CMS ? "CMS" : "Unknown");
      auto combined_result = ctx->decryptAndVerify(m_input, output);
      m_decrypt_result = combined_result.first;
      m_verify_result = combined_result.second;
      if (m_decrypt_result.error())
        {
          MessageBox (NULL, "Decryption failed.", "Failed", MB_OK);
        }
      char buf[2048];
      size_t bRead;
      output.seek (0, SEEK_SET);
      while ((bRead = output.read (buf, 2048)) > 0)
        {
          (*m_body).append(buf, bRead);
        }
      log_debug ("Body is: %s", m_body->c_str());
    }

  if (opt.enable_debug)
    {
       std::stringstream ss;
       ss << m_decrypt_result << '\n' << m_verify_result;
       log_debug ("Decrypt / Verify result: %s", ss.str().c_str());
    }
  Attachment *att = new Attachment ();
  att->write ("Hello attachment", strlen ("Hello attachment"));
  att->set_display_name ("The Attachment.txt");
  m_attachments.push_back (std::shared_ptr<Attachment>(att));
  return std::string();
}

std::shared_ptr<std::string>
MailParser::get_utf8_html_body()
{
  return m_htmlbody;
}

std::shared_ptr<std::string>
MailParser::get_utf8_text_body()
{
  return m_body;
}

std::vector<std::shared_ptr<Attachment> >
MailParser::get_attachments()
{
  return m_attachments;
}
