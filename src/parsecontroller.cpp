/* @file parsecontroller.cpp
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

#include "parsecontroller.h"
#include "attachment.h"
#include "mimedataprovider.h"

#include <gpgme++/context.h>
#include <gpgme++/decryptionresult.h>

#include <sstream>

using namespace GpgME;

#ifdef HAVE_W32_SYSTEM
ParseController::ParseController(LPSTREAM instream, msgtype_t type):
    m_inputprovider  (new MimeDataProvider(instream)),
    m_outputprovider (new MimeDataProvider()),
    m_type (type),
    m_error (false)
{
  log_mime_parser ("%s:%s: Creating parser for stream: %p",
                   SRCNAME, __func__, instream);
}
#endif

ParseController::ParseController(FILE *instream, msgtype_t type):
    m_inputprovider  (new MimeDataProvider(instream)),
    m_outputprovider (new MimeDataProvider()),
    m_type (type),
    m_error (false)
{
  log_mime_parser ("%s:%s: Creating parser for stream: %p",
                   SRCNAME, __func__, instream);
}

ParseController::~ParseController()
{
  log_debug ("%s:%s", SRCNAME, __func__);
  delete m_inputprovider;
  delete m_outputprovider;
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
ParseController::parse()
{
  // Wrap the input stream in an attachment / GpgME Data
  Protocol protocol;
  bool decrypt, verify;
  operation_for_type (m_type, &decrypt, &verify, &protocol);
  auto ctx = Context::createForProtocol (protocol);
  if (!ctx)
    {
      log_error ("%s:%s:Failed to create context. Installation broken.",
                 SRCNAME, __func__);
      // TODO proper error handling
      return std::string("Bad installation");
    }
  ctx->setArmor(true);

  Data output(m_outputprovider);

  Data input;
  if (m_type == MSGTYPE_GPGOL_CLEAR_SIGNED ||
      m_type == MSGTYPE_GPGOL_PGP_MESSAGE)
    {
      /* For clearsigned and PGP Message take the body.
         This does not copy the data. */
      input = Data (m_inputprovider->get_body().c_str(),
                    m_inputprovider->get_body().size(), false);
    }
  else
    {
      input = Data (m_inputprovider);
    }
  log_debug ("%s:%s: decrypt: %i verify: %i with protocol: %s",
             SRCNAME, __func__,
             decrypt, verify,
             protocol == OpenPGP ? "OpenPGP" :
             protocol == CMS ? "CMS" : "Unknown");
  if (decrypt)
    {
      input.seek (0, SEEK_SET);
      auto combined_result = ctx->decryptAndVerify(input, output);
      m_decrypt_result = combined_result.first;
      m_verify_result = combined_result.second;
    }
  else
    {
      const auto sig = m_inputprovider->signature();
      /* Ignore the first two bytes if we did not decrypt. */
      input.seek (2, SEEK_SET);
      if (sig)
        {
          sig->seek (0, SEEK_SET);
          m_verify_result = ctx->verifyDetachedSignature(*sig, input);
        }
      else
        {
          m_verify_result = ctx->verifyOpaqueSignature(input, output);
        }
    }
  log_debug ("%s:%s: decrypt err: %i verify err: %i",
             SRCNAME, __func__, m_decrypt_result.error().code(),
             m_verify_result.error().code());

  if (opt.enable_debug)
    {
       std::stringstream ss;
       ss << m_decrypt_result << '\n' << m_verify_result;
       log_debug ("Decrypt / Verify result: %s", ss.str().c_str());
    }
  /*
  Attachment *att = new Attachment ();
  att->write ("Hello attachment", strlen ("Hello attachment"));
  att->set_display_name ("The Attachment.txt");
  m_attachments.push_back (std::shared_ptr<Attachment>(att));
  */
  return std::string();
}

const std::string
ParseController::get_html_body() const
{
  if (m_outputprovider)
    {
      return m_outputprovider->get_html_body();
    }
  else
    {
      return std::string();
    }
}

const std::string
ParseController::get_body() const
{
  if (m_outputprovider)
    {
      return m_outputprovider->get_body();
    }
  else
    {
      return std::string();
    }
}

std::vector<std::shared_ptr<Attachment> >
ParseController::get_attachments() const
{
  if (m_outputprovider)
    {
      return m_outputprovider->get_attachments();
    }
  else
    {
      return std::vector<std::shared_ptr<Attachment> >();
    }
}
