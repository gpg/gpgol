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

#ifdef HAVE_W32_SYSTEM
#include "common.h"
/* We use UTF-8 internally. */
#undef _
# define _(a) utf8_gettext (a)
#else
# define _(a) a
#endif



const char decrypt_template[] = {
"<html><head></head><body>"
"<table border=\"0\" width=\"100%%\" cellspacing=\"1\" cellpadding=\"1\" bgcolor=\"#0069cc\">"
"<tr>"
"<td bgcolor=\"#0080ff\">"
"<p><span style=\"font-weight:600; background-color:#0080ff;\"><center>%s %s</center><span></p></td></tr>"
"<tr>"
"<td bgcolor=\"#e0f0ff\">"
"<center>"
"<br/>%s"
"</td></tr>"
"</table></body></html>"};

using namespace GpgME;

static bool
expect_no_headers (msgtype_t type)
{
  return type != MSGTYPE_GPGOL_MULTIPART_SIGNED &&
         type != MSGTYPE_GPGOL_OPAQUE_SIGNED &&
         type != MSGTYPE_GPGOL_OPAQUE_ENCRYPTED;
}

#ifdef HAVE_W32_SYSTEM
ParseController::ParseController(LPSTREAM instream, msgtype_t type):
    m_inputprovider  (new MimeDataProvider(instream,
                          expect_no_headers(type))),
    m_outputprovider (new MimeDataProvider()),
    m_type (type)
{
  log_mime_parser ("%s:%s: Creating parser for stream: %p of type %i",
                   SRCNAME, __func__, instream, type);
}
#endif

ParseController::ParseController(FILE *instream, msgtype_t type):
    m_inputprovider  (new MimeDataProvider(instream,
                          expect_no_headers(type))),
    m_outputprovider (new MimeDataProvider()),
    m_type (type)
{
  log_mime_parser ("%s:%s: Creating parser for stream: %p of type %i",
                   SRCNAME, __func__, instream, type);
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

static bool
is_opaque_signed (Data &data)
{
  data.seek (0, SEEK_SET);
  auto id = data.type();
  data.seek (0, SEEK_SET);
  return id == Data::CMSSigned;
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
      if ((!m_decrypt_result.error () &&
          m_verify_result.signatures ().empty() &&
          m_outputprovider->signature ()) ||
          is_opaque_signed (output))
        {
          /* There is a signature in the output. So we have
             to verify it now as an extra step. */
          input = Data (m_outputprovider);
          delete m_inputprovider;
          m_inputprovider = m_outputprovider;
          m_outputprovider = new MimeDataProvider();
          output = Data(m_outputprovider);
          verify = true;
        }
      else
        {
          verify = false;
        }
    }
  if (verify)
    {
      const auto sig = m_inputprovider->signature();
      input.seek (0, SEEK_SET);
      if (sig)
        {
          sig->seek (0, SEEK_SET);
          m_verify_result = ctx->verifyDetachedSignature(*sig, input);
          /* Copy the input to output to do a mime parsing. */
          char buf[4096];
          input.seek (0, SEEK_SET);
          output.seek (0, SEEK_SET);
          size_t nread;
          while ((nread = input.read (buf, 4096)) > 0)
            {
              output.write (buf, nread);
            }
        }
      else
        {
          m_verify_result = ctx->verifyOpaqueSignature(input, output);
        }
    }
  delete ctx;
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

const std::string
ParseController::get_body_charset() const
{
  if (m_outputprovider)
    {
      return m_outputprovider->get_body_charset();
    }
  else
    {
      return std::string();
    }
}

const std::string
ParseController::get_html_charset() const
{
  if (m_outputprovider)
    {
      return m_outputprovider->get_body_charset();
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
