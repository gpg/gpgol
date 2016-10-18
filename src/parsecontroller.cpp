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
#include <gpgme++/key.h>

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

static std::string
format_recipients(GpgME::DecryptionResult result)
{
  std::string msg;
  for (const auto recipient: result.recipients())
    {
      msg += std::string("<br/>0x") + recipient.keyID() + "</a>";
    }
  return msg;
}

static std::string
format_error(GpgME::DecryptionResult result, Protocol protocol)
{
  char *buf;
  bool no_sec = false;
  std::string msg;

  if (result.error ().isCanceled () ||
      result.error ().code () == GPG_ERR_NO_SECKEY)
    {
       msg = _("Decryption canceled or timed out.");
    }

  if (result.error ().code () == GPG_ERR_DECRYPT_FAILED)
    {
      no_sec = true;
      for (const auto &recipient: result.recipients ()) {
        no_sec &= (recipient.status ().code () == GPG_ERR_NO_SECKEY);
      }
    }

  if (no_sec)
    {
      msg = _("No secret key found to decrypt the message."
              "It is encrypted for following keys:");
      msg += format_recipients (result);
    }
  else
    {
      msg = _("Could not decrypt the data.");
    }

  if (gpgrt_asprintf (&buf, decrypt_template,
                      protocol == OpenPGP ? "OpenPGP" : "S/MIME",
                      _("Encrypted message (decryption not possible)"),
                      msg.c_str()) == -1)
    {
      log_error ("%s:%s:Failed to Format error.",
                 SRCNAME, __func__);
      return "Failed to Format error.";
    }
  msg = buf;
  return msg;
}


void
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
      char *buf;
      const char *proto = protocol == OpenPGP ? "OpenPGP" : "S/MIME";
      if (gpgrt_asprintf (&buf, decrypt_template,
                          proto,
                          _("Encrypted message (decryption not possible)"),
                          _("Failed to find GnuPG please ensure that GnuPG or "
                            "Gpg4win is properly installed.")) == -1)
        {
          log_error ("%s:%s:Failed format error.",
                     SRCNAME, __func__);
          /* Should never happen */
          m_error = std::string("Bad installation");
        }
      m_error = buf;
      xfree (buf);
      return;
    }
  ctx->setArmor(true);

  Data output(m_outputprovider);

  Data input (m_inputprovider);
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
      if (m_decrypt_result.error())
        {
          m_error = format_error (m_decrypt_result, protocol);
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

  /* Ensure that the Keys for the signatures are available */
  for (const auto sig: m_verify_result.signatures())
    {
      sig.key(true, true);
    }

  if (opt.enable_debug)
    {
       std::stringstream ss;
       ss << m_decrypt_result << '\n' << m_verify_result;
       log_debug ("Decrypt / Verify result: %s", ss.str().c_str());
    }

  return;
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
