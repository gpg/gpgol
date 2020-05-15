/* @file parsecontroller.cpp
 * @brief Parse a mail and decrypt / verify accordingly
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
#include "config.h"

#include "parsecontroller.h"
#include "attachment.h"
#include "mimedataprovider.h"

#include "keycache.h"

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


const char decrypt_template_html[] = {
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

const char decrypt_template[] = {"%s %s\n\n%s"};

using namespace GpgME;

static bool
expect_no_headers (msgtype_t type)
{
  TSTART;
  TRETURN type != MSGTYPE_GPGOL_MULTIPART_SIGNED;
}

static bool
expect_no_mime (msgtype_t type)
{
  TSTART;
  TRETURN type == MSGTYPE_GPGOL_PGP_MESSAGE ||
         type == MSGTYPE_GPGOL_CLEAR_SIGNED;
}

#ifdef BUILD_TESTS
static void
get_and_print_key_test (const char *fingerprint, GpgME::Protocol proto)
{
  if (!fingerprint)
    {
      STRANGEPOINT;
      return;
    }
  auto ctx = std::unique_ptr<GpgME::Context> (GpgME::Context::createForProtocol
                                              (proto));

  if (!ctx)
    {
      STRANGEPOINT;
      return;
    }
  ctx->setKeyListMode (GpgME::KeyListMode::Local |
                       GpgME::KeyListMode::Signatures |
                       GpgME::KeyListMode::Validate |
                       GpgME::KeyListMode::WithTofu);

  GpgME::Error err;
  const auto newKey = ctx->key (fingerprint, err, false);

  std::stringstream ss;
  ss << newKey;

  log_debug ("Key: %s", ss.str().c_str());
  return;
}
#endif

#ifdef HAVE_W32_SYSTEM
ParseController::ParseController(LPSTREAM instream, msgtype_t type):
    m_inputprovider  (new MimeDataProvider(instream,
                          expect_no_headers(type))),
    m_outputprovider (new MimeDataProvider(expect_no_mime(type))),
    m_type (type),
    m_block_html (false),
    m_second_pass (false)

{
  TSTART;
  memdbg_ctor ("ParseController");
  log_data ("%s:%s: Creating parser for stream: %p of type %i"
                   " expect no headers: %i expect no mime: %i",
                   SRCNAME, __func__, instream, type,
                   expect_no_headers (type), expect_no_mime (type));
  TRETURN;
}
#endif

ParseController::ParseController(FILE *instream, msgtype_t type):
    m_inputprovider  (new MimeDataProvider(instream,
                          expect_no_headers(type))),
    m_outputprovider (new MimeDataProvider(expect_no_mime(type))),
    m_type (type),
    m_block_html (false),
    m_second_pass (false)
{
  TSTART;
  memdbg_ctor ("ParseController");
  log_data ("%s:%s: Creating parser for stream: %p of type %i",
                   SRCNAME, __func__, instream, type);
  TRETURN;
}

ParseController::~ParseController()
{
  TSTART;
  log_debug ("%s:%s", SRCNAME, __func__);
  memdbg_dtor ("ParseController");
  delete m_inputprovider;
  delete m_outputprovider;
  TRETURN;
}

static void
operation_for_type(msgtype_t type, bool *decrypt,
                   bool *verify)
{
  *decrypt = false;
  *verify = false;
  switch (type)
    {
      case MSGTYPE_GPGOL_MULTIPART_ENCRYPTED:
      case MSGTYPE_GPGOL_PGP_MESSAGE:
        *decrypt = true;
        break;
      case MSGTYPE_GPGOL_MULTIPART_SIGNED:
      case MSGTYPE_GPGOL_CLEAR_SIGNED:
        *verify = true;
        break;
      case MSGTYPE_GPGOL_OPAQUE_SIGNED:
        *verify = true;
        break;
      case MSGTYPE_GPGOL_OPAQUE_ENCRYPTED:
        *decrypt = true;
        break;
      default:
        log_error ("%s:%s: Unknown data type: %i",
                   SRCNAME, __func__, type);
    }
}

static bool
is_smime (Data &data)
{
  TSTART;
  data.seek (0, SEEK_SET);
  auto id = data.type();
  data.seek (0, SEEK_SET);
  TRETURN id == Data::CMSSigned || id == Data::CMSEncrypted;
}

static std::string
format_recipients(GpgME::DecryptionResult result, Protocol protocol)
{
  TSTART;
  std::string msg;
  for (const auto recipient: result.recipients())
    {
      auto ctx = Context::createForProtocol(protocol);
      Error e;
      if (!ctx) {
          /* Can't happen */
          TRACEPOINT;
          continue;
      }
      const auto key = ctx->key(recipient.keyID(), e, false);
      delete ctx;
      if (!key.isNull() && key.numUserIDs() && !e) {
        msg += std::string("<br/>") + key.userIDs()[0].id() + " (0x" + recipient.keyID() + ")";
        continue;
      }
      msg += std::string("<br/>") + _("Unknown Key:") + " 0x" + recipient.keyID();
    }
  TRETURN msg;
}

static std::string
format_error(GpgME::DecryptionResult result, Protocol protocol)
{
  TSTART;
  char *buf;
  bool no_sec = false;
  std::string msg;

  if (result.error ().isCanceled () ||
      result.error ().code () == GPG_ERR_NO_SECKEY)
    {
       msg = _("Decryption canceled or timed out.");
    }

  if (result.error ().code () == GPG_ERR_DECRYPT_FAILED ||
      result.error ().code () == GPG_ERR_NO_SECKEY)
    {
      no_sec = true;
      for (const auto &recipient: result.recipients ()) {
        no_sec &= (recipient.status ().code () == GPG_ERR_NO_SECKEY);
      }
    }

  if (no_sec)
    {
      msg = _("No secret key found to decrypt the message. "
              "It is encrypted to the following keys:");
      msg += format_recipients (result, protocol);
    }
  else
    {
      msg = _("Could not decrypt the data: ");

      if (result.isNull ())
        {
          msg += _("Failed to parse the mail.");
        }
      else if (result.isLegacyCipherNoMDC())
        {
          msg += _("Data is not integrity protected. "
                   "Decrypting it could be a security problem. (no MDC)");
        }
      else
        {
          msg += result.error().asString();
        }
    }

  if (gpgrt_asprintf (&buf, opt.prefer_html ? decrypt_template_html :
                      decrypt_template,
                      protocol == OpenPGP ? "OpenPGP" : "S/MIME",
                      _("Encrypted message (decryption not possible)"),
                      msg.c_str()) == -1)
    {
      log_error ("%s:%s:Failed to Format error.",
                 SRCNAME, __func__);
      TRETURN "Failed to Format error.";
    }
  msg = buf;
  memdbg_alloc (buf);
  xfree (buf);
  TRETURN msg;
}

void
ParseController::setSender(const std::string &sender)
{
  TSTART;
  m_sender = sender;
  TRETURN;
}

static void
handle_autocrypt_info (const autocrypt_s &info)
{
  TSTART;
  if (info.pref.size() && info.pref != "mutual")
    {
      log_debug ("%s:%s: Autocrypt preference is %s which is unhandled.",
                 SRCNAME, __func__, info.pref.c_str());
      TRETURN;
    }

#ifndef BUILD_TESTS
  if (!KeyCache::import_pgp_key_data (info.data))
    {
      log_error ("%s:%s: Failed to import", SRCNAME, __func__);
    }
#endif
  /* No need to free b64decoded as gpgme_data_release handles it */

  TRETURN;
}

static bool
is_valid_chksum(const GpgME::Signature &sig)
{
  TSTART;
  const auto sum = sig.summary();
  static unsigned int valid_mask = (unsigned int) (
      GpgME::Signature::Valid |
      GpgME::Signature::Green |
      GpgME::Signature::KeyRevoked |
      GpgME::Signature::KeyExpired |
      GpgME::Signature::SigExpired |
      GpgME::Signature::CrlMissing |
      GpgME::Signature::CrlTooOld |
      GpgME::Signature::TofuConflict );

  bool valid = (sum & valid_mask);
  if (!valid)
    {
      log_debug ("%s:%s: Treating sig as invalid because status is %x",
                 SRCNAME, __func__, sum);
    }
  TRETURN valid;
}

/* Note on stability:

   Experiments have shown that we can have a crash if parse
   Returns at time that is not good for the state of Outlook.

   This happend in my test instance after a delay of > 1s < 3s
   with a < 1% chance :-/

   So if you have really really bad luck this might still crash
   although it usually should be either much quicker or much slower
   (slower e.g. when pinentry is requrired).
*/
void
ParseController::parse(bool offline)
{
  TSTART;
  // Wrap the input stream in an attachment / GpgME Data
  Protocol protocol;
  bool decrypt, verify;

  Data input (m_inputprovider);

  auto inputType = input.type ();

  if (m_autocrypt_info.exists)
    {
      handle_autocrypt_info (m_autocrypt_info);
    }

  if (inputType == Data::Type::PGPSigned)
    {
      verify = true;
      decrypt = false;
    }
  else
    {
      operation_for_type (m_type, &decrypt, &verify);
    }

  if ((m_inputprovider->signature() && is_smime (*m_inputprovider->signature())) ||
      is_smime (input))
    {
      protocol = Protocol::CMS;
      if (m_second_pass)
        {
          log_dbg ("Second pass parsing for CMS. Only doing verify.");
          verify = true;
          decrypt = false;
        }
    }
  else
    {
      protocol = Protocol::OpenPGP;
    }
  auto ctx = std::unique_ptr<Context> (Context::createForProtocol (protocol));
  if (!ctx)
    {
      log_error ("%s:%s:Failed to create context. Installation broken.",
                 SRCNAME, __func__);
      char *buf;
      const char *proto = protocol == OpenPGP ? "OpenPGP" : "S/MIME";
      if (gpgrt_asprintf (&buf, opt.prefer_html ? decrypt_template_html :
                          decrypt_template,
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
      memdbg_alloc (buf);
      m_error = buf;
      xfree (buf);
      TRETURN;
    }

  /* Maybe a different option for this ? */
  if (opt.autoretrieve)
    {
      ctx->setFlag("auto-key-retrieve", "1");
    }

  ctx->setArmor(true);

  if (!m_sender.empty())
    {
      ctx->setSender(m_sender.c_str());
    }
  if (offline)
    {
      ctx->setOffline (true);
    }

  Data output (m_outputprovider);
  log_debug ("%s:%s:%p decrypt: %i verify: %i with protocol: %s sender: %s type: %i",
             SRCNAME, __func__, this,
             decrypt, verify,
             protocol == OpenPGP ? "OpenPGP" :
             protocol == CMS ? "CMS" : "Unknown",
             m_sender.empty() ? "none" : anonstr (m_sender.c_str()), inputType);
  if (decrypt)
    {
      input.seek (0, SEEK_SET);
      TRACEPOINT;
      auto combined_result = ctx->decryptAndVerify(input, output);
      log_debug ("%s:%s:%p decrypt / verify done.",
                 SRCNAME, __func__, this);
      m_decrypt_result = combined_result.first;
      m_verify_result = combined_result.second;

      if ((!m_decrypt_result.error () &&
          m_verify_result.signatures ().empty() &&
          m_outputprovider->signature ()) ||
          is_smime (output) ||
          output.type() == Data::Type::PGPSigned)
        {
          TRACEPOINT;
          /* There is a signature in the output. So we have
             to verify it now as an extra step. */
          input = Data (m_outputprovider);
          delete m_inputprovider;
          m_inputprovider = m_outputprovider;
          m_outputprovider = new MimeDataProvider();
          output = Data(m_outputprovider);
          verify = true;
          TRACEPOINT;
        }
      else
        {
          verify = false;
        }
      TRACEPOINT;
      if (m_decrypt_result.error () || m_decrypt_result.isNull () ||
          m_decrypt_result.error ().isCanceled ())
        {
          m_error = format_error (m_decrypt_result, protocol);
        }
    }
  if (verify)
    {
      TRACEPOINT;
      GpgME::Data *sig = m_inputprovider->signature();
      input.seek (0, SEEK_SET);
      if (sig)
        {
          sig->seek (0, SEEK_SET);
          TRACEPOINT;
          m_verify_result = ctx->verifyDetachedSignature(*sig, input);
          log_debug ("%s:%s:%p verify done.",
                     SRCNAME, __func__, this);
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
          TRACEPOINT;
          m_verify_result = ctx->verifyOpaqueSignature(input, output);
          TRACEPOINT;

          const auto sigs = m_verify_result.signatures();
          bool allBad = sigs.size();
          for (const auto s :sigs)
            {
              if (!(s.summary() & Signature::Red))
                {
                  allBad = false;
                  break;
                }
            }
#ifdef HAVE_W32_SYSTEM
          if (allBad)
            {
              log_debug ("%s:%s:%p inline verify error trying native to utf8.",
                         SRCNAME, __func__, this);


              /* The proper solution would be to take the encoding from
                 the mail / headers. Then convert the wchar body to that
                 encoding. Verify, and convert it after verifcation to
                 UTF-8 which the rest of the code expects.

                 Or native_body from native ACP to InternetCodepage, then
                 verify and convert the output back to utf8 as the rest
                 expects.

                 But as this is clearsigned and we don't really want that.
                 Meh.
                 */
              char *utf8 = native_to_utf8 (input.toString().c_str());
              if (utf8)
                {
                  // Try again after conversion.
                  ctx = std::unique_ptr<Context> (Context::createForProtocol (protocol));
                  ctx->setArmor (true);
                  if (!m_sender.empty())
                    {
                      ctx->setSender(m_sender.c_str());
                    }

                  input = Data (utf8, strlen (utf8));
                  xfree (utf8);

                  // Use a fresh output
                  auto provider = new MimeDataProvider (true);

                  // Warning: The dtor of the Data object touches
                  // the provider. So we have to delete it after
                  // the assignment.
                  output = Data (provider);
                  delete m_outputprovider;
                  m_outputprovider = provider;

                  // Try again
                  m_verify_result = ctx->verifyOpaqueSignature(input, output);
                }
            }
#else
(void)allBad;
#endif
        }
    }
  log_debug ("%s:%s:%p: decrypt err: %i verify err: %i",
             SRCNAME, __func__, this, m_decrypt_result.error().code(),
             m_verify_result.error().code());
  /* If we are called again it is the second pass */
  m_second_pass = true;

  bool has_valid_encrypted_checksum = false;
  /* Ensure that the Keys for the signatures are available
     and if it has a valid encrypted checksum. */
  for (const auto sig: m_verify_result.signatures())
    {
      TRACEPOINT;
      has_valid_encrypted_checksum = is_valid_chksum (sig);

#ifndef BUILD_TESTS
      /* For TOFU we would need to update here even when offline. */
      if (!offline)
        {
          KeyCache::instance ()->update (isig.fingerprint (), protocol);
        }
#endif
      TRACEPOINT;
    }

  if (protocol == Protocol::CMS && decrypt && !m_decrypt_result.error() &&
      !has_valid_encrypted_checksum)
    {
      log_debug ("%s:%s:%p Encrypted S/MIME without checksum. Block HTML.",
                 SRCNAME, __func__, this);
      m_block_html = true;
    }

  /* Import any application/pgp-keys attachments if the option is set. */
  if (opt.autoimport)
    {
      for (const auto &attach: get_attachments())
        {
          if (attach->get_content_type () == "application/pgp-keys")
            {
#ifndef BUILD_TESTS
              KeyCache::import_pgp_key_data (attach->get_data());
#endif
            }
        }
    }

  if (opt.enable_debug & DBG_DATA)
    {
      std::stringstream ss;
      TRACEPOINT;
      ss << m_decrypt_result << '\n' << m_verify_result;
      for (const auto sig: m_verify_result.signatures())
        {
          const auto key = sig.key();
          if (key.isNull())
            {
#ifndef BUILD_TESTS
              ss << '\n' << "Cached key:\n" << KeyCache::instance()->getByFpr(
                  sig.fingerprint(), false);
#else
              get_and_print_key_test (sig.fingerprint (), protocol);
#endif
            }
          else
            {
              ss << '\n' << key;
            }
        }
       log_debug ("Decrypt / Verify result: %s", ss.str().c_str());
    }
  else
    {
      log_debug ("%s:%s:%p Decrypt / verify done errs: %i / %i numsigs: %i.",
                 SRCNAME, __func__, this, m_decrypt_result.error().code(),
                 m_verify_result.error().code(), m_verify_result.numSignatures());
    }
  TRACEPOINT;

  if (m_outputprovider)
    {
      m_outputprovider->finalize ();
    }

  TRETURN;
}

const std::string
ParseController::get_html_body () const
{
  TSTART;
  if (m_outputprovider)
    {
      TRETURN m_outputprovider->get_html_body ();
    }
  else
    {
      TRETURN std::string();
    }
}

const std::string
ParseController::get_body () const
{
  TSTART;
  if (m_outputprovider)
    {
      TRETURN m_outputprovider->get_body ();
    }
  else
    {
      TRETURN std::string();
    }
}

const std::string
ParseController::get_body_charset() const
{
  TSTART;
  if (m_outputprovider)
    {
      TRETURN m_outputprovider->get_body_charset();
    }
  else
    {
      TRETURN std::string();
    }
}

const std::string
ParseController::get_html_charset() const
{
  TSTART;
  if (m_outputprovider)
    {
      TRETURN m_outputprovider->get_html_charset();
    }
  else
    {
      TRETURN std::string();
    }
}

std::vector<std::shared_ptr<Attachment> >
ParseController::get_attachments() const
{
  TSTART;
  if (m_outputprovider)
    {
      TRETURN m_outputprovider->get_attachments();
    }
  else
    {
      TRETURN std::vector<std::shared_ptr<Attachment> >();
    }
}

std::string
ParseController::get_protected_header (const std::string &which) const
{
  TSTART;
  if (m_outputprovider)
    {
      TRETURN m_outputprovider->get_protected_header (which);
    }
  TRETURN std::string ();
}
