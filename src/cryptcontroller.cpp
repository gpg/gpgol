/* @file cryptcontroller.cpp
 * @brief Helper to do crypto on a mail.
 *
 * Copyright (C) 2018 Intevation GmbH
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
#include "cryptcontroller.h"
#include "mail.h"
#include "mapihelp.h"
#include "mimemaker.h"

#include <gpgme++/context.h>
#include <gpgme++/signingresult.h>
#include <gpgme++/encryptionresult.h>

#ifdef HAVE_W32_SYSTEM
#include "common.h"
/* We use UTF-8 internally. */
#undef _
# define _(a) utf8_gettext (a)
#else
# define _(a) a
#endif

#include <sstream>

#define DEBUG_RESOLVER 1

static int
sink_data_write (sink_t sink, const void *data, size_t datalen)
{
  GpgME::Data *d = static_cast<GpgME::Data *>(sink->cb_data);
  d->write (data, datalen);
  return 0;
}


/** We have some C Style cruft in here as this was historically how
  GpgOL worked directly in the MAPI data objects. To reduce the regression
  risk the new object oriented way for crypto reused as much as possible
  from this.
*/
CryptController::CryptController (Mail *mail, bool encrypt, bool sign,
                                  bool doInline, GpgME::Protocol proto):
    m_mail (mail),
    m_encrypt (encrypt),
    m_sign (sign),
    m_inline (doInline),
    m_crypto_success (false),
    m_proto (proto)
{
  log_debug ("%s:%s: CryptController ctor for %p encrypt %i sign %i inline %i.",
             SRCNAME, __func__, mail, encrypt, sign, doInline);
}

CryptController::~CryptController()
{
  log_debug ("%s:%s:%p",
             SRCNAME, __func__, m_mail);
}

int
CryptController::collect_data ()
{
  /* Get the attachment info and the body.  We need to do this before
     creating the engine's filter because sending the cancel to
     the engine with nothing for the engine to process.  Will result
     in an error. This is actually a bug in our engine code but
     we better avoid triggering this bug because the engine
     sometimes hangs.  Fixme: Needs a proper fix. */


  /* Take the Body from the mail if possible. This is a fix for
     GnuPG-Bug-ID: T3614 because the body is not always properly
     updated in MAPI when sending. */
  char *body = m_mail->take_cached_plain_body ();
  if (body && !*body)
    {
      xfree (body);
      body = NULL;
    }

  LPMESSAGE message = get_oom_base_message (m_mail->item ());
  if (!message)
    {
      log_error ("%s:%s: Failed to get base message.",
                 SRCNAME, __func__);
    }

  auto att_table = mapi_create_attach_table (message, 0);
  int n_att_usable = count_usable_attachments (att_table);
  if (!n_att_usable && !body)
    {
      log_debug ("%s:%s: encrypt empty message", SRCNAME, __func__);
    }

  if (n_att_usable && m_inline)
    {
      log_debug ("%s:%s: PGP Inline not supported for attachments."
                 " Using PGP MIME",
                 SRCNAME, __func__);
      m_inline = false;
    }
  else if (m_inline)
    {
      // Inline. Use Body as input and be done.
      m_input.write (body, strlen (body));
      log_debug ("%s:%s: PGP Inline. Using cached body as input.",
                 SRCNAME, __func__);
      gpgol_release (message);
      /* Set the input buffer to start. */
      m_input.seek (0, SEEK_SET);
      return 0;
    }

  /* Set up the sink object to collect the mime structure */
  struct sink_s sinkmem;
  sink_t sink = &sinkmem;
  memset (sink, 0, sizeof *sink);
  sink->cb_data = &m_input;
  sink->writefnc = sink_data_write;

  /* Collect the mime strucutre */
  if (add_body_and_attachments (sink, message, att_table, m_mail,
                                body, n_att_usable))
    {
      log_error ("%s:%s: Collecting body and attachments failed.",
                 SRCNAME, __func__);
      gpgol_release (message);
      return -1;
    }

  /* Message is no longer needed */
  gpgol_release (message);

  /* Set the input buffer to start. */
  m_input.seek (0, SEEK_SET);
  return 0;
}

static void
release_recipient_array (char **recipients)
{
  int idx;

  if (recipients)
    {
      for (idx=0; recipients[idx]; idx++)
        xfree (recipients[idx]);
      xfree (recipients);
    }
}

int
CryptController::parse_keys (GpgME::Data &resolverOutput)
{
  // Todo: Use Data::toString
  std::istringstream ss(resolverOutput.toString());
  std::string line;

  while (std::getline (ss, line))
    {
      if (line == "cancel")
        {
          log_debug ("%s:%s: resolver canceled",
                     SRCNAME, __func__);
          return -1;
        }
      if (line == "unencrypted")
        {
          log_debug ("%s:%s: FIXME resolver wants unencrypted",
                     SRCNAME, __func__);
          return -1;
        }
      std::istringstream lss (line);

      // First is sig or enc
      std::string what;
      std::string how;
      std::string fingerprint;

      lss >> what;
      lss >> how;
      lss >> fingerprint;

      log_debug ("Data what: %s how: %s fingerprint: %s", what.c_str (), how.c_str (), fingerprint.c_str ());
    }
  return -1;
}

int
CryptController::resolve_keys ()
{
  m_recipients.clear();

  std::vector<std::string> args;

  // Collect the arguments
  char *gpg4win_dir = get_gpg4win_dir ();
  if (!gpg4win_dir)
    {
      TRACEPOINT;
      return -1;
    }
  const auto resolver = std::string (gpg4win_dir) + "\\bin\\resolver.exe";
  args.push_back (resolver);

  log_debug ("%s:%s: resolving keys with '%s'",
                 SRCNAME, __func__, resolver.c_str ());

  // We want debug output as OutputDebugString
  args.push_back (std::string ("--debug"));

  if (m_sign)
    {
      args.push_back (std::string ("--sign"));
    }
  const auto cached_sender = m_mail->get_cached_sender ();
  if (cached_sender.empty())
    {
      log_error ("%s:%s: resolve keys without sender.",
                 SRCNAME, __func__);
    }
  else
    {
      args.push_back (std::string ("--sender"));
      args.push_back (cached_sender);
    }

  if (m_encrypt)
    {
      args.push_back (std::string ("--encrypt"));
    }

  if (!opt.autoresolve)
    {
      args.push_back (std::string ("--alwaysShow"));
    }

  // Get the recipients that are cached from OOM
  char **recipients = m_mail->take_cached_recipients ();
  for (size_t i = 0; recipients && recipients[i]; i++)
    {
      args.push_back (GpgME::UserID::addrSpecFromString (recipients[i]));
    }

  release_recipient_array (recipients);

  // Convert our collected vector to c strings
  // It's a bit overhead but should be quick for such small
  // data.
  char **cargs = (char**) xmalloc (sizeof (char*) * (args.size() + 1));
  for (size_t i = 0; i < args.size(); i++)
    {
      gpgrt_asprintf (cargs + i, "%s", args[i].c_str());
    }
  cargs[args.size()] = NULL;

  // Args are prepared. Spawn the resolver.
  auto ctx = GpgME::Context::createForEngine (GpgME::SpawnEngine);

  if (!ctx)
    {
      // can't happen
      release_recipient_array (cargs);
      TRACEPOINT;
      return -1;
    }

  GpgME::Data mystdin (GpgME::Data::null), mystdout, mystderr;

#ifdef DEBUG_RESOLVER
  log_debug ("Spawning args:");
  for (size_t i = 0; cargs && cargs[i]; i++)
    {
      log_debug ("%i: '%s'", i, cargs[i]);
    }
#endif

  GpgME::Error err = ctx->spawn (cargs[0], const_cast <const char**> (cargs),
                                 mystdin, mystdout, mystderr,
                                 (GpgME::Context::SpawnFlags) (
                                  GpgME::Context::SpawnAllowSetFg |
                                  GpgME::Context::SpawnShowWindow));

#ifdef DEBUG_RESOLVER
  log_debug ("Resolver stdout:\n'%s'", mystdout.toString ().c_str ());
  log_debug ("Resolver stderr:\n'%s'", mystderr.toString ().c_str ());
#endif

  release_recipient_array (cargs);

  if (err)
    {
      log_debug ("%s:%s: Resolver spawn finished Err code: %i asString: %s",
                 SRCNAME, __func__, err.code(), err.asString());
    }

  if (parse_keys (mystdout))
    {
      log_debug ("%s:%s: Failed to parse / resolve keys.",
                 SRCNAME, __func__);
      log_debug ("Resolver stdout:\n'%s'", mystdout.toString ().c_str ());
      log_debug ("Resolver stderr:\n'%s'", mystderr.toString ().c_str ());
      return -1;
    }

  return 0;
}

int
CryptController::do_crypto ()
{
  // TODO get recipients and sender and protocol.

  log_debug ("%s:%s:",
             SRCNAME, __func__);
  auto ctx = std::shared_ptr<GpgME::Context> (GpgME::Context::createForProtocol(GpgME::OpenPGP));

  if (!ctx)
    {
      log_error ("%s:%s: Failure to create context.",
                 SRCNAME, __func__);
      return -1;
    }

  if (resolve_keys ())
    {
      log_debug ("%s:%s: Failure to resolve keys.",
                 SRCNAME, __func__);
      return -2;
    }

  if (!m_signer_key.isNull())
    {
      ctx->addSigningKey (m_signer_key);
    }

  ctx->setTextMode (true);
  ctx->setArmor (true);

  if (m_encrypt && m_sign)
    {
      const auto result_pair = ctx->signAndEncrypt (m_recipients,
                                                    m_input,
                                                    m_output,
                                                    GpgME::Context::AlwaysTrust);

      if (result_pair.first.error() || result_pair.second.error())
        {
          log_error ("%s:%s: Encrypt / Sign error %s %s.",
                     SRCNAME, __func__, result_pair.first.error().asString(),
                     result_pair.second.error().asString());
          return -1;
        }

      if (result_pair.first.error().isCanceled() || result_pair.second.error().isCanceled())
        {
          log_debug ("%s:%s: User cancled",
                     SRCNAME, __func__);
          return -2;
        }
    }
  else if (m_encrypt)
    {
      const auto result = ctx->encrypt (m_recipients, m_input,
                                        m_output,
                                        GpgME::Context::AlwaysTrust);
      if (result.error())
        {
          log_error ("%s:%s: Encryption error %s.",
                     SRCNAME, __func__, result.error().asString());
          return -1;
        }
      if (result.error().isCanceled())
        {
          log_debug ("%s:%s: User cancled",
                     SRCNAME, __func__);
          return -2;
        }
    }
  else if (m_sign)
    {
      const auto result = ctx->sign (m_input, m_output,
                                     m_inline ? GpgME::Clearsigned :
                                     GpgME::Detached);
      if (result.error())
        {
          log_error ("%s:%s: Signing error %s.",
                     SRCNAME, __func__, result.error().asString());
          return -1;
        }
      if (result.error().isCanceled())
        {
          log_debug ("%s:%s: User cancled",
                     SRCNAME, __func__);
          return -2;
        }
    }
  else
    {
      // ???
      log_error ("%s:%s: unreachable code reached.",
                 SRCNAME, __func__);
    }


  log_debug ("%s:%s: Crypto done sucessfuly.",
             SRCNAME, __func__);
  m_crypto_success = true;
  return 0;
}

static int
write_data (sink_t sink, GpgME::Data &data)
{
  if (!sink || !sink->writefnc)
    {
      return -1;
    }

  char buf[4096];
  size_t nread;
  data.seek (0, SEEK_SET);
  while ((nread = data.read (buf, 4096)) > 0)
    {
      sink->writefnc (sink, buf, nread);
    }

  return 0;
}

static int
create_sign_attach (sink_t sink, protocol_t protocol,
                    GpgME::Data &signature,
                    GpgME::Data &signedData)
{
  char boundary[BOUNDARYSIZE+1];
  char top_header[BOUNDARYSIZE+200];
  int rc = 0;

  /* Write the top header.  */
  generate_boundary (boundary);
  create_top_signing_header (top_header, sizeof top_header,
                             protocol, 1, boundary,
                             protocol == PROTOCOL_SMIME ? "sha1":"pgp-sha1");

  if ((rc = write_string (sink, top_header)))
    {
      TRACEPOINT;
      return rc;
    }

  /* Write the boundary so that it is not included in the hashing.  */
  if ((rc = write_boundary (sink, boundary, 0)))
    {
      TRACEPOINT;
      return rc;
    }

  /* Write the signed mime structure */
  if ((rc = write_data (sink, signedData)))
    {
      TRACEPOINT;
      return rc;
    }

  /* Write the signature attachment */
  if ((rc = write_boundary (sink, boundary, 0)))
    {
      TRACEPOINT;
      return rc;
    }

  if (protocol == PROTOCOL_OPENPGP)
    {
      rc = write_string (sink,
                         "Content-Type: application/pgp-signature\r\n");
    }
  else
    {
      rc = write_string (sink,
                         "Content-Transfer-Encoding: base64\r\n"
                         "Content-Type: application/pkcs7-signature\r\n");
      /* rc = write_string (sink, */
      /*                    "Content-Type: application/x-pkcs7-signature\r\n" */
      /*                    "\tname=\"smime.p7s\"\r\n" */
      /*                    "Content-Transfer-Encoding: base64\r\n" */
      /*                    "Content-Disposition: attachment;\r\n" */
      /*                    "\tfilename=\"smime.p7s\"\r\n"); */

    }

  if (rc)
    {
      TRACEPOINT;
      return rc;
    }

  if ((rc = write_string (sink, "\r\n")))
    {
      TRACEPOINT;
      return rc;
    }

  // Write the signature data
  if ((rc = write_data (sink, signature)))
    {
      TRACEPOINT;
      return rc;
    }

  // Add an extra linefeed with should not harm.
  if ((rc = write_string (sink, "\r\n")))
    {
      TRACEPOINT;
      return rc;
    }

  /* Write the final boundary.  */
  if ((rc = write_boundary (sink, boundary, 1)))
    {
      TRACEPOINT;
      return rc;
    }

  return rc;
}

static int
create_encrypt_attach (sink_t sink, protocol_t protocol,
                       GpgME::Data &encryptedData)
{
  char boundary[BOUNDARYSIZE+1];
  int rc = create_top_encryption_header (sink, protocol, boundary,
                                         false);
  // From here on use goto failure pattern.
  if (rc)
    {
      log_error ("%s:%s: Failed to create top header.",
                 SRCNAME, __func__);
      return rc;
    }

  rc = write_data (sink, encryptedData);
  if (rc)
    {
      log_error ("%s:%s: Failed to create top header.",
                 SRCNAME, __func__);
      return rc;
    }

  /* Write the final boundary (for OpenPGP) and finish the attachment.  */
  if (*boundary && (rc = write_boundary (sink, boundary, 1)))
    {
      log_error ("%s:%s: Failed to write boundary.",
                 SRCNAME, __func__);
    }
  return rc;
}

int
CryptController::update_mail_mapi ()
{
  log_debug ("%s:%s:", SRCNAME, __func__);

  if (m_inline)
    {
      // Nothing to do for inline.
      return 0;
    }

  LPMESSAGE message = get_oom_base_message (m_mail->item());
  if (!message)
    {
      log_error ("%s:%s: Failed to obtain message.",
                 SRCNAME, __func__);
      return -1;
    }

  mapi_attach_item_t *att_table = mapi_create_attach_table (message, 0);

  // Set up the sink object for our MSOXSMIME attachment.
  struct sink_s sinkmem;
  sink_t sink = &sinkmem;
  memset (sink, 0, sizeof *sink);
  sink->cb_data = &m_input;
  sink->writefnc = sink_data_write;

  LPATTACH attach = create_mapi_attachment (message, sink);
  if (!attach)
    {
      log_error ("%s:%s: Failed to create moss attach.",
                 SRCNAME, __func__);
      gpgol_release (message);
      return -1;
    }

  protocol_t protocol = m_proto == GpgME::CMS ?
                                   PROTOCOL_SMIME :
                                   PROTOCOL_OPENPGP;
  int rc = 0;
  if (m_sign && m_encrypt)
    {
      // FIXME we need some doubling here for S/MIME.
      rc = create_encrypt_attach (sink, protocol, m_output);
    }
  else if (m_encrypt)
    {
      rc = create_encrypt_attach (sink, protocol, m_output);
    }
  else if (m_sign)
    {
      rc = create_sign_attach (sink, protocol, m_output, m_input);
    }

  // Close our attachment
  if (!rc)
    {
      rc = close_mapi_attachment (&attach, sink);
    }

  // Set message class etc.
  if (!rc)
    {
      rc = finalize_message (message, att_table, protocol, m_encrypt ? 1 : 0,
                             false);
    }

  // only on error.
  if (rc)
    {
      cancel_mapi_attachment (&attach, sink);
    }

  // cleanup
  mapi_release_attach_table (att_table);
  gpgol_release (attach);
  gpgol_release (message);

  return rc;
}

std::string
CryptController::get_inline_data ()
{
  std::string ret;
  if (!m_inline)
    {
      return ret;
    }
  m_output.seek (0, SEEK_SET);
  char buf[4096];
  size_t nread;
  while ((nread = m_output.read (buf, 4096)) > 0)
    {
      ret += std::string (buf, nread);
    }
  return ret;
}
