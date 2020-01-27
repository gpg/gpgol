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
#include "cpphelp.h"
#include "cryptcontroller.h"
#include "mail.h"
#include "mapihelp.h"
#include "mimemaker.h"
#include "wks-helper.h"
#include "overlay.h"
#include "keycache.h"
#include "mymapitags.h"
#include "recipient.h"

#include <gpgme++/context.h>
#include <gpgme++/signingresult.h>
#include <gpgme++/encryptionresult.h>

#include "common.h"

#include <sstream>

static int
sink_data_write (sink_t sink, const void *data, size_t datalen)
{
  GpgME::Data *d = static_cast<GpgME::Data *>(sink->cb_data);
  d->write (data, datalen);
  return 0;
}

static int
create_sign_attach (sink_t sink, protocol_t protocol,
                    GpgME::Data &signature,
                    GpgME::Data &signedData,
                    const char *micalg);

/** We have some C Style cruft in here as this was historically how
  GpgOL worked directly in the MAPI data objects. To reduce the regression
  risk the new object oriented way for crypto reused as much as possible
  from this.
*/
CryptController::CryptController (Mail *mail, bool encrypt, bool sign,
                                  GpgME::Protocol proto):
    m_mail (mail),
    m_encrypt (encrypt),
    m_sign (sign),
    m_crypto_success (false),
    m_proto (proto)
{
  TSTART;
  memdbg_ctor ("CryptController");
  log_debug ("%s:%s: CryptController ctor for %p encrypt %i sign %i inline %i.",
             SRCNAME, __func__, mail, encrypt, sign, mail->getDoPGPInline ());
  m_recipients = mail->getCachedRecipients ();
  TRETURN;
}

CryptController::~CryptController()
{
  TSTART;
  memdbg_dtor ("CryptController");
  log_debug ("%s:%s:%p",
             SRCNAME, __func__, m_mail);
  TRETURN;
}

int
CryptController::collect_data ()
{
  TSTART;
  /* Get the attachment info and the body.  We need to do this before
     creating the engine's filter because sending the cancel to
     the engine with nothing for the engine to process.  Will result
     in an error. This is actually a bug in our engine code but
     we better avoid triggering this bug because the engine
     sometimes hangs.  Fixme: Needs a proper fix. */


  /* Take the Body from the mail if possible. This is a fix for
     GnuPG-Bug-ID: T3614 because the body is not always properly
     updated in MAPI when sending. */
  char *body = m_mail->takeCachedPlainBody ();
  if (body && !*body)
    {
      xfree (body);
      body = nullptr;
    }

  LPMESSAGE message = m_mail->isCryptoMail() ?
                      get_oom_base_message (m_mail->item ()) :
                      get_oom_message (m_mail->item ());
  if (!message)
    {
      log_error ("%s:%s: Failed to get message.",
                 SRCNAME, __func__);
    }

  auto att_table = mapi_create_attach_table (message, 0);
  int n_att_usable = count_usable_attachments (att_table);
  if (!n_att_usable && !body)
    {
      if (!m_mail->isDraftEncrypt())
        {
          gpgol_message_box (m_mail->getWindow (),
                             utf8_gettext ("Can't encrypt / sign an empty message."),
                             utf8_gettext ("GpgOL"), MB_OK);
        }
      gpgol_release (message);
      mapi_release_attach_table (att_table);
      xfree (body);
      TRETURN -1;
    }

  bool do_inline = m_mail->getDoPGPInline ();

  if (n_att_usable && do_inline)
    {
      log_debug ("%s:%s: PGP Inline not supported for attachments."
                 " Using PGP MIME",
                 SRCNAME, __func__);
      do_inline = false;
      m_mail->setDoPGPInline (false);
    }
  else if (do_inline)
    {
      /* Inline. Use Body as input.
        We need to collect also our mime structure for S/MIME
        as we don't know yet if we are S/MIME or OpenPGP */
      m_bodyInput.write (body, strlen (body));
      log_debug ("%s:%s: Inline. Caching body.",
                 SRCNAME, __func__);
      /* Set the input buffer to start. */
      m_bodyInput.seek (0, SEEK_SET);
    }

  /* Set up the sink object to collect the mime structure */
  struct sink_s sinkmem;
  sink_t sink = &sinkmem;
  memset (sink, 0, sizeof *sink);
  sink->cb_data = &m_input;
  sink->writefnc = sink_data_write;

  /* Collect the mime strucutre */
  int err = add_body_and_attachments (sink, message, att_table, m_mail,
                                      body, n_att_usable);
  xfree (body);

  if (err)
    {
      log_error ("%s:%s: Collecting body and attachments failed.",
                 SRCNAME, __func__);
      gpgol_release (message);
      mapi_release_attach_table (att_table);
      TRETURN -1;
    }

  /* Message is no longer needed */
  gpgol_release (message);

  mapi_release_attach_table (att_table);

  /* Set the input buffer to start. */
  m_input.seek (0, SEEK_SET);
  TRETURN 0;
}

int
CryptController::lookup_fingerprints (const std::string &sigFpr,
                                      const std::vector<std::string> recpFprs)
{
  TSTART;
  auto ctx = std::shared_ptr<GpgME::Context> (GpgME::Context::createForProtocol (m_proto));

  if (!ctx)
    {
      log_error ("%s:%s: failed to create context with protocol '%s'",
                 SRCNAME, __func__,
                 m_proto == GpgME::CMS ? "smime" :
                 m_proto == GpgME::OpenPGP ? "openpgp" :
                 "unknown");
      TRETURN -1;
    }

  ctx->setKeyListMode (GpgME::Local);
  GpgME::Error err;

  if (!sigFpr.empty()) {
      m_signer_key = ctx->key (sigFpr.c_str (), err, true);
      if (err || m_signer_key.isNull ()) {
          log_error ("%s:%s: failed to lookup key for '%s' with protocol '%s'",
                     SRCNAME, __func__, anonstr (sigFpr.c_str ()),
                     m_proto == GpgME::CMS ? "smime" :
                     m_proto == GpgME::OpenPGP ? "openpgp" :
                     "unknown");
          TRETURN -1;
      }
      // reset context
      ctx = std::shared_ptr<GpgME::Context> (GpgME::Context::createForProtocol (m_proto));
      ctx->setKeyListMode (GpgME::Local);
  }

  if (!recpFprs.size()) {
     TRETURN 0;
  }

  // Convert recipient fingerprints
  char **cRecps = vector_to_cArray (recpFprs);

  err = ctx->startKeyListing (const_cast<const char **> (cRecps));

  if (err) {
      log_error ("%s:%s: failed to start recipient keylisting",
                 SRCNAME, __func__);
      release_cArray (cRecps);
      TRETURN -1;
  }

  do {
      m_enc_keys.push_back(ctx->nextKey(err));
  } while (!err);

  m_enc_keys.pop_back();

  release_cArray (cRecps);

  TRETURN 0;
}


int
CryptController::parse_output (GpgME::Data &resolverOutput)
{
  TSTART;
  // Todo: Use Data::toString
  std::istringstream ss(resolverOutput.toString());
  std::string line;

  std::string sigFpr;
  std::vector<std::string> recpFprs;
  while (std::getline (ss, line))
    {
      rtrim (line);
      if (line == "cancel")
        {
          log_debug ("%s:%s: resolver canceled",
                     SRCNAME, __func__);
          TRETURN -2;
        }
      if (line == "unencrypted")
        {
          log_debug ("%s:%s: FIXME resolver wants unencrypted",
                     SRCNAME, __func__);
          TRETURN -1;
        }
      std::istringstream lss (line);

      // First is sig or enc
      std::string what;
      std::string how;
      std::string fingerprint;

      std::getline (lss, what, ':');
      std::getline (lss, how, ':');
      std::getline (lss, fingerprint, ':');

      if (m_proto == GpgME::UnknownProtocol)
        {
          m_proto = (how == "smime") ? GpgME::CMS : GpgME::OpenPGP;
        }

      if (what == "sig")
        {
          if (!sigFpr.empty ())
            {
              log_error ("%s:%s: multiple signing keys not supported",
                         SRCNAME, __func__);

            }
          sigFpr = fingerprint;
          continue;
        }
      if (what == "enc")
        {
          recpFprs.push_back (fingerprint);
        }
    }

  if (m_sign && sigFpr.empty())
    {
      log_error ("%s:%s: Sign requested but no signing fingerprint - sending unsigned",
                 SRCNAME, __func__);
      m_sign = false;
    }
  if (m_encrypt && !recpFprs.size())
    {
      log_error ("%s:%s: Encrypt requested but no recipient fingerprints",
                 SRCNAME, __func__);
      gpgol_message_box (m_mail->getWindow (),
                         utf8_gettext ("No recipients for encryption selected."),
                         _("GpgOL"), MB_OK);
      TRETURN -2;
    }

  TRETURN lookup_fingerprints (sigFpr, recpFprs);
}

static bool
resolve_through_protocol (const GpgME::Protocol &proto, bool sign,
                          bool encrypt, const std::string &sender,
                          const std::vector<std::string> &recps,
                          std::vector<GpgME::Key> &r_keys,
                          GpgME::Key &r_sig)
{
  TSTART;
  bool sig_ok = true;
  bool enc_ok = true;

  const auto cache = KeyCache::instance();

  if (encrypt)
    {
      r_keys = cache->getEncryptionKeys(recps, proto);
      enc_ok = !r_keys.empty();
    }
  if (sign && enc_ok)
    {
      r_sig = cache->getSigningKey (sender.c_str (), proto);
      sig_ok = !r_sig.isNull();
    }
  TRETURN sig_ok && enc_ok;
}

int
CryptController::resolve_keys_cached()
{
  TSTART;
  // Prepare variables
  const auto cached_sender = m_mail->getSender ();
  std::vector <std::string> recps;
  for (const auto &recp: m_recipients)
    {
      recps.push_back (recp.mbox());
    }

  if (m_encrypt)
    {
      recps.push_back (cached_sender);
    }

  bool resolved = false;
  if (opt.enable_smime && opt.prefer_smime)
    {
      resolved = resolve_through_protocol (GpgME::CMS, m_sign, m_encrypt,
                                           cached_sender, recps, m_enc_keys,
                                           m_signer_key);
      if (resolved)
        {
          log_debug ("%s:%s: Resolved with CMS due to preference.",
                     SRCNAME, __func__);
          m_proto = GpgME::CMS;
        }
    }
  if (!resolved)
    {
      resolved = resolve_through_protocol (GpgME::OpenPGP, m_sign, m_encrypt,
                                           cached_sender, recps, m_enc_keys,
                                           m_signer_key);
      if (resolved)
        {
          log_debug ("%s:%s: Resolved with OpenPGP.",
                     SRCNAME, __func__);
          m_proto = GpgME::OpenPGP;
        }
    }
  if (!resolved && (opt.enable_smime && !opt.prefer_smime))
    {
      resolved = resolve_through_protocol (GpgME::CMS, m_sign, m_encrypt,
                                           cached_sender, recps, m_enc_keys,
                                           m_signer_key);
      if (resolved)
        {
          log_debug ("%s:%s: Resolved with CMS as fallback.",
                     SRCNAME, __func__);
          m_proto = GpgME::CMS;
        }
    }

  if (!resolved)
    {
      log_debug ("%s:%s: Failed to resolve through cache",
                 SRCNAME, __func__);
      m_enc_keys.clear();
      m_signer_key = GpgME::Key();
      m_proto = GpgME::UnknownProtocol;
      TRETURN 1;
    }

  if (!m_enc_keys.empty())
    {
      log_debug ("%s:%s: Encrypting with protocol %s to:",
                 SRCNAME, __func__, to_cstr (m_proto));
    }
  for (const auto &key: m_enc_keys)
    {
      log_debug ("%s", anonstr (key.primaryFingerprint ()));
    }
  if (!m_signer_key.isNull())
    {
      log_debug ("%s:%s: Signing key: %s:%s",
                 SRCNAME, __func__, anonstr (m_signer_key.primaryFingerprint ()),
                 to_cstr (m_signer_key.protocol()));
    }
  TRETURN 0;
}

int
CryptController::resolve_keys ()
{
  TSTART;

  m_enc_keys.clear();

  if (m_mail->isDraftEncrypt() && opt.draft_key)
    {
      const auto key = KeyCache::instance()->getByFpr (opt.draft_key);
      if (key.isNull())
        {
          const char *buf = utf8_gettext ("Failed to encrypt draft.\n\n"
                                          "The configured encryption key for drafts "
                                          "could not be found.\n"
                                          "Please check your configuration or "
                                          "turn off draft encryption in the settings.");
          gpgol_message_box (get_active_hwnd (),
                             buf,
                             _("GpgOL"), MB_OK);
          TRETURN -1;
        }
      log_debug ("%s:%s: resolved draft encryption key protocol is: %s",
                 SRCNAME, __func__, to_cstr (key.protocol()));
      m_proto = key.protocol ();
      m_enc_keys.push_back (key);
      TRETURN 0;
    }

  if (!m_recipients.size())
    {
      /* Should not happen. But we add it for better bug reports. */
      const char *bugmsg = utf8_gettext ("Operation failed.\n\n"
              "This is usually caused by a bug in GpgOL or an error in your setup.\n"
              "Please see https://www.gpg4win.org/reporting-bugs.html "
              "or ask your Administrator for support.");
      char *buf;
      gpgrt_asprintf (&buf, "Failed to resolve recipients.\n\n%s\n", bugmsg);
      memdbg_alloc (buf);
      gpgol_message_box (get_active_hwnd (),
                         buf,
                         _("GpgOL"), MB_OK);
      xfree(buf);
      TRETURN -1;
    }

  if (opt.autoresolve && !opt.alwaysShowApproval && !resolve_keys_cached ())
    {
      log_debug ("%s:%s: resolved keys through the cache",
                 SRCNAME, __func__);
      start_crypto_overlay();
      TRETURN 0;
    }

  std::vector<std::string> args;

  // Collect the arguments
  char *gpg4win_dir = get_gpg4win_dir ();
  if (!gpg4win_dir)
    {
      TRACEPOINT;
      TRETURN -1;
    }
  const auto resolver = std::string (gpg4win_dir) + "\\bin\\resolver.exe";
  args.push_back (resolver);

  log_debug ("%s:%s: resolving keys with '%s'",
             SRCNAME, __func__, resolver.c_str ());

  // We want debug output as OutputDebugString
  args.push_back (std::string ("--debug"));

  // Yes passing it as int is ok.
  auto wnd = m_mail->getWindow ();
  if (wnd)
    {
      // Pass the handle of the active window for raise / overlay.
      args.push_back (std::string ("--hwnd"));
      args.push_back (std::to_string ((int) (intptr_t) wnd));
    }

  // Set the overlay caption
  args.push_back (std::string ("--overlayText"));
  if (m_encrypt)
    {
      args.push_back (std::string (utf8_gettext ("Resolving recipients...")));
    }
  else if (m_sign)
    {
      args.push_back (std::string (utf8_gettext ("Resolving signers...")));
    }

  if (!opt.enable_smime)
    {
      args.push_back (std::string ("--protocol"));
      args.push_back (std::string ("pgp"));
    }

  if (m_sign)
    {
      args.push_back (std::string ("--sign"));
    }
  const auto cached_sender = m_mail->getSender ();
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

  if (!opt.autoresolve || opt.alwaysShowApproval)
    {
      args.push_back (std::string ("--alwaysShow"));
    }

  if (opt.prefer_smime)
    {
      args.push_back (std::string ("--preferred-protocol"));
      args.push_back (std::string ("cms"));
    }

  args.push_back (std::string ("--lang"));
  args.push_back (std::string (gettext_localename ()));

  bool has_smime_override = false;

  if (m_encrypt)
    {
      args.push_back (std::string ("--encrypt"));
      // Get the recipients that are cached from OOM
      for (const auto &recp: m_recipients)
        {
          const auto mbox = recp.mbox ();
          auto overrides =
            KeyCache::instance ()->getOverrides (mbox, GpgME::OpenPGP);
          const auto cms_overrides =
            KeyCache::instance ()->getOverrides (mbox, GpgME::CMS);

          overrides.insert(overrides.end(), cms_overrides.begin(),
                           cms_overrides.end());

          if (overrides.size())
            {
              std::string overrideStr = mbox + ":";
              for (const auto &key: overrides)
                {
                  if (key.isNull())
                    {
                      TRACEPOINT;
                      continue;
                    }
                  has_smime_override |= key.protocol() == GpgME::CMS;
                  overrideStr += key.primaryFingerprint();
                  overrideStr += ",";
                }
              overrideStr.erase(overrideStr.size() - 1, 1);
              args.push_back (std::string ("-o"));
              args.push_back (overrideStr);
            }
          args.push_back (mbox);
        }
    }

  if (!opt.prefer_smime && has_smime_override)
    {
      /* Prefer S/MIME if there was an S/MIME override */
      args.push_back (std::string ("--preferred-protocol"));
      args.push_back (std::string ("cms"));
    }

  // Args are prepared. Spawn the resolver.
  auto ctx = GpgME::Context::createForEngine (GpgME::SpawnEngine);
  if (!ctx)
    {
      // can't happen
      TRACEPOINT;
      TRETURN -1;
    }


  // Convert our collected vector to c strings
  // It's a bit overhead but should be quick for such small
  // data.
  char **cargs = vector_to_cArray (args);
  log_data ("%s:%s: Spawn args:",
            SRCNAME, __func__);
  for (size_t i = 0; cargs && cargs[i]; i++)
    {
      log_data (SIZE_T_FORMAT ": '%s'", i, cargs[i]);
    }

  GpgME::Data mystdin (GpgME::Data::null), mystdout, mystderr;
  GpgME::Error err = ctx->spawn (cargs[0], const_cast <const char**> (cargs),
                                 mystdin, mystdout, mystderr,
                                 (GpgME::Context::SpawnFlags) (
                                  GpgME::Context::SpawnAllowSetFg |
                                  GpgME::Context::SpawnShowWindow));
  // Somehow Qt messes up which window to bring back to front.
  // So we do it manually.
  bring_to_front (wnd);

  // We need to create an overlay while encrypting as pinentry can take a while
  start_crypto_overlay();

  log_data ("Resolver stdout:\n'%s'", mystdout.toString ().c_str ());
  log_data ("Resolver stderr:\n'%s'", mystderr.toString ().c_str ());

  release_cArray (cargs);

  if (err)
    {
      log_debug ("%s:%s: Resolver spawn finished Err code: %i asString: %s",
                 SRCNAME, __func__, err.code(), err.asString());
    }

  int ret = parse_output (mystdout);
  if (ret == -1)
    {
      log_debug ("%s:%s: Failed to parse / resolve keys.",
                 SRCNAME, __func__);
      log_data ("Resolver stdout:\n'%s'", mystdout.toString ().c_str ());
      log_data ("Resolver stderr:\n'%s'", mystderr.toString ().c_str ());
      TRETURN -1;
    }

  TRETURN ret;
}

int
CryptController::do_crypto (GpgME::Error &err, std::string &r_diag)
{
  TSTART;
  log_debug ("%s:%s",
             SRCNAME, __func__);

  if (m_mail->isDraftEncrypt ())
    {
      log_debug ("%s:%s Disabling sign because of draft encrypt",
                 SRCNAME, __func__);
      m_sign = false;
    }

  /* Start a WKS check if necessary. */
  WKSHelper::instance()->start_check (m_mail->getSender ());

  int ret = resolve_keys ();

  /* If we need to send multiple emails we jump back from
     here into the main event loop. Copy the mail object
     and send it out mutiple times. */
  if (ret == -1)
    {
      //error
      log_debug ("%s:%s: Failure to resolve keys.",
                 SRCNAME, __func__);
      TRETURN -1;
    }
  if (ret == -2)
    {
      // Cancel
      TRETURN -2;
    }
  bool do_inline = m_mail->getDoPGPInline ();

  if (m_proto == GpgME::CMS && do_inline)
    {
      log_debug ("%s:%s: Inline for S/MIME not supported. Switching to mime.",
                 SRCNAME, __func__);
      do_inline = false;
      m_mail->setDoPGPInline (false);
      m_bodyInput = GpgME::Data(GpgME::Data::null);
    }

  auto ctx = GpgME::Context::create(m_proto);

  if (!ctx)
    {
      log_error ("%s:%s: Failure to create context.",
                 SRCNAME, __func__);
      gpgol_message_box (m_mail->getWindow (),
                         "Failure to create context.",
                         utf8_gettext ("GpgOL"), MB_OK);
      TRETURN -1;
    }
  if (!m_signer_key.isNull())
    {
      ctx->addSigningKey (m_signer_key);
    }

  ctx->setTextMode (m_proto == GpgME::OpenPGP);
  ctx->setArmor (m_proto == GpgME::OpenPGP);

  if (m_encrypt && m_sign && do_inline)
    {
      // Sign encrypt combined
      const auto result_pair = ctx->signAndEncrypt (m_enc_keys,
                                                    do_inline ? m_bodyInput : m_input,
                                                    m_output,
                                                    GpgME::Context::AlwaysTrust);
      const auto err1 = result_pair.first.error();
      const auto err2 = result_pair.second.error();

      if (err1 || err2)
        {
          log_error ("%s:%s: Encrypt / Sign error %s %s.",
                     SRCNAME, __func__, result_pair.first.error().asString(),
                     result_pair.second.error().asString());
          err = err1 ? err1 : err2;
          GpgME::Data log;
          const auto err3 = ctx->getAuditLog (log,
                                              GpgME::Context::DiagnosticAuditLog);
          if (!err3)
            {
              r_diag = log.toString();
            }
          TRETURN -1;
        }

      if (err1.isCanceled() || err2.isCanceled())
        {
          err = err1.isCanceled() ? err1 : err2;
          log_debug ("%s:%s: User cancled",
                     SRCNAME, __func__);
          TRETURN -2;
        }
    }
  else if (m_encrypt && m_sign)
    {
      // First sign then encrypt
      const auto sigResult = ctx->sign (m_input, m_output,
                                        GpgME::Detached);
      err = sigResult.error();
      if (err)
        {
          log_error ("%s:%s: Signing error %s.",
                     SRCNAME, __func__, sigResult.error().asString());
          GpgME::Data log;
          const auto err3 = ctx->getAuditLog (log,
                                              GpgME::Context::DiagnosticAuditLog);
          if (!err3)
            {
              r_diag = log.toString();
            }
          TRETURN -1;
        }
      if (err.isCanceled())
        {
          log_debug ("%s:%s: User cancled",
                     SRCNAME, __func__);
          TRETURN -2;
        }
      parse_micalg (sigResult);

      // We now have plaintext in m_input
      // The detached signature in m_output

      // Set up the sink object to construct the multipart/signed
      GpgME::Data multipart;
      struct sink_s sinkmem;
      sink_t sink = &sinkmem;
      memset (sink, 0, sizeof *sink);
      sink->cb_data = &multipart;
      sink->writefnc = sink_data_write;

      if (create_sign_attach (sink,
                              m_proto == GpgME::CMS ?
                                         PROTOCOL_SMIME : PROTOCOL_OPENPGP,
                              m_output, m_input, m_micalg.c_str ()))
        {
          TRACEPOINT;
          TRETURN -1;
        }

      // Now we have the multipart throw away the rest.
      m_output = GpgME::Data ();
      m_input = GpgME::Data ();
      multipart.seek (0, SEEK_SET);
      const auto encResult = ctx->encrypt (m_enc_keys, multipart,
                                           m_output,
                                           GpgME::Context::AlwaysTrust);
      err = encResult.error();
      if (err)
        {
          log_error ("%s:%s: Encryption error %s.",
                     SRCNAME, __func__, err.asString());
          GpgME::Data log;
          const auto err3 = ctx->getAuditLog (log,
                                              GpgME::Context::DiagnosticAuditLog);
          if (!err3)
            {
              r_diag = log.toString();
            }
          TRETURN -1;
        }
      if (err.isCanceled())
        {
          log_debug ("%s:%s: User cancled",
                     SRCNAME, __func__);
          TRETURN -2;
        }
      // Now we have encrypted output just treat it like encrypted.
    }
  else if (m_encrypt)
    {
      const auto result = ctx->encrypt (m_enc_keys, do_inline ? m_bodyInput : m_input,
                                        m_output,
                                        GpgME::Context::AlwaysTrust);
      err = result.error();
      if (err)
        {
          log_error ("%s:%s: Encryption error %s.",
                     SRCNAME, __func__, err.asString());
          GpgME::Data log;
          const auto err3 = ctx->getAuditLog (log,
                                              GpgME::Context::DiagnosticAuditLog);
          if (!err3)
            {
              r_diag = log.toString();
            }
          TRETURN -1;
        }
      if (err.isCanceled())
        {
          log_debug ("%s:%s: User cancled",
                     SRCNAME, __func__);
          TRETURN -2;
        }
    }
  else if (m_sign)
    {
      const auto result = ctx->sign (do_inline ? m_bodyInput : m_input, m_output,
                                     do_inline ? GpgME::Clearsigned :
                                     GpgME::Detached);
      err = result.error();
      if (err)
        {
          log_error ("%s:%s: Signing error %s.",
                     SRCNAME, __func__, err.asString());
          GpgME::Data log;
          const auto err3 = ctx->getAuditLog (log,
                                              GpgME::Context::DiagnosticAuditLog);
          if (!err3)
            {
              r_diag = log.toString();
            }
          TRETURN -1;
        }
      if (err.isCanceled())
        {
          log_debug ("%s:%s: User cancled",
                     SRCNAME, __func__);
          TRETURN -2;
        }
      parse_micalg (result);
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

  TRETURN 0;
}

static int
write_data (sink_t sink, GpgME::Data &data)
{
  TSTART;
  if (!sink || !sink->writefnc)
    {
      TRETURN -1;
    }

  char buf[4096];
  size_t nread;
  data.seek (0, SEEK_SET);
  while ((nread = data.read (buf, 4096)) > 0)
    {
      sink->writefnc (sink, buf, nread);
    }

  TRETURN 0;
}

int
create_sign_attach (sink_t sink, protocol_t protocol,
                    GpgME::Data &signature,
                    GpgME::Data &signedData,
                    const char *micalg)
{
  TSTART;
  char boundary[BOUNDARYSIZE+1];
  char top_header[BOUNDARYSIZE+200];
  int rc = 0;

  /* Write the top header.  */
  generate_boundary (boundary);
  create_top_signing_header (top_header, sizeof top_header,
                             protocol, 1, boundary,
                             micalg);

  if ((rc = write_string (sink, top_header)))
    {
      TRACEPOINT;
      TRETURN rc;
    }

  /* Write the boundary so that it is not included in the hashing.  */
  if ((rc = write_boundary (sink, boundary, 0)))
    {
      TRACEPOINT;
      TRETURN rc;
    }

  /* Write the signed mime structure */
  if ((rc = write_data (sink, signedData)))
    {
      TRACEPOINT;
      TRETURN rc;
    }

  /* Write the signature attachment */
  if ((rc = write_boundary (sink, boundary, 0)))
    {
      TRACEPOINT;
      TRETURN rc;
    }

  if (protocol == PROTOCOL_OPENPGP)
    {
      rc = write_string (sink,
                         "Content-Type: application/pgp-signature;\r\n"
                         "\tname=\"" OPENPGP_SIG_NAME "\"\r\n"
                         "Content-Transfer-Encoding: 7Bit\r\n");
    }
  else
    {
      rc = write_string (sink,
                         "Content-Transfer-Encoding: base64\r\n"
                         "Content-Type: application/pkcs7-signature\r\n"
                         "Content-Disposition: inline;\r\n"
                         "\tfilename=\"" SMIME_SIG_NAME "\"\r\n");
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
      TRETURN rc;
    }

  if ((rc = write_string (sink, "\r\n")))
    {
      TRACEPOINT;
      TRETURN rc;
    }

  // Write the signature data
  if (protocol == PROTOCOL_SMIME)
    {
      const std::string sigStr = signature.toString();
      if ((rc = write_b64 (sink, (const void *) sigStr.c_str (), sigStr.size())))
        {
          TRACEPOINT;
          TRETURN rc;
        }
    }
  else if ((rc = write_data (sink, signature)))
    {
      TRACEPOINT;
      TRETURN rc;
    }

  // Add an extra linefeed with should not harm.
  if ((rc = write_string (sink, "\r\n")))
    {
      TRACEPOINT;
      TRETURN rc;
    }

  /* Write the final boundary.  */
  if ((rc = write_boundary (sink, boundary, 1)))
    {
      TRACEPOINT;
      TRETURN rc;
    }

  TRETURN rc;
}

static int
create_encrypt_attach (sink_t sink, protocol_t protocol,
                       GpgME::Data &encryptedData,
                       int exchange_major_version)
{
  TSTART;
  char boundary[BOUNDARYSIZE+1];
  int rc = create_top_encryption_header (sink, protocol, boundary,
                                         false, exchange_major_version);
  // From here on use goto failure pattern.
  if (rc)
    {
      log_error ("%s:%s: Failed to create top header.",
                 SRCNAME, __func__);
      TRETURN rc;
    }

  if (protocol == PROTOCOL_OPENPGP ||
      exchange_major_version >= 15)
    {
      // With exchange 2016 we have to construct S/MIME
      // differently and write the raw data here.
      rc = write_data (sink, encryptedData);
    }
  else
    {
      const auto encStr = encryptedData.toString();
      rc = write_b64 (sink, encStr.c_str(), encStr.size());
    }

  if (rc)
    {
      log_error ("%s:%s: Failed to create top header.",
                 SRCNAME, __func__);
      TRETURN rc;
    }

  /* Write the final boundary (for OpenPGP) and finish the attachment.  */
  if (*boundary && (rc = write_boundary (sink, boundary, 1)))
    {
      log_error ("%s:%s: Failed to write boundary.",
                 SRCNAME, __func__);
    }
  TRETURN rc;
}

int
CryptController::update_mail_mapi ()
{
  TSTART;
  log_debug ("%s:%s", SRCNAME, __func__);

  LPMESSAGE message = get_oom_base_message (m_mail->item());
  if (!message)
    {
      log_error ("%s:%s: Failed to obtain message.",
                 SRCNAME, __func__);
      TRETURN -1;
    }

  if (m_mail->getDoPGPInline ())
    {
      // Nothing to do for inline.
      log_debug ("%s:%s: Inline mail. Setting encoding.",
                 SRCNAME, __func__);

      SPropValue prop;
      prop.ulPropTag = PR_INTERNET_CPID;
      prop.Value.l = 65001;
      if (HrSetOneProp (message, &prop))
        {
          log_error ("%s:%s: Failed to set CPID mapiprop.",
                     SRCNAME, __func__);
        }

      TRETURN 0;
    }

  mapi_attach_item_t *att_table = mapi_create_attach_table (message, 0);

  /* When we forward e.g. a crypto mail we have sent the message
     has a MOSSTEMPL. We need to remove that. T4321 */
  for (ULONG pos=0; att_table && !att_table[pos].end_of_table; pos++)
    {
      if (att_table[pos].attach_type == ATTACHTYPE_MOSSTEMPL)
        {
          log_debug ("%s:%s: Found existing moss attachment at "
                     "pos %i removing it.", SRCNAME, __func__,
                     att_table[pos].mapipos);
          if (message->DeleteAttach (att_table[pos].mapipos, 0,
                                     nullptr, 0) != S_OK)
            {
              log_error ("%s:%s: Failed to remove attachment.",
                         SRCNAME, __func__);
            }

        }
    }

  // Set up the sink object for our MSOXSMIME attachment.
  struct sink_s sinkmem;
  sink_t sink = &sinkmem;
  memset (sink, 0, sizeof *sink);
  sink->cb_data = &m_input;
  sink->writefnc = sink_data_write;

  // For S/MIME encrypted mails we have to use the application/pkcs7-mime
  // content type. Otherwise newer (2016) exchange servers will throw
  // an M2MCVT.StorageError.Exeption (See GnuPG-Bug-Id: T3853 )

  // This means that the conversion / build of the mime structure also
  // happens differently.
  int exchange_major_version = get_ex_major_version_for_addr (
                                        m_mail->getSender ().c_str ());

  std::string overrideMimeTag;
  if (m_proto == GpgME::CMS && m_encrypt && exchange_major_version >= 15)
    {
      log_debug ("%s:%s: CMS Encrypt with Exchange %i activating alternative.",
                 SRCNAME, __func__, exchange_major_version);
      overrideMimeTag = "application/pkcs7-mime";
    }

  LPATTACH attach = create_mapi_attachment (message, sink,
                                            overrideMimeTag.empty() ? nullptr :
                                            overrideMimeTag.c_str());
  if (!attach)
    {
      log_error ("%s:%s: Failed to create moss attach.",
                 SRCNAME, __func__);
      gpgol_release (message);
      TRETURN -1;
    }

  protocol_t protocol = m_proto == GpgME::CMS ?
                                   PROTOCOL_SMIME :
                                   PROTOCOL_OPENPGP;

  int rc = 0;
  /* Do we have override MIME ? */
  const auto overrideMime = m_mail->get_override_mime_data ();
  if (!overrideMime.empty())
    {
      rc = write_string (sink, overrideMime.c_str ());
    }
  else if (m_sign && m_encrypt)
    {
      rc = create_encrypt_attach (sink, protocol, m_output, exchange_major_version);
    }
  else if (m_encrypt)
    {
      rc = create_encrypt_attach (sink, protocol, m_output, exchange_major_version);
    }
  else if (m_sign)
    {
      rc = create_sign_attach (sink, protocol, m_output, m_input, m_micalg.c_str ());
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
                             false, m_mail->isDraftEncrypt (), exchange_major_version);
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

  TRETURN rc;
}

std::string
CryptController::get_inline_data ()
{
  TSTART;
  std::string ret;
  if (!m_mail->getDoPGPInline ())
    {
      TRETURN ret;
    }
  m_output.seek (0, SEEK_SET);
  char buf[4096];
  size_t nread;
  while ((nread = m_output.read (buf, 4096)) > 0)
    {
      ret += std::string (buf, nread);
    }
  TRETURN ret;
}

void
CryptController::parse_micalg (const GpgME::SigningResult &result)
{
  TSTART;
  if (result.isNull())
    {
      TRACEPOINT;
      TRETURN;
    }
  const auto signature = result.createdSignature(0);
  if (signature.isNull())
    {
      TRACEPOINT;
      TRETURN;
    }

  const char *hashAlg = signature.hashAlgorithmAsString ();
  if (!hashAlg)
    {
      TRACEPOINT;
      TRETURN;
    }
  if (m_proto == GpgME::OpenPGP)
    {
      m_micalg = std::string("pgp-") + hashAlg;
    }
  else
    {
      m_micalg = hashAlg;
    }
  std::transform(m_micalg.begin(), m_micalg.end(), m_micalg.begin(), ::tolower);

  log_debug ("%s:%s: micalg is: '%s'.",
             SRCNAME, __func__, m_micalg.c_str ());
  TRETURN;
}

void
CryptController::start_crypto_overlay ()
{
  TSTART;
  auto wid = m_mail->getWindow ();

  std::string text;

  if (m_encrypt)
    {
      text = utf8_gettext ("Encrypting...");
    }
  else if (m_sign)
    {
      text = utf8_gettext ("Signing...");
    }
  m_overlay = std::unique_ptr<Overlay> (new Overlay (wid, text));
  TRETURN;
}
