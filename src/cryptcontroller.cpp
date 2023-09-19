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
#include "recipientmanager.h"
#include "windowmessages.h"

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
  m_signer_keys = mail->getSigningKeys ();
  m_sender = mail->getSender ();
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
  int count = m_mail->plainAttachments ().size ();

  /* Take the Body from the mail if possible. This is a fix for
     GnuPG-Bug-ID: T3614 because the body is not always properly
     updated in MAPI when sending. */
  char *body = m_mail->takeCachedPlainBody ();
  if (body && !*body)
    {
      xfree (body);
      body = nullptr;
    }
  if (!count && !body)
    {
      if (!m_mail->isDraftEncrypt())
        {
          gpgol_message_box (m_mail->getWindow (),
                             utf8_gettext ("Can't encrypt / sign an empty message."),
                             utf8_gettext ("GpgOL"), MB_OK);
        }
      xfree (body);
      TRETURN -1;
    }

  bool do_inline = m_mail->getDoPGPInline ();

  if (count && do_inline)
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
  int err = add_body_and_attachments (sink, m_mail, body);
  xfree (body);

  if (err)
    {
      log_error ("%s:%s: Collecting body and attachments failed.",
                 SRCNAME, __func__);
      TRETURN -1;
    }

  /* Set the input buffer to start. */
  m_input.seek (0, SEEK_SET);
  TRETURN 0;
}

int
CryptController::lookup_fingerprints (const std::vector<std::string> &sigFprs,
                                      const std::vector<std::pair <std::string, std::string> > &recpFprs)
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

  if (sigFprs.size())
    {
      char **cSigners = vector_to_cArray (sigFprs);
      err = ctx->startKeyListing (const_cast<const char **> (cSigners), true);
      if (err)
        {
          log_error ("%s:%s: failed to start signer keylisting",
                     SRCNAME, __func__);
          release_cArray (cSigners);
          TRETURN -1;
        }
      do
        {
          m_signer_keys.push_back(ctx->nextKey(err));
        }
      while (!err);

      release_cArray (cSigners);

      m_signer_keys.pop_back();

      if (m_signer_keys.empty())
        {
          log_error ("%s:%s: failed to lookup key for '%s' with protocol '%s'",
                     SRCNAME, __func__, anonstr (sigFprs[0].c_str ()),
                     m_proto == GpgME::CMS ? "smime" :
                     m_proto == GpgME::OpenPGP ? "openpgp" :
                     "unknown");
          TRETURN -1;
        }
      // reset context
      ctx = std::shared_ptr<GpgME::Context> (GpgME::Context::createForProtocol (m_proto));
      ctx->setKeyListMode (GpgME::Local);
    }

  if (!recpFprs.size())
    {
      TRETURN 0;
    }

  std::vector<std::string> all_fingerprints;
  for (const auto &pair: recpFprs)
    {
      all_fingerprints.push_back (pair.second);
    }

  for (auto &recp: m_recipients)
    {
      std::vector <std::string> fingerprintsToLookup;
      /* Find the matching fingerprints for the recpient. */
      for (const auto &pair: recpFprs)
        {
          if (pair.first.empty())
            {
              log_error ("%s:%s Recipient mbox is empty. Wrong Resolver Version?",
                         SRCNAME, __func__);
            }
          /* We also accept the empty string here for backwards compatibility
             but we should still log an error to be clear. */
          if (pair.first.empty() || pair.first == recp.mbox ())
            {
              if (pair.second.empty ())
                {
                  log_err ("Would have added an empty string!");
                  continue;
                }
              fingerprintsToLookup.push_back (pair.second);
              all_fingerprints.erase(std::remove_if(all_fingerprints.begin(),
                                                    all_fingerprints.end(),
                              [&pair](const std::string &x){return x == pair.second;}),
                              all_fingerprints.end());
            }
        }
      if (!fingerprintsToLookup.size ())
        {
          log_dbg ("No key selected for '%s'",
                   anonstr (recp.mbox().c_str ()));
          continue;
        }
      // Convert recipient fingerprints
      char **cRecps = vector_to_cArray (fingerprintsToLookup);
      err = ctx->startKeyListing (const_cast<const char **> (cRecps));
      if (err)
        {
          log_error ("%s:%s: failed to start recipient keylisting",
                     SRCNAME, __func__);
          release_cArray (cRecps);
          TRETURN -1;
        }

      std::vector <GpgME::Key> keys;
      do {
          const auto key = ctx->nextKey (err);
          if (key.isNull () || err)
            {
              continue;
            }
          keys.push_back (key);
          log_dbg ("Adding '%s' as key for '%s",
                   anonstr (key.primaryFingerprint ()),
                   anonstr (recp.mbox ().c_str ()));
      } while (!err);

      release_cArray (cRecps);
      recp.setKeys (keys);
      // reset context
      ctx = std::shared_ptr<GpgME::Context> (GpgME::Context::createForProtocol (m_proto));
      ctx->setKeyListMode (GpgME::Local);
    }
  if (!all_fingerprints.empty())
    {
      log_error ("%s:%s: BUG: Not all fingerprints could be matched "
                 "to a recipient.",
                 SRCNAME, __func__);
      for (const auto &fpr: all_fingerprints)
        {
          log_debug ("%s:%s Failed to find: %s", SRCNAME, __func__,
                     anonstr (fpr.c_str ()));
        }
      clear_keys ();
      TRETURN -1;
    }

  resolving_done ();

  TRETURN 0;
}


int
CryptController::parse_output (GpgME::Data &resolverOutput)
{
  TSTART;
  std::istringstream ss(resolverOutput.toString());
  std::string line;

  std::vector<std::string> sigFprs;
  std::vector<std::pair<std::string, std::string>> recpFprs;
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
      std::string mbox;

      std::getline (lss, what, ':');
      std::getline (lss, how, ':');
      std::getline (lss, fingerprint, ':');
      std::getline (lss, mbox, ':');

      if (m_proto == GpgME::UnknownProtocol)
        {
          /* TODO: Allow mixed */
          m_proto = (how == "smime") ? GpgME::CMS : GpgME::OpenPGP;
        }

      if (what == "sig")
        {
          sigFprs.push_back (fingerprint);
          continue;
        }
      if (what == "enc")
        {
          recpFprs.push_back(std::make_pair(mbox, fingerprint));
        }
    }

  if (m_sign && sigFprs.empty())
    {
      log_dbg ("%s:%s: Sign requested but no signing fingerprint - sending unsigned",
                 SRCNAME, __func__);
      m_sign = false;
      if (!m_encrypt)
        {
          TRETURN 0;
        }
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

  TRETURN lookup_fingerprints (sigFprs, recpFprs);
}

/* Combines all recipient keys for this operation
   into a list and copys the resolved recipients
   back. */
void
CryptController::resolving_done ()
{
  TSTART;
  m_enc_keys.clear ();

  for (const auto &recp: m_recipients)
    {
      const auto &keys = recp.keys ();
      m_enc_keys.insert (m_enc_keys.end (), keys.begin (), keys.end ());
    }
  /* Copy the resolved recipients back to the mail */
  m_mail->setRecipients (m_recipients);
  m_mail->setSigningKeys (m_signer_keys);
  TRETURN;
}

bool
CryptController::resolve_through_protocol (GpgME::Protocol proto)
{
  TSTART;

  const auto cache = KeyCache::instance ();

  if (m_encrypt)
    {
      for (auto &recp: m_recipients)
        {
          recp.setKeys(cache->getEncryptionKeys(recp.mbox (), proto));
        }
    }
  if (m_sign)
    {
      if (m_sender.empty())
        {
          log_error ("%s:%s: Asked to sign but sender is empty.",
                     SRCNAME, __func__);
          return false;
        }
      const auto key = cache->getSigningKey (m_sender.c_str (), proto);
      if (m_signer_keys.empty ())
        {
          m_signer_keys.push_back (key);
        }
      for (const auto &k: m_signer_keys)
        {
          if (k.protocol () != proto)
            {
              m_signer_keys.push_back (key);
              break;
            }
        }
    }
  TRETURN is_resolved ();
}

/* Check if we can be resolved by a single protocol
   and return it. */
GpgME::Protocol
CryptController::get_resolved_protocol () const
{
  TSTART;
  GpgME::Protocol ret = GpgME::UnknownProtocol;

  bool hasOpenPGPSignKey = false;
  bool hasSMIMESignKey = false;
  if (m_sign)
    {
      for (const auto &sig_key: m_signer_keys)
        {
          hasOpenPGPSignKey |= (!sig_key.isNull() &&
                                sig_key.protocol () == GpgME::OpenPGP);
          hasSMIMESignKey |= (!sig_key.isNull() &&
                              sig_key.protocol () == GpgME::CMS);
        }
    }

  if (m_encrypt)
    {
      for (const auto &recp: m_recipients)
        {
          if (!recp.keys ().size () || recp.keys ()[0].isNull ())
            {
              /* No keys or the first key in the list is null. */
              TRETURN GpgME::UnknownProtocol;
            }
          for (const auto &key: recp.keys())
            {
              if (key.protocol () == GpgME::OpenPGP &&
                  (!m_sign || hasOpenPGPSignKey) &&
                  (ret == GpgME::UnknownProtocol || ret == GpgME::OpenPGP))
                {
                  ret = GpgME::OpenPGP;
                  continue;
                }
              if (key.protocol () == GpgME::CMS &&
                  (!m_sign || hasSMIMESignKey) &&
                  (ret == GpgME::UnknownProtocol || ret == GpgME::CMS))
                {
                  ret = GpgME::CMS;
                  continue;
                }
              // Unresolvable
              TRETURN GpgME::UnknownProtocol;
            }
        }
    }
  TRETURN ret;
}


/* Check that the crypt operation is resolved. This
   supports combined S/MIME and OpenPGP operations. */
bool
CryptController::is_resolved () const
{
  /* Check that we have signing keys if necessary and at
     least one encryption key for each recipient. */
  bool hasOpenPGPSignKey = false;
  bool hasSMIMESignKey = false;
  static bool errMsgShown = false;
  if (m_sign)
    {
      for (const auto &sig_key: m_signer_keys)
        {
          hasOpenPGPSignKey |= (!sig_key.isNull() &&
                                sig_key.protocol () == GpgME::OpenPGP);
          hasSMIMESignKey |= (!sig_key.isNull() &&
                              sig_key.protocol () == GpgME::CMS);
        }
    }

  log_dbg ("Has OpenPGP Sig Key: %i SMIME: %i",
           hasOpenPGPSignKey, hasSMIMESignKey);

  if (!errMsgShown && m_sign && opt.sign_default && opt.enable_smime && opt.prefer_smime)
    {
      std::string msg;
      if (opt.smimeNoCertSigErr)
        {
          msg = opt.smimeNoCertSigErr;
        }
      else
          msg = _("No S/MIME (X509) signing certificate found.\n\n"
                  "Your organization has configured GpgOL to sign outgoing\n"
                  "mails with S/MIME certificates but there is no S/MIME\n"
                  "certificate configured for your mail address.\n\n"
                  "Please ask your Administrators for assistance or switch\n"
                  "to OpenPGP in the next dialog.");
      gpgol_message_box (m_mail->getWindow (), msg.c_str (),
                         _("GpgOL"), MB_OK);
      errMsgShown = true;
    }

  if (m_sign && !hasOpenPGPSignKey && !hasSMIMESignKey)
    {
      return false;
    }

  if (m_encrypt)
    {
      Recipient::dump (m_recipients);
      for (const auto &recp: m_recipients)
        {
          if (!recp.keys ().size () || recp.keys ()[0].isNull ())
            {
              /* No keys or the first key in the list is null. */
              return false;
            }
          /* If we don't sign we need no more checks. */
          if (!m_sign)
            {
              continue;
            }
          for (const auto &key: recp.keys())
            {
              if (key.protocol () == GpgME::OpenPGP && !hasOpenPGPSignKey)
                {
                  return false;
                }
              if (key.protocol () == GpgME::CMS && !hasSMIMESignKey)
                {
                  return false;
                }
            }
        }
    }
  return true;
}

int
CryptController::resolve_keys_cached()
{
  TSTART;
  // Prepare variables
  const auto recps = m_mail->getCachedRecipientAddresses ();
  bool resolved = false;
  if (opt.enable_smime && opt.prefer_smime)
    {
      resolved = resolve_through_protocol (GpgME::CMS);
      if (resolved)
        {
          log_debug ("%s:%s: Resolved with CMS due to preference.",
                     SRCNAME, __func__);
          m_proto = GpgME::CMS;
        }
    }
  if (!resolved)
    {
      resolved = resolve_through_protocol (GpgME::OpenPGP);
      if (resolved)
        {
          log_debug ("%s:%s: Resolved with OpenPGP.",
                     SRCNAME, __func__);
          m_proto = GpgME::OpenPGP;
        }
    }
  if (!resolved && (opt.enable_smime && !opt.prefer_smime))
    {
      resolved = resolve_through_protocol (GpgME::CMS);
      if (resolved)
        {
          log_debug ("%s:%s: Resolved with CMS as fallback.",
                     SRCNAME, __func__);
          m_proto = GpgME::CMS;
        }
    }

  for (const auto &recp: m_recipients)
    {
      log_debug ("Enc Key for: '%s' Type: %i", anonstr (recp.mbox ().c_str ()),
                                                        recp.type ());
      for (const auto &key: recp.keys ())
        {
          log_debug ("%s:%s", to_cstr (key.protocol ()),
                     anonstr (key.primaryFingerprint ()));
        }
    }
  for (const auto &sig_key: m_signer_keys)
    {
      log_debug ("%s:%s: Signing key: %s:%s",
                 SRCNAME, __func__,
                 to_cstr (sig_key.protocol()),
                 anonstr (sig_key.primaryFingerprint ()));
    }

  if (!resolved)
    {
      log_debug ("%s:%s: Failed to resolve through cache",
                 SRCNAME, __func__);
      m_enc_keys.clear ();
      m_signer_keys.clear ();
      m_proto = GpgME::UnknownProtocol;
      TRETURN 1;
    }

  if (m_encrypt)
    {
      log_debug ("%s:%s: Encrypting with protocol %s to:",
                 SRCNAME, __func__, to_cstr (m_proto));
    }
  TRETURN 0;
}

void
CryptController::clear_keys ()
{
  for (auto &recp: m_recipients)
    {
      recp.setKeys (std::vector<GpgME::Key> ());
    }
  m_signer_keys.clear ();
}

int
CryptController::resolve_keys ()
{
  TSTART;

  m_proto = get_resolved_protocol ();
  if (m_proto != GpgME::UnknownProtocol)
    {
      log_debug ("%s:%s: Already resolved by %s Not resolving again.",
                 SRCNAME, __func__, to_cstr (m_proto));
      start_crypto_overlay();
      resolving_done ();
      TRETURN 0;
    }

  m_enc_keys.clear();

  if (m_mail->isDraftEncrypt() && opt.draft_key)
    {
      GpgME::Key key;
      if (opt.draft_key && !strcmp (opt.draft_key, "auto"))
        {
          log_dbg ("Autoselecting draft key first ultimate key.");
          for (const auto &k: KeyCache::instance()->getUltimateKeys ())
            {
              if (k.hasSecret () && k.canEncrypt ())
                {
                  xfree (opt.draft_key);
                  gpgrt_asprintf (&opt.draft_key, "%s", k.primaryFingerprint());
                  log_dbg ("Autoselecting %s as draft encryption key.",
                           opt.draft_key);
                  write_options ();
                  key = k;
                  break;
                }
            }
        }

      if (key.isNull ())
        {
          key = KeyCache::instance()->getByFpr (opt.draft_key);
        }
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
      resolving_done ();
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
CryptController::prepare_crypto()
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

  int ret = 0;

  if (m_mail->copyParent ())
    {
      /* Bypass resolving if we are working on a split mail */
      m_recipients = m_mail->getCachedRecipients ();
      m_signer_keys = m_mail->getSigningKeys ();

      if (m_recipients.size () && m_recipients[0].keys ().size ())
        {
          m_proto = m_recipients[0].keys ()[0].protocol ();
        }
      else if (m_signer_keys.size ())
        {
          m_proto = m_signer_keys[0].protocol ();
        }

      m_enc_keys.clear ();
      for (const auto &recp: m_recipients)
        {
          const auto &keys = recp.keys ();
          m_enc_keys.insert (m_enc_keys.end (), keys.begin (), keys.end ());
        }

      if ((opt.enable_debug & DBG_DATA))
        {
          log_data ("Encrypting to: ");
          Recipient::dump (m_recipients);
        }
    }
  else
    {
      ret = resolve_keys ();
    }

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
  if (!m_encrypt && !m_sign)
    {
      log_dbg ("Nothing left after resolution. Passing unencrypted.");
      TRETURN -4;
    }

  if (!m_mail->copyParent ())
    {
      RecipientManager mngr (m_recipients, m_signer_keys);
      if (mngr.getRequiredMails () > 1)
        {
          log_dbg ("More then one mail required for this recipient selection.");
          /* If we need to send multiple emails we jump back from
             here into the main event loop. Copy the mail object
             and send it out mutiple times. */
          do_in_ui_thread_async (SEND_MULTIPLE_MAILS, m_mail);
          /* Cancel the crypto of this mail this continues
             in Mail::splitAndSend_o */
          wm_unregister_pending_op (m_mail);
          TRETURN -3;
        }
    }
  TRETURN 0;
}

int
CryptController::do_crypto (GpgME::Error &err, std::string &r_diag, bool force)
{
  TSTART;
  if (m_signer_keys.empty () && m_enc_keys.empty ()) {
    log_err ("Do crypto called without prepared keys. Call prepare_crypto "
             "first.");
    TRETURN -1;
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
  if (!m_signer_keys.empty ())
    {
      for (const auto &key: m_signer_keys)
        {
          if (key.protocol () == m_proto)
            {
              ctx->addSigningKey (key);
            }
        }
    }

  ctx->setTextMode (m_proto == GpgME::OpenPGP);
  ctx->setArmor (m_proto == GpgME::OpenPGP);

  /* Prepare to offer a "Forced encryption for S/MIME errors */
  GpgME::Context::EncryptionFlags flags = GpgME::Context::EncryptionFlags::None;
  int errVal = -1;
  /* For openPGP or when force is used we want to use Always Trust */
  if (m_proto == GpgME::OpenPGP || force) {
      if (force)
        {
          log_dbg ("Using alwaysTrust force option");
        }
      flags = GpgME::Context::AlwaysTrust;
  } else if (m_proto == GpgME::CMS && !force) {
      /* This should indicate to the caller that a repeated call with
         force might work */
      errVal = -3;
  }

  if (m_encrypt && m_sign && do_inline)
    {
      // Sign encrypt combined
      const auto result_pair = ctx->signAndEncrypt (m_enc_keys,
                                                    do_inline ? m_bodyInput : m_input,
                                                    m_output,
                                                    flags);
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
          TRETURN errVal;
        }

      if (err1.isCanceled() || err2.isCanceled())
        {
          err = err1.isCanceled() ? err1 : err2;
          log_debug ("%s:%s: User canceled",
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
          TRETURN errVal;
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
      m_output.setEncoding(GpgME::Data::MimeEncoding);
      m_input = GpgME::Data ();
      multipart.seek (0, SEEK_SET);
      const auto encResult = ctx->encrypt (m_enc_keys, multipart,
                                           m_output,
                                           flags);
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
          TRETURN errVal;
        }
      if (err.isCanceled())
        {
          log_debug ("%s:%s: User canceled",
                     SRCNAME, __func__);
          TRETURN -2;
        }
      // Now we have encrypted output just treat it like encrypted.
    }
  else if (m_encrypt)
    {
      m_output.setEncoding(GpgME::Data::MimeEncoding);
      const auto result = ctx->encrypt (m_enc_keys, do_inline ? m_bodyInput : m_input,
                                        m_output,
                                        flags);
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
          TRETURN errVal;
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
          log_debug ("%s:%s: User canceled",
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

  log_debug ("%s:%s: Crypto done successfully.",
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

void
CryptController::stop_crypto_overlay ()
{
  m_overlay = nullptr;
}
