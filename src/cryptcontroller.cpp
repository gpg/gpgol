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

int
CryptController::resolve_keys ()
{
  m_recipients.clear();

  /*XXX Temporary hack part do key resolution here. */
  GpgME::Error err;
  auto ctx = std::shared_ptr<GpgME::Context> (GpgME::Context::createForProtocol(GpgME::OpenPGP));
  const auto key = ctx->key ("EB4C5A5B7AD6C8527F050BAF1ED4F0BC6CFBC912", err, true);

  if (key.isNull())
    {
      log_error ("%s:%s: Failure to resolve keys.",
                 SRCNAME, __func__);
      return -1;
    }

  m_recipients.push_back(key);
  m_signer_key = key;
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

int
CryptController::update_mail_mapi ()
{
  log_debug ("%s:%s:", SRCNAME, __func__);
  return 0;
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
