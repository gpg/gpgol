/* @file cryptcontroller.h
 * @brief Helper to handle sign and encrypt
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
#ifndef CRYPTCONTROLLER_H
#define CRYPTCONTROLLER_H

#ifdef HAVE_CONFIG_H
#include <config.h>
#endif

#include <gpgme++/data.h>

class Mail;
class Overlay;

namespace GpgME
{
  class SigningResult;
  class Error;
} // namespace GpgME

class CryptController
{
public:
  /** @brief Construct a Crypthelper for a Mail object. */
  CryptController (Mail *mail, bool encrypt, bool sign,
                   GpgME::Protocol proto);
  ~CryptController ();

  /** @brief Collect the data from the mail into the internal
      data structures.

      @returns 0 on success.
      */
  int collect_data ();

  /** @brief Does the actual crypto work according to op.
      Can be called in a different thread then the UI Thread.

      An operational error is returned in the passed err
      variable.

      @returns 0 on success.
  */
  int do_crypto (GpgME::Error &err);

  /** @brief Update the MAPI structure of the mail with
    the result. */
  int update_mail_mapi ();

  /** @brief Get an inline body as std::string. */
  std::string get_inline_data ();

  /** @brief Get the protocol. Valid after do_crypto. */
  GpgME::Protocol get_protocol () const { return m_proto; }

  /** @brief check weather something was encrypted. */
  bool is_encrypter () const { return m_encrypt; }

private:
  int resolve_keys ();
  int resolve_keys_cached ();
  int parse_output (GpgME::Data &resolverOutput);
  int lookup_fingerprints (const std::string &sigFpr,
                           const std::vector<std::string> recpFprs);

  void parse_micalg (const GpgME::SigningResult &sResult);

  void start_crypto_overlay ();

private:
  Mail *m_mail;
  GpgME::Data m_input, m_bodyInput, m_signedData, m_output;
  std::string m_micalg;
  bool m_encrypt, m_sign, m_crypto_success;
  GpgME::Protocol m_proto;
  GpgME::Key m_signer_key;
  std::vector<GpgME::Key> m_recipients;
  std::unique_ptr<Overlay> m_overlay;
  std::vector<std::string> m_recipient_addrs;
};

#endif
