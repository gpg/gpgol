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

class CryptController
{
public:
  /** @brief Construct a Crypthelper for a Mail object. */
  CryptController (Mail *mail, bool encrypt, bool sign,
                   bool inlineCrypt, GpgME::Protocol proto);
  ~CryptController ();

  /** @brief Collect the data from the mail into the internal
      data structures.

      @returns 0 on success.
      */
  int collect_data ();

  /** @brief Does the actual crypto work according to op.
      Can be called in a different thread then the UI Thread.

      @returns 0 on success.
  */
  int do_crypto ();

  /** @brief Update the MAPI structure of the mail with
    the result. */
  int update_mail_mapi ();

  /** @brief Get an inline body as std::string. */
  std::string get_inline_data ();

private:
  int resolve_keys ();
  int parse_output (GpgME::Data &resolverOutput);
  int lookup_fingerprints (const std::string &sigFpr,
                           const std::vector<std::string> recpFprs);

private:
  Mail *m_mail;
  GpgME::Data m_input, m_bodyInput, m_smime_intermediate, m_output;
  bool m_encrypt, m_sign, m_inline, m_crypto_success;
  GpgME::Protocol m_proto;
  GpgME::Key m_signer_key;
  std::vector<GpgME::Key> m_recipients;
};

#endif
