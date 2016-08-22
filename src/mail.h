/* @file mail.h
 * @brief High level class to work with Outlook Mailitems.
 *
 *    Copyright (C) 2015, 2016 Intevation GmbH
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
#ifndef MAIL_H
#define MAIL_H

#include "oomhelp.h"
#include "mapihelp.h"

#include <string>
#include <future>

class MailParser;

/** @brief Data wrapper around a mailitem.
 *
 * This class is intended to bundle all that we know about
 * a Mail. Due to the restrictions in Outlook we sometimes may
 * need additional information that is not available at the time
 * like the sender address of an exchange account in the afterWrite
 * event.
 *
 * This class bundles such information and also provides a way to
 * access the event handler of a mail.
 */
class Mail
{
public:
  /** @brief Construct a mail object for the item.
    *
    * This also installs the event sink for this item.
    *
    * The mail object takes ownership of the mailitem
    * reference. Do not Release it! */
  Mail (LPDISPATCH mailitem);

  ~Mail ();

  /** @brief looks for existing Mail objects for the OOM mailitem.

    @returns A reference to an existing mailitem or NULL in case none
    could be found.
  */
  static Mail* get_mail_for_item (LPDISPATCH mailitem);

  /** @brief looks for existing Mail objects.

    @returns A reference to an existing mailitem or NULL in case none
    could be found. Can be used to check if a mail object was destroyed.
  */
  static bool is_mail_valid (const Mail *mail);

  /** @brief wipe the plaintext from all known Mail objects.
    *
    * This is intended as a "cleanup" call to be done on unload
    * to avoid leaking plaintext in case we are deactivated while
    * some mails still have their plaintext inserted.
    *
    * @returns the number of errors that occured.
    */
  static int wipe_all_mails ();

  /** @brief revert all known Mail objects.
    *
    * Similar to wipe but works on MAPI to revert our attachment
    * dance and restore an original MIME mail.
    *
    * @returns the number of errors that occured.
    */
  static int revert_all_mails ();

  /** @brief Reference to the mailitem. Do not Release! */
  LPDISPATCH item () { return m_mailitem; }

  /** @brief Pre process the message. Ususally to be called from BeforeRead.
   *
   * This function assumes that the base message interface can be accessed
   * and calles the MAPI Message handling which changes the message class
   * to enable our own handling.
   *
   * @returns 0 on success.
   */
  int pre_process_message ();

  /** @brief Decrypt / Verify the mail.
   *
   * Sets the needs_wipe and was_encrypted variable.
   *
   * @returns 0 on success. */
  int decrypt_verify ();

  /** @brief do crypto operations as selected by the user.
   *
   * Initiates the crypto operations according to the gpgol
   * draft info flags.
   *
   * @returns 0 on success. */
  int encrypt_sign ();

  /** @brief Necessary crypto operations were completed successfully. */
  bool crypto_successful () { return !needs_crypto() || m_crypt_successful; }

  /** @brief Message should be encrypted and or signed. */
  bool needs_crypto ();

  /** @brief wipe the plaintext from the message and encrypt attachments.
   *
   * @returns 0 on success; */
  int wipe ();

  /** @brief revert the message to the original mail before our changes.
   *
   * @returns 0 on success; */
  int revert ();

  /** @brief update the sender address.
   *
   * For Exchange 2013 at least we don't have any other way to get the
   * senders SMTP address then through the object model. So we have to
   * store the sender address for later events that do not allow us to
   * access the OOM but enable us to work with the underlying MAPI structure.
   *
   * @returns 0 on success */
  int update_sender ();

  /** @brief get sender SMTP address (UTF-8 encoded).
   *
   * If the sender address has not been set through update_sender this
   * calls update_sender before returning the sender.
   *
   * @returns A reference to the utf8 sender address. Or NULL. */
  const char *get_sender ();

  /** @brief get the subject string (UTF-8 encoded).
    *
    * @returns the subject or an empty string. */
  std::string get_subject ();

  /** @brief Is this a crypto mail handled by gpgol.
  *
  * Calling this is only valid after a message has been processed.
  *
  * @returns true if the mail was either signed or encrypted and we processed
  * it.
  */
  bool is_crypto_mail () { return m_processed; }

  /** @brief This mail needs to be actually written.
  *
  * @returns true if the next write event should not be canceled.
  */
  bool needs_save () { return m_needs_save; }

  /** @brief set the needs save state.
  */
  void set_needs_save (bool val) { m_needs_save = val; }

  /** @brief is this mail an S/MIME mail.
    *
    * @returns true for smime messages.
    */
  bool is_smime ();

  /** @brief closes the inspector for this mail
    *
    * @returns true on success.
  */
  int close_inspector ();

  /** @brief get the associated parser.
    only valid while the actual parsing happens. */
  MailParser *parser () { return m_parser; }

  /** To be called from outside once the paser was done.
   In Qt this would be a slot that is called once it is finished
   we hack around that a bit by calling it from our windowmessages
   handler.
  */
  void parsing_done ();
private:
  LPDISPATCH m_mailitem;
  LPDISPATCH m_event_sink;
  bool m_processed,    /* The message has been porcessed by us.  */
       m_needs_wipe,   /* We have added plaintext to the mesage. */
       m_needs_save,   /* A property was changed but not by us. */
       m_crypt_successful, /* We successfuly performed crypto on the item. */
       m_is_smime, /* This is an smime mail. */
       m_is_smime_checked; /* it was checked if this is an smime mail */
  int m_moss_position; /* The number of the original message attachment. */
  char *m_sender;
  msgtype_t m_type; /* Our messagetype as set in mapi */
  MailParser *m_parser;
};
#endif // MAIL_H
