/* @file mail.h
 * @brief High level class to work with Outlook Mailitems.
 *
 *    Copyright (C) 2015 Intevation GmbH
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

  /** @brief wipe the plaintext from all known Mail objects.
    *
    * This is intended as a "cleanup" call to be done on unload
    * to avoid leaking plaintext in case we are deactivated while
    * some mails still have their plaintext inserted.
    *
    * @returns the number of errors that occured.
    */
  static int wipe_all_mails ();

  /** @brief Reference to the mailitem. Do not Release! */
  LPDISPATCH item () { return m_mailitem; }

  /** @brief Process the message. Ususally to be called from BeforeRead.
   *
   * This function assumes that the base message interface can be accessed
   * and calles the MAPI Message handling which creates the GpgOL style
   * attachments and sets up the message class etc.
   *
   * Sets the was_encrypted / processed variables.
   *
   * @returns 0 on success.
   */
  int process_message ();

  /** @brief Replace the body with the plaintext and session decrypts
   * attachments.
   *
   * Sets the needs_wipe variable.
   *
   * @returns 0 on success. */
  int insert_plaintext ();

  /** @brief do crypto operations as selected by the user.
   *
   * Initiates the crypto operations according to the gpgol
   * draft info flags.
   *
   * @returns 0 on success. */
  int do_crypto ();

  /** @brief Necessary crypto operations were completed successfully. */
  bool crypto_successful () { return !needs_crypto() || m_crypt_successful; }

  /** @brief Message should be encrypted and or signed. */
  bool needs_crypto ();

  /** @brief wipe the plaintext from the message and ecnrypt attachments.
   *
   * @returns 0 on success; */
  int wipe ();

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

  /** @brief Is this a crypto mail handled by gpgol.
  *
  * Calling this is only valid after a message has been processed.
  *
  * @returns true if the mail was either signed or encrypted and we processed
  * it.
  */
  bool is_crypto_mail () { return m_processed; }

private:
  LPDISPATCH m_mailitem;
  LPDISPATCH m_event_sink;
  bool m_processed,    /* The message has been porcessed by us.  */
       m_needs_wipe,   /* We have added plaintext to the mesage. */
       m_crypt_successful; /* We successfuly performed crypto on the item. */
  char *m_sender;
};
#endif // MAIL_H
