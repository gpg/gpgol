/* @file mail.h
 * @brief High level class to work with Outlook Mailitems.
 *
 * Copyright (C) 2015, 2016 by Bundesamt für Sicherheit in der Informationstechnik
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
#ifndef MAIL_H
#define MAIL_H

#include "oomhelp.h"
#include "mapihelp.h"
#include "gpgme++/verificationresult.h"
#include "gpgme++/decryptionresult.h"
#include "gpgme++/key.h"

#include <string>
#include <future>

class ParseController;
class CryptController;

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
  enum CryptState
    {
      NoCryptMail,
      NeedsFirstAfterWrite,
      NeedsActualCrypt,
      NeedsUpdateInOOM,
      NeedsSecondAfterWrite,
      NeedsUpdateInMAPI,
      WantsSendInline,
      WantsSendMIME,
    };

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

  /** @brief looks for existing Mail objects in the uuid map.
    Only objects for which set_uid has been called can be found
    in the uid map. Get the Unique ID of a mailitem thorugh get_unique_id

    @returns A reference to an existing mailitem or NULL in case none
    could be found.
  */
  static Mail* get_mail_for_uuid (const char *uuid);

  /** @brief Get the last created mail.

    @returns A reference to the last created mail or null.
  */
  static Mail* get_last_mail ();

  static void invalidate_last_mail ();

  /** @brief looks for existing Mail objects.

    @returns A reference to an existing mailitem or NULL in case none
    could be found. Can be used to check if a mail object was destroyed.
  */
  static bool is_valid_ptr (const Mail *mail);

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

  /** @brief close all known Mail objects.
    *
    * Close our mail with discard changes set to true.
    * This discards the plaintext / attachments. Afterwards
    * it calls save if neccessary to sync back the collected
    * property changes.
    *
    * This is the nicest of our three "Clean plaintext"
    * functions. Will fallback to revert if closing fails.
    * Closed mails are deleted.
    *
    * @returns the number of errors that occured.
    */
  static int close_all_mails ();

  /** @brief locate recipients for all crypto mails
    *
    * To avoid lookups of recipients for non crypto mails we only
    * locate keys when a crypto action is already selected.
    *
    * As the user can do this after recipients were added but
    * we don't know for which mail the crypt button was triggered.
    * we march over all mails and if they are crypto mails we check
    * that the recipents were located.
    */
  static void locate_all_crypto_recipients ();

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

  /** @brief start crypto operations as selected by the user.
   *
   * Initiates the crypto operations according to the gpgol
   * draft info flags.
   *
   * @returns 0 on success. */
  int encrypt_sign_start ();

  /** @brief Necessary crypto operations were completed successfully. */
  bool crypto_successful () { return !needs_crypto() || m_crypt_successful; }

  /** @brief Message should be encrypted and or signed.
    0: No
    1: Encrypt
    2: Sign
    3: Encrypt & Sign
  */
  int needs_crypto ();

  /** @brief wipe the plaintext from the message and encrypt attachments.
   *
   * @returns 0 on success; */
  int wipe (bool force = false);

  /** @brief revert the message to the original mail before our changes.
   *
   * @returns 0 on success; */
  int revert ();

  /** @brief update some data collected from the oom
   *
   * This updates cached values from the OOM that are not available
   * in MAPI events like after Write.
   *
   * For Exchange 2013 at least we don't have any other way to get the
   * senders SMTP address then through the object model. So we have to
   * store the sender address for later events that do not allow us to
   * access the OOM but enable us to work with the underlying MAPI structure.
   *
   * It also updated the is_html_alternative value.
   *
   * @returns 0 on success */
  int update_oom_data ();

  /** @brief get sender SMTP address (UTF-8 encoded).
   *
   * If the sender address has not been set through update_sender this
   * calls update_sender before returning the sender.
   *
   * @returns A reference to the utf8 sender address. Or an empty string. */
  std::string get_sender ();

  /** @brief get sender SMTP address (UTF-8 encoded).
   *
   * Like get_sender but ensures not to touch oom or mapi
   *
   * @returns A reference to the utf8 sender address. Or an empty string. */
  std::string get_cached_sender ();

  /** @brief get the subject string (UTF-8 encoded).
    *
    * @returns the subject or an empty string. */
  std::string get_subject () const;

  /** @brief Is this a crypto mail handled by gpgol.
  *
  * Calling this is only valid after a message has been processed.
  *
  * @returns true if the mail was either signed or encrypted and we processed
  * it.
  */
  bool is_crypto_mail () const;

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

  /** @brief closes the inspector for a mail
    *
    * @returns true on success.
  */
  static int close_inspector (Mail *mail);

  /** @brief get the associated parser.
    only valid while the actual parsing happens. */
  std::shared_ptr<ParseController> parser () { return m_parser; }

  /** @brief get the associated cryptcontroller.
    only valid while the crypting happens. */
  std::shared_ptr<CryptController> crypter () { return m_crypter; }

  /** To be called from outside once the paser was done.
   In Qt this would be a slot that is called once it is finished
   we hack around that a bit by calling it from our windowmessages
   handler.
  */
  void parsing_done ();

  /** Returns true if the mail was verified and has at least one
    signature. Regardless of the validity of the mail */
  bool is_signed () const;

  /** Returns true if the mail is encrypted to at least one
    recipient. Regardless if it could be decrypted. */
  bool is_encrypted () const;

  /** Are we "green" */
  bool is_valid_sig ();

  /** Get UID gets UniqueID property of this mail. Returns
    an empty string if the uid was not set with set uid.*/
  const std::string & get_uuid () const { return m_uuid; }

  /** Returns 0 on success if the mail has a uid alrady or sets
    the uid. Setting only succeeds if the OOM is currently
    accessible. Returns -1 on error. */
  int set_uuid ();

  /** Returns a localized string describing in one or two
    words the crypto status of this mail. */
  std::string get_crypto_summary ();

  /** Returns a localized string describing the detailed
    crypto state of this mail. */
  std::string get_crypto_details ();

  /** Returns a localized string describing a one line
    summary of the crypto state. */
  std::string get_crypto_one_line ();

  /** Get the icon id of the appropiate icon for this mail */
  int get_crypto_icon_id () const;

  /** Get the fingerprint of an associated signature or null
      if it is not signed. */
  const char *get_sig_fpr() const;

  /** Remove all categories of this mail */
  void remove_categories ();

  /** Get the body of the mail */
  std::string get_body () const;

  /** Get the html of the mail */
  std::string get_html_body () const;

  /** Get the recipients recipients is a null
      terminated array of strings. Needs to be freed
      by the caller. */
  char ** get_recipients () const;

  /** Call close with discard changes to discard
      plaintext. returns the value of the oom close
      call. This may have delete the mail if the close
      triggers an unload.
  */
  static int close (Mail *mail);

  /** Try to locate the keys for all recipients */
  void locate_keys();

  /** State variable to check if a close was triggerd by us. */
  void set_close_triggered (bool value);
  bool get_close_triggered () const;

  /** Check if the mail should be sent as html alternative mail.
    Only valid if update_oom_data was called before. */
  bool is_html_alternative () const;

  /** Get the html body. It is updated in update_oom_data.
      Caller takes ownership of the string and has to free it.
  */
  char *take_cached_html_body ();

  /** Get the plain body. It is updated in update_oom_data.
      Caller takes ownership of the string and has to free it.
  */
  char *take_cached_plain_body ();

  /** Get the cached recipients. It is updated in update_oom_data.
      Caller takes ownership of the list and has to free it.
  */
  char **take_cached_recipients ();

  /** Returns 1 if the mail was encrypted, 2 if signed, 3 if both.
      Only valid after decrypt_verify.
  */
  int get_crypto_flags () const;

  /** Returns true if the mail should be encrypted in the
      after write event. */
  bool needs_encrypt () const;
  void set_needs_encrypt (bool val);

  /** Gets the level of the signature. See:
    https://wiki.gnupg.org/EasyGpg2016/AutomatedEncryption for
    a definition of the levels. */
  int get_signature_level () const;

  /** Check if all attachments are hidden and show a warning
    message appropiate to the crypto state if necessary. */
  int check_attachments () const;

  /** Check if the mail should be encrypted "inline" */
  bool should_inline_crypt () const {return m_do_inline;}

  /** Append data to a cached inline body. Helper to do this
     on MAPI level and later add it through OOM */
  void append_to_inline_body (const std::string &data);

  /** Set the inline body as OOM body property. */
  int inline_body_to_body ();

  /** Get the crypt state */
  CryptState crypt_state () const {return m_crypt_state;}

  /** Set the crypt state */
  void set_crypt_state (CryptState state) {m_crypt_state = state;}

  /** Update MAPI data after encryption. */
  void update_crypt_mapi ();

  /** Update OOM data after encryption.

    Checks for plain text leaks and
    does not advance crypt state if body can't be cleaned.
  */
  void update_crypt_oom ();

  /** Enable / Disable the window of this mail.

    When value is false the active window will
    be disabled and the handle stored for a later
    enable. */
  void set_window_enabled (bool value);

  /** Determine if the mail is an inline response.

    Call check_inline_response first to update the state
    from the OOM.

    We need synchronous encryption for inline responses. */
  bool is_inline_response () { return m_is_inline_response; }

  /** Check through OOM if the current mail is an inline
    response.

    Caches the state which can then be queried through
    is_inline_response
  */
  bool check_inline_response ();

  /** Get the window for the mail. Caution! This is only
    really valid in the time that the window is disabled.
    Use with care and can be null or invalid.
  */
  HWND get_window () { return m_window; }

  /** Cleanup any attached crypter object. Useful
    on error. */
  void reset_crypter () { m_crypter = nullptr; }

  /** Set special crypto mime data that should be used as the
    mime structure when sending. */
  void set_override_mime_data (const std::string &data) {m_mime_data = data;}

  /** Get the mime data that should be used when sending. */
  std::string get_override_mime_data () const { return m_mime_data; }

  void update_body ();
private:
  void update_categories ();
  void update_sigstate ();

  LPDISPATCH m_mailitem;
  LPDISPATCH m_event_sink;
  bool m_processed,    /* The message has been porcessed by us.  */
       m_needs_wipe,   /* We have added plaintext to the mesage. */
       m_needs_save,   /* A property was changed but not by us. */
       m_crypt_successful, /* We successfuly performed crypto on the item. */
       m_is_smime, /* This is an smime mail. */
       m_is_smime_checked, /* it was checked if this is an smime mail */
       m_is_signed, /* Mail is signed */
       m_is_valid, /* Mail is valid signed. */
       m_close_triggered, /* We have programtically triggered a close */
       m_is_html_alternative, /* Body Format is not plain text */
       m_needs_encrypt; /* Send was triggered we want to encrypt. */
  int m_moss_position; /* The number of the original message attachment. */
  int m_crypto_flags;
  std::string m_sender;
  char *m_cached_html_body; /* Cached html body. */
  char *m_cached_plain_body; /* Cached plain body. */
  char **m_cached_recipients;
  msgtype_t m_type; /* Our messagetype as set in mapi */
  std::shared_ptr <ParseController> m_parser;
  std::shared_ptr <CryptController> m_crypter;
  GpgME::VerificationResult m_verify_result;
  GpgME::DecryptionResult m_decrypt_result;
  GpgME::Signature m_sig;
  GpgME::UserID m_uid;
  std::string m_uuid;
  std::string m_orig_body;
  bool m_do_inline;
  bool m_is_gsuite; /* Are we on a gsuite account */
  std::string m_inline_body;
  CryptState m_crypt_state;
  HWND m_window;
  bool m_is_inline_response;
  std::string m_mime_data;
};
#endif // MAIL_H
