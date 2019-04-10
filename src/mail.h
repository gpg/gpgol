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
 *
 * Naming conventions of the suffixes:
 *  _o functions that work on OOM and possibly also MAPI.
 *  _m functions that work on MAPI.
 *  _s functions that only work on internal data and are safe to call
 *     from any thread.
 *
 * O and M functions _must_ only be called from the main thread. Use
 * a WindowMessage to signal the Main thread. But be wary. A WindowMessage
 * might be handled while an OOM call in the main thread waits for completion.
 *
 * An example for this is how update_oom_data can work:
 *
 * Main Thread:
 *   call update_oom_data
 *    └> internally invokes an OOM function that might do network access e.g.
 *       to connect to the exchange server to fetch the address.
 *
 *   Counterintutively the Main thread does not return from that function or
 *   blocks for it's completion but handles windowmessages.
 *
 *   After a windowmessage was handled and if the OOM invocation is
 *   completed the invocation returns and normal execution continues.
 *
 *   So if the window message handler's includes for example
 *   also a call to lookup recipients we crash. Note that it's usually
 *   safe to do OOM / MAPI calls from a window message.
 *
 *
 * While this seems impossible, remember that we do not work directly
 * with functions but everything is handled through COM. Without this
 * logic Outlook would probably become unusable because as any long running
 * call to the OOM would block it completely and freeze the UI.
 * (no windowmessages handled).
 *
 * So be wary when accessing the OOM from a Window Message.
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
  static Mail* getMailForItem (LPDISPATCH mailitem);

  /** @brief looks for existing Mail objects in the uuid map.
    Only objects for which set_uid has been called can be found
    in the uid map. Get the Unique ID of a mailitem thorugh get_unique_id

    @returns A reference to an existing mailitem or NULL in case none
    could be found.
  */
  static Mail* getMailForUUID (const char *uuid);

  /** @brief Get the last created mail.

    @returns A reference to the last created mail or null.
  */
  static Mail* getLastMail ();

  /** @brief voids the last mail variable. */
  static void clearLastMail ();

  /** @brief Lock mail deletion.

    Mails are heavily accessed multi threaded. E.g. when locating
    keys. Due to bad timing it would be possible that between
    a check for "is_valid_ptr" to see if a map is still valid
    and the usage of the mail a delete would happen.

    This lock can be used to prevent that. Changes made to the
    mail will of course have no effect as the mail is already in
    the process of beeing unloaded. And calls that access MAPI
    or OOM still might crash. But this at least gurantees that
    the member variables of the mail exist while the lock is taken.

    Use it in your thread like:

      Mail::lockDelete ();
      Mail::isValidPtr (mail);
      mail->set_or_check_something ();
      Mail::unlockDelete ();

      Still be carefull when it is a mapi or oom function.
  */
  static void lockDelete ();
  static void unlockDelete ();

  /** @brief looks for existing Mail objects.

    @returns A reference to an existing mailitem or NULL in case none
    could be found. Can be used to check if a mail object was destroyed.
  */
  static bool isValidPtr (const Mail *mail);

  /** @brief wipe the plaintext from all known Mail objects.
    *
    * This is intended as a "cleanup" call to be done on unload
    * to avoid leaking plaintext in case we are deactivated while
    * some mails still have their plaintext inserted.
    *
    * @returns the number of errors that occured.
    */
  static int wipeAllMails_o ();

  /** @brief revert all known Mail objects.
    *
    * Similar to wipe but works on MAPI to revert our attachment
    * dance and restore an original MIME mail.
    *
    * @returns the number of errors that occured.
    */
  static int revertAllMails_o ();

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
  static int closeAllMails_o ();

  /** @brief closes the inspector for a mail
    *
    * @returns true on success.
  */
  static int closeInspector_o (Mail *mail);

  /** Call close with discard changes to discard
      plaintext. returns the value of the oom close
      call. This may have delete the mail if the close
      triggers an unload.
  */
  static int close (Mail *mail);

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
  static void locateAllCryptoRecipients_o ();

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
  int preProcessMessage_m ();

  /** @brief Decrypt / Verify the mail.
   *
   * Sets the needs_wipe and was_encrypted variable.
   *
   * @returns 0 on success. */
  int decryptVerify_o ();

  /** @brief start crypto operations as selected by the user.
   *
   * Initiates the crypto operations according to the gpgol
   * draft info flags.
   *
   * @returns 0 on success. */
  int encryptSignStart_o ();

  /** @brief Necessary crypto operations were completed successfully. */
  bool wasCryptoSuccessful_m () { return m_crypt_successful || !needs_crypto_m (); }

  /** @brief Message should be encrypted and or signed.
    0: No
    1: Encrypt
    2: Sign
    3: Encrypt & Sign
  */
  int needs_crypto_m () const;

  /** @brief wipe the plaintext from the message and encrypt attachments.
   *
   * @returns 0 on success; */
  int wipe_o (bool force = false);

  /** @brief revert the message to the original mail before our changes.
   *
   * @returns 0 on success; */
  int revert_o ();

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
  int updateOOMData_o ();

  /** @brief get sender SMTP address (UTF-8 encoded).
   *
   * If the sender address has not been set through update_sender this
   * calls update_sender before returning the sender.
   *
   * @returns A reference to the utf8 sender address. Or an empty string. */
  std::string getSender_o ();

  /** @brief get sender SMTP address (UTF-8 encoded).
   *
   * Like get_sender but ensures not to touch oom or mapi
   *
   * @returns A reference to the utf8 sender address. Or an empty string. */
  std::string getSender () const;

  /** @brief get the subject string (UTF-8 encoded).
    *
    * @returns the subject or an empty string. */
  std::string getSubject_o () const;

  /** @brief Is this a crypto mail handled by gpgol.
  *
  * Calling this is only valid after a message has been processed.
  *
  * @returns true if the mail was either signed or encrypted and we processed
  * it.
  */
  bool isCryptoMail () const;

  /** @brief This mail needs to be actually written.
  *
  * @returns true if the next write event should not be canceled.
  */
  bool needsSave () { return m_needs_save; }

  /** @brief set the needs save state.
  */
  void setNeedsSave (bool val) { m_needs_save = val; }

  /** @brief is this mail an S/MIME mail.
    *
    * @returns true for smime messages.
    */
  bool isSMIME_m ();

  /** @brief get the associated parser.
    only valid while the actual parsing happens. */
  std::shared_ptr<ParseController> parser () { return m_parser; }

  /** @brief get the associated cryptcontroller.
    only valid while the crypting happens. */
  std::shared_ptr<CryptController> cryper () { return m_crypter; }

  /** To be called from outside once the paser was done.
   In Qt this would be a slot that is called once it is finished
   we hack around that a bit by calling it from our windowmessages
   handler.
  */
  void parsing_done ();

  /** Returns true if the mail was verified and has at least one
    signature. Regardless of the validity of the mail */
  bool isSigned () const;

  /** Returns true if the mail is encrypted to at least one
    recipient. Regardless if it could be decrypted. */
  bool isEncrypted () const;

  /** Are we "green" */
  bool isValidSig () const;

  /** Get UID gets UniqueID property of this mail. Returns
    an empty string if the uid was not set with set uid.*/
  const std::string & getUUID () const { return m_uuid; }

  /** Returns 0 on success if the mail has a uid alrady or sets
    the uid. Setting only succeeds if the OOM is currently
    accessible. Returns -1 on error. */
  int setUUID_o ();

  /** Returns a localized string describing in one or two
    words the crypto status of this mail. */
  std::string getCryptoSummary () const;

  /** Returns a localized string describing the detailed
    crypto state of this mail. */
  std::string getCryptoDetails_o ();

  /** Returns a localized string describing a one line
    summary of the crypto state. */
  std::string getCryptoOneLine () const;

  /** Get the icon id of the appropiate icon for this mail */
  int getCryptoIconID () const;

  /** Get the fingerprint of an associated signature or null
      if it is not signed. */
  const char *getSigFpr () const;

  /** Remove all categories of this mail */
  void removeCategories_o ();

  /** Get the body of the mail */
  std::string getBody_o () const;

  /** Get the recipients. */
  std::vector<std::string> getRecipients_o () const;

  /** Try to locate the keys for all recipients.
      This also triggers the Addressbook integration, which we
      treat as locate jobs. */
  void locateKeys_o ();

  /** State variable to check if a close was triggerd by us. */
  void setCloseTriggered (bool value);
  bool getCloseTriggered () const;

  /** Check if the mail should be sent as html alternative mail.
    Only valid if update_oom_data was called before. */
  bool isHTMLAlternative () const;

  /** Get the html body. It is updated in update_oom_data.
      Caller takes ownership of the string and has to free it.
  */
  char *takeCachedHTMLBody ();

  /** Get the plain body. It is updated in update_oom_data.
      Caller takes ownership of the string and has to free it.
  */
  char *takeCachedPlainBody ();

  /** Get the cached recipients. It is updated in update_oom_data.*/
  std::vector<std::string> getCachedRecipients ();

  /** Returns 1 if the mail was encrypted, 2 if signed, 3 if both.
      Only valid after decrypt_verify.
  */
  int getCryptoFlags () const;

  /** Returns true if the mail should be encrypted in the
      after write event. */
  bool getNeedsEncrypt () const;
  void setNeedsEncrypt (bool val);

  /** Gets the level of the signature. See:
    https://wiki.gnupg.org/EasyGpg2016/AutomatedEncryption for
    a definition of the levels. */
  int get_signature_level () const;

  /** Check if all attachments are hidden and show a warning
    message appropiate to the crypto state if necessary. */
  int checkAttachments_o () const;

  /** Check if the mail should be encrypted "inline" */
  bool getDoPGPInline () const {return m_do_inline;}

  /** Check if the mail should be encrypted "inline" */
  void setDoPGPInline (bool value) {m_do_inline = value;}

  /** Append data to a cached inline body. Helper to do this
     on MAPI level and later add it through OOM */
  void appendToInlineBody (const std::string &data);

  /** Set the inline body as OOM body property. */
  int inlineBodyToBody_o ();

  /** Get the crypt state */
  CryptState cryptState () const {return m_crypt_state;}

  /** Set the crypt state */
  void setCryptState (CryptState state) {m_crypt_state = state;}

  /** Update MAPI data after encryption. */
  void updateCryptMAPI_m ();

  /** Update OOM data after encryption.

    Checks for plain text leaks and
    does not advance crypt state if body can't be cleaned.
  */
  void updateCryptOOM_o ();

  /** Enable / Disable the window of this mail.

    When the window gets disabled the
    handle is stored for a later enable. */
  void disableWindow_o ();
  void enableWindow ();

  /** Determine if the mail is an inline response.

    Call check_inline_response first to update the state
    from the OOM.

    We need synchronous encryption for inline responses. */
  bool isAsyncCryptDisabled () { return m_async_crypt_disabled; }

  /** Check through OOM if the current mail is an inline
    response.

    Caches the state which can then be queried through
    async_crypt_disabled
  */
  bool check_inline_response ();

  /** Get the window for the mail. Caution! This is only
    really valid in the time that the window is disabled.
    Use with care and can be null or invalid.
  */
  HWND getWindow () { return m_window; }

  /** Cleanup any attached crypter object. Useful
    on error. */
  void resetCrypter () { m_crypter = nullptr; }

  /** Set special crypto mime data that should be used as the
    mime structure when sending. */
  void setOverrideMIMEData (const std::string &data) {m_mime_data = data;}

  /** Get the mime data that should be used when sending. */
  std::string get_override_mime_data () const { return m_mime_data; }
  bool hasOverrideMimeData() const { return !m_mime_data.empty(); }

  /** Set if this is a forward of a crypto mail. */
  void setIsForwardedCryptoMail (bool value) { m_is_forwarded_crypto_mail = value; }
  bool is_forwarded_crypto_mail () { return m_is_forwarded_crypto_mail; }

  /** Set if this is a reply of a crypto mail. */
  void setIsReplyCryptoMail (bool value) { m_is_reply_crypto_mail = value; }
  bool is_reply_crypto_mail () { return m_is_reply_crypto_mail; }

  /** Remove the hidden GpgOL attachments. This is needed when forwarding
    without encryption so that our attachments are not included in the forward.
    Returns 0 on success. Works in OOM. */
  int removeOurAttachments_o ();

  /** Remove all attachments. Including our own. This is needed for
    forwarding of unsigned S/MIME mails (Efail).
    Returns 0 on success. Works in OOM. */
  int removeAllAttachments_o ();

  /** Check both OOM and MAPI if the body is either empty or
    encrypted. Won't abort on OOM or MAPI errors, so it can be
    used in both states. But will return false if a body
    was detected or in the OOM the MAPI Base Message. This
    is intended as a saveguard before sending a mail.

    This function should not be used to detected the necessity
    of encryption and is only an extra check to catch unexpected
    errors.
    */
  bool hasCryptedOrEmptyBody_o ();

  void updateBody_o ();

  /** Set if this mail looks like the send again of a crypto mail.
      This will mean that after it is decrypted it is treated
      like an unencrypted mail so that it can be encrypted again
      or sent unencrypted.
      */
  void setIsSendAgain (bool value) { m_is_send_again = value; }


  /* Attachment removal state variables. */
  bool attachmentRemoveWarningDisabled () { return m_disable_att_remove_warning; }

  /* Gets the string dump of the verification result. */
  std::string getVerificationResultDump ();

  /* Block loading HTML content */
  void setBlockHTML (bool value);
  bool isBlockHTML () const { return m_block_html; }

  /* Remove automatic loading of HTML references setting. */
  void setBlockStatus_m ();

  /* Crypto options (sign/encrypt) have been set manually. */
  void setCryptoSelectedManually (bool v) { m_manual_crypto_opts = v; }
  // bool is_crypto_selected_manually () const { return m_manual_crypto_opts; }

  /* Reference that a resolver thread is running for this mail. */
  void incrementLocateCount ();

  /* To be called when a resolver thread is done. If there are no running
     resolver threads we can check the recipients to see if we should
     toggle / untoggle the secure state.
     */
  void decrementLocateCount ();

  /* Check if the keys can be resolved automatically and trigger
   * setting the crypto flags accordingly.
   */
  void autosecureCheck ();

  /* Set if a mail should be secured (encrypted and signed)
   *
   * Only save to call from a place that may access mapi.
   */
  void setDoAutosecure_m (bool value);

  /* Install an event handler for the folder of this mail. */
  void installFolderEventHandler_o ();

  /* Marker for a "Move" of this mail */
  bool passWrite () { return m_pass_write; }
  void setPassWrite(bool value) { m_pass_write = value; }

  /* Releases the current item ref obtained in update oom data */
  void releaseCurrentItem ();
  /* Gets an additional reference for GetInspector.CurrentItem */
  void refCurrentItem ();

  /* Get the storeID for this mail */
  std::string storeID() { return m_store_id; }

  /* Remove encryption permanently. */
  void decryptPermanently_o ();

  /* Prepare for encrypt / sign. Updates data. */
  void prepareCrypto_o ();

  /* State variable to check if we are about to encrypt a draft. */
  void setIsDraftEncrypt (bool value) { m_is_draft_encrypt = value; }
  bool isDraftEncrypt () { return m_is_draft_encrypt; }
private:
  void updateCategories_o ();
  void updateSigstate ();

  LPDISPATCH m_mailitem;
  LPDISPATCH m_event_sink;
  LPDISPATCH m_currentItemRef;
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
  std::string m_sent_on_behalf;
  char *m_cached_html_body; /* Cached html body. */
  char *m_cached_plain_body; /* Cached plain body. */
  std::vector<std::string> m_cached_recipients;
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
  bool m_async_crypt_disabled;
  std::string m_mime_data;
  bool m_is_forwarded_crypto_mail; /* Is this a forward of a crypto mail */
  bool m_is_reply_crypto_mail; /* Is this a reply to a crypto mail */
  bool m_is_send_again; /* Is this a send again of a crypto mail */
  bool m_disable_att_remove_warning; /* Should not warn about attachment removal. */
  bool m_block_html; /* Force blocking of html content. e.g for unsigned S/MIME mails. */
  bool m_manual_crypto_opts; /* Crypto options (sign/encrypt) have been set manually. */
  bool m_first_autosecure_check; /* This is the first autoresolve check */
  int m_locate_count; /* The number of key locates pending for this mail. */
  bool m_pass_write; /* Danger the next write will be passed. This is for closed mails */
  bool m_locate_in_progress; /* Simplified state variable for locate */
  std::string m_store_id; /* Store id for categories */
  std::string m_verify_category; /* The category string for the verify result */
  bool m_is_junk; /* Mail is in the junk folder */
  bool m_is_draft_encrypt; /* Mail is a draft that should be encrypted */
};
#endif // MAIL_H
