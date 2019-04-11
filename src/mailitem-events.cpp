/* mailitem-events.h - Event handling for mails.
 * Copyright (C) 2015 by Bundesamt für Sicherheit in der Informationstechnik
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

#include "config.h"
#include "common.h"
#include "eventsink.h"
#include "eventsinks.h"
#include "mymapi.h"
#include "mymapitags.h"
#include "oomhelp.h"
#include "ocidl.h"
#include "windowmessages.h"
#include "mail.h"
#include "mapihelp.h"
#include "gpgoladdin.h"
#include "wks-helper.h"

#undef _
#define _(a) utf8_gettext (a)

const wchar_t *prop_blacklist[] = {
  L"Body",
  L"HTMLBody",
  L"To", /* Somehow this is done when a mail is opened */
  L"CC", /* Ditto */
  L"BCC", /* Ditto */
  L"Categories",
  L"UnRead",
  L"OutlookVersion",
  L"OutlookInternalVersion",
  L"ReceivedTime",
  L"InternetCodepage",
  L"ConversationIndex",
  L"Subject",
  L"SentOnBehalfOfName",
  NULL };

typedef enum
  {
    AfterWrite = 0xFC8D,
    AttachmentAdd = 0xF00B,
    AttachmentRead = 0xF00C,
    AttachmentRemove = 0xFBAE,
    BeforeAttachmentAdd = 0xFBB0,
    BeforeAttachmentPreview = 0xFBAF,
    BeforeAttachmentRead = 0xFBAB,
    BeforeAttachmentSave = 0xF00D,
    BeforeAttachmentWriteToTempFile = 0xFBB2,
    BeforeAutoSave = 0xFC02,
    BeforeCheckNames = 0xF00A,
    BeforeDelete = 0xFA75,
    BeforeRead = 0xFC8C,
    Close = 0xF004,
    CustomAction = 0xF006,
    CustomPropertyChange = 0xF008,
    Forward = 0xF468,
    Open = 0xF003,
    PropertyChange = 0xF009,
    Read = 0xF001,
    ReadComplete = 0xFC8F,
    Reply = 0xF466,
    ReplyAll = 0xF467,
    Send = 0xF005,
    Unload = 0xFBAD,
    Write = 0xF002
  } MailEvent;

/* Mail Item Events */
BEGIN_EVENT_SINK(MailItemEvents, IDispatch)
/* We are still in the class declaration */

private:
  Mail * m_mail; /* The mail object related to this mailitem */
};

MailItemEvents::MailItemEvents() :
    m_object(NULL),
    m_pCP(NULL),
    m_cookie(0),
    m_ref(1),
    m_mail(NULL)
{
}

MailItemEvents::~MailItemEvents()
{
  if (m_pCP)
    m_pCP->Unadvise(m_cookie);
  if (m_object)
    gpgol_release (m_object);
}

static bool propchangeWarnShown = false;
static bool attachRemoveWarnShown = false;
static bool addinsLogged = false;

static DWORD WINAPI
do_delayed_locate (LPVOID arg)
{
  Sleep(100);
  do_in_ui_thread (RECIPIENT_ADDED, arg);
  return 0;
}

/* The main Invoke function. The return value of this
   function does not appear to have any effect on outlook
   although I have read in an example somewhere that you
   should return S_OK so that outlook continues to handle
   the event I have not yet seen any effect by returning
   error values here and no MSDN documentation about the
   return values.
*/
EVENT_SINK_INVOKE(MailItemEvents)
{
  USE_INVOKE_ARGS
  TSTART;
  if (!m_mail)
    {
      m_mail = Mail::getMailForItem (m_object);
      if (!m_mail)
        {
          log_error ("%s:%s: mail event without mail object known. Bug.",
                     SRCNAME, __func__);
          TRETURN S_OK;
        }
    }

  bool is_reply = false;
  switch(dispid)
    {
      case BeforeAutoSave:
        {
          log_oom ("%s:%s: BeforeAutoSave : %p",
                   SRCNAME, __func__, m_mail);
          if (parms->cArgs != 1 || parms->rgvarg[0].vt != (VT_BOOL | VT_BYREF))
           {
             /* This happens in the weird case */
             log_debug ("%s:%s: Uncancellable BeforeAutoSave.",
                        SRCNAME, __func__);
             TBREAK;
           }

          if (m_mail->isCryptoMail() && !m_mail->decryptedSuccessfully ())
            {
              *(parms->rgvarg[0].pboolVal) = VARIANT_TRUE;
              log_debug ("%s:%s: Autosave for not successfuly decrypted mail."
                         "Cancel it.",
                         SRCNAME, __func__);
              TBREAK;
            }

          if (opt.draft_key && (m_mail->needs_crypto_m () & 1) &&
              !m_mail->isDraftEncrypt())
            {
              log_debug ("%s:%s: Draft encryption for autosave starting now.",
                         SRCNAME, __func__);
              m_mail->setIsDraftEncrypt (true);
              m_mail->prepareCrypto_o ();
            }
          TRETURN S_OK;
        }
      case Open:
        {
          log_oom ("%s:%s: Open : %p",
                         SRCNAME, __func__, m_mail);
          int draft_flags = 0;
          if (!opt.encrypt_default && !opt.sign_default)
            {
              TRETURN S_OK;
            }
          LPMESSAGE message = get_oom_base_message (m_object);
          if (!message)
            {
              log_error ("%s:%s: Failed to get message.",
                         SRCNAME, __func__);
              TBREAK;
            }
          if (opt.encrypt_default)
            {
              draft_flags = 1;
            }
          if (opt.sign_default)
            {
              draft_flags += 2;
            }
          set_gpgol_draft_info_flags (message, draft_flags);
          gpgol_release (message);
          TBREAK;
        }
      case BeforeRead:
        {
          log_oom ("%s:%s: BeforeRead : %p",
                         SRCNAME, __func__, m_mail);
          if (GpgolAddin::get_instance ()->isShutdown())
            {
              log_debug ("%s:%s: Ignoring read after shutdown.",
                         SRCNAME, __func__);
              TBREAK;
            }

          if (m_mail->preProcessMessage_m ())
            {
              log_error ("%s:%s: Pre process message failed.",
                         SRCNAME, __func__);
            }
          TBREAK;
        }
      case Read:
        {
          log_oom ("%s:%s: Read : %p",
                         SRCNAME, __func__, m_mail);
          if (!addinsLogged)
            {
              // We do it here as this nearly always comes and we want to remove
              // as much as possible from the startup time.
              log_addins ();
              addinsLogged = true;
            }
          if (!m_mail->isCryptoMail ())
            {
              log_debug ("%s:%s: Non crypto mail %p opened. Updating sigstatus.",
                         SRCNAME, __func__, m_mail);
              /* Ensure that no wrong sigstatus is shown */
              CloseHandle(CreateThread (NULL, 0, delayed_invalidate_ui, (LPVOID) 300, 0,
                                        NULL));
              TBREAK;
            }
          if (m_mail->setUUID_o ())
            {
              log_debug ("%s:%s: Failed to set uuid.",
                         SRCNAME, __func__);
              delete m_mail; /* deletes this, too */
              TRETURN S_OK;
            }
          if (m_mail->decryptVerify_o ())
            {
              log_error ("%s:%s: Decrypt message failed.",
                         SRCNAME, __func__);
            }
          if (!opt.enable_smime && m_mail->isSMIME_m ())
            {
              /* We want to save the mail when it's an smime mail and smime
                 is disabled to revert it. */
              log_debug ("%s:%s: S/MIME mail but S/MIME is disabled."
                         " Need save.",
                         SRCNAME, __func__);
              m_mail->setNeedsSave (true);
            }
          TBREAK;
        }
      case PropertyChange:
        {
          if (!parms || parms->cArgs != 1 ||
              parms->rgvarg[0].vt != VT_BSTR ||
              !parms->rgvarg[0].bstrVal)
            {
              log_error ("%s:%s: Unexpected params.",
                         SRCNAME, __func__);
              TBREAK;
            }
          const wchar_t *prop_name = parms->rgvarg[0].bstrVal;
          if (!m_mail->isCryptoMail ())
            {
              if (m_mail->hasOverrideMimeData())
                {
                  /* This is a mail created by us. Ignore propchanges. */
                  TBREAK;
                }
              if (!wcscmp (prop_name, L"To") /* ||
                  !wcscmp (prop_name, L"BCC") ||
                  !wcscmp (prop_name, L"CC")
                  Testing shows that Outlook always sends these three in a row
                  */)
                {
                  if (opt.autosecure || (m_mail->needs_crypto_m () & 1))
                    {
                      /* XXX Racy race. This is a fix for crashes
                         that happend if a resolved recipient is copied an pasted.
                         If we then access the recipients object in the Property
                         Change event we crash. Thus we do the delay dance. */
                      HANDLE thread = CreateThread (NULL, 0, do_delayed_locate,
                                                    (LPVOID) m_mail, 0,
                                                    NULL);
                      CloseHandle(thread);
                    }
                }
              TBREAK;
            }
          for (const wchar_t **cur = prop_blacklist; *cur; cur++)
            {
              if (!wcscmp (prop_name, *cur))
                {
                  log_oom ("%s:%s: Message %p propchange: %ls discarded.",
                           SRCNAME, __func__, m_object, prop_name);
                  TRETURN S_OK;
                }
            }
          log_oom ("%s:%s: Message %p propchange: %ls.",
                   SRCNAME, __func__, m_object, prop_name);

          if (!wcscmp (prop_name, L"SendUsingAccount"))
            {
              bool sent = get_oom_bool (m_object, "Sent");
              if (sent)
                {
                  log_debug ("%s:%s: Ignoring SendUsingAccount change for sent %p ",
                             SRCNAME, __func__, m_object);
                  TRETURN S_OK;
                }
              log_debug ("%s:%s: Message %p looks like send again.",
                        SRCNAME, __func__, m_object);
              m_mail->setIsSendAgain (true);
              TRETURN S_OK;
            }
          if (is_draft_mail (m_object))
            {
              log_oom ("%s:%s: Change allowed for draft",
                       SRCNAME, __func__);
              TRETURN S_OK;
            }

          /* We have tried several scenarios to handle propery changes.
             Only save the property in MAPI and call MAPI SaveChanges
             worked and did not leak plaintext but this caused outlook
             still to break the attachments of PGP/MIME Mails into two
             attachments and add them as winmail.dat so other clients
             are broken.

             Alternatively reverting the mail, saving the property and
             then decrypt again also worked a bit but there were some
             weird side effects and breakages. But this has the usual
             problem of a revert that the mail is created by outlook and
             e.g. multipart/signed signatures from most MUA's are broken.

             Some things to try out might be the close approach and then
             another open or a selection change. But for now we just warn.

             As a workardound a user should make property changes when
             the mail was not read by us. */
          if (propchangeWarnShown)
            {
              TRETURN S_OK;
            }

          wchar_t *title = utf8_to_wchar (_("Sorry, that's not possible, yet"));
          char *fmt;
          gpgrt_asprintf (&fmt, _("GpgOL has prevented the change to the \"%s\" property.\n"
                                  "Property changes are not yet handled for crypto messages.\n\n"
                                  "To workaround this limitation please change the property when the "
                                  "message is not open in any window and not selected in the "
                                  "messagelist.\n\nFor example by right clicking but not selecting the message.\n"),
                          wchar_to_utf8(prop_name));
          memdbg_alloc (fmt);
          wchar_t *msg = utf8_to_wchar (fmt);
          xfree (fmt);
          MessageBoxW (get_active_hwnd(), msg, title,
                       MB_ICONINFORMATION | MB_OK);
          xfree (msg);
          xfree (title);
          propchangeWarnShown = true;
          TRETURN S_OK;
        }
      case CustomPropertyChange:
        {
          log_oom ("%s:%s: CustomPropertyChange : %p",
                         SRCNAME, __func__, m_mail);
          /* TODO */
          TBREAK;
        }
      case Send:
        {
          /* This is the only event where we can cancel the send of a
             mailitem. But it is too early for us to encrypt as the MAPI
             structures are not yet filled. Crypto based on the
             Outlook Object Model data did not work as the messages
             were only sent out empty. See 2b376a48 for a try of
             this.

             This is why we store send_seen and invoke a save which
             may result in an error but only after triggering all the
             behavior we need -> filling mapi structures and invoking the
             AfterWrite handler where we encrypt.

             If this encryption is successful and we pass the send
             as then the encrypted data is sent.
           */
          log_oom ("%s:%s: Send : %p",
                         SRCNAME, __func__, m_mail);
          if (!m_mail->needs_crypto_m () && m_mail->cryptState () == Mail::NoCryptMail)
            {
             log_debug ("%s:%s: No crypto neccessary. Passing send for %p obj %p",
                        SRCNAME, __func__, m_mail, m_object);
             TBREAK;
            }

          if (parms->cArgs != 1 || parms->rgvarg[0].vt != (VT_BOOL | VT_BYREF))
           {
             log_debug ("%s:%s: Uncancellable send event.",
                        SRCNAME, __func__);
             TBREAK;
           }

          if (m_mail->cryptState () == Mail::NoCryptMail &&
              m_mail->needs_crypto_m ())
            {
              log_debug ("%s:%s: Send event for crypto mail %p saving and starting.",
                         SRCNAME, __func__, m_mail);

              m_mail->prepareCrypto_o ();

              // Save the Mail
              invoke_oom_method (m_object, "Save", NULL);

              if (!m_mail->isAsyncCryptDisabled ())
                {
                  // The afterwrite in the save should have triggered
                  // the encryption. We cancel send for our asyncness.
                  // Cancel send
                  *(parms->rgvarg[0].pboolVal) = VARIANT_TRUE;
                  TBREAK;
                }
              else
                {
                  if (m_mail->cryptState () == Mail::NoCryptMail)
                    {
                      // Crypto failed or was canceled
                      log_debug ("%s:%s: Message %p mail %p cancelling send - "
                                 "Crypto failed or canceled.",
                                 SRCNAME, __func__, m_object, m_mail);
                      *(parms->rgvarg[0].pboolVal) = VARIANT_TRUE;
                      TBREAK;
                    }
                  // For inline response we can't trigger send programatically
                  // so we do the encryption in sync.
                  if (m_mail->cryptState () == Mail::NeedsUpdateInOOM)
                    {
                      m_mail->updateCryptOOM_o ();
                    }
                  if (m_mail->cryptState () == Mail::NeedsSecondAfterWrite)
                    {
                      m_mail->setCryptState (Mail::WantsSendMIME);
                    }
                  if (m_mail->getDoPGPInline () && m_mail->cryptState () != Mail::WantsSendInline)
                    {
                      log_debug ("%s:%s: Message %p mail %p cancelling send - "
                                 "Invalid state.",
                                 SRCNAME, __func__, m_object, m_mail);
                      gpgol_bug (m_mail->getWindow (),
                                 ERR_INLINE_BODY_INV_STATE);
                      *(parms->rgvarg[0].pboolVal) = VARIANT_TRUE;
                      TBREAK;
                    }
                }
            }

          if (m_mail->cryptState () == Mail::WantsSendInline)
            {
              if (!m_mail->hasCryptedOrEmptyBody_o ())
                {
                  log_debug ("%s:%s: Message %p mail %p cancelling send - "
                             "not encrypted or not empty body detected.",
                             SRCNAME, __func__, m_object, m_mail);
                  gpgol_bug (m_mail->getWindow (),
                             ERR_WANTS_SEND_INLINE_BODY);
                  m_mail->setCryptState (Mail::NoCryptMail);
                  *(parms->rgvarg[0].pboolVal) = VARIANT_TRUE;
                  TBREAK;
                }
              log_debug ("%s:%s: Passing send event for no-mime message %p.",
                         SRCNAME, __func__, m_object);
              WKSHelper::instance()->allow_notify (1000);
              TBREAK;
            }

          if (m_mail->cryptState () == Mail::WantsSendMIME)
            {
              if (!m_mail->hasCryptedOrEmptyBody_o ())
                {
/* The safety checks here trigger too often. Somehow for some
   users the body is not empty after the encryption but when
   it is sent it is still sent with the crypto content because
   the encrypted MIME Structure is used because it is
   correct in MAPI land.

   For safety reasons enabling the checks might be better but
   until we figure out why for some users the body replacement
   does not work we have to disable them. Otherwise GpgOL
   is unusuable for such users. GnuPG-Bug-Id: T3875
*/
#define DISABLE_SAFTEY_CHECKS
#ifndef DISABLE_SAFTEY_CHECKS
                  gpgol_bug (m_mail->getWindow (),
                             ERR_WANTS_SEND_MIME_BODY);
                  log_debug ("%s:%s: Message %p mail %p cancelling send mime - "
                             "not encrypted or not empty body detected.",
                             SRCNAME, __func__, m_object, m_mail);
                  m_mail->setCryptState (Mail::NoCryptMail);
                  *(parms->rgvarg[0].pboolVal) = VARIANT_TRUE;
                  TBREAK;
#else
                  log_debug ("%s:%s: Message %p mail %p - "
                             "not encrypted or not empty body detected - MIME.",
                             SRCNAME, __func__, m_object, m_mail);
#endif
                }
              /* Now we adress T3656 if Outlooks internal S/MIME is somehow
               * mixed in (even if it is enabled and then disabled) it might
               * cause strange behavior in that it sends the plain message
               * and not the encrypted message. Tests have shown that we can
               * bypass that by calling submit message on our base
               * message.
               *
               * We do this conditionally as our other way of using OOM
               * to send is proven to work and we don't want to mess
               * with it.
               */
              // Get the Message class.
              HRESULT hr;
              LPSPropValue propval = NULL;

              // It's important we use the _not_ base message here.
              LPMESSAGE message = get_oom_message (m_object);
              hr = HrGetOneProp ((LPMAPIPROP)message, PR_MESSAGE_CLASS_A, &propval);
              gpgol_release (message);
              if (FAILED (hr))
                {
                  log_error ("%s:%s: HrGetOneProp() failed: hr=%#lx\n",
                             SRCNAME, __func__, hr);
                  gpgol_release (message);
                  TBREAK;
                }
              if (propval->Value.lpszA && !strstr (propval->Value.lpszA, "GpgOL"))
                {
                  // Does not have a message class by us.
                  log_debug ("%s:%s: Message %p - No GpgOL Message class after encryption. cls is: '%s'",
                             SRCNAME, __func__, m_object, propval->Value.lpszA);
                  log_debug ("%s:%s: Message %p - Activating T3656 Workaround",
                             SRCNAME, __func__, m_object);
                  message = get_oom_base_message (m_object);
                  if (message)
                    {
                      // It's important we use the _base_ message here.
                      mapi_save_changes (message, FORCE_SAVE);
                      message->SubmitMessage(0);
                      gpgol_release (message);
                      // Close the composer and trigger unloads
                      CloseHandle(CreateThread (NULL, 0, close_mail, (LPVOID) m_mail, 0,
                                                NULL));
                    }
                  else
                    {
                      gpgol_bug (nullptr,
                                 ERR_GET_BASE_MSG_FAILED);
                    }
                  // Cancel send
                  *(parms->rgvarg[0].pboolVal) = VARIANT_TRUE;
                }
              MAPIFreeBuffer (propval);
              if (*(parms->rgvarg[0].pboolVal) == VARIANT_TRUE)
                {
                  TBREAK;
                }
              log_debug ("%s:%s: Passing send event for mime-encrypted message %p.",
                         SRCNAME, __func__, m_object);
              WKSHelper::instance()->allow_notify (1000);
              TBREAK;
            }
          else
            {
              log_debug ("%s:%s: Message %p cancelling send - "
                         "crypto or second save failed.",
                         SRCNAME, __func__, m_object);
              *(parms->rgvarg[0].pboolVal) = VARIANT_TRUE;
            }
          TRETURN S_OK;
        }
      case Write:
        {
          log_oom ("%s:%s: Write : %p",
                         SRCNAME, __func__, m_mail);
          /* This is a bit strange. We sometimes get multiple write events
             without a read in between. When we access the message in
             the second event it fails and if we cancel the event outlook
             crashes. So we have keep the m_needs_wipe state variable
             to keep track of that. */
          if (parms->cArgs != 1 || parms->rgvarg[0].vt != (VT_BOOL | VT_BYREF))
           {
             /* This happens in the weird case */
             log_debug ("%s:%s: Uncancellable write event.",
                        SRCNAME, __func__);
             TBREAK;
           }

          if (m_mail->passWrite())
            {
              log_debug ("%s:%s: Passing write because passNextWrite was set for %p",
                         SRCNAME, __func__, m_mail);
              TBREAK;
            }

          if (m_mail->isCryptoMail () && !m_mail->needsSave ())
            {
              if (opt.draft_key && (m_mail->needs_crypto_m () & 1) &&
                  is_draft_mail (m_object) && m_mail->decryptedSuccessfully ())
                {
                  if (m_mail->cryptState () == Mail::NeedsFirstAfterWrite ||
                      m_mail->cryptState () == Mail::NeedsSecondAfterWrite)
                    {
                      log_debug ("%s:%s: re-encryption in progress. Passing.",
                                 SRCNAME, __func__);
                      TBREAK;
                    }
                  /* This is the case for a modified draft */
                  log_debug ("%s:%s: Draft re-encryption starting now.",
                             SRCNAME, __func__);
                  m_mail->setIsDraftEncrypt (true);
                  m_mail->prepareCrypto_o ();
                  /* Passing write to trigger encrypt in after write */
                  TBREAK;
                }

              Mail *last_mail = Mail::getLastMail ();
              if (Mail::isValidPtr (last_mail))
                {
                  /* We want to identify here if there was a mail created that
                     should receive the contents of this mail. For this we check
                     for a write in the same loop as a mail creation.
                     Now when switching from one mail to another this is also what
                     happens. The new mail is loaded and the old mail is written.
                     To distinguish the two we check that the new mail does not have
                     an entryID, a Subject and No Size. Maybe just size or entryID
                     would be enough but better save then sorry.

                     Security consideration: Worst case we pass the write here but
                     an unload follows before we get the scheduled revert. This
                     would leak plaintext. But does not happen in our tests.

                     Similarly if we crash or Outlook is closed before we see this
                     revert. But as we immediately revert after the write this should
                     also not happen. */
                  const std::string lastSubject = last_mail->getSubject_o ();
                  char *lastEntryID = get_oom_string (last_mail->item (), "EntryID");
                  int lastSize = get_oom_int (last_mail->item (), "Size");
                  std::string lastEntryStr;
                  if (lastEntryID)
                    {
                      lastEntryStr = lastEntryID;
                      xfree (lastEntryID);
                    }

                  if (!lastSize && !lastEntryStr.size () && !lastSubject.size ())
                    {
                      log_debug ("%s:%s: Write in the same loop as empty load."
                                 " Pass but schedule revert.",
                                 SRCNAME, __func__);

                      /* This might be a forward. So don't invalidate yet. */

                      // Mail::clearLastMail ();

                      do_in_ui_thread_async (REVERT_MAIL, m_mail);
                      TRETURN S_OK;
                    }
                }
              /* We cancel the write event to stop outlook from excessively
                 syncing our changes.
                 if smime support is disabled and we still have an smime
                 mail we also don't want to cancel the write event
                 to enable reverting this mails.
                 */
              *(parms->rgvarg[0].pboolVal) = VARIANT_TRUE;
              log_debug ("%s:%s: Canceling write event.",
                         SRCNAME, __func__);
              TRETURN S_OK;
            }

          if (m_mail->isCryptoMail () && m_mail->needsSave () &&
              m_mail->revert_o ())
            {
              /* An error cleaning the mail should not happen normally.
                 But just in case there is an error we cancel the
                 write here. */
              log_debug ("%s:%s: Failed to remove plaintext. Canceling.",
                         SRCNAME, __func__);
              *(parms->rgvarg[0].pboolVal) = VARIANT_TRUE;
              TRETURN S_OK;
            }

          if (!m_mail->isCryptoMail () && m_mail->is_forwarded_crypto_mail () &&
              !m_mail->needs_crypto_m () && m_mail->cryptState () == Mail::NoCryptMail)
            {
              /* We are sure now that while this is a forward of an encrypted
               * mail that the forward should not be signed or encrypted. So
               * it's not constructed by us. We need to remove our attachments
               * though so that they are not included in the forward. */
              log_debug ("%s:%s: Writing unencrypted forward of crypt mail. "
                         "Removing attachments. mail: %p item: %p",
                         SRCNAME, __func__, m_mail, m_object);
              if (m_mail->removeOurAttachments_o ())
                {
                  // Worst case we forward some encrypted data here not
                  // a security problem, so let it pass.
                  log_error ("%s:%s: Failed to remove our attachments.",
                             SRCNAME, __func__);
                }
              /* Remove marker because we did this now. */
              m_mail->setIsForwardedCryptoMail (false);
            }

          if (m_mail->isDraftEncrypt () &&
              m_mail->cryptState () != Mail::NeedsFirstAfterWrite &&
              m_mail->cryptState () != Mail::NeedsSecondAfterWrite)
            {
              log_debug ("%s:%s: Canceling write because draft encrypt is in"
                         " progress.",
                         SRCNAME, __func__);
              *(parms->rgvarg[0].pboolVal) = VARIANT_TRUE;
              TRETURN S_OK;
            }

          if (opt.draft_key && (m_mail->needs_crypto_m () & 1) &&
              !m_mail->isDraftEncrypt())
            {
              log_debug ("%s:%s: Draft encryption starting now.",
                         SRCNAME, __func__);
              m_mail->setIsDraftEncrypt (true);
              m_mail->prepareCrypto_o ();
            }

          log_debug ("%s:%s: Passing write event. %i %i",
                     SRCNAME, __func__, m_mail->isDraftEncrypt(), m_mail->cryptState());
          m_mail->setNeedsSave (false);
          TBREAK;
        }
      case AfterWrite:
        {
          log_oom ("%s:%s: AfterWrite : %p",
                         SRCNAME, __func__, m_mail);
          if (m_mail->cryptState () == Mail::NeedsFirstAfterWrite)
            {
              /* Seen the first after write. Advance the state */
              m_mail->setCryptState (Mail::NeedsActualCrypt);
              if (m_mail->encryptSignStart_o ())
                {
                  log_debug ("%s:%s: Encrypt sign start failed.",
                             SRCNAME, __func__);
                  m_mail->setCryptState (Mail::NoCryptMail);
                }
              TRETURN S_OK;
            }
          if (m_mail->cryptState () == Mail::NeedsSecondAfterWrite)
            {
              m_mail->setCryptState (Mail::NeedsUpdateInMAPI);
              m_mail->updateCryptMAPI_m ();
              log_debug ("%s:%s: Second after write done.",
                         SRCNAME, __func__);
              TRETURN S_OK;
            }
          TBREAK;
        }
      case Close:
        {
          log_oom ("%s:%s: Close : %p",
                         SRCNAME, __func__, m_mail);
          if (m_mail->isCryptoMail () && !is_draft_mail (m_object))
            {
              /* Close. This happens when an Opened mail is closed.
                 To prevent the question of wether or not to save the changes
                 (Which would save the decrypted data without an event to
                 prevent it) we cancel the close and then either close it
                 with discard changes or revert / save it.
                 Contrary to documentation we can invoke close from
                 close.
                 */
              if (parms->cArgs != 1 || parms->rgvarg[0].vt != (VT_BOOL | VT_BYREF))
                {
                  /* This happens in the weird case */
                  log_debug ("%s:%s: Uncancellable close event.",
                             SRCNAME, __func__);
                  TBREAK;
                }
              if (m_mail->getCloseTriggered ())
                {
                  /* Our close with discard changes, pass through */
                  m_mail->setCloseTriggered (false);
                  TRETURN S_OK;
                }
              *(parms->rgvarg[0].pboolVal) = VARIANT_TRUE;
              log_oom ("%s:%s: Canceling close event.",
                             SRCNAME, __func__);
              if (Mail::close(m_mail))
                {
                  log_debug ("%s:%s: Close request failed.",
                             SRCNAME, __func__);
                }
            }
          TRETURN S_OK;
        }
      case Unload:
        {
          log_oom ("%s:%s: Unload : %p",
                         SRCNAME, __func__, m_mail);
          log_debug ("%s:%s: Removing Mail for message: %p.",
                     SRCNAME, __func__, m_object);
          delete m_mail;
          log_oom ("%s:%s: deletion done",
                         SRCNAME, __func__);
          memdbg_dump ();
          TRETURN S_OK;
        }
      case ReplyAll:
      case Reply:
          is_reply = true;
          /* fall through */
      case Forward:
        {
          log_oom ("%s:%s: %s : %p",
                         SRCNAME, __func__, is_reply ? "reply" : "forward", m_mail);
          int draft_flags = 0;
          if (opt.encrypt_default)
            {
              draft_flags = 1;
            }
          if (opt.sign_default)
            {
              draft_flags += 2;
            }
          bool is_crypto_mail = m_mail->isCryptoMail ();

          /* If it is a crypto mail and the settings should not be taken
           * from the crypto mail and always encrypt / sign is on. Or
           * If it is not a crypto mail and we have automaticalls sign_encrypt. */
          if ((is_crypto_mail && !opt.reply_crypt && draft_flags) ||
              (!is_crypto_mail && draft_flags))
            {
              /* Check if we can use the dispval */
                if (parms->cArgs == 2 && parms->rgvarg[1].vt == (VT_DISPATCH) &&
                    parms->rgvarg[0].vt == (VT_BOOL | VT_BYREF))
                {
                  LPMESSAGE msg = get_oom_base_message (parms->rgvarg[1].pdispVal);
                  if (msg)
                    {
                      set_gpgol_draft_info_flags (msg, draft_flags);
                      gpgol_release (msg);
                    }
                  else
                    {
                      log_error ("%s:%s: Failed to get base message.",
                                 SRCNAME, __func__);
                    }
                }
              else
                {
                  log_error ("%s:%s: Unexpected parameters.",
                             SRCNAME, __func__);
                }
            }

          if (!is_crypto_mail)
            {
              /* Replys to non crypto mails do not interest us anymore. */
              TBREAK;
            }

          Mail *last_mail = Mail::getLastMail ();
          if (Mail::isValidPtr (last_mail))
            {
              /* We want to identify here if there was a mail created that
                 should receive the contents of this mail. For this we check
                 for a forward in the same loop as a mail creation.

                 We need to do it this complicated and can't just use
                 get_mail_for_item because the mailitem pointer we get here
                 is a different one then the one with which the mail was loaded.
              */
              char *lastEntryID = get_oom_string (last_mail->item (), "EntryID");
              int lastSize = get_oom_int (last_mail->item (), "Size");
              std::string lastEntryStr;
              if (lastEntryID)
                {
                  lastEntryStr = lastEntryID;
                  xfree (lastEntryID);
                }

              if (!lastSize && !lastEntryStr.size ())
                {
                  if (!is_reply)
                    {
                      log_debug ("%s:%s: Forward in the same loop as empty "
                                 "load Marking %p (item %p) as forwarded.",
                                 SRCNAME, __func__, last_mail,
                                 last_mail->item ());

                      last_mail->setIsForwardedCryptoMail (true);
                    }
                  else
                    {
                      log_debug ("%s:%s: Reply in the same loop as empty "
                                 "load Marking %p (item %p) as reply.",
                                 SRCNAME, __func__, last_mail,
                                 last_mail->item ());
                    }
                  if (m_mail->isBlockHTML ())
                    {
                      std::string buf;
                      /** TRANSLATORS: Part of a warning dialog that disallows
                        reply and forward with contents */
                      buf = is_reply ? _("You are replying to an unsigned S/MIME "
                                         "email.") :
                                       _("You are forwarding an unsigned S/MIME "
                                         "email.");
                      buf +="\n\n";
                      buf += _("In this version of S/MIME an attacker could "
                               "use the missing signature to have you "
                               "decrypt contents from a different, otherwise "
                               "completely unrelated email and place it in the "
                               "quote so they can get hold of it.\n"
                               "This is why we only allow quoting to be done manually.");
                      buf += "\n\n";
                      buf += _("Please copy the relevant contents and insert "
                               "them into the new email.");

                      gpgol_message_box (get_active_hwnd (), buf.c_str(),
                                         _("GpgOL"), MB_OK);

                      do_in_ui_thread_async (CLEAR_REPLY_FORWARD, last_mail, 1000);
                    }
                }
              // We can now invalidate the last mail
              Mail::clearLastMail ();
            }

          log_oom ("%s:%s: Reply Forward ReplyAll: %p",
                         SRCNAME, __func__, m_mail);
          if (!opt.reply_crypt)
            {
              TBREAK;
            }
          int crypto_flags = 0;
          if (!(crypto_flags = m_mail->getCryptoFlags ()))
            {
              TBREAK;
            }
          if (parms->cArgs != 2 || parms->rgvarg[1].vt != (VT_DISPATCH) ||
              parms->rgvarg[0].vt != (VT_BOOL | VT_BYREF))
            {
              /* This happens in the weird case */
              log_debug ("%s:%s: Unexpected args %i %x %x named: %i",
                         SRCNAME, __func__, parms->cArgs, parms->rgvarg[0].vt, parms->rgvarg[1].vt,
                         parms->cNamedArgs);
              TBREAK;
            }
          LPMESSAGE msg = get_oom_base_message (parms->rgvarg[1].pdispVal);
          if (!msg)
            {
              log_debug ("%s:%s: Failed to get base message",
                         SRCNAME, __func__);
              TBREAK;
            }
          set_gpgol_draft_info_flags (msg, crypto_flags);
          gpgol_release (msg);
          TBREAK;
        }
      case AttachmentRemove:
        {
          log_oom ("%s:%s: AttachmentRemove: %p",
                         SRCNAME, __func__, m_mail);
          if (!m_mail->isCryptoMail () || attachRemoveWarnShown ||
              m_mail->attachmentRemoveWarningDisabled ())
            {
              TRETURN S_OK;
            }
          gpgol_message_box (get_active_hwnd (),
                             _("Attachments are part of the crypto message.\nThey "
                               "can't be permanently removed and will be shown again the next "
                               "time this message is opened."),
                             _("Sorry, that's not possible, yet"), MB_OK);
          attachRemoveWarnShown = true;
          TRETURN S_OK;
        }

      default:
        log_oom ("%s:%s: Message:%p Unhandled Event: %lx \n",
                       SRCNAME, __func__, m_object, dispid);
    }
  TRETURN S_OK;
}
END_EVENT_SINK(MailItemEvents, IID_MailItemEvents)
