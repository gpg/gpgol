/* mailitem-events.h - Event handling for mails.
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

#include "config.h"
#include "common.h"
#include "eventsink.h"
#include "eventsinks.h"
#include "mymapi.h"
#include "oomhelp.h"
#include "ocidl.h"
#include "windowmessages.h"
#include "mail.h"
#include "mapihelp.h"

const wchar_t *prop_blacklist[] = {
  L"Body",
  L"HTMLBody",
  L"To", /* Somehow this is done when a mail is opened */
  L"CC", /* Ditto */
  L"BCC", /* Ditto */
  L"Categories",
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
    Reply = 0xFC8F,
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
  bool m_send_seen;   /* The message is about to be submitted */
  bool m_decrypt_after_write;
  bool m_ignore_unloads;
  bool m_ignore_next_unload;
};

MailItemEvents::MailItemEvents() :
    m_object(NULL),
    m_pCP(NULL),
    m_cookie(0),
    m_ref(1),
    m_mail(NULL),
    m_send_seen (false),
    m_decrypt_after_write(false),
    m_ignore_unloads(false),
    m_ignore_next_unload(false)
{
}

MailItemEvents::~MailItemEvents()
{
  if (m_pCP)
    m_pCP->Unadvise(m_cookie);
  if (m_object)
    gpgol_release (m_object);
}

static DWORD WINAPI
request_decrypt (LPVOID arg)
{
  log_debug ("%s:%s: requesting decrypt again for: %p",
             SRCNAME, __func__, arg);
  if (do_in_ui_thread (REQUEST_DECRYPT, arg))
    {
      log_debug ("%s:%s: second decrypt failed for: %p",
                 SRCNAME, __func__, arg);
    }
  return 0;
}

static DWORD WINAPI
request_close (LPVOID arg)
{
  log_debug ("%s:%s: requesting close for: %s",
             SRCNAME, __func__, (char*) arg);
  if (do_in_ui_thread (REQUEST_CLOSE, arg))
    {
      log_debug ("%s:%s: close request failed for: %s",
                 SRCNAME, __func__, (char*) arg);
    }
  return 0;
}

static bool propchangeWarnShown = false;

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
  if (!m_mail)
    {
      m_mail = Mail::get_mail_for_item (m_object);
      if (!m_mail)
        {
          log_error ("%s:%s: mail event without mail object known. Bug.",
                     SRCNAME, __func__);
          return S_OK;
        }
    }
  switch(dispid)
    {
      case Open:
        {
          log_oom_extra ("%s:%s: Open : %p",
                         SRCNAME, __func__, m_mail);
          LPMESSAGE message;
          int draft_flags = 0;
          if (!opt.encrypt_default && !opt.sign_default)
            {
              return S_OK;
            }
          message = get_oom_base_message (m_object);
          if (!message)
            {
              log_error ("%s:%s: Failed to get message.",
                         SRCNAME, __func__);
              break;
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

          if (m_mail->is_crypto_mail())
            {
              m_ignore_unloads = true;
            }
        }
      case BeforeRead:
        {
          log_oom_extra ("%s:%s: BeforeRead : %p",
                         SRCNAME, __func__, m_mail);
          if (m_mail->pre_process_message ())
            {
              log_error ("%s:%s: Pre process message failed.",
                         SRCNAME, __func__);
            }
          break;
        }
      case Read:
        {
          log_oom_extra ("%s:%s: Read : %p",
                         SRCNAME, __func__, m_mail);
          if (!m_mail->is_crypto_mail())
            {
              /* Not for us. */
              break;
            }
          if (m_mail->set_uuid ())
            {
              log_debug ("%s:%s: Failed to set uuid.",
                         SRCNAME, __func__);
              delete m_mail; /* deletes this, too */
              return S_OK;
            }
          if (m_mail->decrypt_verify ())
            {
              log_error ("%s:%s: Decrypt message failed.",
                         SRCNAME, __func__);
            }
          if (!opt.enable_smime && m_mail->is_smime ())
            {
              /* We want to save the mail when it's an smime mail and smime
                 is disabled to revert it. */
              m_mail->set_needs_save (true);
            }
          break;
        }
      case PropertyChange:
        {
          const wchar_t *prop_name;
          if (!m_mail->is_crypto_mail ())
            {
              break;
            }
          if (!parms || parms->cArgs != 1 ||
              parms->rgvarg[0].vt != VT_BSTR ||
              !parms->rgvarg[0].bstrVal)
            {
              log_error ("%s:%s: Unexpected params.",
                         SRCNAME, __func__);
              break;
            }
          prop_name = parms->rgvarg[0].bstrVal;
          for (const wchar_t **cur = prop_blacklist; *cur; cur++)
            {
              if (!wcscmp (prop_name, *cur))
                {
                  log_oom ("%s:%s: Message %p propchange: %ls discarded.",
                           SRCNAME, __func__, m_object, prop_name);
                  return S_OK;
                }
            }
          log_oom ("%s:%s: Message %p propchange: %ls.",
                   SRCNAME, __func__, m_object, prop_name);

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

             Close -> discard changes -> then setting the property and
             then saving also works but then the mail is closed / unloaded
             and we can't decrypt again.

             Some things to try out might be the close approach and then
             another open or a selection change. But for now we just warn.

             As a workardound a user should make property changes when
             the mail was not read by us. */
          if (propchangeWarnShown)
            {
              return S_OK;
            }

          wchar_t *title = utf8_to_wchar (_("Sorry, that's not possible, yet"));
          char *fmt;
          gpgrt_asprintf (&fmt, _("GpgOL has prevented the change to the \"%s\" property.\n"
                                  "Property changes are not yet handled for crypto messages.\n\n"
                                  "To workaround this limitation please change the property when the"
                                  "message is not open in any window and not selected in the"
                                  "messagelist.\n\nFor example by right clicking but not selecting the message.\n"),
                          wchar_to_utf8(prop_name));
          wchar_t *msg = utf8_to_wchar (fmt);
          xfree (fmt);
          MessageBoxW (get_active_hwnd(), msg, title,
                       MB_ICONINFORMATION | MB_OK);
          xfree (msg);
          xfree (title);
          propchangeWarnShown = true;
          return S_OK;
        }
      case CustomPropertyChange:
        {
          log_oom_extra ("%s:%s: CustomPropertyChange : %p",
                         SRCNAME, __func__, m_mail);
          /* TODO */
          break;
        }
      case Send:
        {
          /* This is the only event where we can cancel the send of an
             mailitem. But it is too early for us to encrypt as the MAPI
             structures are not yet filled. Crypto based on the
             Outlook Object Model data did not work as the messages
             were only sent out empty. See 2b376a48 for a try of
             this.

             The we store send_seend and invoke a save which will result
             in an error but only after triggering all the behavior
             we need -> filling mapi structures and invoking the
             AfterWrite handler where we encrypt.

             If this encryption is successful and we pass the send
             as then the encrypted data is sent.
           */
          log_oom_extra ("%s:%s: Send : %p",
                         SRCNAME, __func__, m_mail);
          if (parms->cArgs != 1 || parms->rgvarg[0].vt != (VT_BOOL | VT_BYREF))
           {
             log_debug ("%s:%s: Uncancellable send event.",
                        SRCNAME, __func__);
             break;
           }
          m_mail->update_sender ();
          m_send_seen = true;
          invoke_oom_method (m_object, "Save", NULL);
          if (m_mail->crypto_successful ())
            {
               log_debug ("%s:%s: Passing send event for message %p.",
                          SRCNAME, __func__, m_object);
               m_send_seen = false;
               break;
            }
          else
            {
              log_debug ("%s:%s: Message %p cancelling send crypto failed.",
                         SRCNAME, __func__, m_object);
              *(parms->rgvarg[0].pboolVal) = VARIANT_TRUE;
            }
          return S_OK;
        }
      case Write:
        {
          log_oom_extra ("%s:%s: Write : %p",
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
             break;
           }

          if (m_mail->is_crypto_mail () && !m_mail->needs_save ())
            {
              /* We cancel the write event to stop outlook from excessively
                 syncing our changes.
                 if smime support is disabled and we still have an smime
                 mail we also don't want to cancel the write event
                 to enable reverting this mails.
                 */
              *(parms->rgvarg[0].pboolVal) = VARIANT_TRUE;
              log_debug ("%s:%s: Canceling write event.",
                         SRCNAME, __func__);
              return S_OK;
            }
          log_debug ("%s:%s: Passing write event.",
                     SRCNAME, __func__);
          m_mail->set_needs_save (false);
          break;
        }
      case AfterWrite:
        {
          log_oom_extra ("%s:%s: AfterWrite : %p",
                         SRCNAME, __func__, m_mail);
          if (m_send_seen)
            {
              m_send_seen = false;
              m_mail->encrypt_sign ();
              return S_OK;
            }
          else if (m_decrypt_after_write)
            {
              char *uuid = strdup (m_mail->get_uuid ().c_str());
              HANDLE thread = CreateThread (NULL, 0, request_decrypt,
                                            (LPVOID) uuid, 0, NULL);
              CloseHandle (thread);
              m_decrypt_after_write = false;
            }
          break;
        }
      case Close:
        {
          log_oom_extra ("%s:%s: Close : %p",
                         SRCNAME, __func__, m_mail);
          if (m_mail->is_crypto_mail ())
            {
              /* Close. This happens when an Opened mail is closed.
                 To prevent the question of wether or not to save the changes
                 (Which would save the decrypted data without an event to
                 prevent it) we cancel the close and then either close it
                 with discard changes or revert / save it.
                 This happens with a window message as we can't invoke close from
                 close.

                 But as a side effect the mail, if opened in the explorer still will
                 be reverted, too. So shown as empty. To prevent that
                 we request a decrypt in the AfterWrite event which checks if the
                 message is opened in the explorer. If not it destroys the mail.

                 Evil Hack: Outlook sends an Unload event after the message is closed
                 This is not true our Internal Object is kept alive if it is opened
                 in the explorer. So we ignore the unload event and then check in
                 the window message handler that checks for decrypt again if the
                 mail is currently open in the active explorer. If not we delete our
                 Mail object so that the message is released.
                 */
              if (parms->cArgs != 1 || parms->rgvarg[0].vt != (VT_BOOL | VT_BYREF))
                {
                  /* This happens in the weird case */
                  log_debug ("%s:%s: Uncancellable close event.",
                             SRCNAME, __func__);
                  break;
                }
              *(parms->rgvarg[0].pboolVal) = VARIANT_TRUE;
              log_oom_extra ("%s:%s: Canceling close event.",
                             SRCNAME, __func__);
              m_decrypt_after_write = true;
              m_ignore_unloads = false;
              m_ignore_next_unload = true;

              char *uuid = strdup (m_mail->get_uuid ().c_str());
              HANDLE thread = CreateThread (NULL, 0, request_close,
                                            (LPVOID) uuid, 0, NULL);
              CloseHandle (thread);
            }
        }
      case Unload:
        {
          log_oom_extra ("%s:%s: Unload : %p",
                         SRCNAME, __func__, m_mail);
          /* Unload. Experiments have shown that this does not
             mean a mail is actually unloaded in Outlook. E.g.
             If it was open in an inspector and then closed we
             see an unload event but the mail is still shown in
             the explorer. Fun. On the other hand if a message
             was opened and the explorer selection changes
             we also get an unload but the mail is still open.

             Really we still get events after the unload and
             can make changes to the object.

             In case the mail was opened m_ignore_unloads is set
             to true so the mail is not removed when the message
             selection changes. As close invokes decrypt_again
             the mail object is removed there when the explorer
             selection changed.

             In case the mail was closed m_ignore_next_unload
             is set so only the Unload thad follows the canceled
             close is ignored and not the unload that comes from
             our then triggered close (save / discard).


             This is horribly hackish and feels wrong. But it
             works.
             */
          if (m_ignore_unloads || m_ignore_next_unload)
            {
              if (m_ignore_next_unload)
                {
                  m_ignore_next_unload = false;
                }
              log_debug ("%s:%s: Ignoring unload for message: %p.",
                         SRCNAME, __func__, m_object);
            }
          else
            {
              log_debug ("%s:%s: Removing Mail for message: %p.",
                         SRCNAME, __func__, m_object);
              delete m_mail;
            }
          return S_OK;
        }
      default:
        log_oom_extra ("%s:%s: Message:%p Unhandled Event: %lx \n",
                       SRCNAME, __func__, m_object, dispid);
    }
  return S_OK;
}
END_EVENT_SINK(MailItemEvents, IID_MailItemEvents)
