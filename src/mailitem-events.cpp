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
};

MailItemEvents::MailItemEvents() :
    m_object(NULL),
    m_pCP(NULL),
    m_cookie(0),
    m_ref(1),
    m_mail(NULL),
    m_send_seen (false)
{
}

MailItemEvents::~MailItemEvents()
{
  if (m_pCP)
    m_pCP->Unadvise(m_cookie);
  if (m_object)
    m_object->Release();
}

static DWORD WINAPI
request_send (LPVOID arg)
{
  log_debug ("%s:%s: requesting send for: %p",
             SRCNAME, __func__, arg);
  if (do_in_ui_thread (REQUEST_SEND_MAIL, arg))
    {
      MessageBox (NULL,
                  "Error while requesting send of message.\n"
                  "Please press the send button again.",
                  _("GpgOL"),
                  MB_ICONINFORMATION|MB_OK);
    }
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
      case BeforeRead:
        {
          if (m_mail->process_message ())
            {
              log_error ("%s:%s: Process message failed.",
                         SRCNAME, __func__);
            }
          break;
        }
      case Read:
        {
          if (m_mail->insert_plaintext ())
            {
              log_error ("%s:%s: Failed to insert plaintext into oom.",
                         SRCNAME, __func__);
            }
          break;
        }
      case Send:
        {
          /* This is the only event where we can cancel the send of an
             mailitem. But it is too early for us to encrypt as the MAPI
             structures are not yet filled (and we don't seem to have a way
             to trigger this and it is likely to be impossible)

             So the first send event is canceled but we save that we have
             seen it in m_send_seen. We then trigger a Save of that item.
             The Save causes the Item to be written and we have a chance
             to Encrypt it in the AfterWrite event.

             If this encryption is successful and we see a send again
             we let it pass as then the encrypted data is sent.

             The value of m_send_seen is set to false in this case as
             we consumed the original send that we canceled. */
          if (parms->cArgs != 1 || parms->rgvarg[0].vt != (VT_BOOL | VT_BYREF))
           {
             log_debug ("%s:%s: Uncancellable send event.",
                        SRCNAME, __func__);
             break;
           }
          if (m_mail->crypto_successful ())
            {
               log_debug ("%s:%s: Passing send event for message %p.",
                          SRCNAME, __func__, m_object);
               m_send_seen = false;
               break;
            }
          m_mail->update_sender ();
          m_send_seen = true;
          log_debug ("%s:%s: Message %p cancelling send to let us do crypto.",
                     SRCNAME, __func__, m_object);
          *(parms->rgvarg[0].pboolVal) = VARIANT_TRUE;
          invoke_oom_method (m_object, "Save", NULL);

          return S_OK;
        }
      case Write:
        {
          /* This is a bit strange. We sometimes get multiple write events
             without a read in between. When we access the message in
             the second event it fails and if we cancel the event outlook
             crashes. So we have keep the m_needs_wipe state variable
             to keep track of that. */
          if (parms->cArgs != 1 || parms->rgvarg[0].vt != (VT_BOOL | VT_BYREF))
           {
             /* This happens in the weird case */
             log_oom ("%s:%s: Uncancellable write event.",
                      SRCNAME, __func__);
             break;
           }

          if (m_mail->wipe ())
            {
              /* An error cleaning the mail should not happen normally.
                 But just in case there is an error we cancel the
                 write here. */
              log_debug ("%s:%s: Failed to remove plaintext.",
                         SRCNAME, __func__);
              *(parms->rgvarg[0].pboolVal) = VARIANT_TRUE;
              return E_ABORT;
            }
          break;
        }
      case AfterWrite:
        {
          if (m_send_seen)
            {
              m_send_seen = false;
              m_mail->do_crypto ();
              if (m_mail->crypto_successful ())
                {
                  /* We can't trigger a Send event in the current state.
                     Appearently Outlook locks some methods in some events.
                     So we Create a new thread that will sleep a bit before
                     it requests to send the item again. */
                  CreateThread (NULL, 0, request_send, (LPVOID) m_object, 0,
                                NULL);
                }
              return S_OK;
            }
          break;
        }
      case Unload:
        {
          log_debug ("%s:%s: Removing Mail for message: %p.",
                     SRCNAME, __func__, m_object);
          delete m_mail;
          return S_OK;
        }
      default:
        log_oom_extra ("%s:%s: Message:%p Unhandled Event: %lx \n",
                       SRCNAME, __func__, m_object, dispid);
    }
  return S_OK;
}
END_EVENT_SINK(MailItemEvents, IID_MailItemEvents)
