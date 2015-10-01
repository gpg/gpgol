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

#include "common.h"
#include "eventsink.h"
#include "eventsinks.h"
#include "mymapi.h"
#include "message.h"
#include "oomhelp.h"
#include "ocidl.h"
#include "attachment.h"
#include "mapihelp.h"
#include "gpgoladdin.h"
#include "windowmessages.h"

/* TODO Add a proper / l10n encrypted thing message. */
static const char * ENCRYPTED_MESSAGE_BODY = \
"This message is encrypted. Please install or activate GpgOL"\
" to decrypt this message.";

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
  bool m_send_seen,   /* The message is about to be submitted */
       m_want_html,    /* Encryption of HTML is desired. */
       m_processed,    /* The message has been porcessed by us.  */
       m_needs_wipe,   /* We have added plaintext to the mesage. */
       m_was_encrypted, /* The original message was encrypted.  */
       m_crypt_successful; /* We successfuly performed crypto on the item. */

  HRESULT handle_before_read();
  HRESULT handle_read();
};

MailItemEvents::MailItemEvents() :
    m_object(NULL),
    m_pCP(NULL),
    m_cookie(0),
    m_ref(1),
    m_send_seen(false),
    m_want_html(false),
    m_processed(false),
    m_crypt_successful(false)
{
}

MailItemEvents::~MailItemEvents()
{
  if (m_pCP)
    m_pCP->Unadvise(m_cookie);
  if (m_object)
    m_object->Release();
}

HRESULT
MailItemEvents::handle_read()
{
  int err;
  int is_html, was_protected = 0;
  char *body = NULL;
  LPMESSAGE message = get_oom_message (m_object);
  if (!message)
    {
      log_error ("%s:%s: Failed to get message \n",
                 SRCNAME, __func__);
      return S_OK;
    }
  err = mapi_get_gpgol_body_attachment (message, &body, NULL,
                                        &is_html, &was_protected);
  message->Release ();
  if (err || !body)
    {
      log_error ("%s:%s: Failed to get body attachment of \n",
                 SRCNAME, __func__);
      return S_OK;
    }
  if (put_oom_string (m_object, is_html ? "HTMLBody" : "Body", body))
    {
      log_error ("%s:%s: Failed to modify body of item. \n",
                 SRCNAME, __func__);
    }

  xfree (body);

  if (unprotect_attachments (m_object))
    {
      log_error ("%s:%s: Failed to unprotect attachments. \n",
                 SRCNAME, __func__);
    }

  return S_OK;
}

/* Before read is the time where we can access the underlying
   base message. So this is where we create our attachment. */
HRESULT
MailItemEvents::handle_before_read()
{
  int err;
  LPMESSAGE message = get_oom_base_message (m_object);
  if (!message)
    {
      log_error ("%s:%s: Failed to get base message.",
                 SRCNAME, __func__);
      return S_OK;
    }
  log_oom_extra ("%s:%s: GetBaseMessage OK.",
                 SRCNAME, __func__);
  err = message_incoming_handler (message, NULL,
                                  false);
  m_processed = (err == 1) || (err == 2);
  m_was_encrypted = err == 2;

  log_debug ("%s:%s: incoming handler status: %i",
             SRCNAME, __func__, err);
  message->Release ();
  return S_OK;
}


static int
do_crypto_on_item (LPDISPATCH mailitem)
{
  int err = -1,
      flags = 0;
  LPMESSAGE message = get_oom_base_message (mailitem);
  if (!message)
    {
      log_error ("%s:%s: Failed to get base message.",
                 SRCNAME, __func__);
      return err;
    }
  flags = get_gpgol_draft_info_flags (message);
  if (flags == 3)
    {
      log_debug ("%s:%s: Sign / Encrypting message",
                 SRCNAME, __func__);
      err = message_sign_encrypt (message, PROTOCOL_UNKNOWN,
                                  NULL);
    }
  else if (flags == 2)
    {
      err = message_sign (message, PROTOCOL_UNKNOWN,
                          NULL);
    }
  else if (flags == 1)
    {
      err = message_encrypt (message, PROTOCOL_UNKNOWN,
                             NULL);
    }
  else
    {
      log_debug ("%s:%s: Unknown flags for crypto: %i",
                 SRCNAME, __func__, flags);
    }
  log_debug ("%s:%s: Status: %i",
             SRCNAME, __func__, err);
  message->Release ();
  return err;
}


DWORD WINAPI
request_send (LPVOID arg)
{
  int not_sent = 1;
  int tries = 0;
  do
    {
      /* Outlook needs to handle the message some more to unblock
         calls to Send. Lets give it 50ms before we send it again. */
      Sleep (50);
      log_debug ("%s:%s: requesting send for: %p",
                 SRCNAME, __func__, arg);
      not_sent = do_in_ui_thread (REQUEST_SEND_MAIL, arg);
      tries++;
    } while (not_sent && tries < 50);
  if (tries == 50)
    {
      // Hum should not happen but I rather avoid
      // an endless loop in that case.
      // TODO show error message.
    }
  return 0;
}

static bool
needs_crypto (LPDISPATCH mailitem)
{
  LPMESSAGE message = get_oom_message (mailitem);
  bool ret;
  if (!message)
    {
      log_error ("%s:%s: Failed to get message.",
                 SRCNAME, __func__);
      return false;
    }
  ret = get_gpgol_draft_info_flags (message);
  message->Release ();
  return ret;
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
  switch(dispid)
    {
      case BeforeRead:
        {
          return handle_before_read();
        }
      case Read:
        {
          if (m_processed)
            {
              m_needs_wipe = m_was_encrypted;
              handle_read();
            }
          return S_OK;
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
          if (!needs_crypto (m_object) || m_crypt_successful)
            {
               log_debug ("%s:%s: Passing send event for message %p.",
                          SRCNAME, __func__, m_object);
               m_send_seen = false;
               break;
            }
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
          if (m_processed && m_needs_wipe && !m_send_seen)
            {
              log_debug ("%s:%s: Message %p removing plaintext from Message.",
                         SRCNAME, __func__, m_object);
              if (put_oom_string (m_object, "HTMLBody",
                                  ENCRYPTED_MESSAGE_BODY) ||
                  put_oom_string (m_object, "Body", ENCRYPTED_MESSAGE_BODY) ||
                  protect_attachments (m_object))
                {
                  /* An error cleaning the mail should not happen normally.
                     But just in case there is an error we cancel the
                     write here. */
                  log_debug ("%s:%s: Failed to remove plaintext.",
                             SRCNAME, __func__);
                  *(parms->rgvarg[0].pboolVal) = VARIANT_TRUE;
                  return E_ABORT;
                }
              m_needs_wipe = false;
            }
          break;
        }
      case AfterWrite:
        {
          if (m_send_seen)
            {
              m_send_seen = false;
              m_crypt_successful = !do_crypto_on_item (m_object);
              if (m_crypt_successful)
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
      default:
        log_oom_extra ("%s:%s: Message:%p Unhandled Event: %lx \n",
                       SRCNAME, __func__, m_object, dispid);
    }
  return S_OK;
}
END_EVENT_SINK(MailItemEvents, IID_MailItemEvents)
