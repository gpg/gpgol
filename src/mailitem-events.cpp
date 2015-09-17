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
       m_was_encrypted; /* The original message was encrypted.  */

  HRESULT handle_before_read();
  HRESULT handle_read();
  HRESULT handle_after_write();
};

MailItemEvents::MailItemEvents() :
    m_object(NULL),
    m_pCP(NULL),
    m_cookie(0),
    m_ref(1),
    m_send_seen(false),
    m_want_html(false),
    m_processed(false)
{
/* The event sink default dtor closes this for us. */
EVENT_SINK_DEFAULT_DTOR(MailItemEvents)


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
  err = message_incoming_handler (message, get_oom_context_window (m_object),
                                  false);
  m_processed = (err == 1) || (err == 2);
  m_was_encrypted = err == 2;

  log_debug ("%s:%s: incoming handler status: %i",
             SRCNAME, __func__, err);
  message->Release ();
}

HRESULT
MailItemEvents::handle_after_write()
{
  int err;
  LPMESSAGE message = get_oom_base_message (m_object);
  if (!message)
    {
      log_error ("%s:%s: Failed to get base message.",
                 SRCNAME, __func__);
      return S_OK;
    }
  log_debug ("%s:%s: Sign / Encrypting message",
             SRCNAME, __func__);
  /* TODO check for message flags to determine */
  err = message_sign_encrypt (message, PROTOCOL_UNKNOWN,
                              get_oom_context_window (m_object));
  log_debug ("%s:%s: Sign / Encryption status: %i",
             SRCNAME, __func__, err);
  message->Release ();
  if (err)
    {
      // TODO: I think we can still cancel the send
      // on the MAPI level in case of errors
      // but we have to get at the messagestore to
      // do that.
    }
  return S_OK;
}

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
              return handle_read();
            }
          return S_OK;
        }
      case Send:
        {
          m_send_seen = true;
          return S_OK;
        }
      case AfterWrite:
        {
          if (m_send_seen)
            {
              return handle_after_write();
            }
        }
      default:
        log_debug ("%s:%s: Unhandled Event: %lx \n",
                       SRCNAME, __func__, dispid);
    }
  return S_OK;
}
END_EVENT_SINK(MailItemEvents, IID_MailItemEvents)
