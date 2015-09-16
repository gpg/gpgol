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

/* Mail Item Events */
BEGIN_EVENT_SINK(MailItemEvents, IDispatch)
EVENT_SINK_DEFAULT_CTOR(MailItemEvents)
EVENT_SINK_DEFAULT_DTOR(MailItemEvents)
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

EVENT_SINK_INVOKE(MailItemEvents)
{
  USE_INVOKE_ARGS
  switch(dispid)
    {
      case BeforeRead:
        {
          LPMESSAGE message = get_oom_base_message (m_object);
          if (message)
            {
              int ret;
              log_oom_extra ("%s:%s: GetBaseMessage OK.",
                             SRCNAME, __func__);
              ret = message_incoming_handler (message, NULL, false);
              log_debug ("%s:%s: incoming handler status: %i",
                         SRCNAME, __func__, ret);
              message->Release ();
            }
          break;
        }
      case ReadComplete:
        {
          break;
        }
      case AfterWrite:
        {
          LPMESSAGE message = get_oom_base_message (m_object);
          if (message)
            {
              int ret;
              log_debug ("%s:%s: Sign / Encrypting message",
                         SRCNAME, __func__);
              ret = message_sign_encrypt (message, PROTOCOL_UNKNOWN, NULL);
              log_debug ("%s:%s: Sign / Encryption status: %i",
                         SRCNAME, __func__, ret);
              message->Release ();
              if (ret)
                {
                  // VARIANT_BOOL *cancel = parms->rgvarg[0].pboolVal;
                  // *cancel = VARIANT_TRUE;
                  /* TODO inform the user that sending was canceled */
                }
            }
          else
            {
              log_error ("%s:%s: Failed to get base message.",
                         SRCNAME, __func__);
              break;
            }
        }
      default:
        log_debug ("%s:%s: Unhandled Event: %lx \n",
                       SRCNAME, __func__, dispid);
    }
  return S_OK;
}
END_EVENT_SINK(MailItemEvents, IID_MailItemEvents)
