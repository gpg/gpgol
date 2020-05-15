/* application-events.cpp - Event handling for the application.
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

/* The event handler classes defined in this file follow the
   general pattern that they implment the IDispatch interface
   through the eventsink macros and handle event invocations
   in their invoke methods.
*/
#ifdef HAVE_CONFIG_H
#include <config.h>
#endif

#include "eventsink.h"
#include "ocidl.h"
#include "common.h"
#include "oomhelp.h"
#include "mail.h"
#include "gpgoladdin.h"
#include "windowmessages.h"

/* Application Events */
BEGIN_EVENT_SINK(ApplicationEvents, IDispatch)
EVENT_SINK_DEFAULT_CTOR(ApplicationEvents)
EVENT_SINK_DEFAULT_DTOR(ApplicationEvents)
typedef enum
  {
    AdvancedSearchComplete = 0xFA6A,
    AdvancedSearchStopped = 0xFA6B,
    AttachmentContextMenuDisplay = 0xFB3E,
    BeforePrint = 0xFC8E,
    BeforeFolderSharingDialog = 0xFC01,
    ContextMenuClose = 0xFBA6,
    FolderContextMenuDisplay = 0xFB42,
    ItemContextMenuDisplay = 0xFB41,
    ItemLoad = 0xFBA7,
    ItemSend = 0xF002,
    MAPILogonComplete = 0xFA90,
    NewMail = 0xF003,
    NewMailEx = 0xFAB5,
    OptionsPagesAdd = 0xF005,
    Quit = 0xF007,
    Reminder = 0xF004,
    ShortcutContextMenuDisplay = 0xFB44,
    Startup = 0xF006,
    StoreContextMenuDisplay = 0xFB43,
    ViewContextMenuDisplay = 0xFB40
  } ApplicationEvent;

static bool beforePrintSeen;

EVENT_SINK_INVOKE(ApplicationEvents)
{
  USE_INVOKE_ARGS
  switch(dispid)
    {
      case BeforePrint:
        {
          log_debug ("%s:%s: BeforePrint seen.",
                     SRCNAME, __func__);
          beforePrintSeen = true;
          break;
        }
      case ItemLoad:
        {
          TSTART;
          if (g_ignore_next_load)
            {
              log_debug ("%s:%s: Ignore ItemLoad because ignore next "
                         "load is true", SRCNAME, __func__);
              g_ignore_next_load = false;
              TBREAK;
            }
          LPDISPATCH mailItem;
          /* The mailItem should be the first argument */
          if (!parms || parms->cArgs != 1 ||
              parms->rgvarg[0].vt != VT_DISPATCH)
            {
              log_error ("%s:%s: ItemLoad with unexpected Arguments.",
                         SRCNAME, __func__);
              TBREAK;
            }

          log_debug ("%s:%s: ItemLoad event. Getting object.",
                     SRCNAME, __func__);
          mailItem = get_object_by_id (parms->rgvarg[0].pdispVal,
                                       IID_MailItem);
          if (!mailItem)
            {
              beforePrintSeen = false;
              log_debug ("%s:%s: ItemLoad is not for a mail.",
                         SRCNAME, __func__);
              TBREAK;
            }
          log_debug ("%s:%s: Creating mail object for item: %p",
                     SRCNAME, __func__, mailItem);
          auto mail = new Mail (mailItem);
          if (beforePrintSeen)
            {
              mail->setIsPrint (true);
              log_dbg ("Mail pointer %p item: %p is print", mail, mailItem);
            }
          beforePrintSeen = false;
          do_in_ui_thread_async (INVALIDATE_LAST_MAIL, nullptr);
          TBREAK;
        }
      case Quit:
        {
          log_debug ("%s:%s: Quit event. Shutting down",
                     SRCNAME, __func__);
          TBREAK;
        }
      default:
        log_oom ("%s:%s: Unhandled Event: %lx \n",
                       SRCNAME, __func__, dispid);
    }
  /* We always return S_OK even on error so that outlook
     continues to handle the event and is not disturbed
     by our errors. There shouldn't be errors in here
     anyway if everything works as documented. */
  return S_OK;
}
END_EVENT_SINK(ApplicationEvents, IID_ApplicationEvents_11)
