/* application-events.cpp - Event handling for the application.
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

/* The event handler classes defined in this file follow the
   general pattern that they implment the IDispatch interface
   through the eventsink macros and handle event invocations
   in their invoke methods.
*/
#include "eventsink.h"
#include "eventsinks.h"
#include "ocidl.h"
#include "common.h"
#include "oomhelp.h"

/* Application Events */
BEGIN_EVENT_SINK(ApplicationEvents, IDispatch)
EVENT_SINK_DEFAULT_CTOR(ApplicationEvents)
EVENT_SINK_DEFAULT_DTOR(ApplicationEvents)
typedef enum
  {
    AdvancedSearchComplete = 0xFA6A,
    AdvancedSearchStopped = 0xFA6B,
    AttachmentContextMenuDisplay = 0xFB3E,
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

EVENT_SINK_INVOKE(ApplicationEvents)
{
  USE_INVOKE_ARGS
  switch(dispid)
    {
      case ItemLoad:
        {
          LPDISPATCH mailItem;
          LPDISPATCH mailEventSink;
          /* The mailItem should be the first argument */
          if (parms->cArgs != 1 || parms->rgvarg[0].vt != VT_DISPATCH)
            {
              log_error ("%s:%s: ItemLoad with unexpected Arguments.",
                         SRCNAME, __func__);
              break;
            }

          mailItem = get_object_by_id (parms->rgvarg[0].pdispVal,
                                       IID_MailItem);
          if (!mailItem)
            {
              log_error ("%s:%s: ItemLoad event without mailitem.",
                         SRCNAME, __func__);
              break;
            }
          mailEventSink = install_MailItemEvents_sink (mailItem);
          /* TODO figure out what we need to do with the event sink.
             Does it need to be Released at some point? What happens
             on unload? */
          if (!mailEventSink)
            {
              log_error ("%s:%s: Failed to install MailItemEvents sink.",
                         SRCNAME, __func__);
            }
          mailItem->Release ();
          break;
        }
      default:
        log_oom_extra ("%s:%s: Unhandled Event: %lx \n",
                       SRCNAME, __func__, dispid);
    }
  /* We always return S_OK even on error so that outlook
     continues to handle the event and is not disturbed
     by our errors. There shouldn't be errors in here
     anyway if everything works as documented. */
  return S_OK;
}
END_EVENT_SINK(ApplicationEvents, IID_ApplicationEvents)
