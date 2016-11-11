/* explorer-events.cpp - Event handling for the application.
 *    Copyright (C) 2016 Intevation GmbH
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

/* Explorer Events */
BEGIN_EVENT_SINK(ExplorerEvents, IDispatch)
EVENT_SINK_DEFAULT_CTOR(ExplorerEvents)
EVENT_SINK_DEFAULT_DTOR(ExplorerEvents)
typedef enum
  {
    Activate = 0xF001,
    AttachmentSelectionChange = 0xFC79,
    BeforeFolderSwitch = 0xF003,
    BeforeItemCopy = 0xFA0E,
    BeforeItemCut = 0xFA0F,
    BeforeItemPaste = 0xFA10,
    BeforeMaximize = 0xFA11,
    BeforeMinimize = 0xFA12,
    BeforeMove = 0xFA13,
    BeforeSize = 0xFA14,
    BeforeViewSwitch = 0xF005,
    Close = 0xF008,
    Deactivate = 0xF006,
    DisplayModeChange = 0xFC98,
    FolderSwitch = 0xF002,
    InlineResponse = 0xFC92,
    InlineResponseClose = 0xFC96,
    SelectionChange = 0xF007,
    ViewSwitch = 0xF004
  } ExplorerEvent;

static DWORD WINAPI
invalidate_ui (LPVOID)
{
  /* We sleep here a bit to prevent invalidtion immediately
     after the selection change before we have started processing
     the mail. */
  Sleep (1000);
  do_in_ui_thread (INVALIDATE_UI, nullptr);
  return 0;
}

EVENT_SINK_INVOKE(ExplorerEvents)
{
  USE_INVOKE_ARGS
  switch(dispid)
    {
      case SelectionChange:
        {
          log_oom_extra ("%s:%s: Selection change in explorer: %p",
                         SRCNAME, __func__, this);
          HANDLE thread = CreateThread (NULL, 0, invalidate_ui, (LPVOID) this, 0,
                                        NULL);

          if (!thread)
            {
              log_error ("%s:%s: Failed to create invalidate_ui thread.",
                         SRCNAME, __func__);
            }
          else
            {
              CloseHandle (thread);
            }
          break;
        }
      case Close:
        {
          log_oom_extra ("%s:%s: Deleting event handler: %p",
                         SRCNAME, __func__, this);

          remove_explorer (m_object);
          delete this;
          return S_OK;
        }
      default:
        break;
#if 0
        log_oom_extra ("%s:%s: Unhandled Event: %lx \n",
                       SRCNAME, __func__, dispid);
#endif
    }
  return S_OK;
}
END_EVENT_SINK(ExplorerEvents, IID_ExplorerEvents)
