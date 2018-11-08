/* inspector-events.cpp - Event handling for an inspector.
 * Copyright (C) 2018 Intevation GmBH
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

/* Folder Events */
BEGIN_EVENT_SINK(InspectorEvents, IDispatch)
EVENT_SINK_DEFAULT_CTOR(InspectorEvents)
EVENT_SINK_DEFAULT_DTOR(InspectorEvents)
typedef enum
  {
    Activate = 0xF001,
    AttachmentSelectionChange = 0xFC79,
    BeforeMaximize = 0xFA11,
    BeforeMinimize = 0xFA12,
    BeforeMove = 0xFA13,
    BeforeSize = 0xFA14,
    Close = 0xF008,
    Deactivate = 0xF006,
    PageChange = 0xFBF4,
  } InspectorEvent;

EVENT_SINK_INVOKE(InspectorEvents)
{
  USE_INVOKE_ARGS
  switch(dispid)
    {
      case Close:
        {
          TSTART;
          log_oom ("%s:%s Close event in inspector %p",
                   SRCNAME, __func__, this);
          /* Get the mail object belonging to us */
          auto mailitem = MAKE_SHARED (get_oom_object (m_object, "CurrentItem"));

          if (!mailitem)
            {
              STRANGEPOINT;
              TBREAK;
            }

          char *uid = get_unique_id (mailitem.get (), 0, nullptr);
          mailitem = nullptr;
          if (!uid)
            {
              log_debug ("%s:%s: Failed to find uid for %p \n",
                         SRCNAME, __func__, mailitem);
              TBREAK;
            }
          auto mail = Mail::getMailForUUID (uid);
          xfree (uid);

          if (!mail)
            {
              STRANGEPOINT;
              TBREAK;
            }

          if (!mail->isCryptoMail())
            {
              STRANGEPOINT;
              TBREAK;
            }

          if (mail->getCloseTriggered ())
            {
              log_debug ("%s:%s: Passing our close for item %p",
                         SRCNAME, __func__, mail);
              TBREAK;
            }
          mail->setCloseTriggered (true);
          Mail::closeInspector_o (mail);
          Mail::close (mail);
        }
      default:
        break;
#if 1
        log_oom ("%s:%s: Unhandled Event: %lx \n",
                       SRCNAME, __func__, dispid);
#endif
    }
  return S_OK;
}
END_EVENT_SINK(InspectorEvents, IID_InspectorEvents)
