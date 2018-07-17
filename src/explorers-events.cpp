/* explorer-events.cpp - Event handling for the application.
 * Copyright (C) 2016 by Bundesamt für Sicherheit in der Informationstechnik
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
#include "eventsinks.h"
#include "ocidl.h"
#include "common.h"
#include "oomhelp.h"
#include "mail.h"
#include "gpgoladdin.h"

/* Explorers Events */
BEGIN_EVENT_SINK(ExplorersEvents, IDispatch)
EVENT_SINK_DEFAULT_CTOR(ExplorersEvents)
EVENT_SINK_DEFAULT_DTOR(ExplorersEvents)
typedef enum
  {
    NewExplorer = 0xF001,
  } ExplorersEvent;

/* Don't confuse with ExplorerEvents. ExplorerEvents is
   the actual event sink for explorer events. This just
   ensures that we create such a sink for each new explorer. */
EVENT_SINK_INVOKE(ExplorersEvents)
{
  USE_INVOKE_ARGS
  switch(dispid)
    {
      case NewExplorer:
        {
          if (parms->cArgs != 1 || !(parms->rgvarg[0].vt & VT_DISPATCH))
            {
              log_debug ("%s:%s: No explorer in new Explorer.",
                         SRCNAME, __func__);
              break;
            }
          LPDISPATCH expSink = install_ExplorerEvents_sink (parms->rgvarg[0].pdispVal);
          if (!expSink)
            {
              log_error ("%s:%s: Failed to install Explorer event sink.",
                         SRCNAME, __func__);
              break;
            }
          GpgolAddin::get_instance()->registerExplorerSink (expSink);
        }
      default:
        break;
    }
  return S_OK;
}
END_EVENT_SINK(ExplorersEvents, IID_ExplorersEvents)
