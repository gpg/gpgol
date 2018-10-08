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

/*
   We need to avoid UI invalidations as much as possible as invalidations
   can trigger reloads of mails and at a bad time can crash us.

   So we only invalidate the UI after we have handled the read event of
   a mail and again after decrypt / verify.

   The problem is that we also need to update the UI when mails are
   unselected so we don't show "Secure" if nothing is selected.

   On a switch from one Mail to another we see two selection changes.
   One for the unselect the other for the select immediately after
   each other.

   When we just have an unselect we see only one selection change.

   So after we detect the unselect we switch the state in our
   explorerMap to unselect seen and start a WatchDog thread.

   That thread sleeps for 500ms and then checks if the state
   was switched to select seen in the meantime. If
   not it triggers the UI Invalidation in the GUI thread.
   */
#include <map>

typedef enum
  {
    WatchDogActive = 0x01,
    UnselectSeen = 0x02,
    SelectSeen = 0x04,
  } SelectionState;

std::map<LPDISPATCH, int> s_explorerMap;

gpgrt_lock_t explorer_map_lock = GPGRT_LOCK_INITIALIZER;

static bool
hasSelection (LPDISPATCH explorer)
{
  LPDISPATCH selection = get_oom_object (explorer, "Selection");

  if (!selection)
    {
      TRACEPOINT;
      return false;
    }

  int count = get_oom_int (selection, "Count");
  gpgol_release (selection);

  if (count)
    {
      return true;
    }
  return false;
}

static DWORD WINAPI
start_watchdog (LPVOID arg)
{
  LPDISPATCH explorer = (LPDISPATCH) arg;

  Sleep (500);
  gpgol_lock (&explorer_map_lock);

  auto it = s_explorerMap.find (explorer);

  if (it == s_explorerMap.end ())
    {
      log_error ("%s:%s: Watchdog for unknwon explorer %p",
                 SRCNAME, __func__, explorer);
      gpgol_unlock (&explorer_map_lock);
      return 0;
    }

  if ((it->second & SelectSeen))
    {
      log_oom ("%s:%s: Cancel watchdog as we have seen a select %p",
                     SRCNAME, __func__, explorer);
      it->second = SelectSeen;
    }
  else if ((it->second & UnselectSeen))
    {
      log_debug ("%s:%s: Deteced unselect invalidating UI.",
                 SRCNAME, __func__);
      it->second = UnselectSeen;
      gpgol_unlock (&explorer_map_lock);
      do_in_ui_thread (INVALIDATE_UI, nullptr);
      return 0;
    }
  gpgol_unlock (&explorer_map_lock);

  return 0;
}

static void
changeSeen (LPDISPATCH explorer)
{
  gpgol_lock (&explorer_map_lock);

  auto it = s_explorerMap.find (explorer);

  if (it == s_explorerMap.end ())
    {
      it = s_explorerMap.insert (std::make_pair (explorer, 0)).first;
    }

  auto state = it->second;
  bool has_selection = hasSelection (explorer);

  if (has_selection)
    {
      it->second = (state & WatchDogActive) + SelectSeen;
      log_oom ("%s:%s: Seen select for %p",
                     SRCNAME, __func__, explorer);
    }
  else
    {
      if ((it->second & WatchDogActive))
        {
          log_oom ("%s:%s: Seen unselect for %p but watchdog exists.",
                         SRCNAME, __func__, explorer);
        }
      else
        {
          CloseHandle (CreateThread (NULL, 0, start_watchdog, (LPVOID) explorer,
                                     0, NULL));
        }
      it->second = UnselectSeen + WatchDogActive;
    }
  gpgol_unlock (&explorer_map_lock);
}

EVENT_SINK_INVOKE(ExplorerEvents)
{
  USE_INVOKE_ARGS
  switch(dispid)
    {
      case SelectionChange:
        {
          log_oom ("%s:%s: Selection change in explorer: %p",
                         SRCNAME, __func__, this);
          changeSeen (m_object);
          break;
        }
      case Close:
        {
          log_oom ("%s:%s: Deleting event handler: %p",
                         SRCNAME, __func__, this);

          GpgolAddin::get_instance ()->unregisterExplorerSink (this);
          gpgol_lock (&explorer_map_lock);
          s_explorerMap.erase (m_object);
          gpgol_unlock (&explorer_map_lock);
          delete this;
          return S_OK;
        }
      default:
        break;
#if 0
        log_oom ("%s:%s: Unhandled Event: %lx \n",
                       SRCNAME, __func__, dispid);
#endif
    }
  return S_OK;
}
END_EVENT_SINK(ExplorerEvents, IID_ExplorerEvents)
