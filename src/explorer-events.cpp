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
#include "mymapitags.h"

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

typedef struct
  {
    int state;
    LPSPropValue propEntryId;
  } exInfo, *pExInfo;

std::map<LPDISPATCH, pExInfo> s_explorerMap;

gpgrt_lock_t explorer_map_lock = GPGRT_LOCK_INITIALIZER;

static bool
hasSelection (LPDISPATCH explorer, pExInfo pEntry)
{
  TSTART;
  LPDISPATCH selection = get_oom_object (explorer, "Selection");

  if (!selection)
    {
      TRACEPOINT;
      TRETURN false;
    }

  int count = get_oom_int (selection, "Count");
  LPDISPATCH selectitem = NULL, mailitem = NULL;
  bool selected = false;
  if (count==1)
    {
      selected = hasMailitemEventReadBeenCalled ();
      log_debug ("%s:%s: ReadEvent %s been called",
            SRCNAME, __func__, selected ? "HAS": "has NOT");
      g_ignore_next_load = true;
      selectitem = get_oom_object (selection, "Item(1)");
      mailitem = get_object_by_id (selectitem, IID_MailItem);
      if (!mailitem)
      {
        log_debug ("%s:%s: New selection is no mail, return false.",
            SRCNAME, __func__);
        selected = false;
      }
      else
      {
        LPMESSAGE msg = get_oom_message(mailitem);
        HRESULT hr = HrGetOneProp ((LPMAPIPROP)msg, PR_ENTRYID, &pEntry->propEntryId);
        if (!FAILED (hr) && PROP_TYPE (pEntry->propEntryId->ulPropTag) == PT_BINARY)
        {
          size_t keylen = pEntry->propEntryId->Value.bin.cb;
          void *key = pEntry->propEntryId->Value.bin.lpb;
          log_hexdump (key, keylen, "%s: %20s=", __func__, "Item(1) ENTRYID");
        }
        else if (!FAILED (hr))
        {
           MAPIFreeBuffer(pEntry->propEntryId);
           log_debug ("%s:%s: HrGetOneProp(%s) returned non binary property: hr=%#lx\n",
                      SRCNAME, __func__, "PR_ENTRYID", PROP_TYPE (pEntry->propEntryId->ulPropTag));
        }
        else
        {
          log_debug ("%s:%s: HrGetOneProp(%s) failed: hr=%#lx\n",
                    SRCNAME, __func__, "PR_ENTRYID", hr);
        }
        gpgol_release(msg);
      }
      gpgol_release (mailitem);
      gpgol_release (selectitem);
    }
    else
    {
      log_debug ("%s:%s: %d Items selected return false to show insecure",
            SRCNAME, __func__, count);

        selected = false; // We can't show the security level for none/more than one => show insecure
    }

  gpgol_release (selection);
  TRETURN selected;
}

static DWORD WINAPI
start_watchdog (LPVOID arg)
{
  TSTART;
  LPDISPATCH explorer = (LPDISPATCH) arg;

  Sleep (500);
  gpgol_lock (&explorer_map_lock);

  auto it = s_explorerMap.find (explorer);

  if (it == s_explorerMap.end ())
    {
      log_error ("%s:%s: Watchdog for unknwon explorer %p",
                 SRCNAME, __func__, explorer);
      gpgol_unlock (&explorer_map_lock);
      TRETURN 0;
    }

  if ((it->second->state & SelectSeen))
    {
      log_oom ("%s:%s: Cancel watchdog as we have seen a select %p",
                     SRCNAME, __func__, explorer);
      it->second->state = SelectSeen;
    }
  else if ((it->second->state & UnselectSeen))
    {
      log_debug ("%s:%s: Deteced unselect invalidating UI.",
                 SRCNAME, __func__);
      it->second->state = UnselectSeen;
      gpgol_unlock (&explorer_map_lock);
      do_in_ui_thread (INVALIDATE_UI, nullptr);
      TRETURN 0;
    }
  gpgol_unlock (&explorer_map_lock);

  TRETURN 0;
}

static void
changeSeen (LPDISPATCH explorer)
{
  TSTART;
  auto view = get_oom_object_s (explorer, "CurrentView");
  if (view && get_object_name_s (view.get ()) == "_PeopleView")
    {
      log_oom ("Selection change in people view. Invalidating.");
      gpgoladdin_invalidate_ui ();
      TRETURN;
    }

  gpgol_lock (&explorer_map_lock);

  auto it = s_explorerMap.find (explorer);

  if (it == s_explorerMap.end ())
    {
      pExInfo pStateInfo = (pExInfo) xmalloc(sizeof(exInfo));
      pStateInfo->state = 0;
      pStateInfo->propEntryId = NULL;
      it = s_explorerMap.insert (std::make_pair (explorer, pStateInfo)).first;
    }

  auto state = it->second->state;
  bool has_selection = false;
  if (it->second->propEntryId != NULL)
  {
    size_t keylen = it->second->propEntryId->Value.bin.cb;
    void *key = it->second->propEntryId->Value.bin.lpb;
    log_hexdump (key, keylen, "%s: %20s=", __func__, "Explorer selected Item ENTRYID");
    MAPIFreeBuffer(it->second->propEntryId);
    it->second->propEntryId = NULL;
  }
  else
  {
    has_selection = hasSelection (explorer, it->second);
  }

  if (has_selection)
    {
      it->second->state = (state & WatchDogActive) + SelectSeen;
      log_oom ("%s:%s: Seen select for %p",
                     SRCNAME, __func__, explorer);
    }
  else
    {
      if ((it->second->state & WatchDogActive))
        {
          log_oom ("%s:%s: Seen unselect for %p but watchdog exists.",
                         SRCNAME, __func__, explorer);
        }
      else
        {
          CloseHandle (CreateThread (NULL, 0, start_watchdog, (LPVOID) explorer,
                                     0, NULL));
        }
      it->second->state = UnselectSeen + WatchDogActive;
    }
  gpgol_unlock (&explorer_map_lock);
  TRETURN;
}

EVENT_SINK_INVOKE(ExplorerEvents)
{
  TSTART;
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
          auto it = s_explorerMap.find (m_object);
          if (it != s_explorerMap.end ())
          {
             MAPIFreeBuffer(it->second->propEntryId);
             xfree(it->second);
          }
          s_explorerMap.erase (m_object);
          gpgol_unlock (&explorer_map_lock);
          delete this;
          TRETURN S_OK;
        }
      default:
        break;
#if 0
        log_oom ("%s:%s: Unhandled Event: %lx \n",
                       SRCNAME, __func__, dispid);
#endif
    }
  TRETURN S_OK;
}
END_EVENT_SINK(ExplorerEvents, IID_ExplorerEvents)
