/* folder-events.cpp - Event handling for a folder.
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
#include "windowmessages.h"
#include "mymapitags.h"

/* Folder Events */
BEGIN_EVENT_SINK(FolderEvents, IDispatch)
EVENT_SINK_DEFAULT_CTOR(FolderEvents)
EVENT_SINK_DEFAULT_DTOR(FolderEvents)
typedef enum
  {
    BeforeItemMove = 0xFBA9,
    BeforeFolderMove = 0xFBA8,
  } FolderEvent;

EVENT_SINK_INVOKE(FolderEvents)
{
  USE_INVOKE_ARGS
  switch(dispid)
    {
      case BeforeItemMove:
        {
          TSTART;
          log_oom ("%s:%s: Item Move in folder: %p",
                         SRCNAME, __func__, this);

          /* Parameters should be
             disp item   Represents the Outlook item that is to be moved or deleted.
             disp folder Represents the folder to which the item is being moved.
                         If null the message will be deleted.
             bool cancel Move should be canceled.

             Remember that the order is inverted. */
          if (!(parms->cArgs == 3 && parms->rgvarg[1].vt == (VT_DISPATCH) &&
                parms->rgvarg[2].vt == (VT_DISPATCH) &&
                parms->rgvarg[0].vt == (VT_BOOL | VT_BYREF)))
            {
              log_error ("%s:%s: Invalid args.",
                         SRCNAME, __func__);
              TBREAK;
            }

          if (!parms->rgvarg[1].pdispVal)
            {
              log_oom ("%s:%s: Passing delete",
                             SRCNAME, __func__);
              TBREAK;
            }

          LPDISPATCH mailitem = parms->rgvarg[2].pdispVal;

          char *name = get_object_name (mailitem);
          if (!name || strcmp (name, "_MailItem"))
            {
              log_debug ("%s:%s: move is not about a mailitem.",
                         SRCNAME, __func__);
              xfree (name);
              TBREAK;
            }
          xfree (name);

          char *uid = get_unique_id (mailitem, 0, nullptr);
          if (!uid)
            {
              LPMESSAGE msg = get_oom_base_message (mailitem);
              uid = mapi_get_uid (msg);
              gpgol_release (msg);
              if (!uid)
                {
                  log_debug ("%s:%s: Failed to get uid for %p",
                             SRCNAME, __func__, mailitem);
                  TBREAK;
                }
            }

          Mail *mail = Mail::getMailForUUID (uid);
          if (!mail)
            {
              auto retV = Mail::searchMailsByUUID(uid);
              if (!retV.empty())
                {
                  mail = retV.front();
                }
              else
                {
                  log_debug ("%s:%s: Failed to find mail %p in map.",
                            SRCNAME, __func__, mailitem);
                }
                TRACEPOINT;
            }
          TRACEPOINT;
          xfree (uid);

          if (!mail)
            {
              log_error ("%s:%s: Failed to find mail for uuid",
                         SRCNAME, __func__);
              TBREAK;
            }
          if (mail->isCryptoMail ())
            {
              log_debug ("%s:%s: Detected move of crypto mail. %p Closing",
                          SRCNAME, __func__, mail);

              size_t entryIDLen = 0;
              char *entryID = nullptr;
              char *old_class = nullptr;
              if (mail->isSMIME_m ())
                {
                  LPMESSAGE msg = get_oom_message (mail->item ());
                  if (msg)
                    {
                      entryID = mapi_get_binary_prop (msg, PR_ENTRYID,
                                                      &entryIDLen);
                      old_class = mapi_get_message_class (msg);
                      gpgol_release (msg);
                    }
                }

              if (mail->close (false))
                {
                  log_error ("%s:%s: Failed to close.",
                             SRCNAME, __func__);
                  xfree (entryID);
                  xfree (old_class);
                  TBREAK;
                }
              /* Beware: The mail object might be destroyed now. */

              if (!entryID || !old_class)
                {
                  /* This is not an S/MIME mail so we are done. */
                  TBREAK;
                }

              auto target = (LPMAPIFOLDER) get_oom_iunknown (
                      parms->rgvarg[1].pdispVal, "MAPIOBJECT");
              if (!target)
                {
                  log_error ("%s:%s: Failed to obtain target folder.",
                             SRCNAME, __func__);
                  xfree (entryID);
                  xfree (old_class);
                  TBREAK;
                }
              memdbg_addRef (target);

              auto *data = (wm_after_move_data_t *)
                xmalloc (sizeof (wm_after_move_data_t));

              data->target_folder = target;
              data->entry_id = entryID;
              data->entry_id_len = entryIDLen;
              data->old_class = old_class;

              do_in_ui_thread_async (AFTER_MOVE, data, 500);
            }
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
END_EVENT_SINK(FolderEvents, IID_FolderEvents)
