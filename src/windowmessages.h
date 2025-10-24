/* windowmessages.h - Helper functions for Window message exchange.
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
#ifndef WINDOWMESSAGES_H
#define WINDOWMESSAGES_H

#include <windows.h>

#include "config.h"
#include "mapihelp.h"

#include <gpg-error.h>

class Mail;

/** Window Message handling for GpgOL.
  In Outlook only one thread has access to the Outlook Object model
  and this is the UI Thread. We can work in other threads but
  to do something with outlooks data we neet to be in the UI Thread.
  So we create a hidden Window in this thread and use the fact
  that SendMessage handles Window messages in the thread where the
  Window was created.
  This way we can go back to interactct with the Outlook from another
  thread without working with COM Multithreading / Marshaling.

  The Responder Window should be initalized on startup.
  */
typedef enum _gpgol_wmsg_type
{
  UNKNOWN = 1100, /* A large offset to avoid conflicts */
  INVALIDATE_UI, /* The UI should be invalidated. */
  PARSING_DONE, /* A mail was parsed. Data should be a pointer
                      to the mail object. */
  RECIPIENT_ADDED, /* A recipient was added. Data should be ptr
                      to mail */
  CLOSE, /* Close the message in the next event loop. */
  CRYPTO_DONE, /* Sign / Encrypt done. */
  WKS_NOTIFY, /* Show a WKS Notification. */
  BRING_TO_FRONT, /* Bring the active Outlook window to the front. */
  INVALIDATE_LAST_MAIL,
  REVERT_MAIL,
  CLEAR_REPLY_FORWARD,
  DO_AUTO_SECURE,
  DONT_AUTO_SECURE,
  CONFIG_KEY_DONE,
  DECRYPT,
  AFTER_MOVE,
  SEND_MULTIPLE_MAILS,
  SEND,
  SHOW_PREVIEW, /* Show mail contents before a verify is done */
  SELECT_MAIL,
  /* External API, keep it stable! */
  EXT_API_CLOSE = 1301,
  EXT_API_CLOSE_ALL = 1302,
  EXT_API_DECRYPT = 1303,
} gpgol_wmsg_type;

typedef struct
{
  void *data; /* Pointer to arbitrary data depending on msg type */
  gpgol_wmsg_type wmsg_type; /* Type of the msg. */
  int err; /* Set to true on error */
  int delay;
} wm_ctx_t;

typedef struct
{
  LPMAPIFOLDER target_folder;
  char *entry_id;
  char *old_class;
  size_t entry_id_len;
} wm_after_move_data_t;

/** Create and register the responder window.
  The responder window should be */
HWND
create_responder_window ();

/** Uses send_msg_to_ui_thread to execute the request
  in the ui thread.  Returns the result. */
int
do_in_ui_thread (gpgol_wmsg_type type, void *data);

/** Send a message to the UI thread but returns
    immediately without waiting for the execution.

    The delay is used in the detached thread to delay
    the sending of the actual message. */
void
do_in_ui_thread_async (gpgol_wmsg_type type, void *data, int delay = 0);

/** Create our filter before outlook Window Messages. */
HHOOK
create_message_hook();

/** Block ui invalidation. The idea here is to further reduce
    UI invalidations because depending on timing they might crash.
    So we try to block invalidation for as long as it is a bad time
    for us. */
void blockInv ();

/** Unblock ui invalidation */
void unblockInv ();

DWORD WINAPI
delayed_invalidate_ui (LPVOID minsleep_ms = 0);

DWORD WINAPI
close_mail (LPVOID);

void
wm_register_pending_op (Mail *mail);

void
wm_unregister_pending_op (Mail *mail);

void
wm_abort_pending_ops ();
#endif // WINDOWMESSAGES_H
