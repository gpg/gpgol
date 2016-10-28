/* windowmessages.h - Helper functions for Window message exchange.
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
#ifndef WINDOWMESSAGES_H
#define WINDOWMESSAGES_H

#include <windows.h>

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
  UNKNOWN = 0,
  PARSING_DONE = 2, /* A mail was parsed. Data should be a pointer
                      to the mail object. */
  REQUEST_DECRYPT = 3
} gpgol_wmsg_type;

typedef struct
{
  void *data; /* Pointer to arbitrary data depending on msg type */
  gpgol_wmsg_type wmsg_type; /* Type of the msg. */
  int err; /* Set to true on error */
} wm_ctx_t;

/** Create and register the responder window.
  The responder window should be */
HWND
create_responder_window ();

/** Send a message to the UI thread through the responder Window.
  Returns 0 on success. */
int
send_msg_to_ui_thread (wm_ctx_t *ctx);

/** Uses send_msg_to_ui_thread to execute the request
  in the ui thread.  Returns the result. */
int
do_in_ui_thread (gpgol_wmsg_type type, void *data);
#endif // WINDOWMESSAGES_H
