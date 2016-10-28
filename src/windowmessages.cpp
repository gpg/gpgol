/* @file windowmessages.h
 * @brief Helper class to work with the windowmessage handler thread.
 *
 *    Copyright (C) 2015, 2016 Intevation GmbH
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
#include "windowmessages.h"

#include "common.h"
#include "oomhelp.h"
#include "mail.h"
#include "gpgoladdin.h"

#include <stdio.h>

#define RESPONDER_CLASS_NAME "GpgOLResponder"

/* Singleton window */
static HWND g_responder_window = NULL;

LONG_PTR WINAPI
gpgol_window_proc (HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
  if (message == WM_USER + 1)
    {
      wm_ctx_t *ctx = (wm_ctx_t *) lParam;
      log_debug ("%s:%s: Recieved user msg: %i",
                 SRCNAME, __func__, ctx->wmsg_type);
      switch (ctx->wmsg_type)
        {
          case (PARSING_DONE):
            {
              auto mail = (Mail*) ctx->data;
              if (!Mail::is_valid_ptr (mail))
                {
                  log_debug ("%s:%s: Parsing done for mail which is gone.",
                             SRCNAME, __func__);
                  break;
                }
              mail->parsing_done();
              break;
            }
          case (REQUEST_DECRYPT):
            {
              char *uuid = (char *) ctx->data;
              auto mail = Mail::get_mail_for_uuid (uuid);
              if (!mail)
                {
                  log_debug ("%s:%s: Decrypt again for uuid which is gone.",
                             SRCNAME, __func__);
                  xfree (uuid);
                  break;
                }
              /* Check if we are still in the active explorer. */
              LPDISPATCH mailitem = get_oom_object (GpgolAddin::get_instance()->get_application (),
                                                    "ActiveExplorer.Selection.Item(1)");
              if (!mailitem)
                {
                  log_debug ("%s:%s: Decrypt again but no selected mailitem.",
                             SRCNAME, __func__);
                  xfree (uuid);
                  delete mail;
                  break;
                }

              char *active_uuid = get_unique_id (mailitem, 0);
              if (!active_uuid || strcmp (active_uuid, uuid))
                {
                  log_debug ("%s:%s: UUID mismatch",
                             SRCNAME, __func__);
                  xfree (uuid);
                  delete mail;
                  break;
                }
              xfree (uuid);
              xfree (active_uuid);

              mail->decrypt_verify ();
              break;
            }

          break;
          default:
            log_debug ("Unknown msg");
        }
        return DefWindowProc(hWnd, message, wParam, lParam);
    }

  return DefWindowProc(hWnd, message, wParam, lParam);
}

HWND
create_responder_window ()
{
  size_t cls_name_len = strlen(RESPONDER_CLASS_NAME) + 1;
  char cls_name[cls_name_len];
  if (g_responder_window)
    {
      return g_responder_window;
    }
  /* Create Window wants a mutable string as the first parameter */
  snprintf (cls_name, cls_name_len, "%s", RESPONDER_CLASS_NAME);

  WNDCLASS windowClass;
  windowClass.style = CS_GLOBALCLASS | CS_DBLCLKS;
  windowClass.lpfnWndProc = gpgol_window_proc;
  windowClass.cbClsExtra = 0;
  windowClass.cbWndExtra = 0;
  windowClass.hInstance = (HINSTANCE) GetModuleHandle(NULL);
  windowClass.hIcon = 0;
  windowClass.hCursor = 0;
  windowClass.hbrBackground = 0;
  windowClass.lpszMenuName  = 0;
  windowClass.lpszClassName = cls_name;
  RegisterClass(&windowClass);
  g_responder_window = CreateWindow (cls_name, RESPONDER_CLASS_NAME, 0, 0, 0,
                                     0, 0, 0, (HMENU) 0,
                                     (HINSTANCE) GetModuleHandle(NULL), 0);
  return g_responder_window;
}

int
send_msg_to_ui_thread (wm_ctx_t *ctx)
{
  size_t cls_name_len = strlen(RESPONDER_CLASS_NAME) + 1;
  char cls_name[cls_name_len];
  snprintf (cls_name, cls_name_len, "%s", RESPONDER_CLASS_NAME);

  HWND responder = FindWindow (cls_name, RESPONDER_CLASS_NAME);
  if (!responder)
  {
    log_error ("%s:%s: Failed to find responder window.",
               SRCNAME, __func__);
    return -1;
  }
  SendMessage (responder, WM_USER + 1, 0, (LPARAM) ctx);
  return 0;
}

int
do_in_ui_thread (gpgol_wmsg_type type, void *data)
{
  wm_ctx_t ctx = {NULL, UNKNOWN, 0};
  ctx.wmsg_type = type;
  ctx.data = data;
  if (send_msg_to_ui_thread (&ctx))
    {
      return -1;
    }
  return ctx.err;
}
