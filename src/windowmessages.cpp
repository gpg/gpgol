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

              char *active_uuid = get_unique_id (mailitem, 0, nullptr);
              if (!active_uuid || strcmp (active_uuid, uuid))
                {
                  log_debug ("%s:%s: UUID mismatch",
                             SRCNAME, __func__);
                  xfree (uuid);
                  delete mail;
                  break;
                }
              log_debug ("%s:%s: Decrypting %s again",
                         SRCNAME, __func__, uuid);
              xfree (uuid);
              xfree (active_uuid);

              mail->decrypt_verify ();
              break;
            }
          case (REQUEST_CLOSE):
            {
              char *uuid = (char *) ctx->data;
              auto mail = Mail::get_mail_for_uuid (uuid);
              if (!mail)
                {
                  log_debug ("%s:%s: Close request for uuid which is gone.",
                             SRCNAME, __func__);
                  break;
                }
              if (mail->close())
                {
                  log_debug ("%s:%s: Close request failed.",
                             SRCNAME, __func__);
                }
              ctx->wmsg_type = REQUEST_DECRYPT;
              gpgol_window_proc (hWnd, message, wParam, (LPARAM) ctx);
              break;
            }
          case (INVALIDATE_UI):
            {
              log_debug ("%s:%s: Invalidating UI",
                         SRCNAME, __func__);
              gpgoladdin_invalidate_ui();
              log_debug ("%s:%s: Invalidation done",
                         SRCNAME, __func__);
              break;
            }
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

static std::vector <HWND> explorers;

void
add_explorer_window (HWND hwnd)
{
  explorers.push_back (hwnd);
}

void remove_explorer_window (HWND hwnd)
{
  explorers.erase(std::remove(explorers.begin(),
                              explorers.end(),
                              hwnd),
                  explorers.end());
}

LRESULT CALLBACK
gpgol_hook(int code, WPARAM wParam, LPARAM lParam)
{
/* Once we are in the close events we don't have enough
   control to revert all our changes so we have to do it
   with this nice little hack by catching the WM_CLOSE message
   before it reaches outlook. */
  LPCWPSTRUCT cwp = (LPCWPSTRUCT) lParam;

  switch (cwp->message)
    {
      case WM_CLOSE:
      {
        if (std::find(explorers.begin(), explorers.end(), cwp->hwnd) == explorers.end())
          {
            /* Not an explorer window */
            break;
          }
        log_debug ("%s:%s: WM_CLOSE windowmessage for explorer. "
                   "Closing all mails.",
                   SRCNAME, __func__);
        Mail::close_all_mails();
      }
     case WM_SYSCOMMAND:
       if (cwp->wParam == SC_CLOSE)
        {
          log_debug ("%s:%s: SC_CLOSE syscommand. Closing all mails.",
                     SRCNAME, __func__);
          Mail::close_all_mails();
        }
       break;
     default:
       break;
    }
  return CallNextHookEx (NULL, code, wParam, lParam);
}

/* Create the message hook for outlook's windowmessages
   we are especially interested in WM_QUIT to do cleanups
   and prevent the "Item has changed" question. */
HHOOK
create_message_hook()
{
  return SetWindowsHookEx (WH_CALLWNDPROC,
                           gpgol_hook,
                           NULL,
                           GetCurrentThreadId());
}
