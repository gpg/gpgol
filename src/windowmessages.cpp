/* @file windowmessages.h
 * @brief Helper class to work with the windowmessage handler thread.
 *
 * Copyright (C) 2015, 2016 by Bundesamt für Sicherheit in der Informationstechnik
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
          case (RECIPIENT_ADDED):
            {
              auto mail = (Mail*) ctx->data;
              if (!Mail::is_valid_ptr (mail))
                {
                  log_debug ("%s:%s: Recipient add for mail which is gone.",
                             SRCNAME, __func__);
                  break;
                }
              mail->locate_keys();
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
          case (CLOSE):
            {
              auto mail = (Mail*) ctx->data;
              if (!Mail::is_valid_ptr (mail))
                {
                  log_debug ("%s:%s: Close for mail which is gone.",
                             SRCNAME, __func__);
                  break;
                }
              Mail::close (mail);
              break;
            }
          case (CRYPTO_DONE):
            {
              auto mail = (Mail*) ctx->data;
              if (!Mail::is_valid_ptr (mail))
                {
                  log_debug ("%s:%s: Crypto done for mail which is gone.",
                             SRCNAME, __func__);
                  break;
                }
              // modify the mail.
              if (mail->crypt_state (Mail::NeedsUpdateInOOM))
                {
                  mail->update_crypt_oom();
                }
              if (mail->crypt_state (Mail::NeedsSecondAfterWrite))
                {
                  // Save the Mail
                  log_debug ("%s:%s: Crypto done for %p Invoking second save.",
                             SRCNAME, __func__, mail);
                  invoke_oom_method (mail->item (), "Save", NULL);
                  log_debug ("%s:%s: Second save done for %p Invoking second send.",
                             SRCNAME, __func__, mail);
                }
              // Finaly this should pass.
              invoke_oom_method (mail->item (), "Send", NULL);
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

static std::vector <LPDISPATCH> explorers;

void
add_explorer (LPDISPATCH explorer)
{
  explorers.push_back (explorer);
}

void remove_explorer (LPDISPATCH explorer)
{
  explorers.erase(std::remove(explorers.begin(),
                              explorers.end(),
                              explorer),
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
        HWND lastChild = NULL;
        for (const auto explorer: explorers)
          {
            /* Casting to LPOLEWINDOW and calling GetWindow
               succeeded in Outlook 2016 but always returned
               the number 1. So we need this hack. */
            char *caption = get_oom_string (explorer, "Caption");
            if (!caption)
              {
                log_debug ("%s:%s: No caption.",
                           SRCNAME, __func__);
                continue;
              }
            /* rctrl_renwnd32 is the window class of outlook. */
            HWND hwnd = FindWindowExA(NULL, lastChild, "rctrl_renwnd32",
                                      caption);
            xfree (caption);
            lastChild = hwnd;
            if (hwnd == cwp->hwnd)
              {
                log_debug ("%s:%s: WM_CLOSE windowmessage for explorer. "
                           "Closing all mails.",
                           SRCNAME, __func__);
                Mail::close_all_mails();
                break;
              }
          }
        break;
      }
     case WM_SYSCOMMAND:
        /*
         This comes to often and when we are closed from the icon
         we also get WM_CLOSE
       if (cwp->wParam == SC_CLOSE)
        {
          log_debug ("%s:%s: SC_CLOSE syscommand. Closing all mails.",
                     SRCNAME, __func__);
          Mail::close_all_mails();
        } */
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

GPGRT_LOCK_DEFINE(invalidate_lock);
static bool invalidation_in_progress;

DWORD WINAPI
delayed_invalidate_ui (LPVOID)
{
  if (invalidation_in_progress)
    {
      log_debug ("%s:%s: Invalidation canceled as it is in progress.",
                 SRCNAME, __func__);
      return 0;
    }
  gpgrt_lock_lock(&invalidate_lock);
  invalidation_in_progress = true;
  /* We sleep here a bit to prevent invalidation immediately
     after the selection change before we have started processing
     the mail. */
  Sleep (500);
  do_in_ui_thread (INVALIDATE_UI, nullptr);
  invalidation_in_progress = false;
  gpgrt_lock_unlock(&invalidate_lock);
  return 0;
}

DWORD WINAPI
close_mail (LPVOID mail)
{
  do_in_ui_thread (CLOSE, mail);
  return 0;
}
