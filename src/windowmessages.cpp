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
#include "wks-helper.h"

#include <stdio.h>

#define RESPONDER_CLASS_NAME "GpgOLResponder"

/* Singleton window */
static HWND g_responder_window = NULL;
static int invalidation_blocked = 0;

LONG_PTR WINAPI
gpgol_window_proc (HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
//  log_debug ("WMG: %x", (unsigned int) message);
  if (message == WM_USER + 42)
    {
      wm_ctx_t *ctx = (wm_ctx_t *) lParam;
      log_debug ("%s:%s: Recieved user msg: %i",
                 SRCNAME, __func__, ctx->wmsg_type);
      switch (ctx->wmsg_type)
        {
          case (PARSING_DONE):
            {
              auto mail = (Mail*) ctx->data;
              if (!Mail::isValidPtr (mail))
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
              if (!Mail::isValidPtr (mail))
                {
                  log_debug ("%s:%s: Recipient add for mail which is gone.",
                             SRCNAME, __func__);
                  break;
                }
              mail->locateKeys_o ();
              break;
            }
          case (REVERT_MAIL):
            {
              auto mail = (Mail*) ctx->data;
              if (!Mail::isValidPtr (mail))
                {
                  log_debug ("%s:%s: Revert mail for mail which is gone.",
                             SRCNAME, __func__);
                  break;
                }

              mail->setNeedsSave (true);
              /* Some magic here. Accessing any existing inline body cements
                 it. Otherwise updating the body through the revert also changes
                 the body of a inline mail. */
              char *inlineBody = get_inline_body ();
              xfree (inlineBody);

              // Does the revert.
              log_debug ("%s:%s: Revert mail. Invoking save.",
                         SRCNAME, __func__);
              invoke_oom_method (mail->item (), "Save", NULL);
              log_debug ("%s:%s: Revert mail. Save done. Updating body..",
                         SRCNAME, __func__);
              mail->updateBody_o ();
              log_debug ("%s:%s: Revert mail done.",
                         SRCNAME, __func__);
              break;
            }
          case (INVALIDATE_UI):
            {
              if (!invalidation_blocked)
                {
                  log_debug ("%s:%s: Invalidating UI",
                             SRCNAME, __func__);
                  gpgoladdin_invalidate_ui();
                  log_debug ("%s:%s: Invalidation done",
                             SRCNAME, __func__);
                }
              else
                {
                  log_debug ("%s:%s: Received invalidation msg while blocked."
                             " Ignoring it",
                             SRCNAME, __func__);
                }
              break;
            }
          case (INVALIDATE_LAST_MAIL):
            {
              log_debug ("%s:%s: clearing last mail",
                         SRCNAME, __func__);
              Mail::clearLastMail ();
              break;
            }
          case (CLOSE):
            {
              auto mail = (Mail*) ctx->data;
              if (!Mail::isValidPtr (mail))
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
              if (!Mail::isValidPtr (mail))
                {
                  log_debug ("%s:%s: Crypto done for mail which is gone.",
                             SRCNAME, __func__);
                  break;
                }
              // modify the mail.
              if (mail->cryptState () == Mail::NeedsUpdateInOOM)
                {
                  // Save the Mail
                  log_debug ("%s:%s: Crypto done for %p updating oom.",
                             SRCNAME, __func__, mail);
                  mail->updateCryptOOM_o ();
                }
              if (mail->cryptState () == Mail::NeedsSecondAfterWrite)
                {
                  invoke_oom_method (mail->item (), "Save", NULL);
                  log_debug ("%s:%s: Second save done for %p Invoking second send.",
                             SRCNAME, __func__, mail);
                }
              // Finaly this should pass.
              invoke_oom_method (mail->item (), "Send", NULL);
              log_debug ("%s:%s:  Send for %p completed.",
                         SRCNAME, __func__, mail);
              mail->releaseCurrentItem();
              break;
            }
          case (BRING_TO_FRONT):
            {
              HWND wnd = get_active_hwnd ();
              if (wnd)
                {
                  log_debug ("%s:%s: Bringing window %p to front.",
                             SRCNAME, __func__, wnd);
                  bring_to_front (wnd);
                }
              else
                {
                  log_debug ("%s:%s: No active window found for bring to front.",
                             SRCNAME, __func__);
                }
              break;
            }
          case (WKS_NOTIFY):
            {
              WKSHelper::instance ()->notify ((const char *) ctx->data);
              xfree (ctx->data);
              break;
            }
          case (CLEAR_REPLY_FORWARD):
            {
              auto mail = (Mail*) ctx->data;
              if (!Mail::isValidPtr (mail))
                {
                  log_debug ("%s:%s: Clear reply forward for mail which is gone.",
                             SRCNAME, __func__);
                  break;
                }
              mail->wipe_o (true);
              mail->removeAllAttachments_o ();
              break;
            }
          case (DO_AUTO_SECURE):
            {
              auto mail = (Mail*) ctx->data;
              if (!Mail::isValidPtr (mail))
                {
                  log_debug ("%s:%s: DO_AUTO_SECURE for mail which is gone.",
                             SRCNAME, __func__);
                  break;
                }
              mail->setDoAutosecure_m (true);
              break;
            }
          case (DONT_AUTO_SECURE):
            {
              auto mail = (Mail*) ctx->data;
              if (!Mail::isValidPtr (mail))
                {
                  log_debug ("%s:%s: DO_AUTO_SECURE for mail which is gone.",
                             SRCNAME, __func__);
                  break;
                }
              mail->setDoAutosecure_m (false);
              break;
            }
          default:
            log_debug ("%s:%s: Unknown msg %x",
                       SRCNAME, __func__, ctx->wmsg_type);
        }
        return 0;
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

static int
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
  SendMessage (responder, WM_USER + 42, 0, (LPARAM) ctx);
  return 0;
}

int
do_in_ui_thread (gpgol_wmsg_type type, void *data)
{
  wm_ctx_t ctx = {NULL, UNKNOWN, 0, 0};
  ctx.wmsg_type = type;
  ctx.data = data;

  log_debug ("%s:%s: Sending message of type %i",
             SRCNAME, __func__, type);

  if (send_msg_to_ui_thread (&ctx))
    {
      return -1;
    }
  return ctx.err;
}

static DWORD WINAPI
do_async (LPVOID arg)
{
  wm_ctx_t *ctx = (wm_ctx_t*) arg;
  log_debug ("%s:%s: Do async with type %i after %i ms",
             SRCNAME, __func__, ctx ? ctx->wmsg_type : -1,
             ctx->delay);
  if (ctx->delay)
    {
      Sleep (ctx->delay);
    }
  send_msg_to_ui_thread (ctx);
  xfree (ctx);
  return 0;
}

void
do_in_ui_thread_async (gpgol_wmsg_type type, void *data, int delay)
{
  wm_ctx_t *ctx = (wm_ctx_t *) xcalloc (1, sizeof (wm_ctx_t));
  ctx->wmsg_type = type;
  ctx->data = data;
  ctx->delay = delay;

  CloseHandle (CreateThread (NULL, 0, do_async, (LPVOID) ctx, 0, NULL));
}

LRESULT CALLBACK
gpgol_hook(int code, WPARAM wParam, LPARAM lParam)
{
/* Once we are in the close events we don't have enough
   control to revert all our changes so we have to do it
   with this nice little hack by catching the WM_CLOSE message
   before it reaches outlook. */
  LPCWPSTRUCT cwp = (LPCWPSTRUCT) lParam;

  /* What we do here is that we catch all WM_CLOSE messages that
     get to Outlook. Then we check if the last open Explorer
     is the target of the close. In set case we start our shutdown
     routine before we pass the WM_CLOSE to outlook */
  switch (cwp->message)
    {
      case WM_CLOSE:
      {
        HWND lastChild = NULL;
        log_debug ("%s:%s: Got WM_CLOSE",
                   SRCNAME, __func__);
        if (!GpgolAddin::get_instance() || !GpgolAddin::get_instance ()->get_application())
          {
            TRACEPOINT;
            break;
          }
        LPDISPATCH explorers = get_oom_object (GpgolAddin::get_instance ()->get_application(),
                                               "Explorers");

        if (!explorers)
          {
            log_error ("%s:%s: No explorers object",
                       SRCNAME, __func__);
            break;
          }
        int count = get_oom_int (explorers, "Count");

        if (count != 1)
          {
            log_debug ("%s:%s: More then one explorer. Not shutting down.",
                       SRCNAME, __func__);
            gpgol_release (explorers);
            break;
          }

        LPDISPATCH explorer = get_oom_object (explorers, "Item(1)");
        gpgol_release (explorers);

        if (!explorer)
          {
            TRACEPOINT;
            break;
          }

        /* Casting to LPOLEWINDOW and calling GetWindow
           succeeded in Outlook 2016 but always returned
           the number 1. So we need this hack. */
        char *caption = get_oom_string (explorer, "Caption");
        gpgol_release (explorer);
        if (!caption)
          {
            log_debug ("%s:%s: No caption.",
                       SRCNAME, __func__);
            break;
          }
        /* rctrl_renwnd32 is the window class of outlook. */
        HWND hwnd = FindWindowExA(NULL, lastChild, "rctrl_renwnd32",
                                  caption);
        xfree (caption);
        lastChild = hwnd;
        if (hwnd == cwp->hwnd)
          {
            log_debug ("%s:%s: WM_CLOSE windowmessage for explorer. "
                       "Shutting down.",
                       SRCNAME, __func__);
            GpgolAddin::get_instance ()->shutdown();
            break;
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
          GpgolAddin::get_instance ()->shutdown();
        } */
       break;
     default:
//       log_debug ("WM: %x", (unsigned int) cwp->message);
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

GPGRT_LOCK_DEFINE (invalidate_lock);

static bool invalidation_in_progress;

DWORD WINAPI
delayed_invalidate_ui (LPVOID minsleep)
{
  if (invalidation_in_progress)
    {
      log_debug ("%s:%s: Invalidation canceled as it is in progress.",
                 SRCNAME, __func__);
      return 0;
    }
  TRACEPOINT;
  invalidation_in_progress = true;
  gpgrt_lock_lock(&invalidate_lock);

  int sleep_ms = (intptr_t)minsleep;
  Sleep (sleep_ms);
  int i = 0;
  while (invalidation_blocked)
    {
      Sleep (100);
      i++;

      if (i % 10 == 0)
        {
          log_debug ("%s:%s: Waiting for invalidation.",
                     SRCNAME, __func__);
        }

      /* Do we need an abort statement here? */
    }
  do_in_ui_thread (INVALIDATE_UI, nullptr);
  TRACEPOINT;
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

void
blockInv()
{
  invalidation_blocked++;
  log_oom_extra ("%s:%s: Invalidation block count %i",
                 SRCNAME, __func__, invalidation_blocked);
}

void
unblockInv()
{
  invalidation_blocked--;
  log_oom_extra ("%s:%s: Invalidation block count %i",
                 SRCNAME, __func__, invalidation_blocked);

  if (invalidation_blocked < 0)
    {
      log_error ("%s:%s: Invalidation block mismatch",
                 SRCNAME, __func__);
      invalidation_blocked = 0;
    }
}
