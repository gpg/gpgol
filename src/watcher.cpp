/* watcher.cpp
 *	Copyright (C) 2005 g10 Code GmbH
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

#error not anymore used.

#include <config.h>
#include <windows.h>
#include <stdio.h>
#include <gpgme.h>

#include "mymapi.h"
#include "myexchext.h"
#include "mymapitags.h"
#include "common.h"


/* Exchange callback context to retrieve the last message. */
static LPEXCHEXTCALLBACK g_cb = NULL;

/* The current hook handle. */
static HHOOK             g_cbt_hook = NULL;

/* MAPI message and storage handle. */
static LPMESSAGE         g_msg = NULL;
static LPMDB             g_mdb = NULL;
static HWND              g_creat_wnd = 0;

static HWND
find_message_window2 (HWND parent)
{
  HWND child;

  if (!parent)
    return NULL;

  child = GetWindow (parent, GW_CHILD);
  while (child)
    {
      char buf[1024+1];
      HWND w;
      size_t len;
      const char *s;
      
      memset (buf, 0, sizeof (buf));
      GetWindowText (child, buf, sizeof (buf)-1);
      len = strlen (buf);
      if (len > 22
          && (s = strstr (buf, "-----BEGIN PGP "))
          &&  (!strncmp (s+15, "MESSAGE-----", 12)
               || !strncmp (s+15, "SIGNED MESSAGE-----", 19)))
        return child;
      w = find_message_window2 (child);
      if (w)
	{
	  log_debug ("%s: watcher found message window: %p\n", __func__, w);
	  return w;
	}
      
      child = GetNextWindow (child, GW_HWNDNEXT);	
    }

  return NULL;
}


/* Decrypt a single message. */
static void
decrypt_message (HWND hwnd, LPMESSAGE msg)
{
  GpgMsg *m = CreateGpgMsg (msg);
  m->setExchangeCallback ((void *)g_cb);
  m->decrypt (hwnd, 0);
  delete m;
  UlRelease (msg);
  msg = NULL;
}

/* This hook procedure allows to monitor all windows which are created
 * and/or activated.
 * We use it to find the new opened message window and start automatically
 * the decryption code. The pre-condition is, that the new window has the
 * proper window class and a child window with valid OpenPGP data.
 *
 * We cannot use the HCBT_CREATEWND alone because at the point the hook
 * procedure is called, the window is only about to be created but the
 * procedure has not finished yet. Plus the child windows are not yet
 * created. Thus we check if the new window has the proper class. If
 * this window is then activated, we call the decryption code and reset
 * the window handle so the decryption is only called once.
 */
static LRESULT CALLBACK
cbt_proc (int code, WPARAM w, LPARAM l)
{
  char wclass[128];
  HWND msgwnd, hwnd;

  if (code < 0)
    return CallNextHookEx (g_cbt_hook, code, w, l);
  
  hwnd = (HWND)w;
  if (code == HCBT_CREATEWND)
    {
      GetClassName (hwnd, wclass, 127);
      if (strstr (wclass, "rctrl_renwnd32"))
	{
	  g_creat_wnd = hwnd;
	  log_debug ("%s: watch for window %p\n", __func__, hwnd);
	}
      
    }
  
  if (code == HCBT_ACTIVATE && g_creat_wnd == hwnd)
    {
      log_debug ("%s: %p == %p?\n", __func__, g_creat_wnd, hwnd);
      g_creat_wnd = NULL;
      msgwnd = find_message_window2 (hwnd);
      if (msgwnd && g_msg)
	{
	  log_debug ("%s: decrypt_message(%p, %p)\n", __func__, hwnd, g_msg);
	  decrypt_message (hwnd, g_msg);
	  UlRelease (g_mdb); 
	  g_mdb = NULL;
	}
    }

    return CallNextHookEx (g_cbt_hook, code, w, l);
}


/* Initialize the CBT hook. */
extern "C" int
watcher_init_hook (void)
{
  if (g_cbt_hook != NULL)
    return 0;
  g_cbt_hook = SetWindowsHookEx (WH_CBT, cbt_proc, glob_hinst, 0);
  if (!g_cbt_hook)
    {
      log_debug ("%s: SetWindowsHookEx failed ec=%d\n", 
		 __func__, (int)GetLastError ());
      return -1;
    }  
  return 0;
}


/* Remove the CBT hook. */
extern "C" int
watcher_free_hook (void)
{
  if (g_msg != NULL) 
    {
      UlRelease (g_msg);
      g_msg = NULL;
    }
  if (g_mdb != NULL) 
    {
      UlRelease (g_mdb);
      g_mdb = NULL;
    }
  if (g_cbt_hook != NULL)
    {
      UnhookWindowsHookEx (g_cbt_hook);
      g_cbt_hook = NULL;
    }  
  return 0;
}


/* Set the Exchange callback context. */
extern "C" void
watcher_set_callback_ctx (void *cb)
{     
  HRESULT hr;
  
  g_cb = (LPEXCHEXTCALLBACK)cb;

  if (g_msg != NULL)
    UlRelease (g_msg);
  if (g_mdb != NULL)
    UlRelease (g_mdb);
  hr = g_cb->GetObject (&g_mdb, (LPMAPIPROP *)&g_msg);
  if (FAILED (hr)) 
    {
      log_debug ("%s: GetObject() failed ec=%lx\n", __func__, hr);
      g_mdb = NULL;
      g_msg = NULL;
    }
}
