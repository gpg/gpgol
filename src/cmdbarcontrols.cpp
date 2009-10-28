/* cmdbarcontrols.cpp - Code to handle the CommandBarControls
 *	Copyright (C) 2009 g10 Code GmbH
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

#ifdef HAVE_CONFIG_H
#include <config.h>
#endif

#include <windows.h>
#include <olectl.h>

#include "common.h"
#include "oomhelp.h"
#include "cmdbarcontrols.h"
#include "inspectors.h"
#include "engine.h"

#include "eventsink.h"


/* Subclass of CommandBarButtonEvents so that we can hook into the
   Click event.  */
BEGIN_EVENT_SINK(GpgolCommandBarButtonEvents, IOOMCommandBarButtonEvents)
  STDMETHOD (Click) (THIS_ LPDISPATCH, PBOOL);
EVENT_SINK_DEFAULT_DTOR(GpgolCommandBarButtonEvents)
EVENT_SINK_INVOKE(GpgolCommandBarButtonEvents)
{
  HRESULT hr;
  (void)lcid; (void)riid; (void)result; (void)exepinfo; (void)argerr;

  if (dispid == 1 && (flags & DISPATCH_METHOD))
    {
      if (!parms) 
        hr = DISP_E_PARAMNOTOPTIONAL;
      else if (parms->cArgs != 2)
        hr = DISP_E_BADPARAMCOUNT;
      else if (parms->rgvarg[0].vt != (VT_BOOL|VT_BYREF)
               || parms->rgvarg[1].vt != VT_DISPATCH)
        hr = DISP_E_BADVARTYPE;
      else
        {
          BOOL cancel_default = !!*parms->rgvarg[0].pboolVal;
          hr = Click (parms->rgvarg[1].pdispVal, (PBOOL)&cancel_default);
          *parms->rgvarg[0].pboolVal = (cancel_default 
                                        ? VARIANT_TRUE:VARIANT_FALSE);
        }
    }
  else
    hr = DISP_E_MEMBERNOTFOUND;
  return hr;
}
END_EVENT_SINK(GpgolCommandBarButtonEvents, IID_IOOMCommandBarButtonEvents)




static int
tagcmp (const char *a, const char *b)
{
  return strncmp (a, b, strlen (b));
}


/* This is the event sink for a button click.  */
STDMETHODIMP
GpgolCommandBarButtonEvents::Click (LPDISPATCH button, PBOOL cancel_default)
{
  char *tag;

  (void)cancel_default;
  log_debug ("%s:%s: Called", SRCNAME, __func__);

  {
    char *tmp = get_object_name (button);
    log_debug ("%s:%s: button is %p (%s)", 
               SRCNAME, __func__, button, tmp? tmp:"(null)");
    xfree (tmp);
  }

  tag = get_oom_string (button, "Tag");
  log_debug ("%s:%s: button's tag is (%s)", 
             SRCNAME, __func__, tag? tag:"(null)");
  if (!tag)
    ;
  else if (!tagcmp (tag, "GpgOL_Start_Key_Manager"))
    {
      /* FIXME: We don't have the current window handle.  */
      if (engine_start_keymanager (NULL))
         MessageBox (NULL, _("Could not start certificate manager"),
                     _("GpgOL"), MB_ICONERROR|MB_OK);
    }
  else if (!tagcmp (tag, "GpgOL_Inspector_Crypto_Info"))
    {
      /* FIXME: We should invoke the decrypt/verify again. */
      update_inspector_crypto_info (button);
#if 0 /* This is the code we used to use.  */
      log_debug ("%s:%s: command CryptoState called\n", SRCNAME, __func__);
      hr = eecb->GetObject (&mdb, (LPMAPIPROP *)&message);
      if (SUCCEEDED (hr))
        {
          if (message_incoming_handler (message, hwnd, true))
            message_display_handler (eecb, hwnd);
	}
      else
        log_debug_w32 (hr, "%s:%s: command CryptoState failed", 
                       SRCNAME, __func__);
      ul_release (message, __func__, __LINE__);
      ul_release (mdb, __func__, __LINE__);
#endif

    }

  xfree (tag);
  return S_OK;
}


