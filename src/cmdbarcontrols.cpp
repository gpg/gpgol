/* cmdbarcontrols.cpp - Code to handle the CommandBarControls
 * Copyright (C) 2009 g10 Code GmbH
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
#include "explorers.h"
#include "engine.h"

#include "eventsink.h"


/* Subclass of CommandBarButtonEvents so that we can hook into the
   Click event.  */
BEGIN_EVENT_SINK(GpgolCommandBarButtonEvents, IOOMCommandBarButtonEvents)
  STDMETHOD (Click) (THIS_ LPDISPATCH, PBOOL);
EVENT_SINK_DEFAULT_CTOR(GpgolCommandBarButtonEvents)
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




/* This is the event sink for a button click.  */
STDMETHODIMP
GpgolCommandBarButtonEvents::Click (LPDISPATCH button, PBOOL cancel_default)
{
  char *tag;
  int instid;

  (void)cancel_default;

  tag = get_oom_string (button, "Tag");
  instid = get_oom_int (button, "InstanceId");
  {
    char *tmp = get_object_name (button);
    log_debug ("%s:%s: button is %p (%s) tag is (%s) instid %d", 
               SRCNAME, __func__, button, 
               tmp? tmp:"(null)",
               tag? tag:"(null)", instid);
    xfree (tmp);
  }

  if (!tag)
    ;
  else if (!tagcmp (tag, "GpgOL_Inspector"))
    {
      proc_inspector_button_click (button, tag, instid);
    }
  else if (!tagcmp (tag, "GpgOL_Start_Key_Manager"))
    {
      /* FIXME: We don't have the current window handle.  */
      if (engine_start_keymanager (NULL))
         MessageBox (NULL, _("Could not start certificate manager"),
                     _("GpgOL"), MB_ICONERROR|MB_OK);
    }
  else if (!tagcmp (tag, "GpgOL_Revert_Folder"))
    {
      run_explorer_revert_folder (button);
    }

  xfree (tag);
  return S_OK;
}


