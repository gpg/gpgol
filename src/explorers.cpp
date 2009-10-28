/* explorers.cpp - Code to handle the OOM Explorers
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
#include "explorers.h"
#include "dialogs.h"  /* IDB_xxx */
#include "cmdbarcontrols.h"

#include "eventsink.h"

BEGIN_EVENT_SINK(GpgolExplorersEvents, IOOMExplorersEvents)
  STDMETHOD (NewExplorer) (THIS_ LPOOMEXPLORER);
EVENT_SINK_DEFAULT_DTOR(GpgolExplorersEvents)
EVENT_SINK_INVOKE(GpgolExplorersEvents)
{
  HRESULT hr;
  (void)lcid; (void)riid; (void)result; (void)exepinfo; (void)argerr;

  if (dispid == 0xf001 && (flags & DISPATCH_METHOD))
    {
      if (!parms) 
        hr = DISP_E_PARAMNOTOPTIONAL;
      else if (parms->cArgs != 1)
        hr = DISP_E_BADPARAMCOUNT;
      else if (parms->rgvarg[0].vt != VT_DISPATCH)
        hr = DISP_E_BADVARTYPE;
      else
        hr = NewExplorer ((LPOOMEXPLORER)parms->rgvarg[0].pdispVal);
    }
  else
    hr = DISP_E_MEMBERNOTFOUND;
  return hr;
}
END_EVENT_SINK(GpgolExplorersEvents, IID_IOOMExplorersEvents)




/* The method called by outlook for each new explorer.  Note that
   Outlook sometimes reuses Inspectro objects thus this event is not
   an indication for a newly opened Explorer.  */
STDMETHODIMP
GpgolExplorersEvents::NewExplorer (LPOOMEXPLORER explorer)
{
  log_debug ("%s:%s: Called", SRCNAME, __func__);

  add_explorer_controls (explorer);
  return S_OK;
}


/* Add buttons to the explorer.  */
void
add_explorer_controls (LPOOMEXPLORER explorer)
{
  LPDISPATCH pObj, pDisp, pTmp;

  log_debug ("%s:%s: Enter", SRCNAME, __func__);

  /* In theory we should take a lock here to avoid a race between the
     test for a new control and the creation.  However, we are not
     called from a second thread.  */

  /* Check that our controls do not already exist.  */
  pObj = get_oom_object (explorer, "CommandBars");
  if (!pObj)
    {
      log_debug ("%s:%s: CommandBars not found", SRCNAME, __func__);
      return;
    }
  pDisp = get_oom_control_bytag (pObj, "GpgOL_Start_Key_Manager");
  pObj->Release ();
  pObj = NULL;
  if (pDisp)
    {
      pDisp->Release ();
      log_debug ("%s:%s: Leave (Controls are already added)",
                 SRCNAME, __func__);
      return;
    }

  /* Fixme: To avoid an extra lookup we should use the CommandBars
     object to start the search for the controls.  I tried this but
     Outlooked crashed.  Quite likely my error but I was not able to
     find the problem.  */

  /* Create the Start-Key-Manager menu entry.  */
  pDisp = get_oom_object (explorer,
                          "CommandBars.FindControl(,30007).get_Controls");
  if (!pDisp)
    log_debug ("%s:%s: Menu Popup Extras not found\n", SRCNAME, __func__);
  else
    {
      pTmp = add_oom_button (pDisp);
      pDisp->Release ();
      pDisp = pTmp;
      if (pDisp)
        {
          put_oom_string (pDisp, "Tag", "GpgOL_Start_Key_Manager");
          put_oom_int (pDisp, "Style", msoButtonIconAndCaption );
          put_oom_string (pDisp, "Caption", _("GnuPG Certificate &Manager"));
          put_oom_string (pDisp, "TooltipText",
                          _("Open the certificate manager"));
          put_oom_icon (pDisp, IDB_KEY_MANAGER, 16);
          
          install_GpgolCommandBarButtonEvents_sink (pDisp);
          pDisp->Release ();
        }
    }

  /* Create the Start-Key-Manager toolbar icon.  Not ethat we need to
     use a different tag name here.  If we won't do that event sink
     would be called twice.  */
  pDisp = get_oom_object (explorer,
                          "CommandBars.Item(Standard).get_Controls");
  if (!pDisp)
    log_debug ("%s:%s: CommandBar \"Standard\" not found\n",
               SRCNAME, __func__);
  else
    {
      pTmp = add_oom_button (pDisp);
      pDisp->Release ();
      pDisp = pTmp;
      if (pDisp)
        {
          put_oom_string (pDisp, "Tag", "GpgOL_Start_Key_Manager@t");
          put_oom_int (pDisp, "Style", msoButtonIcon );
          put_oom_string (pDisp, "TooltipText",
                          _("Open the certificate manager"));
          put_oom_icon (pDisp, IDB_KEY_MANAGER, 16);
          
          install_GpgolCommandBarButtonEvents_sink (pDisp);
          /* Fixme:  store the event sink object somewhere.  */
          pDisp->Release ();
        }
    }
  
  log_debug ("%s:%s: Leave", SRCNAME, __func__);
}

