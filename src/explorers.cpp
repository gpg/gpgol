/* explorers.cpp - Code to handle the OOM Explorers
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
#include "explorers.h"
#include "dialogs.h"  /* IDB_xxx */
#include "cmdbarcontrols.h"
#include "revert.h"

#include "eventsink.h"

BEGIN_EVENT_SINK(GpgolExplorersEvents, IOOMExplorersEvents)
  STDMETHOD (NewExplorer) (THIS_ LPOOMEXPLORER);
EVENT_SINK_DEFAULT_CTOR(GpgolExplorersEvents)
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
  LPDISPATCH controls, button, obj;

  log_debug ("%s:%s: Enter", SRCNAME, __func__);

  /* In theory we should take a lock here to avoid a race between the
     test for a new control and the creation.  However, we are not
     called from a second thread.  */

  /* Check that our controls do not already exist.  */
  obj = get_oom_object (explorer, "CommandBars");
  if (!obj)
    {
      log_debug ("%s:%s: CommandBars not found", SRCNAME, __func__);
      return;
    }
  button = get_oom_control_bytag (obj, "GpgOL_Start_Key_Manager");
  gpgol_release (obj);
  obj = NULL;
  if (button)
    {
      gpgol_release (button);
      log_debug ("%s:%s: Leave (Controls are already added)",
                 SRCNAME, __func__);
      return;
    }

  /* Fixme: To avoid an extra lookup we should use the CommandBars
     object to start the search for the controls.  I tried this but
     Outlooked crashed.  Quite likely my error but I was not able to
     find the problem.  */

  /* Create the Start-Key-Manager menu entries.  */
  controls = get_oom_object (explorer,
                             "CommandBars.FindControl(,30007).get_Controls");
  if (!controls)
    log_debug ("%s:%s: Menu Popup Extras not found\n", SRCNAME, __func__);
  else
    {
      button = add_oom_button (controls);
      if (button)
        {
          put_oom_string (button, "Tag", "GpgOL_Start_Key_Manager");
          put_oom_bool (button, "BeginGroup", true);
          put_oom_int (button, "Style", msoButtonIconAndCaption );
          put_oom_string (button, "Caption", _("GnuPG Certificate &Manager"));
          put_oom_icon (button, IDB_KEY_MANAGER_16, 16);
          
          install_GpgolCommandBarButtonEvents_sink (button);
          /* Fixme:  Save the returned object for an Unadvice.  */
          gpgol_release (button);
        }

      button = add_oom_button (controls);
      if (button)
        {
          put_oom_string (button, "Tag", "GpgOL_Revert_Folder");
          put_oom_int (button, "Style", msoButtonCaption );
          put_oom_string (button, "Caption",
                          _("Remove GpgOL flags from this folder"));
          
          install_GpgolCommandBarButtonEvents_sink (button);
          /* Fixme:  Save the returned object for an Unadvice.  */
          gpgol_release (button);
        }

      gpgol_release (controls);
      controls = NULL;
    }

  /* Create the Start-Key-Manager toolbar icon.  Not ethat we need to
     use a different tag name here.  If we won't do that event sink
     would be called twice.  */
  controls = get_oom_object (explorer,
                          "CommandBars.Item(Standard).get_Controls");
  if (!controls)
    log_debug ("%s:%s: CommandBar \"Standard\" not found\n",
               SRCNAME, __func__);
  else
    {
      button = add_oom_button (controls);
      if (button)
        {
          put_oom_string (button, "Tag", "GpgOL_Start_Key_Manager@t");
          put_oom_int (button, "Style", msoButtonIcon );
          put_oom_string (button, "TooltipText",
                          _("Open the certificate manager"));
          put_oom_icon (button, IDB_KEY_MANAGER_16, 16);
          
          install_GpgolCommandBarButtonEvents_sink (button);
          /* Fixme:  store the event sink object somewhere.  */
          gpgol_release (button);
        }
      gpgol_release (controls);
    }
  
  log_debug ("%s:%s: Leave", SRCNAME, __func__);
}



void
run_explorer_revert_folder (LPDISPATCH button)
{
  LPDISPATCH obj;

  log_debug ("%s:%s: Enter", SRCNAME, __func__);

  /* Notify the user that the general GpgOl function will be disabled
     when calling this function.  */
  if ( opt.disable_gpgol
       || (MessageBox 
           (NULL/* FIXME: need the hwnd */,
            _("You are about to start the process of reversing messages "
              "created by GpgOL to prepare deinstalling of GpgOL. "
              "Running this command will put GpgOL into a disabled state "
              "so that messages are not anymore processed by GpgOL.\n"
              "\n"
              "You should convert all folders one after the other with "
              "this command, close Outlook and then deinstall GpgOL.\n"
              "\n"
              "Note that if you start Outlook again with GpgOL still "
              "being installed, GpgOL will again process messages."),
            _("GpgOL"), MB_ICONWARNING|MB_OKCANCEL) == IDOK))
    {
      if ( MessageBox 
           (NULL /* Fixme: need hwnd */,
            _("Do you want to revert this folder?"),
            _("GpgOL"), MB_ICONQUESTION|MB_YESNO) == IDYES )
        {
          if (!opt.disable_gpgol)
            opt.disable_gpgol = 1;

          obj = get_oom_object (button, 
                                "get_Parent.get_Parent.get_Parent.get_Parent"
                                ".get_CurrentFolder");
          if (obj)
            {
              gpgol_folder_revert (obj);
              gpgol_release (obj);
            }
        }
    }
}
