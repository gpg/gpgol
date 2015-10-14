/* inspectors.cpp - Code to handle OOM Inspectors
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
#include "myexchext.h"
#include "mapihelp.h"
#include "message.h"
#include "dialogs.h"  /* IDB_xxx */
#include "cmdbarcontrols.h"
#include "eventsink.h"
#include "inspectors.h"
#include "mailitem.h"
#include "revert.h"


/* Event sink for an Inspectors collection object.  */
BEGIN_EVENT_SINK(GpgolInspectorsEvents, IOOMInspectorsEvents)
  STDMETHOD (NewInspector) (THIS_ LPOOMINSPECTOR);
EVENT_SINK_DEFAULT_CTOR(GpgolInspectorsEvents)
EVENT_SINK_DEFAULT_DTOR(GpgolInspectorsEvents)
EVENT_SINK_INVOKE(GpgolInspectorsEvents)
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
        hr = NewInspector ((LPOOMINSPECTOR)parms->rgvarg[0].pdispVal);
    }
  else
    hr = DISP_E_MEMBERNOTFOUND;
  return hr;
}
END_EVENT_SINK(GpgolInspectorsEvents, IID_IOOMInspectorsEvents)



/* Event sink for an Inspector object.  */
typedef struct GpgolInspectorEvents GpgolInspectorEvents;
typedef GpgolInspectorEvents *LPGPGOLINSPECTOREVENTS;

BEGIN_EVENT_SINK(GpgolInspectorEvents, IOOMInspectorEvents)
  STDMETHOD_ (void, Activate) (THIS_);
  STDMETHOD_ (void, Close) (THIS_);
  STDMETHOD_ (void, Deactivate) (THIS_);
  bool m_first_activate_seen;
EVENT_SINK_CTOR(GpgolInspectorEvents)
{
  m_first_activate_seen = false;
}
EVENT_SINK_DEFAULT_DTOR(GpgolInspectorEvents)
EVENT_SINK_INVOKE(GpgolInspectorEvents)
{
  HRESULT hr = S_OK;
  (void)lcid; (void)riid; (void)result; (void)exepinfo; (void)argerr;
  
  if ((dispid == 0xf001 || dispid == 0xf006 || dispid == 0xf008)
      && (flags & DISPATCH_METHOD))
    {
      if (!parms) 
        hr = DISP_E_PARAMNOTOPTIONAL;
      else if (parms->cArgs != 0)
        hr = DISP_E_BADPARAMCOUNT;
      else if (dispid == 0xf001)
        Activate ();
      else if (dispid == 0xf006)
        Deactivate ();
      else if (dispid == 0xf008)
        Close ();
    }
  else
    hr = DISP_E_MEMBERNOTFOUND;
  return hr;
}
END_EVENT_SINK(GpgolInspectorEvents, IID_IOOMInspectorEvents)



/* A linked list as a simple collection of button.  */
struct button_list_s
{
  struct button_list_s *next;
  LPDISPATCH sink;
  LPDISPATCH button;
  int instid;
  char tag[1];         /* Variable length string.  */
};
typedef struct button_list_s *button_list_t;


/* To avoid messing around with the OOM (adding extra objects as user
   properties and such), we keep our own information structure about
   inspectors. */
struct inspector_info_s
{
  /* We are pretty lame and keep all inspectors in a linked list.  */
  struct inspector_info_s *next;

  /* The event sink object.  This is used by the event methods to
     locate the inspector object.  */
  LPGPGOLINSPECTOREVENTS eventsink;

  /* The inspector object.  */
  LPOOMINSPECTOR inspector;

  /* The Window handle of the inspector.  */
  HWND hwnd;

  /* A list of all the buttons.  */
  button_list_t buttons;

};
typedef struct inspector_info_s *inspector_info_t;

/* The list of all inspectors and a lock for it.  */
static inspector_info_t all_inspectors;
static HANDLE all_inspectors_lock;


static void add_inspector_controls (LPOOMINSPECTOR inspector);
static void update_crypto_info (LPDISPATCH inspector);



/* Initialize this module.  Returns 0 on success.  Called once by dllinit.  */
int
initialize_inspectors (void)
{
  SECURITY_ATTRIBUTES sa;
  
  memset (&sa, 0, sizeof sa);
  sa.bInheritHandle = FALSE;
  sa.lpSecurityDescriptor = NULL;
  sa.nLength = sizeof sa;
  all_inspectors_lock = CreateMutex (&sa, FALSE, NULL);
  return !all_inspectors_lock;
}


/* Acquire the all_inspectors_lock.  No error is returned because we
   can't do anything anyway.  */
static void
lock_all_inspectors (void)
{
  int ec = WaitForSingleObject (all_inspectors_lock, INFINITE);
  if (ec != WAIT_OBJECT_0)
    {
      log_error ("%s:%s: waiting on mutex failed: ec=%#x\n", 
                 SRCNAME, __func__, ec);
      fatal_error ("%s:%s: waiting on mutex failed: ec=%#x\n", 
                   SRCNAME, __func__, ec);
    }
}


/* Release the all_inspectors_lock.  No error is returned because this
   is a fatal error anyway and there is no way to clean up.  */
static void
unlock_all_inspectors (void)
{
  if (!ReleaseMutex (all_inspectors_lock))
    {
      log_error_w32 (-1, "%s:%s: ReleaseMutex failed", SRCNAME, __func__);
      fatal_error ("%s:%s: ReleaseMutex failed", SRCNAME, __func__);
    }
}


/* Add SINK and BUTTON to the list at LISTADDR.  The list takes
   ownership of SINK and BUTTON, thus the caller may not use OBJ or
   OBJ2 after this call.  TAG must be given without the '#' marked
   suffix.  */
static void
move_to_button_list (button_list_t *listaddr, 
                     LPDISPATCH sink, LPDISPATCH button, const char *tag)
{
  button_list_t item;
  int instid;

  if (!tag)
    tag = "";

  instid = button? get_oom_int (button, "InstanceId"): 0;

  log_debug ("%s:%s: sink=%p btn=%p tag=(%s) instid=%d",
             SRCNAME, __func__, sink, button, tag, instid);

  item = (button_list_t)xcalloc (1, sizeof *item + strlen (tag));
  item->sink = sink;
  item->button = button;
  item->instid = instid;
  strcpy (item->tag, tag);
  item->next = *listaddr;
  *listaddr = item;
}


static HWND
find_ole_window (LPOOMINSPECTOR inspector)
{
  HRESULT hr;
  LPOLEWINDOW olewndw = NULL;
  HWND hwnd = NULL;

  hr = inspector->QueryInterface (IID_IOleWindow, (void**)&olewndw);
  if (hr != S_OK || !olewndw)
    {
      log_error ("%s:%s: IOleWindow not found: hr=%#lx", SRCNAME, __func__, hr);
      return NULL;
    }

  hr = olewndw->GetWindow (&hwnd);
  if (hr != S_OK || !hwnd)
    {
      log_error ("%s:%s: IOleWindow->GetWindow failed: hr=%#lx", 
                 SRCNAME, __func__, hr);
      hwnd = NULL;
    }
  olewndw->Release ();
  log_debug ("%s:%s: inspector %p has hwnd=%p",
             SRCNAME, __func__, inspector, hwnd);
  return hwnd;
}



/* Register the inspector object INSPECTOR with its event SINK.  */
static void
register_inspector (LPGPGOLINSPECTOREVENTS sink, LPOOMINSPECTOR inspector)
{
  inspector_info_t item;
  HWND hwnd;

  log_debug ("%s:%s: Called (sink=%p, inspector=%p)",
             SRCNAME, __func__, sink, inspector);
  hwnd = find_ole_window (inspector);
  item = (inspector_info_t)xcalloc (1, sizeof *item);
  lock_all_inspectors ();

  sink->AddRef ();
  item->eventsink = sink;

  inspector->AddRef ();
  item->inspector = inspector;

  item->hwnd = hwnd;

  item->next = all_inspectors;
  all_inspectors = item;

  unlock_all_inspectors ();
}


/* Deregister the inspector with the event SINK.  */
static void
deregister_inspector (LPGPGOLINSPECTOREVENTS sink)
{
  inspector_info_t r, rprev;
  button_list_t ol, ol2;

  log_debug ("%s:%s: Called (sink=%p)", SRCNAME, __func__, sink);

  if (!sink)
    return;

  lock_all_inspectors ();
  for (r = all_inspectors, rprev = NULL; r; rprev = r, r = r->next)
    if (r->eventsink == sink)
      {
        if (!rprev)
          all_inspectors = r->next;
        else
          rprev->next = r->next;
        r->next = NULL;
        break; 
      }
  unlock_all_inspectors ();
  if (!r)
    {
      log_error ("%s:%s: inspector not registered", SRCNAME, __func__);
      return;
    }

  detach_GpgolInspectorEvents_sink (r->eventsink);

  for (ol = r->buttons; ol; ol = ol2)
    {
      ol2 = ol->next;
      if (ol->sink)
        {
          detach_GpgolCommandBarButtonEvents_sink (ol->sink);
          ol->sink->Release ();
        }
      if (ol->button)
        {
          del_oom_button (ol->button);
          ol->button->Release ();
        }
      xfree (ol);
    }

  r->inspector->Release ();
  r->eventsink->Release ();

  xfree (r);
}


/* Return the inspector info for INSPECTOR.  On success the
   ALL_INSPECTORS list is locked and thus the caller should call
   unlock_all_inspectors as soon as possible.  On error NULL is
   returned and the caller must *not* call unlock_all_inspectors.  */
static inspector_info_t
get_inspector_info (LPOOMINSPECTOR inspector)
{
  inspector_info_t r;

  if (!inspector)
    return NULL;

  lock_all_inspectors ();
  for (r = all_inspectors; r; r = r->next)
    if (r->inspector == inspector)
      return r;
  unlock_all_inspectors ();
  return NULL;
}


/* Return the button with TAG and assigned to INSPECTOR.  TAG must be
   given without the suffix.  Returns NULL if not found.  */
static LPDISPATCH
get_button (LPDISPATCH inspector, const char *tag)
{
  LPDISPATCH result = NULL;
  inspector_info_t iinfo;
  button_list_t ol;

  lock_all_inspectors ();

  for (iinfo = all_inspectors; iinfo; iinfo = iinfo->next)
    if (iinfo->inspector == inspector)
      {
        for (ol = iinfo->buttons; ol; ol = ol->next)
          if (ol->tag && !strcmp (ol->tag, tag))
            {
              result = ol->button;
              if (result)
                result->AddRef ();
              break;
            }
        break;
      }
  
  unlock_all_inspectors ();
  return result;
}


/* Search through all objects and find the inspector which has a
   button with the instance id INSTID.  Returns NULL if not found.  */
static LPDISPATCH
get_inspector_from_instid (int instid)
{
  LPDISPATCH result = NULL;
  inspector_info_t iinfo;
  button_list_t ol;

  lock_all_inspectors ();

  for (iinfo = all_inspectors; iinfo; iinfo = iinfo->next)
    for (ol = iinfo->buttons; ol; ol = ol->next)
      if (ol->instid == instid)
        {
          result = iinfo->inspector;
          if (result)
            result->AddRef ();
          break;
        }
  
  unlock_all_inspectors ();
  return result;
}


/* Search through all objects and find the inspector which has a
   the window handle HWND.  Returns NULL if not found.  */
LPDISPATCH
get_inspector_from_hwnd (HWND hwnd)
{
  LPDISPATCH result = NULL;
  inspector_info_t iinfo;

  lock_all_inspectors ();

  for (iinfo = all_inspectors; iinfo; iinfo = iinfo->next)
    if (iinfo->hwnd == hwnd)
        {
          result = iinfo->inspector;
          if (result)
            result->AddRef ();
          break;
        }
  
  unlock_all_inspectors ();
  return result;
}


/* The method called by outlook for each new inspector.  Note that
   Outlook sometimes reuses Inspectro objects thus this event is not
   an indication for a newly opened Inspector.  */
STDMETHODIMP
GpgolInspectorsEvents::NewInspector (LPOOMINSPECTOR inspector)
{
  LPDISPATCH obj;

  log_debug ("%s:%s: Called (this=%p, inspector=%p)",
             SRCNAME, __func__, this, inspector);

  /* It has been said that INSPECTOR here a "weak" reference. This may
     mean that the object has not been fully initialized.  So better
     take some care here and also take an additional reference.  */

  inspector->AddRef ();
  obj = install_GpgolInspectorEvents_sink ((LPDISPATCH)inspector);
  if (obj)
    {
      register_inspector ((LPGPGOLINSPECTOREVENTS)obj, inspector);
      obj->Release ();
    }
  inspector->Release ();
  return S_OK;
}


/* The method is called by an inspector before closing.  */
STDMETHODIMP_(void)
GpgolInspectorEvents::Close (void)
{
  log_debug ("%s:%s: Called (this=%p)", SRCNAME, __func__, this );

  /* Deregister the inspector.  */
  deregister_inspector (this);
  /* We don't release ourself because we already dropped the initial
     reference after doing a register_inspector.  */
}


/* The method called by an inspector on activation.  */
STDMETHODIMP_(void)
GpgolInspectorEvents::Activate (void)
{
  LPOOMINSPECTOR inspector;
  LPDISPATCH obj;

  log_debug ("%s:%s: Called (this=%p, inspector=%p)", 
             SRCNAME, __func__, this, m_object);
  
  /* Note: It is easier to use the registered inspector object than to
     find the inspector object in the ALL_INSPECTORS list.  The
     ALL_INSPECTORS list primarly useful to keep track of additional
     information, not directly related to the event sink. */
  if (!m_object)
    {
      log_error ("%s:%s: Object not set", SRCNAME, __func__);
      return;
    }
  inspector = (LPOOMINSPECTOR)m_object;
  inspector->AddRef ();

  /* If this is the first activate for the inspector, we add the
     controls.  We do it only now to be sure that everything has been
     initialized.  Doing that in GpgolInspectorsEvents::NewInspector
     is not suggested due to claims from some mailing lists.  */ 
  if (!m_first_activate_seen)
    {
      m_first_activate_seen = true;
      add_inspector_controls (inspector);
      obj = get_oom_object (inspector, "get_CurrentItem");
      if (obj)
        {
          // LPDISPATCH obj2 = install_GpgolItemEvents_sink (obj);
          // if (obj2)
          //   obj2->Release ();
          obj->Release ();
        }
    }
  
  update_crypto_info (inspector);
  inspector->Release ();
}


/* The method called by an inspector on dectivation.  */
STDMETHODIMP_(void)
GpgolInspectorEvents::Deactivate (void)
{
  log_debug ("%s:%s: Called (this=%p)", SRCNAME, __func__, this);

}


/* Check whether we are in composer or read mode.  */
static bool
is_inspector_in_composer_mode (LPDISPATCH inspector)
{
  LPDISPATCH obj;
  bool in_composer;
  
  obj = get_oom_object (inspector, "get_CurrentItem");
  if (obj)
    {
      /* We are in composer mode if the "Sent" property is false and
         the class is 43.  */
      in_composer = (!get_oom_bool (obj, "Sent") 
                     && get_oom_int (obj, "Class") == 43);
      obj->Release ();
    }
  else
    in_composer = false;
  return in_composer;
}


/* Get the flags from the inspector; i.e. whether to sign or encrypt a
   message.  Returns 0 on success.  */
int
get_inspector_composer_flags (LPDISPATCH inspector,
                              bool *r_sign, bool *r_encrypt)
{
  LPDISPATCH button;
  int rc = 0;

  button = get_button (inspector, "GpgOL_Inspector_Sign");
  if (!button)
    {
      log_error ("%s:%s: Sign button not found", SRCNAME, __func__);
      rc = -1;
    }
  else
    {
      *r_sign = get_oom_int (button, "State") == msoButtonDown;
      button->Release ();
    }

  button = get_button (inspector, "GpgOL_Inspector_Encrypt");
  if (!button)
    {
      log_error ("%s:%s: Encrypt button not found", SRCNAME, __func__);
      rc = -1;
    }
  else
    {
      *r_encrypt = get_oom_int (button, "State") == msoButtonDown;
      button->Release ();
    }
  
  if (!rc)
    log_debug ("%s:%s: sign=%d encrypt=%d",
               SRCNAME, __func__, *r_sign, *r_encrypt);
  return rc;
}


static int
set_one_button (LPDISPATCH inspector, const char *tag, bool down)
{
  LPDISPATCH button;
  int rc = 0;

  button = get_button (inspector, tag);
  if (!button)
    {
      log_error ("%s:%s: `%s' not found", SRCNAME, __func__, tag);
      rc = -1;
    }
  else
    {
      if (put_oom_int (button, "State", down? msoButtonDown : msoButtonUp))
        rc = -1;
      button->Release ();
    }
  return rc;
}


/* Set the flags for the inspector; i.e. whether to sign or encrypt a
   message.  Returns 0 on success.  */
int
set_inspector_composer_flags (LPDISPATCH inspector, bool sign, bool encrypt)
{
  int rc = 0;

  if (set_one_button (inspector, "GpgOL_Inspector_Sign", sign))
    rc = -1;
  if (set_one_button (inspector, "GpgOL_Inspector_Sign@t", sign))
    rc = -1;
  if (set_one_button (inspector, "GpgOL_Inspector_Encrypt", encrypt))
    rc = -1;
  if (set_one_button (inspector, "GpgOL_Inspector_Encrypt@t", encrypt))
    rc = -1;

  return rc;
}


/* Helper to make the tag unique.  */
static const char *
add_tag (LPDISPATCH control, const char *value)
{
  int instid;
  
  char buf[256];
  
  instid = get_oom_int (control, "InstanceId");
  snprintf (buf, sizeof buf, "%s#%d", value, instid);
  put_oom_string (control, "Tag", buf);
  return value;
}


/* Add all the controls.  */
static void
add_inspector_controls (LPOOMINSPECTOR inspector)
{
  static
  LPDISPATCH obj, controls, button;
  inspector_info_t inspinfo;
  button_list_t buttonlist = NULL;
  const char *tag;
  int in_composer;

  log_debug ("%s:%s: Enter", SRCNAME, __func__);

  /* Check whether we are in composer or read mode.  */
  in_composer = is_inspector_in_composer_mode (inspector);

  /* Add buttons to the Format menu but only in composer mode.  */
  if (in_composer)
    {
      controls = get_oom_object 
        (inspector, "CommandBars.FindControl(,30006).get_Controls");
      if (!controls)
        log_debug ("%s:%s: Menu Popup Format not found\n", SRCNAME, __func__);
      else
        {
          button = opt.disable_gpgol? NULL : add_oom_button (controls);
          if (button)
            {
              tag = add_tag (button, "GpgOL_Inspector_Encrypt");
              put_oom_bool (button, "BeginGroup", true);
              put_oom_int (button, "Style", msoButtonIconAndCaption );
              put_oom_string (button, "Caption",
                              _("&encrypt message with GnuPG"));
              put_oom_icon (button, IDB_ENCRYPT_16, 16);
              put_oom_int (button, "State",
                           opt.encrypt_default? msoButtonDown: msoButtonUp);
              
              obj = install_GpgolCommandBarButtonEvents_sink (button);
              move_to_button_list (&buttonlist, obj, button, tag);
            }
          
          button = opt.disable_gpgol? NULL : add_oom_button (controls);
          if (button)
            {
              tag = add_tag (button, "GpgOL_Inspector_Sign");
              put_oom_int (button, "Style", msoButtonIconAndCaption );
              put_oom_string (button, "Caption", _("&sign message with GnuPG"));
              put_oom_icon (button, IDB_SIGN_16, 16);
              put_oom_int (button, "State",
                           opt.sign_default? msoButtonDown: msoButtonUp);
              
              obj = install_GpgolCommandBarButtonEvents_sink (button);
              move_to_button_list (&buttonlist, obj, button, tag);
            }
          
          controls->Release ();
        }
    }
  

  /* Add buttons to the Extra menu.  */
  controls = get_oom_object (inspector,
                             "CommandBars.FindControl(,30007).get_Controls");
  if (!controls)
    log_debug ("%s:%s: Menu Popup Extras not found\n", SRCNAME, __func__);
  else
    {
      button = in_composer? NULL : add_oom_button (controls);
      if (button)
        {
          tag = add_tag (button, "GpgOL_Inspector_Verify");
          put_oom_int (button, "Style", msoButtonIconAndCaption );
          put_oom_string (button, "Caption", _("GpgOL Decrypt/Verify"));
          put_oom_icon (button, IDB_DECRYPT_VERIFY_16, 16);
          
          obj = install_GpgolCommandBarButtonEvents_sink (button);
          move_to_button_list (&buttonlist, obj, button, tag);
        }

      button = opt.enable_debug? add_oom_button (controls) : NULL;
      if (button)
        {
          tag = add_tag (button, "GpgOL_Inspector_Debug-0");
          put_oom_int (button, "Style", msoButtonCaption );
          put_oom_string (button, "Caption",
                          "GpgOL Debug-0 (display crypto info)");
          
          obj = install_GpgolCommandBarButtonEvents_sink (button);
          move_to_button_list (&buttonlist, obj, button, tag);
        }

      button = opt.enable_debug? add_oom_button (controls) : NULL;
      if (button)
        {
          tag = add_tag (button, "GpgOL_Inspector_Debug-1");
          put_oom_int (button, "Style", msoButtonCaption );
          put_oom_string (button, "Caption",
                          "GpgOL Debug-1 (not used)");
          
          obj = install_GpgolCommandBarButtonEvents_sink (button);
          move_to_button_list (&buttonlist, obj, button, tag);
        }

      button = opt.enable_debug? add_oom_button (controls) : NULL;
      if (button)
        {
          tag = add_tag (button, "GpgOL_Inspector_Debug-2");
          put_oom_int (button, "Style", msoButtonCaption );
          put_oom_string (button, "Caption",
                          "GpgOL Debug-2 (change message class)");
          
          obj = install_GpgolCommandBarButtonEvents_sink (button);
          move_to_button_list (&buttonlist, obj, button, tag);
        }

      controls->Release ();
    }


  /* Create the toolbar buttons.  */
  controls = get_oom_object (inspector,
                             "CommandBars.Item(Standard).get_Controls");
  if (!controls)
    log_error ("%s:%s: CommandBar \"Standard\" not found\n",
               SRCNAME, __func__);
  else
    {
      button = (opt.disable_gpgol || !in_composer
                ? NULL : add_oom_button (controls));
      if (button)
        {
          tag = add_tag (button, "GpgOL_Inspector_Encrypt@t");
          put_oom_int (button, "Style", msoButtonIcon );
          put_oom_string (button, "Caption", _("Encrypt message with GnuPG"));
          put_oom_icon (button, IDB_ENCRYPT_16, 16);
          put_oom_int (button, "State", msoButtonMixed );
          put_oom_int (button, "State",
                       opt.encrypt_default? msoButtonDown: msoButtonUp);
          
          obj = install_GpgolCommandBarButtonEvents_sink (button);
          move_to_button_list (&buttonlist, obj, button, tag);
        }

      button = (opt.disable_gpgol || !in_composer
                ? NULL : add_oom_button (controls));
      if (button)
        {
          tag = add_tag (button, "GpgOL_Inspector_Sign@t");
          put_oom_int (button, "Style", msoButtonIcon);
          put_oom_string (button, "Caption", _("Sign message with GnuPG"));
          put_oom_icon (button, IDB_SIGN_16, 16);
          put_oom_int (button, "State", msoButtonDown);
          put_oom_int (button, "State",
                       opt.sign_default? msoButtonDown: msoButtonUp);
          
          obj = install_GpgolCommandBarButtonEvents_sink (button);
          move_to_button_list (&buttonlist, obj, button, tag);
        }

      button = in_composer? NULL : add_oom_button (controls);
      if (button)
        {
          tag = add_tag (button, "GpgOL_Inspector_Crypto_Info");
          put_oom_int (button, "Style", msoButtonIcon);
          
          obj = install_GpgolCommandBarButtonEvents_sink (button);
          move_to_button_list (&buttonlist, obj, button, tag);
        }

      controls->Release ();
    }


  /* Save the buttonlist.  */
  inspinfo = get_inspector_info (inspector);
  if (inspinfo)
    {
      inspinfo->buttons = buttonlist;
      unlock_all_inspectors ();
    }
  else
    {
      button_list_t ol, ol2;

      log_error ("%s:%s: inspector not registered", SRCNAME, __func__);
      for (ol = buttonlist; ol; ol = ol2)
        {
          ol2 = ol->next;
          if (ol->sink)
            ol->sink->Release ();
          if (ol->button)
            ol->button->Release ();
          xfree (ol);
        }
    }
    
  log_debug ("%s:%s: Leave", SRCNAME, __func__);
}


/* Update the crypto info icon.  */
static void
update_crypto_info (LPDISPATCH inspector)
{
  HRESULT hr;
  LPDISPATCH button;
  const char *tooltip = "";
  int iconrc = -1;

  button = get_button (inspector, "GpgOL_Inspector_Crypto_Info");
  if (!button)
    {
      log_error ("%s:%s: Crypto Info button not found", SRCNAME, __func__);
      return;
    }

  if (!is_inspector_in_composer_mode (inspector))
    {
      LPDISPATCH obj;
      LPUNKNOWN unknown;
      LPMESSAGE message = NULL;

      obj = get_oom_object (inspector, "get_CurrentItem");
      if (obj)
        {
          unknown = get_oom_iunknown (obj, "MAPIOBJECT");
          if (!unknown)
            log_error ("%s:%s: error getting MAPI object", SRCNAME, __func__);
          else
            {
              hr = unknown->QueryInterface (IID_IMessage, (void**)&message);
              if (hr != S_OK || !message)
                {
                  message = NULL;
                  log_error ("%s:%s: error getting IMESSAGE: hr=%#lx",
                             SRCNAME, __func__, hr);
                }
              unknown->Release ();
            }
          obj->Release ();
        }
      if (message)
        {
          int is_encrypted = 0;
          int is_signed = 0;
          
          switch (mapi_get_message_type (message))
            {
            case MSGTYPE_GPGOL_MULTIPART_ENCRYPTED:
            case MSGTYPE_GPGOL_OPAQUE_ENCRYPTED:
            case MSGTYPE_GPGOL_PGP_MESSAGE:
              is_encrypted = 1;
              if ( mapi_test_sig_status (message) )
                is_signed = 1;
              break;
            case MSGTYPE_GPGOL:
            case MSGTYPE_SMIME:
            case MSGTYPE_UNKNOWN:
              break;
            default:
              is_signed = 1;
              break;
            }
          
          if (is_signed && is_encrypted)
            {
              tooltip =  _("This is a signed and encrypted message.\n"
                           "Click for more information. ");
              iconrc = IDB_DECRYPT_VERIFY_16;
            }
          else if (is_signed)
            {
              tooltip =  _("This is a signed message.\n"
                           "Click for more information. ");
              iconrc = IDB_VERIFY_16;
            }
          else if (is_encrypted)
            {
              tooltip =  _("This is an encrypted message.\n"
                           "Click for more information. ");
              iconrc = IDB_DECRYPT_16;
            }
          
          message->Release ();
        }
    }

  put_oom_string (button, "TooltipText", tooltip);
  if (iconrc != -1)
    put_oom_icon (button, iconrc, 16);
  put_oom_bool (button, "Visible", (iconrc != -1));
  button->Release ();
}


/* Return the MAPI message object of then inspector from a button's
   instance id. */
static LPMESSAGE
get_message_from_button (unsigned long instid, LPDISPATCH *r_inspector)
{
  HRESULT hr;
  LPDISPATCH inspector, obj;
  LPUNKNOWN  unknown;
  LPMESSAGE message = NULL;
  
  if (r_inspector)
    *r_inspector = NULL;
  inspector = get_inspector_from_instid (instid);
  if (inspector)
    {
      obj = get_oom_object (inspector, "get_CurrentItem");
      if (!obj)
        log_error ("%s:%s: error getting CurrentItem", SRCNAME, __func__);
      else
        {
          unknown = get_oom_iunknown (obj, "MAPIOBJECT");
          if (!unknown)
            log_error ("%s:%s: error getting MAPI object", SRCNAME, __func__);
          else
            {
              hr = unknown->QueryInterface (IID_IMessage, (void**)&message);
              if (hr != S_OK || !message)
                {
                  message = NULL;
                  log_error ("%s:%s: error getting IMESSAGE: hr=%#lx",
                             SRCNAME, __func__, hr);
                }
              unknown->Release ();
            }
          obj->Release ();
        }
      if (r_inspector)
        *r_inspector = inspector;
      else
        inspector->Release ();
    }
  return message;
}


/* Toggle a button and return the new state.  */
static void
toggle_button (LPDISPATCH button, const char *tag, int instid)
{
  int state;
  char tag2[256];
  char *p;
  LPDISPATCH inspector;
  
  inspector = get_inspector_from_instid (instid);
  if (!inspector)
    {
      log_debug ("%s:%s: inspector not found", SRCNAME, __func__);
      return;
    }

  state = get_oom_int (button, "State");
  log_debug ("%s:%s: button `%s' state is %d", SRCNAME, __func__, tag, state);
  state = (state == msoButtonUp)? msoButtonDown : msoButtonUp;
  put_oom_int (button, "State", state);

  /* Toggle the other button.  */
  mem2str (tag2, tag, sizeof tag2 - 2);
  p = strchr (tag2, '#');
  if (p)
    *p = 0;  /* Strip the instance id suffix.  */
  if (*tag2 && tag2[1] && !strcmp (tag2+strlen(tag2)-2, "@t"))
    tag2[strlen(tag2)-2] = 0; /* Remove the "@t".  */
  else
    strcat (tag2, "@t");      /* Append a "@t".  */
  
  log_debug ("%s:%s: setting `%s' state to %d", SRCNAME, __func__, tag2, state);
  set_one_button (inspector, tag2, state);
  inspector->Release ();
}


/* Called for a click on an inspector button.  BUTTON is the button
   object and TAG is the tag value (which is guaranteed not to be
   NULL).  INSTID is the instance ID of the button. */
void
proc_inspector_button_click (LPDISPATCH button, const char *tag, int instid)
{
  LPMESSAGE message;
  HWND hwnd = NULL; /* Fixme  */

  if (!tagcmp (tag, "GpgOL_Inspector_Encrypt"))
    {  
      toggle_button (button, tag, instid);
    }
  else if (!tagcmp (tag, "GpgOL_Inspector_Sign"))
    {  
      toggle_button (button, tag, instid);
    }
  else if (!tagcmp (tag, "GpgOL_Inspector_Verify")
           || !tagcmp (tag, "GpgOL_Inspector_Crypto_Info"))
    {
      LPDISPATCH inspector;

      message = get_message_from_button (instid, &inspector);
      if (message)
        {
          if (message_incoming_handler (message, hwnd, true))
            message_display_handler (message, inspector, hwnd);
          message->Release ();
        }
      if (inspector)
        {
          update_crypto_info (inspector);
          inspector->Release ();
        }
    }
  else if (!tagcmp (tag, "GpgOL_Inspector_Debug-0"))
    {
      log_debug ("%s:%s: command Debug0 (showInfo) called\n",
                 SRCNAME, __func__);
      message = get_message_from_button (instid, NULL);
      if (message)
        {
          message_show_info (message, hwnd);
          message->Release ();
        }
    }
  else if (!tagcmp (tag, "GpgOL_Inspector_Debug-1"))
    {
      log_debug ("%s:%s: command Debug1 (not used) called\n",
                 SRCNAME, __func__);
    }
  else if (!tagcmp (tag, "GpgOL_Inspector_Debug-2"))
    {
      log_debug ("%s:%s: command Debug2 (change message class) called", 
                 SRCNAME, __func__);
      message = get_message_from_button (instid, NULL);
      if (message)
        {
          /* We sync here. */
          mapi_change_message_class (message, 1);
          message->Release ();
        }
    }
  else if (!tagcmp (tag, "GpgOL_Inspector_Debug-3"))
    {
      log_debug ("%s:%s: command Debug3 (revert_message_class) called", 
                 SRCNAME, __func__);
      message = get_message_from_button (instid, NULL);
      if (message)
        {
          int rc = gpgol_message_revert (message, 1, 
                                         KEEP_OPEN_READWRITE|FORCE_SAVE);
          log_debug ("%s:%s: gpgol_message_revert returns %d\n", 
                     SRCNAME, __func__, rc);
          message->Release ();
        }
    }

}
