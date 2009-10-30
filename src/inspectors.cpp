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
#include "inspectors.h"
#include "dialogs.h"  /* IDB_xxx */
#include "cmdbarcontrols.h"

#include "eventsink.h"


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
  unsigned long m_serialno;
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
  unsigned long serialno; /* of the inspector.  */
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

  /* The inspector object.  */
  LPOOMINSPECTOR inspector;

  /* Our serial number for the inspector.  */
  unsigned long serialno;

  /* The event sink object.  This is used by the event methods to
     locate the inspector object.  */
  LPOOMINSPECTOREVENTS eventsink;

  /* A list of all the buttons.  */
  button_list_t buttons;

};
typedef struct inspector_info_s *inspector_info_t;

/* The list of all inspectors and a lock for it.  */
static inspector_info_t all_inspectors;
static HANDLE all_inspectors_lock;


static void add_inspector_controls (LPOOMINSPECTOR inspector, 
                                    unsigned long serialno);
static void update_crypto_info (LPDISPATCH button);



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


/* Return a new serial number for an inspector.  These serial numbers
   are used to make the button tags unique.  */
static unsigned long
create_inspector_serial (void)
{
  static long serial;
  long n;

  /* Avoid returning 0 because we use that value as Nil.  */
  while (!(n = InterlockedIncrement (&serial)))
    ;
  return (unsigned long)n;
}


/* Add SINK and BUTTON to the list at LISTADDR.  The list takes
   ownership of SINK and BUTTON, thus the caller may not use OBJ or
   OBJ2 after this call.  If TAG must be given without the
   serialnumber suffix.  SERIALNO is the serialno of the correspnding
   inspector.  */
static void
move_to_button_list (button_list_t *listaddr, 
                     LPDISPATCH sink, LPDISPATCH button, 
                     const char *tag, unsigned long serialno)
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
  item->serialno = serialno;
  strcpy (item->tag, tag);
  item->next = *listaddr;
  *listaddr = item;
}


/* Register the inspector object INSPECTOR along with its event SINK.  */
static void
register_inspector (LPOOMINSPECTOR inspector, LPGPGOLINSPECTOREVENTS sink)
{
  inspector_info_t item;

  item = (inspector_info_t)xcalloc (1, sizeof *item);
  lock_all_inspectors ();
  inspector->AddRef ();
  item->inspector = inspector;
  item->serialno = sink->m_serialno = create_inspector_serial ();

  sink->AddRef ();
  item->eventsink = sink;

  item->next = all_inspectors;
  all_inspectors = item;

  unlock_all_inspectors ();
}


/* Deregister the inspector with the event SINK.  */
static void
deregister_inspector (LPOOMINSPECTOREVENTS sink)
{
  inspector_info_t r, rprev;
  button_list_t ol, ol2;

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
  r->eventsink->Release ();
  r->eventsink = NULL;
  if (r->inspector)
    r->inspector->Release ();

  for (ol = r->buttons; ol; ol = ol2)
    {
      ol2 = ol->next;
      if (ol->sink)
        ol->sink->Release ();
      if (ol->button)
        ol->button->Release ();
      xfree (ol);
    }

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


/* Return the serialno of INSPECTOR or 0 if not found.  */
static unsigned long
get_serialno (LPDISPATCH inspector)
{
  unsigned int result = 0;
  inspector_info_t iinfo;

  /* FIXME: This might not bet reliable.  We merely compare the
     pointer and not something like an Instance Id.  We should check
     whether this is sufficient or whether to track the inspectors
     with different hack.  For example we could add an invisible menu
     entry and scan for that entry to get the serial number serial
     number of it.  A better option would be to add a custom property
     to the inspector, but that seems not supported - we could of
     course add it to a button then. */ 
  lock_all_inspectors ();

  for (iinfo = all_inspectors; iinfo; iinfo = iinfo->next)
    if (iinfo->inspector == inspector)
      {
        result = iinfo->serialno;
        break;
      }

  unlock_all_inspectors ();
  return result;
}


/* Return the button with TAG and assigned to the isnpector with
   SERIALNO.  Return NULL if not found.  */
static LPDISPATCH
get_button (unsigned long serialno, const char *tag)
{
  LPDISPATCH result = NULL;
  inspector_info_t iinfo;
  button_list_t ol;

  lock_all_inspectors ();

  for (iinfo = all_inspectors; iinfo; iinfo = iinfo->next)
    if (iinfo->serialno == serialno)
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
   button with the instanceId INSTID.  The find the button with TAG in
   that inspector and return it.  Caller must release the returned
   button.  Returns NULL if not found.  */
static LPDISPATCH
get_button_by_instid_and_tag (int instid, const char *tag)
{
  LPDISPATCH result = NULL;
  inspector_info_t iinfo;
  button_list_t ol;

  // log_debug ("%s:%s: inst=%d tag=(%s)",SRCNAME, __func__, instid, tag);
  
  lock_all_inspectors ();

  for (iinfo = all_inspectors; iinfo; iinfo = iinfo->next)
    for (ol = iinfo->buttons; ol; ol = ol->next)
      if (ol->instid == instid)
        {
          /* Found the inspector.  Now look for the tag.  */
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




/* The method called by outlook for each new inspector.  Note that
   Outlook sometimes reuses Inspectro objects thus this event is not
   an indication for a newly opened Inspector.  */
STDMETHODIMP
GpgolInspectorsEvents::NewInspector (LPOOMINSPECTOR inspector)
{
  LPDISPATCH obj;
  log_debug ("%s:%s: Called", SRCNAME, __func__);

  /* It has been said that INSPECTOR here a "weak" reference. This may
     mean that the object has not been fully initialized.  So better
     take some care here and also take an additional reference.  */
  inspector->AddRef ();
  obj = install_GpgolInspectorEvents_sink ((LPDISPATCH)inspector);
  if (obj)
    {
      register_inspector (inspector, (LPGPGOLINSPECTOREVENTS)obj);
      obj->Release ();
    }
  inspector->Release ();
  return S_OK;
}


/* The method is called by an inspector before closing.  */
STDMETHODIMP_(void)
GpgolInspectorEvents::Close (void)
{
  log_debug ("%s:%s: Called", SRCNAME, __func__);
  /* Deregister the inspector and free outself.  */
  deregister_inspector (this);
  this->Release ();
}


/* The method called by an inspector on activation.  */
STDMETHODIMP_(void)
GpgolInspectorEvents::Activate (void)
{
  LPDISPATCH obj, button;
  LPOOMINSPECTOR inspector;

  log_debug ("%s:%s: Called", SRCNAME, __func__);

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

  /* If this is the first activate for the inspector, we add the
     controls.  We do it only now to be sure that everything has been
     initialized.  Doing that in GpgolInspectorsEvents::NewInspector
     is not suggested due to claims from some mailing lists.  */ 
  if (!m_first_activate_seen)
    {
      m_first_activate_seen = true;
      add_inspector_controls (inspector, m_serialno);
    }
  
  /* Update the crypt info.  */
  obj = get_oom_object (inspector, "CommandBars");
  if (!obj)
    log_error ("%s:%s: CommandBars not found", SRCNAME, __func__);
  else
    {
      button = get_button (m_serialno, "GpgOL_Inspector_Crypto_Info");
      obj->Release ();
      if (button)
        {
          update_crypto_info (button);
          button->Release ();
        }
    }
}


/* The method called by an inspector on dectivation.  */
STDMETHODIMP_(void)
GpgolInspectorEvents::Deactivate (void)
{
  log_debug ("%s:%s: Called", SRCNAME, __func__);

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
  unsigned long serialno;

  serialno = get_serialno (inspector);
  if (!serialno)
    {
      log_error ("%s:%s: S/n not found", SRCNAME, __func__);
      return -1;
    }

  button = get_button (serialno, "GpgOL_Inspector_Sign");
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

  button = get_button (serialno, "GpgOL_Inspector_Encrypt");
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


/* Set the flags for the inspector; i.e. whether to sign or encrypt a
   message.  Returns 0 on success.  */
int
set_inspector_composer_flags (LPDISPATCH inspector, bool sign, bool encrypt)
{
  LPDISPATCH button;
  int rc = 0;
  unsigned long serialno;

  serialno = get_serialno (inspector);
  if (!serialno)
    {
      log_error ("%s:%s: S/n not found", SRCNAME, __func__);
      return -1;
    }

  button = get_button (serialno, "GpgOL_Inspector_Sign");
  if (!button)
    {
      log_error ("%s:%s: Sign button not found", SRCNAME, __func__);
      rc = -1;
    }
  else
    {
      if (put_oom_int (button, "State", sign? msoButtonDown : msoButtonUp))
        rc = -1;
      button->Release ();
    }

  button = get_button (serialno, "GpgOL_Inspector_Encrypt");
  if (!button)
    {
      log_error ("%s:%s: Encrypt button not found", SRCNAME, __func__);
      rc = -1;
    }
  else
    {
      if (put_oom_int (button, "State", encrypt? msoButtonDown : msoButtonUp))
        rc = -1;
      button->Release ();
    }
  
  return rc;
}


/* Helper to make the tag unique.  */
static const char *
add_tag (LPDISPATCH control, unsigned long serialno, const char *value)
{
  char buf[256];
  
  snprintf (buf, sizeof buf, "%s#%lu", value, serialno);
  put_oom_string (control, "Tag", buf);
  return value;
}


/* Add all the controls.  */
static void
add_inspector_controls (LPOOMINSPECTOR inspector, unsigned long serialno)
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
              tag = add_tag (button, serialno, "GpgOL_Inspector_Encrypt");
              put_oom_bool (button, "BeginGroup", true);
              put_oom_int (button, "Style", msoButtonIconAndCaption );
              put_oom_string (button, "Caption",
                              _("&encrypt message with GnuPG"));
              put_oom_icon (button, IDB_ENCRYPT, 16);
              
              obj = install_GpgolCommandBarButtonEvents_sink (button);
              move_to_button_list (&buttonlist, obj, button, tag, serialno);
            }
          
          button = opt.disable_gpgol? NULL : add_oom_button (controls);
          if (button)
            {
              tag = add_tag (button, serialno, "GpgOL_Inspector_Sign");
              put_oom_int (button, "Style", msoButtonIconAndCaption );
              put_oom_string (button, "Caption", _("&sign message with GnuPG"));
              put_oom_icon (button, IDB_SIGN, 16);
              
              obj = install_GpgolCommandBarButtonEvents_sink (button);
              move_to_button_list (&buttonlist, obj, button, tag, serialno);
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
          tag = add_tag (button, serialno, "GpgOL_Inspector_Verify");
          put_oom_int (button, "Style", msoButtonIconAndCaption );
          put_oom_string (button, "Caption", _("GpgOL Decrypt/Verify"));
          put_oom_icon (button, IDB_DECRYPT_VERIFY, 16);
          
          obj = install_GpgolCommandBarButtonEvents_sink (button);
          move_to_button_list (&buttonlist, obj, button, tag, serialno);
        }

      button = opt.enable_debug? add_oom_button (controls) : NULL;
      if (button)
        {
          tag = add_tag (button, serialno, "GpgOL_Inspector_Debug-0");
          put_oom_int (button, "Style", msoButtonCaption );
          put_oom_string (button, "Caption",
                          "GpgOL Debug-0 (display crypto info)");
          
          obj = install_GpgolCommandBarButtonEvents_sink (button);
          move_to_button_list (&buttonlist, obj, button, tag, serialno);
        }

      button = opt.enable_debug? add_oom_button (controls) : NULL;
      if (button)
        {
          tag = add_tag (button, serialno, "GpgOL_Inspector_Debug-1");
          put_oom_int (button, "Style", msoButtonCaption );
          put_oom_string (button, "Caption",
                          "GpgOL Debug-1 (open_inspector)");
          
          obj = install_GpgolCommandBarButtonEvents_sink (button);
          move_to_button_list (&buttonlist, obj, button, tag, serialno);
        }

      button = opt.enable_debug? add_oom_button (controls) : NULL;
      if (button)
        {
          tag = add_tag (button, serialno,"GpgOL_Inspector_Debug-2");
          put_oom_int (button, "Style", msoButtonCaption );
          put_oom_string (button, "Caption",
                          "GpgOL Debug-2 (change message class)");
          
          obj = install_GpgolCommandBarButtonEvents_sink (button);
          move_to_button_list (&buttonlist, obj, button, tag, serialno);
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
          tag = add_tag (button, serialno, "GpgOL_Inspector_Encrypt@t");
          put_oom_int (button, "Style", msoButtonIcon );
          put_oom_string (button, "Caption", _("Encrypt message with GnuPG"));
          put_oom_icon (button, IDB_ENCRYPT, 16);
          put_oom_int (button, "State", msoButtonMixed );
          
          obj = install_GpgolCommandBarButtonEvents_sink (button);
          move_to_button_list (&buttonlist, obj, button, tag, serialno);
        }

      button = (opt.disable_gpgol || !in_composer
                ? NULL : add_oom_button (controls));
      if (button)
        {
          tag = add_tag (button, serialno, "GpgOL_Inspector_Sign@t");
          put_oom_int (button, "Style", msoButtonIcon);
          put_oom_string (button, "Caption", _("Sign message with GnuPG"));
          put_oom_icon (button, IDB_SIGN, 16);
          put_oom_int (button, "State", msoButtonDown);
          
          obj = install_GpgolCommandBarButtonEvents_sink (button);
          move_to_button_list (&buttonlist, obj, button, tag, serialno);
        }

      button = in_composer? NULL : add_oom_button (controls);
      if (button)
        {
          tag = add_tag (button, serialno, "GpgOL_Inspector_Crypto_Info");
          put_oom_int (button, "Style", msoButtonIcon);
          
          obj = install_GpgolCommandBarButtonEvents_sink (button);
          move_to_button_list (&buttonlist, obj, button, tag, serialno);
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
update_crypto_info (LPDISPATCH button)
{
  LPDISPATCH inspector;
  char *msgcls = NULL;;
  const char *s;
  int in_composer = 0;
  int is_encrypted = 0;
  int is_signed = 0;
  const char *tooltip;
  int iconrc;

  /* FIXME: We should store the information retrieved by old
     versions via mapi_get_message_type and mapi_test_sig_status
     in UserProperties and use them instead of a direct lookup of
     the messageClass.  */

  inspector = get_oom_object (button, "get_Parent.get_Parent.get_CurrentItem");
  if (inspector)
    {
      msgcls = get_oom_string (inspector, "MessageClass");
      in_composer = is_inspector_in_composer_mode (inspector);
      inspector->Release ();
    }
  if (msgcls)
    {
      log_debug ("%s:%s: message class is `%s'", SRCNAME, __func__, msgcls);
      if (!strncmp (msgcls, "IPM.Note.GpgOL", 14) 
          && (!msgcls[14] || msgcls[14] == '.'))
        {
          s = msgcls + 14;
          if (!*s)
            ;
          else if (!strcmp (s, ".MultipartSigned"))
            is_signed = 1;
          else if (!strcmp (s, ".MultipartEncrypted"))
            is_encrypted = 1;
          else if (!strcmp (s, ".OpaqueSigned"))
            is_signed = 1;
          else if (!strcmp (s, ".OpaqueEncrypted"))
            is_encrypted = 1;
          else if (!strcmp (s, ".ClearSigned"))
            is_signed = 1;
          else if (!strcmp (s, ".PGPMessage"))
            is_encrypted = 1;
        }
      
      /*FIXME: check something like mail_test_sig_status to see
        whether it is an encrypted and signed message.  */
      
      if (in_composer)
        {
          tooltip = "";
          iconrc = -1;
        }
      else if (is_signed && is_encrypted)
        {
          tooltip =  _("This is a signed and encrypted message.\n"
                       "Click for more information. ");
          iconrc = IDB_DECRYPT_VERIFY;
        }
      else if (is_signed)
        {
          tooltip =  _("This is a signed message.\n"
                       "Click for more information. ");
          iconrc = IDB_VERIFY;
        }
      else if (is_encrypted)
        {
          tooltip =  _("This is an encrypted message.\n"
                       "Click for more information. ");
          iconrc = IDB_DECRYPT;
        }
      else
        {
          tooltip = "";
          iconrc = -1;
        }
      
      put_oom_string (button, "TooltipText", tooltip);
      if (iconrc != -1)
        put_oom_icon (button, iconrc, 16);
      put_oom_bool (button, "Visible", (iconrc != -1));
      
      xfree (msgcls);
    }
}


/* Toggle a button and return the new state.  */
static int
toggle_button (LPDISPATCH button, const char *tag, int instid)
{
  int state = get_oom_int (button, "State");
  char tag2[256];
  char *p;
  LPDISPATCH button2;

  log_debug ("%s:%s: button `%s' state is %d", SRCNAME, __func__, tag, state);
  state = (state == msoButtonUp)? msoButtonDown : msoButtonUp;
  put_oom_int (button, "State", state);

  /* Toggle the other button.  */
  mem2str (tag2, tag, sizeof tag2 - 2);
  p = strchr (tag2, '#');
  if (p)
    *p = 0;  /* Strip the serialno suffix.  */
  if (*tag2 && tag2[1] && !strcmp (tag2+strlen(tag2)-2, "@t"))
    tag2[strlen(tag2)-2] = 0; /* Remove the "@t".  */
  else
    strcat (tag2, "@t");      /* Append a "@t".  */

  button2 = get_button_by_instid_and_tag (instid, tag2);
  if (!button2)
    log_debug ("%s:%s: button `%s' not found", SRCNAME, __func__, tag2);
  else
    {
      put_oom_int (button2, "State", state);
      button2->Release ();
    }
  return state;
}


/* Called for a click on an inspector button.  BUTTON is the button
   object and TAG is the tag value (which is guaranteed not to be
   NULL).  INSTID is the instance ID of the button. */
void
proc_inspector_button_click (LPDISPATCH button, const char *tag, int instid)
{
  if (!tagcmp (tag, "GpgOL_Inspector_Encrypt"))
    {  
      toggle_button (button, tag, instid);
    }
  else if (!tagcmp (tag, "GpgOL_Inspector_Sign"))
    {  
      toggle_button (button, tag, instid);
    }
  else if (!tagcmp (tag, "GpgOL_Inspector_Verify"))
    {
      /* FIXME: We need to invoke decrypt/verify again. */
    }
  else if (!tagcmp (tag, "GpgOL_Inspector_Crypto_Info"))
    {
      /* FIXME: We should invoke the decrypt/verify again. */
      update_crypto_info (button);
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
  else if (!tagcmp (tag, "GpgOL_Inspector_Debug-0"))
    {
      /* Show crypto info.  */
      log_debug ("%s:%s: command Debug0 (showInfo) called\n",
                 SRCNAME, __func__);
      // hr = eecb->GetObject (&mdb, (LPMAPIPROP *)&message);
      // if (SUCCEEDED (hr))
      //   {
      //     message_show_info (message, hwnd);
      //   }
      // ul_release (message, __func__, __LINE__);
      // ul_release (mdb, __func__, __LINE__);
    }
  else if (!tagcmp (tag, "GpgOL_Inspector_Debug-1"))
    {
      /* Open inspector.  */
      log_debug ("%s:%s: command Debug1 (open inspector) called\n",
                 SRCNAME, __func__);
      // hr = eecb->GetObject (&mdb, (LPMAPIPROP *)&message);
      // if (SUCCEEDED (hr))
      //   {
      //     open_inspector (eecb, message);
      //   }
      // ul_release (message, __func__, __LINE__);
      // ul_release (mdb, __func__, __LINE__);
    }
  else if (!tagcmp (tag, "GpgOL_Inspector_Debug-2"))
    {
      /* Change message class.  */
      log_debug ("%s:%s: command Debug2 (change message class) called", 
                 SRCNAME, __func__);
      // hr = eecb->GetObject (&mdb, (LPMAPIPROP *)&message);
      // if (SUCCEEDED (hr))
      //   {
      //     /* We sync here. */
      //     mapi_change_message_class (message, 1);
      //   }
      // ul_release (message, __func__, __LINE__);
      // ul_release (mdb, __func__, __LINE__);
    }
  else if (!tagcmp (tag, "GpgOL_Inspector_Debug-3"))
    {
      log_debug ("%s:%s: command Debug3 (revert_message_class) called", 
                 SRCNAME, __func__);
      // hr = eecb->GetObject (&mdb, (LPMAPIPROP *)&message);
      // if (SUCCEEDED (hr))
      //   {
      //     int rc = gpgol_message_revert (message, 1, 
      //                                    KEEP_OPEN_READWRITE|FORCE_SAVE);
      //     log_debug ("%s:%s: gpgol_message_revert returns %d\n", 
      //                SRCNAME, __func__, rc);
      //   }
      // ul_release (message, __func__, __LINE__);
      // ul_release (mdb, __func__, __LINE__);
    }



}
