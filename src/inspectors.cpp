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
BEGIN_EVENT_SINK(GpgolInspectorEvents, IOOMInspectorEvents)
  STDMETHOD_ (void, Activate) (THIS_);
  STDMETHOD_ (void, Close) (THIS_);
  STDMETHOD_ (void, Deactivate) (THIS_);
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



/* To avoid messing around with the OOM (adding extra objects as user
   properties and such), we keep our own information structure about
   inspectors. */
struct inspector_info_s
{
  /* We are pretty lame and keep all inspectors in a linked list.  */
  struct inspector_info_s *next;

  /* The inspector object.  */
  LPOOMINSPECTOR inspector;

  /* The event sink object.  This is used by the event methods to
     locate the inspector object.  */
  LPOOMINSPECTOREVENTS eventsink;

  /* The event sink object for the crypto info button.  */
  LPDISPATCH crypto_info_sink;

};
typedef struct inspector_info_s *inspector_info_t;

/* The list of all inspectors and a lock for it.  */
static inspector_info_t all_inspectors;
static HANDLE all_inspectors_lock;

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


/* Register the inspector object INSPECTOR along with its event SINK.  */
static void
register_inspector (LPOOMINSPECTOR inspector, LPOOMINSPECTOREVENTS sink)
{
  inspector_info_t item;

  item = (inspector_info_t)xcalloc (1, sizeof *item);
  lock_all_inspectors ();
  inspector->AddRef ();
  item->inspector = inspector;

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
  if (r->crypto_info_sink)
    r->crypto_info_sink->Release ();
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
      register_inspector (inspector, (LPOOMINSPECTOREVENTS)obj);
      obj->Release ();
      add_inspector_controls (inspector);
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

  /* Update the crypt info.  */
  obj = get_oom_object (m_object, "CommandBars");
  if (!obj)
    log_error ("%s:%s: CommandBars not found", SRCNAME, __func__);
  else
    {
      button = get_oom_control_bytag (obj, "GpgOL_Inspector_Crypto_Info");
      obj->Release ();
      if (button)
        {
          update_inspector_crypto_info (button);
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





void
add_inspector_controls (LPOOMINSPECTOR inspector)
{
  LPDISPATCH pObj, pDisp, pTmp;
  inspector_info_t inspinfo;

  log_debug ("%s:%s: Enter", SRCNAME, __func__);

  /* In theory we should take a lock here to avoid a race between the
     test for a new control and the creation.  However, we are not
     called from a second thread.  FIXME: We might want to use
     inspector_info insteasd to check this.  However this requires us
     to keep a lock per inspector. */

  /* Check that our controls do not already exist.  */
  pObj = get_oom_object (inspector, "CommandBars");
  if (!pObj)
    {
      log_debug ("%s:%s: CommandBars not found", SRCNAME, __func__);
      return;
    }
  pDisp = get_oom_control_bytag (pObj, "GpgOL_Inspector_Crypto_Info");
  pObj->Release ();
  pObj = NULL;
  if (pDisp)
    {
      pDisp->Release ();
      log_debug ("%s:%s: Leave (Controls are already added)",
                 SRCNAME, __func__);
      return;
    }

  /* Create our controls.  */
  pDisp = get_oom_object (inspector,
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
          put_oom_string (pDisp, "Tag", "GpgOL_Inspector_Crypto_Info");
          put_oom_int (pDisp, "Style", msoButtonIcon );
          put_oom_string (pDisp, "TooltipText",
                          _("Indicates the crypto status of the message"));
          put_oom_icon (pDisp, IDB_SIGN, 16);
          
          pObj = install_GpgolCommandBarButtonEvents_sink (pDisp);
          pDisp->Release ();
          inspinfo = get_inspector_info (inspector);
          if (inspinfo)
            {
              inspinfo->crypto_info_sink = pObj;
              unlock_all_inspectors ();
            }
          else
            {
              log_error ("%s:%s: inspector not registered", SRCNAME, __func__);
              pObj->Release (); /* Get rid of the sink.  */
            }
        }
    }

  /* FIXME: Menuswe need to add: */
        // (opt.disable_gpgol || not_a_gpgol_message)?
        //         "" : _("GpgOL Decrypt/Verify"), &m_nCmdCryptoState,
        // opt.enable_debug? "GpgOL Debug-0 (display crypto info)":"", 
        //         &m_nCmdDebug0,


    
  log_debug ("%s:%s: Leave", SRCNAME, __func__);
}


/* Update the crypto info icon.  */
void
update_inspector_crypto_info (LPDISPATCH button)
{
  LPDISPATCH obj;
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
  obj = get_oom_object (button, "get_Parent.get_Parent.get_CurrentItem");
  if (obj)
    {
      msgcls = get_oom_string (obj, "MessageClass");
      /* If "Sent" is false we are in composer mode  */
      in_composer = !get_oom_bool (obj, "Sent");
      obj->Release ();
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

