/* ol-ext-callback.cpp - Code to use the IOutlookExtCallback.
 *	Copyright (C) 2004, 2005, 2007 g10 Code GmbH
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

#include "mymapi.h"
#include "mymapitags.h"
#include "myexchext.h"
#include "display.h"
#include "common.h"
#include "msgcache.h"
#include "mapihelp.h"

#include "olflange-def.h"
#include "olflange.h"
#include "ol-ext-callback.h"

#define TRACEPOINT() do { log_debug ("%s:%s:%d: tracepoint\n", \
                                     SRCNAME, __func__, __LINE__); \
                        } while (0)


/* Wrapper around UlRelease with error checking. */
static void 
ul_release (LPVOID punk, const char *func, int lnr)
{
  ULONG res;
  
  if (!punk)
    return;
  res = UlRelease (punk);
  if (opt.enable_debug & DBG_MEMORY)
    log_debug ("%s:%s:%d: UlRelease(%p) had %lu references\n", 
               SRCNAME, func, lnr, punk, res);
}



/* Locate a property using the provided callback LPEECB and traverse
   down to the last element of the dot delimited NAME.  Returns the
   Dispatch object and if R_DISPID is not NULL, the dispatch-id of the
   last part.  Returns NULL on error.  The traversal implictly starts
   at the object returned by the outlook application callback. */
LPDISPATCH
find_outlook_property (LPEXCHEXTCALLBACK lpeecb,
                       const char *name, DISPID *r_dispid)
{
  HRESULT hr;
  LPOUTLOOKEXTCALLBACK pCb;
  LPUNKNOWN pObj;
  LPDISPATCH pDisp;
  DISPID dispid;
  wchar_t *wname;
  const char *s;

  log_debug ("%s:%s: looking for `%s'\n", SRCNAME, __func__, name);

  pCb = NULL;
  pObj = NULL;
  lpeecb->QueryInterface (IID_IOutlookExtCallback, (LPVOID*)&pCb);
  if (pCb)
    pCb->GetObject (&pObj);
  for (; pObj && (s = strchr (name, '.')) && s != name; name = s + 1)
    {
      VARIANT vtResult;
      char namepart[100];
      size_t n;
      char *p, *pend;
      BSTR parmstr = NULL;

      /* Our loop expects that all objects except for the last one are
         of class IDispatch.  This is pretty reasonable. */
      pObj->QueryInterface (IID_IDispatch, (LPVOID*)&pDisp);
      if (!pDisp)
        return NULL;

      n = s - name;
      if (n >= sizeof namepart)
        n = sizeof namepart - 1;
      strncpy (namepart, name, n);
      namepart[n] = 0;

      /* Parse a parameter (like "Print" in "Item(Print)").  */
      p = strchr (namepart, '(');
      if (p)
        {
          *p++ = 0;
          pend = strchr (p, ')');
          if (pend)
            *pend = 0;
          wname = utf8_to_wchar (p);
          if (wname)
            {
              log_debug ("   parm(%s)=(%s)\n", namepart, p);
              parmstr = SysAllocString (wname);
              xfree (wname);
            }
        }

      wname = utf8_to_wchar (namepart);
      if (!wname)
        {
          if (parmstr)
            SysFreeString (parmstr);
          return NULL;
        }

      hr = pDisp->GetIDsOfNames(IID_NULL, &wname, 1,
                                LOCALE_SYSTEM_DEFAULT, &dispid);
      xfree (wname);
      log_debug ("   dispid(%s)=%d  (hr=0x%x)\n",
                 namepart, (int)dispid, (unsigned int)hr);

      vtResult.pdispVal = NULL;
      if (parmstr)
        {
          DISPPARAMS dispparams;
          VARIANT aVariant[4];

          dispparams.rgvarg = aVariant;
          dispparams.rgvarg[0].vt = VT_BSTR;
          dispparams.rgvarg[0].bstrVal = parmstr;
          dispparams.cArgs = 1;
          dispparams.cNamedArgs = 0;
          hr = pDisp->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                              DISPATCH_METHOD|DISPATCH_PROPERTYGET,
                              &dispparams, &vtResult, NULL, NULL);
          SysFreeString (parmstr);
        }
      else
        {
          DISPPARAMS dispparamsNoArgs = {NULL, NULL, 0, 0};

          hr = pDisp->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                              DISPATCH_METHOD, &dispparamsNoArgs,
                              &vtResult, NULL, NULL);
        }

      pObj = vtResult.pdispVal;
      /* FIXME: Check that the class of the returned object is as
         expected.  To do this we better let GetIdsOfNames also return
         the ID of "Class". */
      log_debug ("   %s=%p vt=%d (hr=0x%x)\n",
                 namepart, pObj, vtResult.vt, (unsigned int)hr);
      pDisp->Release ();
      pDisp = NULL;
      if (vtResult.vt != VT_DISPATCH)
        return NULL;
      /* Fixme: Do we need to release pObj? */
    }
  if (!pObj || !*name)
    return NULL;

  pObj->QueryInterface (IID_IDispatch, (LPVOID*)&pDisp);
  if (!pDisp)
    return NULL;
  wname = utf8_to_wchar (name);
  if (!wname)
    {
      pDisp->Release ();
      return NULL;
    }
      
  hr = pDisp->GetIDsOfNames (IID_NULL, &wname, 1,
                             LOCALE_SYSTEM_DEFAULT, &dispid);
  xfree (wname);
  log_debug ("   dispid(%s)=%d  (hr=0x%x)\n", name, (int)dispid, (int)hr);
  if (r_dispid)
    *r_dispid = dispid;

  log_debug ("%s:%s:    got IDispatch=%p dispid=%u\n",
  	     SRCNAME, __func__, pDisp, (unsigned int)dispid);
  return pDisp;
}


int
put_outlook_property (void *pEECB, const char *key, const char *value)
{
  int result = -1;
  HRESULT hr;
  LPMDB pMDB = NULL;
  LPMESSAGE pMessage = NULL;
  LPDISPATCH pDisp;
  DISPID dispid;
  DISPID dispid_put = DISPID_PROPERTYPUT;
  DISPPARAMS dispparams;
  VARIANT aVariant;
  if (!pEECB)
    return -1;

  hr = ((LPEXCHEXTCALLBACK)pEECB)->GetObject (&pMDB, (LPMAPIPROP *)&pMessage);
  if (FAILED (hr))
    log_debug ("%s:%s: getObject failed: hr=%#lx\n", SRCNAME, __func__, hr);
  else if ( (pDisp = find_outlook_property ((LPEXCHEXTCALLBACK)pEECB,
                                            key, &dispid)))
    {
      BSTR abstr;

      dispparams.cNamedArgs = 1;
      dispparams.rgdispidNamedArgs = &dispid_put;
      dispparams.cArgs = 1;
      dispparams.rgvarg = &aVariant;
      {
        wchar_t *tmp = utf8_to_wchar (value);
        abstr = SysAllocString (tmp);
        xfree (tmp);
      }
      if (!abstr)
        log_error ("%s:%s: SysAllocString failed\n", SRCNAME, __func__);
      else
        {
          dispparams.rgvarg[0].vt = VT_BSTR;
          dispparams.rgvarg[0].bstrVal = abstr;
          hr = pDisp->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                              DISPATCH_PROPERTYPUT, &dispparams,
                              NULL, NULL, NULL);
          log_debug ("%s:%s: PROPERTYPUT(%s) result -> %#lx\n",
                     SRCNAME, __func__, key, hr);
          SysFreeString (abstr);
        }
      
      pDisp->Release ();
      pDisp = NULL;
      result = 0;
    }

  ul_release (pMessage, __func__, __LINE__);
  ul_release (pMDB, __func__, __LINE__);
  return result;
}

int
put_outlook_property_int (void *pEECB, const char *key, int value)
{
  int result = -1;
  HRESULT hr;
  LPMDB pMDB = NULL;
  LPMESSAGE pMessage = NULL;
  LPDISPATCH pDisp;
  DISPID dispid;
  DISPID dispid_put = DISPID_PROPERTYPUT;
  DISPPARAMS dispparams;
  VARIANT aVariant;

  if (!pEECB)
    return -1;

  hr = ((LPEXCHEXTCALLBACK)pEECB)->GetObject (&pMDB, (LPMAPIPROP *)&pMessage);
  if (FAILED (hr))
    log_debug ("%s:%s: getObject failed: hr=%#lx\n", SRCNAME, __func__, hr);
  else if ( (pDisp = find_outlook_property ((LPEXCHEXTCALLBACK)pEECB,
                                            key, &dispid)))
    {
      dispparams.cNamedArgs = 1;
      dispparams.rgdispidNamedArgs = &dispid_put;
      dispparams.cArgs = 1;
      dispparams.rgvarg = &aVariant;
      dispparams.rgvarg[0].vt = VT_I4;
      dispparams.rgvarg[0].intVal = value;
      hr = pDisp->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
			  DISPATCH_PROPERTYPUT, &dispparams,
			  NULL, NULL, NULL);
      log_debug ("%s:%s: PROPERTYPUT(%s) result -> %#lx\n",
                 SRCNAME, __func__, key, hr);

      pDisp->Release ();
      pDisp = NULL;
      result = 0;
    }

  ul_release (pMessage, __func__, __LINE__);
  ul_release (pMDB, __func__, __LINE__);
  return result;
}


/* Return an Outlook OO property named KEY.  This needs to be some
   kind of string. PEECP is required to indificate the context.  On
   error NULL is returned.   It is usually used with "Body". */
char *
get_outlook_property (void *pEECB, const char *key)
{
  char *result = NULL;
  HRESULT hr;
  LPDISPATCH pDisp;
  DISPID dispid;
  DISPPARAMS dispparamsNoArgs = {NULL, NULL, 0, 0};
  VARIANT aVariant;

  if (!pEECB)
    return NULL;

  pDisp = find_outlook_property ((LPEXCHEXTCALLBACK)pEECB, key, &dispid);
  if (!pDisp)
    return NULL;

  aVariant.bstrVal = NULL;
  hr = pDisp->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                      DISPATCH_PROPERTYGET, &dispparamsNoArgs,
                      &aVariant, NULL, NULL);
  if (hr != S_OK)
    log_debug ("%s:%s: retrieving `%s' failed: %#lx",
               SRCNAME, __func__, key, hr);
  else if (aVariant.vt != VT_BSTR)
    log_debug ("%s:%s: `%s' is not a string (%d)",
                           SRCNAME, __func__, key, aVariant.vt);
  else if (aVariant.bstrVal)
    {
      result = wchar_to_utf8 (aVariant.bstrVal);
      //log_debug ("%s:%s: `%s' is `%s'",
      //           SRCNAME, __func__, key, result);
      /* From MSDN (Invoke): It is up to the caller to free the return value.*/
      SysFreeString (aVariant.bstrVal);
    }

  pDisp->Release();
  pDisp = NULL;

  return result;
}


/* Check whether the preview pane is visisble.  Returns:
   -1 := Don't know.
    0 := No
    1 := Yes.
 */
int
is_preview_pane_visible (LPEXCHEXTCALLBACK eecb)
{
  HRESULT hr;      
  LPDISPATCH pDisp;
  DISPID dispid;
  DISPPARAMS dispparams;
  VARIANT aVariant, rVariant;
      
  pDisp = find_outlook_property (eecb,
                                 "Application.ActiveExplorer.IsPaneVisible",
                                 &dispid);
  if (!pDisp)
    {
      log_debug ("%s:%s: ActiveExplorer.IsPaneVisible NOT found\n",
                 SRCNAME, __func__);
      return -1;
    }

  dispparams.rgvarg = &aVariant;
  dispparams.rgvarg[0].vt = VT_INT;
  dispparams.rgvarg[0].intVal = 3; /* olPreview */
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  rVariant.bstrVal = NULL;
  hr = pDisp->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                      DISPATCH_METHOD, &dispparams,
                      &rVariant, NULL, NULL);
  pDisp->Release();
  pDisp = NULL;
  if (hr == S_OK && rVariant.vt != VT_BOOL)
    {
      log_debug ("%s:%s: invoking IsPaneVisible succeeded but vt is %d",
                 SRCNAME, __func__, rVariant.vt);
      if (rVariant.vt == VT_BSTR && rVariant.bstrVal)
        SysFreeString (rVariant.bstrVal);
      return -1;
    }
  if (hr != S_OK)
    {
      log_debug ("%s:%s: invoking IsPaneVisible failed: %#lx",
                 SRCNAME, __func__, hr);
      return -1;
    }
  
  return !!rVariant.boolVal;
  
}


/* Set the preview pane to visible if visble is true or to invisible
   if visible is false.  */
void
show_preview_pane (LPEXCHEXTCALLBACK eecb, int visible)
{
  HRESULT hr;      
  LPDISPATCH pDisp;
  DISPID dispid;
  DISPPARAMS dispparams;
  VARIANT aVariant[2];
      
  pDisp = find_outlook_property (eecb,
                                 "Application.ActiveExplorer.ShowPane",
                                 &dispid);
  if (!pDisp)
    {
      log_debug ("%s:%s: ActiveExplorer.ShowPane NOT found\n",
                 SRCNAME, __func__);
      return;
    }

  dispparams.rgvarg = aVariant;
  dispparams.rgvarg[0].vt = VT_BOOL;
  dispparams.rgvarg[0].boolVal = !!visible;
  dispparams.rgvarg[1].vt = VT_INT;
  dispparams.rgvarg[1].intVal = 3; /* olPreview */
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  hr = pDisp->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                      DISPATCH_METHOD, &dispparams,
                      NULL, NULL, NULL);
  pDisp->Release();
  pDisp = NULL;
  if (hr != S_OK)
    log_debug ("%s:%s: invoking ShowPane(%d) failed: %#lx",
               SRCNAME, __func__, visible, hr);
}



void
show_event_object (LPEXCHEXTCALLBACK eecb, const char *description)
{
#if 1
  (void)eecb;
  (void)description;
#else
  HRESULT hr;
  LPOUTLOOKEXTCALLBACK outlook_cb;
  LPUNKNOWN obj;
  LPDISPATCH disp;
  LPTYPEINFO tinfo;
  BSTR bstrname;
  char *name;

  log_debug ("%s:%s: looking for origin of event: %s\n", 
             SRCNAME, __func__, description);

  outlook_cb = NULL;
  eecb->QueryInterface(IID_IOutlookExtCallback, (void **)&outlook_cb);
  if (!outlook_cb)
    {
      log_debug ("%s%s: no outlook callback found\n", SRCNAME, __func__);
      return;
    }
		
  obj = NULL;
  outlook_cb->GetObject (&obj);
  if (!obj)
    {
      log_debug ("%s%s: no object found for event\n", SRCNAME, __func__);
      outlook_cb->Release ();
      return;
    }

  disp = NULL;
  obj->QueryInterface (IID_IDispatch, (void **)&disp);
  obj->Release ();
  obj = NULL;
  if (!disp)
    {
      log_debug ("%s%s: no dispatcher found for event\n", SRCNAME, __func__);
      outlook_cb->Release ();
      return;
    }

  tinfo = NULL;
  disp->GetTypeInfo (0, 0, &tinfo);
  if (!tinfo)
    {
      log_debug ("%s%s: no typeinfo found for dispatcher\n", 
                 SRCNAME, __func__);
      disp->Release ();
      outlook_cb->Release ();
      return;
    }

  bstrname = NULL;
  hr = tinfo->GetDocumentation (MEMBERID_NIL, &bstrname, 0, 0, 0);
  if (hr || !bstrname)
    log_debug ("%s%s: GetDocumentation failed: hr=%#lx\n", 
               SRCNAME, __func__, hr);

  if (bstrname)
    {
      name = wchar_to_utf8 (bstrname);
      SysFreeString (bstrname);
      log_debug ("%s:%s: event fired by item type `%s'\n",
                 SRCNAME, __func__, name);
      xfree (name);
    }

  disp->Release ();
  outlook_cb->Release ();
#endif
}


/* 
   Test code
 */
void
add_oom_command_button (LPEXCHEXTCALLBACK eecb)
{
  HRESULT hr;      
  LPUNKNOWN pObj;
  LPDISPATCH pDisp;
  DISPID dispid;
  DISPPARAMS dispparams;
  VARIANT aVariant[5];
  VARIANT rVariant;
  const char name[] = 
    "Application.ActiveExplorer.CommandBars.Item(Standard).FindControl";

  log_debug ("%s:%s: ENTER", SRCNAME, __func__);
  pDisp = find_outlook_property (eecb, name, &dispid);
  if (!pDisp)
    {
      log_debug ("%s:%s: %s NOT found\n",
                 SRCNAME, __func__, name);
      return;
    }

  dispparams.rgvarg = aVariant;
  dispparams.rgvarg[0].vt = VT_ERROR;
  dispparams.rgvarg[0].scode = DISP_E_PARAMNOTFOUND; 
  dispparams.rgvarg[1].vt = VT_ERROR;
  dispparams.rgvarg[1].scode = DISP_E_PARAMNOTFOUND; 
  dispparams.rgvarg[2].vt = VT_INT;
  dispparams.rgvarg[2].intVal = 4; /* Print button.  */
  dispparams.rgvarg[3].vt = VT_ERROR;
  dispparams.rgvarg[3].scode = DISP_E_PARAMNOTFOUND; 
  dispparams.cArgs = 4;
  dispparams.cNamedArgs = 0;
  rVariant.bstrVal = NULL;
  hr = pDisp->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                      DISPATCH_METHOD, &dispparams,
                      &rVariant, NULL, NULL);
  pDisp->Release();
  pDisp = NULL;
  if (hr != S_OK)
    {
      log_debug ("%s:%s: invoking %s failed: %#lx",
                 SRCNAME, __func__, name, hr);
      return ;
    }

  log_debug ("%s:%s: invoking %s succeeded; vt is %d",
             SRCNAME, __func__, name, rVariant.vt);
  if (rVariant.vt == VT_DISPATCH)
    {
      log_debug ("%s:%s:   rVariant.pdispVal=%p", 
                 SRCNAME, __func__, rVariant.pdispVal);
      pObj = rVariant.pdispVal;
      pObj->QueryInterface (IID_IDispatch, (LPVOID*)&pDisp);
      log_debug ("%s:%s:   queryinterface=%p", 
                 SRCNAME, __func__, pDisp);
      if (pDisp)
        {
          DISPPARAMS dispparamsNoArgs = {NULL, NULL, 0, 0};
          VARIANT rVariant2;
          char *result;
          wchar_t *wname;

          wname = utf8_to_wchar ("Caption");
          if (!wname)
            return;

          hr = pDisp->GetIDsOfNames(IID_NULL, &wname, 1,
                                    LOCALE_SYSTEM_DEFAULT, &dispid);
          xfree (wname);
          log_debug ("   dispid(%s)=%d  (hr=0x%x)\n",
                        "Caption", (int)dispid, (unsigned int)hr);
          
          rVariant2.bstrVal = NULL;
          hr = pDisp->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                              DISPATCH_PROPERTYGET, &dispparamsNoArgs,
                              &rVariant2, NULL, NULL);
          if (hr != S_OK)
            log_debug ("%s:%s: retrieving dispid failed: %#lx",
                       SRCNAME, __func__, hr);
          else if (rVariant2.vt != VT_BSTR)
            log_debug ("%s:%s: id is not a string but vt %d",
                           SRCNAME, __func__, rVariant2.vt);
          else if (rVariant2.bstrVal)
            {
              result = wchar_to_utf8 (rVariant2.bstrVal);
              log_debug ("%s:%s: `id is `%s'", SRCNAME, __func__,  result);
              SysFreeString (rVariant2.bstrVal);
              xfree (result);
            }
        }
    }
  else
    log_debug ("%s:%s: ERROR: unexpected vt", SRCNAME, __func__);
    
  if (rVariant.vt == VT_BSTR && rVariant.bstrVal)
    SysFreeString (rVariant.bstrVal);
}


