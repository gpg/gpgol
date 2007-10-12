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
/* FIXME: Duplicated code.  */
static void 
ul_release (LPVOID punk)
{
  ULONG res;
  
  if (!punk)
    return;
  res = UlRelease (punk);
//   log_debug ("%s UlRelease(%p) had %lu references\n", __func__, punk, res);
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
      DISPPARAMS dispparamsNoArgs = {NULL, NULL, 0, 0};
      VARIANT vtResult;

      /* Our loop expects that all objects except for the last one are
         of class IDispatch.  This is pretty reasonable. */
      pObj->QueryInterface (IID_IDispatch, (LPVOID*)&pDisp);
      if (!pDisp)
        return NULL;
      
      wname = utf8_to_wchar2 (name, s-name);
      if (!wname)
        return NULL;

      hr = pDisp->GetIDsOfNames(IID_NULL, &wname, 1,
                                LOCALE_SYSTEM_DEFAULT, &dispid);
      xfree (wname);
      //log_debug ("   dispid(%.*s)=%d  (hr=0x%x)\n",
      //           (int)(s-name), name, dispid, hr);
      vtResult.pdispVal = NULL;
      hr = pDisp->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                          DISPATCH_METHOD, &dispparamsNoArgs,
                          &vtResult, NULL, NULL);
      pObj = vtResult.pdispVal;
      /* FIXME: Check that the class of the returned object is as
         expected.  To do this we better let GetIdsOfNames also return
         the ID of "Class". */
      //log_debug ("%s:%s: %.*s=%p  (hr=0x%x)\n",
      //           SRCNAME, __func__, (int)(s-name), name, pObj, hr);
      pDisp->Release ();
      pDisp = NULL;
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
  //log_debug ("   dispid(%s)=%d  (hr=0x%x)\n", name, dispid, hr);
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

  ul_release (pMessage);
  ul_release (pMDB);
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

  ul_release (pMessage);
  ul_release (pMDB);
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
      log_debug ("%s:%s: `%s' is `%s'",
                 SRCNAME, __func__, key, result);
      /* From MSDN (Invoke): It is up to the caller to free the return value.*/
      SysFreeString (aVariant.bstrVal);
    }

  pDisp->Release();
  pDisp = NULL;

  return result;
}



