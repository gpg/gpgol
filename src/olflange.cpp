/* olflange.cpp - Flange between Outlook and the GpgMsg class
 *	Copyright (C) 2001 G Data Software AG, http://www.gdata.de
 *	Copyright (C) 2004, 2005 g10 Code GmbH
 * 
 * This file is part of GPGol.
 * 
 * GPGol is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2 of the License, or (at your option) any later version.
 * 
 * GPGol is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Lesser General Public License for more details.
 * 
 * You should have received a copy of the GNU Lesser General Public
 * License along with this program; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA
 * 02110-1301, USA.
 */

#ifdef HAVE_CONFIG_H
#include <config.h>
#endif

#include <windows.h>

#ifndef INITGUID
#define INITGUID
#endif

#include <initguid.h>
#include "mymapi.h"
#include "mymapitags.h"
#include "myexchext.h"
#include "display.h"
#include "intern.h"
#include "gpgmsg.hh"
#include "msgcache.h"
#include "engine.h"

#include "olflange-ids.h"
#include "olflange-def.h"
#include "olflange.h"
#include "attach.h"

#define CLSIDSTR_GPGOL   "{42d30988-1a3a-11da-c687-000d6080e735}"
DEFINE_GUID(CLSID_GPGOL, 0x42d30988, 0x1a3a, 0x11da, 
            0xc6, 0x87, 0x00, 0x0d, 0x60, 0x80, 0xe7, 0x35);


#define TRACEPOINT() do { log_debug ("%s:%s:%d: tracepoint\n", \
                                     SRCNAME, __func__, __LINE__); \
                        } while (0)


bool g_initdll = FALSE;

static HWND show_window_hierarchy (HWND parent, int level);


/* Registers this module as an Exchange extension. This basically updates
   some Registry entries. */
STDAPI 
DllRegisterServer (void)
{    
  HKEY hkey, hkey2;
  CHAR szKeyBuf[MAX_PATH+1024];
  CHAR szEntry[MAX_PATH+512];
  TCHAR szModuleFileName[MAX_PATH];
  DWORD dwTemp = 0;
  long ec;

  /* Get server location. */
  if (!GetModuleFileName(glob_hinst, szModuleFileName, MAX_PATH))
    return E_FAIL;

  lstrcpy (szKeyBuf, "Software\\Microsoft\\Exchange\\Client\\Extensions");
  lstrcpy (szEntry, "4.0;");
  lstrcat (szEntry, szModuleFileName);
  lstrcat (szEntry, ";1"); /* Entry point ordinal. */
  /* Context information string:
     pos       context
     1 	EECONTEXT_SESSION
     2 	EECONTEXT_VIEWER
     3 	EECONTEXT_REMOTEVIEWER
     4 	EECONTEXT_SEARCHVIEWER
     5 	EECONTEXT_ADDRBOOK
     6 	EECONTEXT_SENDNOTEMESSAGE
     7 	EECONTEXT_READNOTEMESSAGE
     8 	EECONTEXT_SENDPOSTMESSAGE
     9 	EECONTEXT_READPOSTMESSAGE
     10 EECONTEXT_READREPORTMESSAGE
     11 EECONTEXT_SENDRESENDMESSAGE
     12 EECONTEXT_PROPERTYSHEETS
     13 EECONTEXT_ADVANCEDCRITERIA
     14 EECONTEXT_TASK
  */
  lstrcat (szEntry, ";11000111111100"); 
  ec = RegCreateKeyEx (HKEY_LOCAL_MACHINE, szKeyBuf, 0, NULL, 
                       REG_OPTION_NON_VOLATILE,
                       KEY_ALL_ACCESS, NULL, &hkey, NULL);
  if (ec != ERROR_SUCCESS) 
    {
      log_debug ("DllRegisterServer failed\n");
      return E_ACCESSDENIED;
    }
    
  dwTemp = lstrlen (szEntry) + 1;
  RegSetValueEx (hkey, "GPGol", 0, REG_SZ, (BYTE*) szEntry, dwTemp);

  /* To avoid conflicts with the old G-DATA plugin and older vesions
     of this Plugin, we remove the key used by these versions. */
  RegDeleteValue (hkey, "GPG Exchange");

  /* Set outlook update flag. */
  strcpy (szEntry, "4.0;Outxxx.dll;7;000000000000000;0000000000;OutXXX");
  dwTemp = lstrlen (szEntry) + 1;
  RegSetValueEx (hkey, "Outlook Setup Extension",
                 0, REG_SZ, (BYTE*) szEntry, dwTemp);
  RegCloseKey (hkey);
    
  hkey = NULL;
  lstrcpy (szKeyBuf, "Software\\GNU\\GPGol");
  RegCreateKeyEx (HKEY_CURRENT_USER, szKeyBuf, 0, NULL,
                  REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hkey, NULL);
  if (hkey != NULL)
    RegCloseKey (hkey);

  hkey = NULL;
  strcpy (szKeyBuf, "CLSID\\" CLSIDSTR_GPGOL );
  ec = RegCreateKeyEx (HKEY_CLASSES_ROOT, szKeyBuf, 0, NULL,
                  REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hkey, NULL);
  if (ec != ERROR_SUCCESS) 
    {
      fprintf (stderr, "creating key `%s' failed: ec=%#lx\n", szKeyBuf, ec);
      return E_ACCESSDENIED;
    }

  strcpy (szEntry, "GPGol - The GPG Outlook Plugin");
  dwTemp = strlen (szEntry) + 1;
  RegSetValueEx (hkey, NULL, 0, REG_SZ, (BYTE*)szEntry, dwTemp);

  strcpy (szKeyBuf, "InprocServer32");
  ec = RegCreateKeyEx (hkey, szKeyBuf, 0, NULL, REG_OPTION_NON_VOLATILE,
                       KEY_ALL_ACCESS, NULL, &hkey2, NULL);
  if (ec != ERROR_SUCCESS) 
    {
      fprintf (stderr, "creating key `%s' failed: ec=%#lx\n", szKeyBuf, ec);
      RegCloseKey (hkey);
      return E_ACCESSDENIED;
    }
  strcpy (szEntry, szModuleFileName);
  dwTemp = strlen (szEntry) + 1;
  RegSetValueEx (hkey2, NULL, 0, REG_SZ, (BYTE*)szEntry, dwTemp);

  strcpy (szEntry, "Neutral");
  dwTemp = strlen (szEntry) + 1;
  RegSetValueEx (hkey2, "ThreadingModel", 0, REG_SZ, (BYTE*)szEntry, dwTemp);

  RegCloseKey (hkey2);
  RegCloseKey (hkey);


  log_debug ("DllRegisterServer succeeded\n");
  return S_OK;
}


/* Unregisters this module as an Exchange extension. */
STDAPI 
DllUnregisterServer (void)
{
  HKEY hkey;
  CHAR buf[512];
  DWORD ntemp;
  long res;

  strcpy (buf, "Software\\Microsoft\\Exchange\\Client\\Extensions");
  /* create and open key and subkey */
  res = RegCreateKeyEx (HKEY_LOCAL_MACHINE, buf, 0, NULL, 
			REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 
			NULL, &hkey, NULL);
  if (res != ERROR_SUCCESS) 
    {
      log_debug ("DllUnregisterServer: access denied.\n");
      return E_ACCESSDENIED;
    }
  RegDeleteValue (hkey, "GPGol");
  
  /* set outlook update flag */  
  strcpy (buf, "4.0;Outxxx.dll;7;000000000000000;0000000000;OutXXX");
  ntemp = strlen (buf) + 1;
  RegSetValueEx (hkey, "Outlook Setup Extension", 0, 
		 REG_SZ, (BYTE*) buf, ntemp);
  RegCloseKey (hkey);
  
  return S_OK;
}

/* Wrapper around UlRelease with error checking. */
static void 
ul_release (LPVOID punk)
{
  ULONG res;
  
  if (!punk)
    return;
  res = UlRelease (punk);
//   log_debug ("%s UlRelease(%p) had %lu references\n", __func__, punk, res);
}



/* DISPLAY a MAPI property. */
static void
show_mapi_property (LPMESSAGE message, ULONG prop, const char *propname)
{
  HRESULT hr;
  LPSPropValue lpspvFEID = NULL;
  size_t keylen;
  void *key;

  if (!message)
    return; /* No message: Nop. */

  hr = HrGetOneProp ((LPMAPIPROP)message, prop, &lpspvFEID);
  if (FAILED (hr))
    {
      log_debug ("%s: HrGetOneProp(%s) failed: hr=%#lx\n",
                 __func__, propname, hr);
      return;
    }
    
  if ( PROP_TYPE (lpspvFEID->ulPropTag) != PT_BINARY )
    {
      log_debug ("%s: HrGetOneProp(%s) returned unexpected property type\n",
                 __func__, propname);
      MAPIFreeBuffer (lpspvFEID);
      return;
    }
  keylen = lpspvFEID->Value.bin.cb;
  key = lpspvFEID->Value.bin.lpb;
  log_hexdump (key, keylen, "%s: %20s=", __func__, propname);
  MAPIFreeBuffer (lpspvFEID);
}



/* Locate a property using the current callback LPEECB and traverse
   down to the last element in the dot delimited NAME.  Returns the
   Dispatch object and if R_DISPID is not NULL, the dispatch-id of the
   last part.  Returns NULL on error.  The traversal implictly starts
   at the object returned by the outlook application callback. */
static LPDISPATCH
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


/* Return Outlook's Application object. */
/* FIXME: We should be able to fold most of the code into
   find_outlook_property. */
#if 0 /* Not used as of now. */
static LPUNKNOWN
get_outlook_application_object (LPEXCHEXTCALLBACK lpeecb)
{
  LPOUTLOOKEXTCALLBACK pCb = NULL;
  LPDISPATCH pDisp = NULL;
  LPUNKNOWN pUnk = NULL;

  lpeecb->QueryInterface (IID_IOutlookExtCallback, (LPVOID*)&pCb);
  if (pCb)
    pCb->GetObject (&pUnk);
  if (pUnk)
    {
      pUnk->QueryInterface (IID_IDispatch, (LPVOID*)&pDisp);
      pUnk->Release();
      pUnk = NULL;
    }

  if (pDisp)
    {
      WCHAR *name = L"Class";
      DISPPARAMS dispparamsNoArgs = {NULL, NULL, 0, 0};
      DISPID dispid;
      VARIANT vtResult;

      pDisp->GetIDsOfNames(IID_NULL, &name, 1,
                           LOCALE_SYSTEM_DEFAULT, &dispid);
      vtResult.pdispVal = NULL;
      pDisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                    DISPATCH_PROPERTYGET, &dispparamsNoArgs,
                    &vtResult, NULL, NULL);
      log_debug ("%s:%s: Outlookcallback returned object of class=%d\n",
                   SRCNAME, __func__, vtResult.intVal);
    }
  if (pDisp)
    {
      WCHAR *name = L"Application";
      DISPPARAMS dispparamsNoArgs = {NULL, NULL, 0, 0};
      DISPID dispid;
      VARIANT vtResult;
      
      pDisp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
      //log_debug ("   dispid(Application)=%d\n", dispid);
      vtResult.pdispVal = NULL;
      pDisp->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                     DISPATCH_METHOD, &dispparamsNoArgs,
                     &vtResult, NULL, NULL);
      pUnk = vtResult.pdispVal;
      //log_debug ("%s:%s: Outlook.Application=%p\n",
      //             SRCNAME, __func__, pUnk);
      pDisp->Release();
      pDisp = NULL;
    }
  return pUnk;
}
#endif /* commented */


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


/* Retuirn an Outlook OO property anmed KEY.  This needs to be some
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
      /* FIXME: Do we need to free the string returned in  AVARIANT? */
    }

  pDisp->Release();
  pDisp = NULL;

  return result;
}



/* The entry point which Exchange calls.  This is called for each
   context entry. Creates a new CGPGExchExt object every time so each
   context will get its own CGPGExchExt interface. */
EXTERN_C LPEXCHEXT __stdcall
ExchEntryPoint (void)
{
  log_debug ("%s:%s: creating new CGPGExchExt object\n", SRCNAME, __func__);
  return new CGPGExchExt;
}


/* Constructor of CGPGExchExt

   Initializes members and creates the interface objects for the new
   context.  Does the DLL initialization if it has not been done
   before. */
CGPGExchExt::CGPGExchExt (void)
{ 
  m_lRef = 1;
  m_lContext = 0;
  m_hWndExchange = 0;
  m_gpgEncrypt = FALSE;
  m_gpgSign = FALSE;
  m_pExchExtMessageEvents = new CGPGExchExtMessageEvents (this);
  m_pExchExtCommands = new CGPGExchExtCommands (this);
  m_pExchExtPropertySheets = new CGPGExchExtPropertySheets (this);
  m_pExchExtAttachedFileEvents = new CGPGExchExtAttachedFileEvents (this);
  if (!m_pExchExtMessageEvents || !m_pExchExtCommands
      || !m_pExchExtPropertySheets || !m_pExchExtAttachedFileEvents)
    out_of_core ();
  
  if (!g_initdll)
    {
      if (opt.compat.auto_decrypt)
        watcher_init_hook ();
      read_options ();
      op_init ();
      g_initdll = TRUE;
      log_debug ("%s:%s: first time initialization done\n",
                 SRCNAME, __func__);
    }
}


/*  Uninitializes the DLL in the session context. */
CGPGExchExt::~CGPGExchExt (void) 
{
  log_debug ("%s:%s: cleaning up CGPGExchExt object; "
             "context=0x%lx (%s)\n", SRCNAME, __func__, 
             m_lContext,
             (m_lContext == EECONTEXT_SESSION?           "Session":
              m_lContext == EECONTEXT_VIEWER?            "Viewer":
              m_lContext == EECONTEXT_REMOTEVIEWER?      "RemoteViewer":
              m_lContext == EECONTEXT_SEARCHVIEWER?      "SearchViewer":
              m_lContext == EECONTEXT_ADDRBOOK?          "AddrBook" :
              m_lContext == EECONTEXT_SENDNOTEMESSAGE?   "SendNoteMessage" :
              m_lContext == EECONTEXT_READNOTEMESSAGE?   "ReadNoteMessage" :
              m_lContext == EECONTEXT_SENDPOSTMESSAGE?   "SendPostMessage" :
              m_lContext == EECONTEXT_READPOSTMESSAGE?   "ReadPostMessage" :
              m_lContext == EECONTEXT_READREPORTMESSAGE? "ReadReportMessage" :
              m_lContext == EECONTEXT_SENDRESENDMESSAGE? "SendResendMessage" :
              m_lContext == EECONTEXT_PROPERTYSHEETS?    "PropertySheets" :
              m_lContext == EECONTEXT_ADVANCEDCRITERIA?  "AdvancedCriteria" :
              m_lContext == EECONTEXT_TASK?              "Task" : ""));
  
  if (m_lContext == EECONTEXT_SESSION)
    {
      if (g_initdll)
        {
	  watcher_free_hook ();
          op_deinit ();
          write_options ();
          g_initdll = FALSE;
          log_debug ("%s:%s: DLL closed down\n", SRCNAME, __func__);
	}	
    }

}


/* Called by Exchange to retrieve an object pointer for a named
   interface.  This is a standard COM method.  REFIID is the ID of the
   interface and PPVOBJ will get the address of the object pointer if
   this class defines the requested interface.  Return value: S_OK if
   the interface is supported, otherwise E_NOINTERFACE. */
STDMETHODIMP 
CGPGExchExt::QueryInterface(REFIID riid, LPVOID *ppvObj)
{
  HRESULT hr = S_OK;
  
  *ppvObj = NULL;
  
  if ((riid == IID_IUnknown) || (riid == IID_IExchExt)) 
    {
      *ppvObj = (LPUNKNOWN) this;
    }
  else if (riid == IID_IExchExtMessageEvents) 
    {
      *ppvObj = (LPUNKNOWN) m_pExchExtMessageEvents;
      m_pExchExtMessageEvents->SetContext (m_lContext);
    }
  else if (riid == IID_IExchExtCommands) 
    {
      *ppvObj = (LPUNKNOWN)m_pExchExtCommands;
      m_pExchExtCommands->SetContext (m_lContext);
    }
  else if (riid == IID_IExchExtPropertySheets) 
    {
      if (m_lContext != EECONTEXT_PROPERTYSHEETS)
	return E_NOINTERFACE;
      *ppvObj = (LPUNKNOWN) m_pExchExtPropertySheets;
    }
  else if (riid == IID_IExchExtAttachedFileEvents)
    {
      *ppvObj = (LPUNKNOWN)m_pExchExtAttachedFileEvents;
    }  
  else
    hr = E_NOINTERFACE;
  
  /* On success we need to bump up the reference counter for the
   requested object. */
  if (*ppvObj)
    ((LPUNKNOWN)*ppvObj)->AddRef();
  
  /*log_debug("QueryInterface %d\n", __LINE__);*/
    return hr;
}


/* Called once for each new context. Checks the Exchange extension
   version number and the context.  Returns: S_OK if the extension
   should used in the requested context, otherwise S_FALSE.  PEECB is
   a pointer to Exchange extension callback function.  LCONTEXT is the
   context code at time of being called. LFLAGS carries flags to
   indicate whether the extension should be installed modal.
*/
STDMETHODIMP 
CGPGExchExt::Install(LPEXCHEXTCALLBACK pEECB, ULONG lContext, ULONG lFlags)
{
  ULONG lBuildVersion;
  ULONG lActualVersion;
  ULONG lVirtualVersion;

  /* Save the context in an instance variable. */
  m_lContext = lContext;

  log_debug ("%s:%s: context=0x%lx (%s) flags=0x%lx\n", SRCNAME, __func__, 
               lContext,
               (lContext == EECONTEXT_SESSION?           "Session":
                lContext == EECONTEXT_VIEWER?            "Viewer":
                lContext == EECONTEXT_REMOTEVIEWER?      "RemoteViewer":
                lContext == EECONTEXT_SEARCHVIEWER?      "SearchViewer":
                lContext == EECONTEXT_ADDRBOOK?          "AddrBook" :
                lContext == EECONTEXT_SENDNOTEMESSAGE?   "SendNoteMessage" :
                lContext == EECONTEXT_READNOTEMESSAGE?   "ReadNoteMessage" :
                lContext == EECONTEXT_SENDPOSTMESSAGE?   "SendPostMessage" :
                lContext == EECONTEXT_READPOSTMESSAGE?   "ReadPostMessage" :
                lContext == EECONTEXT_READREPORTMESSAGE? "ReadReportMessage" :
                lContext == EECONTEXT_SENDRESENDMESSAGE? "SendResendMessage" :
                lContext == EECONTEXT_PROPERTYSHEETS?    "PropertySheets" :
                lContext == EECONTEXT_ADVANCEDCRITERIA?  "AdvancedCriteria" :
                lContext == EECONTEXT_TASK?              "Task" : ""),
               lFlags);
  
  /* Check version. */
  log_debug ("GPGol: this is %s\n", PACKAGE_STRING);
  pEECB->GetVersion (&lBuildVersion, EECBGV_GETBUILDVERSION);
  pEECB->GetVersion (&lActualVersion, EECBGV_GETACTUALVERSION);
  pEECB->GetVersion (&lVirtualVersion, EECBGV_GETVIRTUALVERSION);
  log_debug ("GPGol: detected Outlook build version 0x%lx (%lu.%lu)\n",
             lBuildVersion,
             (lBuildVersion & EECBGV_BUILDVERSION_MAJOR_MASK) >> 16,
             (lBuildVersion & EECBGV_BUILDVERSION_MINOR_MASK));
  log_debug ("GPGol:                 actual version 0x%lx (%u.%u.%u.%u)\n",
             lActualVersion, 
             (unsigned int)((lActualVersion >> 24) & 0xff),
             (unsigned int)((lActualVersion >> 16) & 0xff),
             (unsigned int)((lActualVersion >> 8) & 0xff),
             (unsigned int)(lActualVersion & 0xff));
  log_debug ("GPGol:                virtual version 0x%lx (%u.%u.%u.%u)\n",
             lVirtualVersion, 
             (unsigned int)((lVirtualVersion >> 24) & 0xff),
             (unsigned int)((lVirtualVersion >> 16) & 0xff),
             (unsigned int)((lVirtualVersion >> 8) & 0xff),
             (unsigned int)(lVirtualVersion & 0xff));

  if (EECBGV_BUILDVERSION_MAJOR
      != (lBuildVersion & EECBGV_BUILDVERSION_MAJOR_MASK))
    {
      log_debug ("%s:%s: invalid version 0x%lx\n",
                   SRCNAME, __func__, lBuildVersion);
      return S_FALSE;
    }
  if ((lBuildVersion & EECBGV_BUILDVERSION_MAJOR_MASK) < 13
      ||(lBuildVersion & EECBGV_BUILDVERSION_MINOR_MASK) < 1573)
    {
      static int shown;
      HWND hwnd;
      
      if (!shown)
        {
          shown = 1;
          
          if (FAILED(pEECB->GetWindow (&hwnd)))
            hwnd = NULL;
          MessageBox (hwnd,
                      _("This version of Outlook is too old!\n\n"
                        "At least versions of Outlook 2003 older than SP2 "
                        "exhibit crashes when sending messages and messages "
                        "might get stuck in the outgoing queue.\n\n"
                        "Please update at least to SP2 before trying to send "
                        "a message"),
                      "GPGol", MB_ICONSTOP|MB_OK);
        }
    }
  

  /* Check context. */
  if (   lContext == EECONTEXT_PROPERTYSHEETS
      || lContext == EECONTEXT_SENDNOTEMESSAGE
      || lContext == EECONTEXT_SENDPOSTMESSAGE
      || lContext == EECONTEXT_SENDRESENDMESSAGE
      || lContext == EECONTEXT_READNOTEMESSAGE
      || lContext == EECONTEXT_READPOSTMESSAGE
      || lContext == EECONTEXT_READREPORTMESSAGE
      || lContext == EECONTEXT_VIEWER)
    {
//       LPUNKNOWN pApplication = get_outlook_application_object (pEECB);
//       log_debug ("%s:%s: pApplication=%p\n",
//                    SRCNAME, __func__, pApplication);
      return S_OK;
    }
  
  log_debug ("%s:%s: can't handle this context\n", SRCNAME, __func__);
  return S_FALSE;
}



CGPGExchExtMessageEvents::CGPGExchExtMessageEvents 
                                              (CGPGExchExt *pParentInterface)
{ 
  m_pExchExt = pParentInterface;
  m_lRef = 0; 
  m_bOnSubmitActive = FALSE;
  m_want_html = FALSE;
}


STDMETHODIMP 
CGPGExchExtMessageEvents::QueryInterface (REFIID riid, LPVOID FAR *ppvObj)
{   
    *ppvObj = NULL;
    if (riid == IID_IExchExtMessageEvents) {
        *ppvObj = (LPVOID)this;
        AddRef();
        return S_OK;
    }
    if (riid == IID_IUnknown) {
        *ppvObj = (LPVOID)m_pExchExt;  
        m_pExchExt->AddRef();
        return S_OK;
    }
    return E_NOINTERFACE;
}


/* Called from Exchange on reading a message.  Returns: S_FALSE to
   signal Exchange to continue calling extensions.  PEECB is a pointer
   to the IExchExtCallback interface. */
STDMETHODIMP 
CGPGExchExtMessageEvents::OnRead (LPEXCHEXTCALLBACK pEECB) 
{
  LPMDB pMDB = NULL;
  LPMESSAGE pMessage = NULL;

  log_debug ("%s:%s: received\n", SRCNAME, __func__);
  pEECB->GetObject (&pMDB, (LPMAPIPROP *)&pMessage);
  show_mapi_property (pMessage, PR_CONVERSATION_INDEX,"PR_CONVERSATION_INDEX");
  ul_release (pMessage);
  ul_release (pMDB);

  return S_FALSE;
}


/* Called by Exchange after a message has been read.  Returns: S_FALSE
   to signal Exchange to continue calling extensions.  PEECB is a
   pointer to the IExchExtCallback interface. LFLAGS are some flags. */
STDMETHODIMP 
CGPGExchExtMessageEvents::OnReadComplete (LPEXCHEXTCALLBACK pEECB,
                                          ULONG lFlags)
{
  log_debug ("%s:%s: received\n", SRCNAME, __func__);

  /* The preview_info stuff does not work because for some reasons we
     can't update the window.  Thus disabled for now. */
  if (opt.preview_decrypt /*|| !opt.compat.no_preview_info*/)
    {
      HRESULT hr;
      HWND hWnd = NULL;
      LPMESSAGE pMessage = NULL;
      LPMDB pMDB = NULL;

      if (FAILED (pEECB->GetWindow (&hWnd)))
        hWnd = NULL;
      hr = pEECB->GetObject (&pMDB, (LPMAPIPROP *)&pMessage);
      if (SUCCEEDED (hr))
        {
          GpgMsg *m = CreateGpgMsg (pMessage);
          m->setExchangeCallback ((void*)pEECB);
          m->setPreview (1);
          /* If preview decryption has been requested, do so.  If not,
             pass true as the second arg to let the fucntion display a
             hint on what kind of message this is. */
          m->decrypt (hWnd, !opt.preview_decrypt);
          delete m;
 	}
      ul_release (pMessage);
      ul_release (pMDB);
    }


#if 0
    {
      HWND hWnd = NULL;

      if (FAILED (pEECB->GetWindow (&hWnd)))
        hWnd = NULL;
      else
        show_window_hierarchy (hWnd, 0);
    }
#endif

  return S_FALSE;
}


/* Called by Exchange when a message will be written. Returns: S_FALSE
   to signal Exchange to continue calling extensions.  PEECB is a
   pointer to the IExchExtCallback interface. */
STDMETHODIMP 
CGPGExchExtMessageEvents::OnWrite (LPEXCHEXTCALLBACK pEECB)
{
  log_debug ("%s:%s: received\n", SRCNAME, __func__);

  HRESULT hr;
  LPDISPATCH pDisp;
  DISPID dispid;
  VARIANT aVariant;
  DISPPARAMS dispparamsNoArgs = {NULL, NULL, 0, 0};
  HWND hWnd = NULL;


  /* If we are going to encrypt, check that the BodyFormat is
     something we support.  This helps avoiding surprise by sending
     out unencrypted messages. */
  if (m_pExchExt->m_gpgEncrypt || m_pExchExt->m_gpgSign)
    {
      pDisp = find_outlook_property (pEECB, "BodyFormat", &dispid);
      if (!pDisp)
        {
          log_debug ("%s:%s: BodyFormat not found\n", SRCNAME, __func__);
          m_bWriteFailed = TRUE;	
          return E_FAIL;
        }
      
      aVariant.bstrVal = NULL;
      hr = pDisp->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                          DISPATCH_PROPERTYGET, &dispparamsNoArgs,
                          &aVariant, NULL, NULL);
      if (hr != S_OK)
        {
          log_debug ("%s:%s: retrieving BodyFormat failed: %#lx",
                     SRCNAME, __func__, hr);
          m_bWriteFailed = TRUE;	
          pDisp->Release();
          return E_FAIL;
        }
  
      if (aVariant.vt != VT_INT && aVariant.vt != VT_I4)
        {
          log_debug ("%s:%s: BodyFormat is not an integer (%d)",
                     SRCNAME, __func__, aVariant.vt);
          m_bWriteFailed = TRUE;	
          pDisp->Release();
          return E_FAIL;
        }
  
      if (aVariant.intVal == 1)
        m_want_html = 0;
      else if (aVariant.intVal == 2)
        m_want_html = 1;
      else
        {

          log_debug ("%s:%s: BodyFormat is %d",
                     SRCNAME, __func__, aVariant.intVal);
          
          if (FAILED(pEECB->GetWindow (&hWnd)))
            hWnd = NULL;
          MessageBox (hWnd,
                      _("Sorry, we can only encrypt plain text messages and\n"
                      "no RTF messages. Please make sure that only the text\n"
                      "format has been selected."),
                      "GPGol", MB_ICONERROR|MB_OK);

          m_bWriteFailed = TRUE;	
          pDisp->Release();
          return E_FAIL;
        }
      pDisp->Release();
      
    }
  
  
  return S_FALSE;
}


/* Called by Exchange when the data has been written to the message.
   Encrypts and signs the message if the options are set.  PEECB is a
   pointer to the IExchExtCallback interface.  Returns: S_FALSE to
   signal Exchange to continue calling extensions.  We return E_FAIL
   to signals Exchange an error; the message will not be sent.  Note
   that we don't return S_OK because this would mean that we rolled
   back the write operation and that no further extensions should be
   called. */
STDMETHODIMP 
CGPGExchExtMessageEvents::OnWriteComplete (LPEXCHEXTCALLBACK pEECB,
                                           ULONG lFlags)
{
  log_debug ("%s:%s: received\n", SRCNAME, __func__);

  HRESULT hrReturn = S_FALSE;
  LPMESSAGE msg = NULL;
  LPMDB pMDB = NULL;
  HWND hWnd = NULL;
  int rc;

  if (lFlags & (EEME_FAILED|EEME_COMPLETE_FAILED))
    return S_FALSE; /* We don't need to rollback anything in case
                       other extensions flagged a failure. */
          
  if (!m_bOnSubmitActive) /* The user is just saving the message. */
    return S_FALSE;
  
  if (m_bWriteFailed)     /* Operation failed already. */
    return S_FALSE;
  
  if (FAILED(pEECB->GetWindow (&hWnd)))
    hWnd = NULL;

  HRESULT hr = pEECB->GetObject (&pMDB, (LPMAPIPROP *)&msg);
  if (SUCCEEDED (hr))
    {
      SPropTagArray proparray;

      GpgMsg *m = CreateGpgMsg (msg);
      m->setExchangeCallback ((void*)pEECB);
      if (m_pExchExt->m_gpgEncrypt && m_pExchExt->m_gpgSign)
        rc = m->signEncrypt (hWnd, m_want_html);
      if (m_pExchExt->m_gpgEncrypt && !m_pExchExt->m_gpgSign)
        rc = m->encrypt (hWnd, m_want_html);
      if (!m_pExchExt->m_gpgEncrypt && m_pExchExt->m_gpgSign)
        rc = m->sign (hWnd, m_want_html);
      else
        rc = 0;
      delete m;

      /* If we are encrypting we need to make sure that the other
         format gets deleted and is not actually sent in the clear.
         Note that this other format is always HTML because we have
         moved that into an attachment and kept PR_BODY.  It seems
         that OL always creates text and HTML if HTML has been
         selected. */
      /* ARGHH: This seems to delete also the PR_BODY for some reasonh
         - need to disable this safe net. */
//       if (m_pExchExt->m_gpgEncrypt)
//         {
//           log_debug ("%s:%s: deleting possible extra property PR_BODY_HTML\n",
//                      SRCNAME, __func__);
//           proparray.cValues = 1;
//           proparray.aulPropTag[0] = PR_BODY_HTML;
//           msg->DeleteProps (&proparray, NULL);
//         }
     
 
      if (rc)
        {
          hrReturn = E_FAIL;
          m_bWriteFailed = TRUE;	

          /* Due to a bug in Outlook the error is ignored and the
             message sent out anyway.  Thus we better delete the stuff
             now. */
          if (m_pExchExt->m_gpgEncrypt)
            {
              log_debug ("%s:%s: deleting property PR_BODY due to error\n",
                         SRCNAME, __func__);
              proparray.cValues = 1;
              proparray.aulPropTag[0] = PR_BODY;
              hr = msg->DeleteProps (&proparray, NULL);
              if (hr != S_OK)
                log_debug ("%s:%s: DeleteProps failed: hr=%#lx\n",
                           SRCNAME, __func__, hr);
              /* FIXME: We should delete the attachments too. 
                 We really, really should do this!!!          */
            }
          
        }
    }

  ul_release (msg);
  ul_release (pMDB);

  return hrReturn;
}

/* Called by Exchange when the user selects the "check names" command.
   PEECB is a pointer to the IExchExtCallback interface.  Returns
   S_FALSE to signal Exchange to continue calling extensions. */
STDMETHODIMP 
CGPGExchExtMessageEvents::OnCheckNames(LPEXCHEXTCALLBACK pEECB)
{
  log_debug ("%s:%s: received\n", SRCNAME, __func__);
  return S_FALSE;
}


/* Called by Exchange when "check names" command is complete.
   PEECB is a pointer to the IExchExtCallback interface.  Returns
   S_FALSE to signal Exchange to continue calling extensions. */
STDMETHODIMP 
CGPGExchExtMessageEvents::OnCheckNamesComplete (LPEXCHEXTCALLBACK pEECB,
                                                ULONG lFlags)
{
  log_debug ("%s:%s: received\n", SRCNAME, __func__);
  return S_FALSE;
}


/* Called by Exchange before the message data will be written and
   submitted to MAPI.  PEECB is a pointer to the IExchExtCallback
   interface.  Returns S_FALSE to signal Exchange to continue calling
   extensions. */
STDMETHODIMP 
CGPGExchExtMessageEvents::OnSubmit (LPEXCHEXTCALLBACK pEECB)
{
  log_debug ("%s:%s: received\n", SRCNAME, __func__);
  m_bOnSubmitActive = TRUE;
  m_bWriteFailed = FALSE;
  return S_FALSE;
}


/* Called by Echange after the message has been submitted to MAPI.
   PEECB is a pointer to the IExchExtCallback interface. */
STDMETHODIMP_ (VOID) 
CGPGExchExtMessageEvents::OnSubmitComplete (LPEXCHEXTCALLBACK pEECB,
                                            ULONG lFlags)
{
  log_debug ("%s:%s: received\n", SRCNAME, __func__);
  m_bOnSubmitActive = FALSE; 
}



CGPGExchExtCommands::CGPGExchExtCommands (CGPGExchExt* pParentInterface)
{ 
  m_pExchExt = pParentInterface; 
  m_lRef = 0; 
  m_lContext = 0; 
  m_nCmdEncrypt = 0;  
  m_nCmdSign = 0; 
  m_nToolbarButtonID1 = 0; 
  m_nToolbarButtonID2 = 0; 
  m_nToolbarBitmap1 = 0;
  m_nToolbarBitmap2 = 0; 
  m_hWnd = NULL; 
}



STDMETHODIMP 
CGPGExchExtCommands::QueryInterface (REFIID riid, LPVOID FAR * ppvObj)
{
    *ppvObj = NULL;
    if ((riid == IID_IExchExtCommands) || (riid == IID_IUnknown)) {
        *ppvObj = (LPVOID)this;
        AddRef ();
        return S_OK;
    }
    return E_NOINTERFACE;
}


static HWND
show_window_hierarchy (HWND parent, int level)
{
  HWND child;

  child = GetWindow (parent, GW_CHILD);
  while (child)
    {
      char buf[1024+1];
      char name[200];
      int nname;
      char *pname;
      
      memset (buf, 0, sizeof (buf));
      GetWindowText (child, buf, sizeof (buf)-1);
      nname = GetClassName (child, name, sizeof (name)-1);
      if (nname)
        pname = name;
      else
        pname = NULL;
      log_debug ("XXX %*shwnd=%p (%s) `%s'", level*2, "", child,
                 pname? pname:"", buf);
      show_window_hierarchy (child, level+1);
      child = GetNextWindow (child, GW_HWNDNEXT);	
    }

  return NULL;
}


/* Called by Exchange to install commands and toolbar buttons.  Returns
   S_FALSE to signal Exchange to continue calling extensions. */
STDMETHODIMP 
CGPGExchExtCommands::InstallCommands (
	LPEXCHEXTCALLBACK pEECB, // The Exchange Callback Interface.
	HWND hWnd,               // The window handle to the main window
                                 // of context.
	HMENU hMenu,             // The menu handle to main menu of context.
	UINT FAR * pnCommandIDBase,  // The base command id.
	LPTBENTRY pTBEArray,     // The array of toolbar button entries.
	UINT nTBECnt,            // The count of button entries in array.
	ULONG lFlags)            // reserved
{
  HRESULT hr;
  HMENU hMenuTools;
  m_hWnd = hWnd;
  LPDISPATCH pDisp;
  DISPID dispid;
  DISPID dispid_put = DISPID_PROPERTYPUT;
  DISPPARAMS dispparams;
  VARIANT aVariant;
  int force_encrypt = 0;
  
  log_debug ("%s:%s: context=0x%lx (%s) flags=0x%lx\n", SRCNAME, __func__, 
             m_lContext,
             (m_lContext == EECONTEXT_SESSION?           "Session"          :
              m_lContext == EECONTEXT_VIEWER?            "Viewer"           :
              m_lContext == EECONTEXT_REMOTEVIEWER?      "RemoteViewer"     :
              m_lContext == EECONTEXT_SEARCHVIEWER?      "SearchViewer"     :
              m_lContext == EECONTEXT_ADDRBOOK?          "AddrBook"         :
              m_lContext == EECONTEXT_SENDNOTEMESSAGE?   "SendNoteMessage"  :
              m_lContext == EECONTEXT_READNOTEMESSAGE?   "ReadNoteMessage"  :
              m_lContext == EECONTEXT_SENDPOSTMESSAGE?   "SendPostMessage"  :
              m_lContext == EECONTEXT_READPOSTMESSAGE?   "ReadPostMessage"  :
              m_lContext == EECONTEXT_READREPORTMESSAGE? "ReadReportMessage":
              m_lContext == EECONTEXT_SENDRESENDMESSAGE? "SendResendMessage":
              m_lContext == EECONTEXT_PROPERTYSHEETS?    "PropertySheets"   :
              m_lContext == EECONTEXT_ADVANCEDCRITERIA?  "AdvancedCriteria" :
              m_lContext == EECONTEXT_TASK?              "Task" : ""),
             lFlags);


  /* Outlook 2003 sometimes displays the plaintext sometimes the
     orginal undecrypted text when doing a Reply.  This seems to
     depend on the size of the message; my guess it that only short
     messages are locally saved in the process and larger ones are
     fetyched again from the backend - or the other way around.
     Anyway, we can't rely on that and thus me make sure to update the
     Body object right here with our own copy of the plaintext.  To
     match the text we use the ConversationIndex property.

     Unfortunately there seems to be no way of resetting the Saved
     property after updating the body, thus even without entering a
     single byte the user will be asked when cancelling a reply
     whether he really wants to do that.  

     Note, that we can't optimize the code here by first reading the
     body because this would pop up the securiy window, telling the
     user that someone is trying to read this data.
  */
  if (m_lContext == EECONTEXT_SENDNOTEMESSAGE)
    {
      LPMDB pMDB = NULL;
      LPMESSAGE pMessage = NULL;
      
      /*  Note that for read and send the object returned by the
          outlook extension callback is of class 43 (MailItem) so we
          only need to ask for Body then. */
      hr = pEECB->GetObject (&pMDB, (LPMAPIPROP *)&pMessage);
      if (FAILED(hr))
        log_debug ("%s:%s: getObject failed: hr=%#lx\n", SRCNAME,__func__,hr);
      else if ( !opt.compat.no_msgcache)
        {
          const char *body;
          char *key = NULL;
          size_t keylen = 0;
          void *refhandle = NULL;
     
          pDisp = find_outlook_property (pEECB, "ConversationIndex", &dispid);
          if (pDisp)
            {
              DISPPARAMS dispparamsNoArgs = {NULL, NULL, 0, 0};

              aVariant.bstrVal = NULL;
              hr = pDisp->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                                  DISPATCH_PROPERTYGET, &dispparamsNoArgs,
                                  &aVariant, NULL, NULL);
              if (hr != S_OK)
                log_debug ("%s:%s: retrieving ConversationIndex failed: %#lx",
                           SRCNAME, __func__, hr);
              else if (aVariant.vt != VT_BSTR)
                log_debug ("%s:%s: ConversationIndex is not a string (%d)",
                           SRCNAME, __func__, aVariant.vt);
              else if (aVariant.bstrVal)
                {
                  char *p;

                  key = wchar_to_utf8 (aVariant.bstrVal);
                  log_debug ("%s:%s: ConversationIndex is `%s'",
                           SRCNAME, __func__, key);
                  /* The key is a hex string.  Convert it to binary. */
                  for (keylen=0,p=key; hexdigitp(p) && hexdigitp(p+1); p += 2)
                    ((unsigned char*)key)[keylen++] = xtoi_2 (p);
                  
                  /* FIXME: Do we need to free the string returned in
                     AVARIANT?  Check at other places too. */
                }

              pDisp->Release();
              pDisp = NULL;
            }
          
          if (key && keylen
              && (body = msgcache_get (key, keylen, &refhandle)) 
              && (pDisp = find_outlook_property (pEECB, "Body", &dispid)))
            {
#if 1
              dispparams.cNamedArgs = 1;
              dispparams.rgdispidNamedArgs = &dispid_put;
              dispparams.cArgs = 1;
              dispparams.rgvarg = &aVariant;
              dispparams.rgvarg[0].vt = VT_LPWSTR;
              dispparams.rgvarg[0].bstrVal = utf8_to_wchar (body);
              hr = pDisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                                 DISPATCH_PROPERTYPUT, &dispparams,
                                 NULL, NULL, NULL);
              xfree (dispparams.rgvarg[0].bstrVal);
              log_debug ("%s:%s: PROPERTYPUT(body) result -> %#lx\n",
                         SRCNAME, __func__, hr);
#else
              log_debug ("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX");
              show_window_hierarchy (hWnd, 0);
#endif
              pDisp->Release();
              pDisp = NULL;
              
              /* Because we found the plaintext in the cache we can assume
                 that the orginal message has been encrypted and thus we
                 now set a flag to make sure that by default the reply
                 gets encrypted too. */
              force_encrypt = 1;
            }
          msgcache_unref (refhandle);
          xfree (key);
        }
      
      ul_release (pMessage);
      ul_release (pMDB);
    }



  /* XXX: factor out common code */
  if (m_lContext == EECONTEXT_READNOTEMESSAGE)
    {
      int nTBIndex;
      HWND hwndToolbar = NULL;

      if (opt.compat.auto_decrypt)
        watcher_set_callback_ctx ((void *)pEECB);
      pEECB->GetMenuPos (EECMDID_ToolsCustomizeToolbar, &hMenuTools,
                         NULL, NULL, 0);
      AppendMenu (hMenuTools, MF_SEPARATOR, 0, NULL);
	
      AppendMenu (hMenuTools, MF_BYPOSITION | MF_STRING,
                  *pnCommandIDBase, _("&Decrypt and verify message"));

      m_nCmdEncrypt = *pnCommandIDBase;
      (*pnCommandIDBase)++;
	
      for (nTBIndex = nTBECnt-1; nTBIndex > -1; --nTBIndex)
        {	
          if (EETBID_STANDARD == pTBEArray[nTBIndex].tbid)
            {
              hwndToolbar = pTBEArray[nTBIndex].hwnd;		
              m_nToolbarButtonID1 = pTBEArray[nTBIndex].itbbBase;
              pTBEArray[nTBIndex].itbbBase++;
              break;		
            }	
        }

      if (hwndToolbar)
        {
          TBADDBITMAP tbab;
          tbab.hInst = glob_hinst;
          tbab.nID = IDB_DECRYPT;
          m_nToolbarBitmap1 = SendMessage(hwndToolbar, TB_ADDBITMAP,
                                          1, (LPARAM)&tbab);
          m_nToolbarButtonID2 = pTBEArray[nTBIndex].itbbBase;
          pTBEArray[nTBIndex].itbbBase++;
        }
    }

  if (m_lContext == EECONTEXT_SENDNOTEMESSAGE) 
    {
      int nTBIndex;
      HWND hwndToolbar = NULL;

      pEECB->GetMenuPos(EECMDID_ToolsCustomizeToolbar, &hMenuTools,
                        NULL, NULL, 0);
      AppendMenu(hMenuTools, MF_SEPARATOR, 0, NULL);
	
      AppendMenu(hMenuTools, MF_STRING,
                 *pnCommandIDBase, _("GPG &encrypt message"));

      m_nCmdEncrypt = *pnCommandIDBase;
      (*pnCommandIDBase)++;

      AppendMenu(hMenuTools, MF_STRING,
                 *pnCommandIDBase, _("GPG &sign message"));

      m_nCmdSign = *pnCommandIDBase;
      (*pnCommandIDBase)++;

      for (nTBIndex = nTBECnt-1; nTBIndex > -1; --nTBIndex)
        {
          if (EETBID_STANDARD == pTBEArray[nTBIndex].tbid)
            {
              hwndToolbar = pTBEArray[nTBIndex].hwnd;
              m_nToolbarButtonID1 = pTBEArray[nTBIndex].itbbBase;
              pTBEArray[nTBIndex].itbbBase++;
              break;	
            }
        }

      if (hwndToolbar) 
        {
          TBADDBITMAP tbab;
          tbab.hInst = glob_hinst;
          tbab.nID = IDB_ENCRYPT;
          m_nToolbarBitmap1 = SendMessage (hwndToolbar, TB_ADDBITMAP,
                                           1, (LPARAM)&tbab);
          m_nToolbarButtonID2 = pTBEArray[nTBIndex].itbbBase;
          pTBEArray[nTBIndex].itbbBase++;
          tbab.nID = IDB_SIGN;
          m_nToolbarBitmap2 = SendMessage (hwndToolbar, TB_ADDBITMAP,
                                           1, (LPARAM)&tbab);
        }

      m_pExchExt->m_gpgEncrypt = opt.encrypt_default;
      m_pExchExt->m_gpgSign    = opt.sign_default;
      if (force_encrypt)
        m_pExchExt->m_gpgEncrypt = true;
    }

  if (m_lContext == EECONTEXT_VIEWER) 
    {
      int nTBIndex;
      HWND hwndToolbar = NULL;
      
      pEECB->GetMenuPos (EECMDID_ToolsCustomizeToolbar, &hMenuTools,
                         NULL, NULL, 0);
      AppendMenu (hMenuTools, MF_SEPARATOR, 0, NULL);
      
      AppendMenu (hMenuTools, MF_BYPOSITION | MF_STRING,
                  *pnCommandIDBase, _("GPG Key &Manager"));

      m_nCmdEncrypt = *pnCommandIDBase;
      (*pnCommandIDBase)++;	

      for (nTBIndex = nTBECnt-1; nTBIndex > -1; --nTBIndex)
        {
          if (EETBID_STANDARD == pTBEArray[nTBIndex].tbid) 
            {
              hwndToolbar = pTBEArray[nTBIndex].hwnd;
              m_nToolbarButtonID1 = pTBEArray[nTBIndex].itbbBase;
              pTBEArray[nTBIndex].itbbBase++;
              break;	
            }
        }
      if (hwndToolbar)
        {
          TBADDBITMAP tbab;
          tbab.hInst = glob_hinst;
          tbab.nID = IDB_KEY_MANAGER;
          m_nToolbarBitmap1 = SendMessage(hwndToolbar, TB_ADDBITMAP,
                                          1, (LPARAM)&tbab);
        }	
    }
  return S_FALSE;
}


/* Called by Exchange when a user selects a command.  Return value:
   S_OK if command is handled, otherwise S_FALSE. */
STDMETHODIMP 
CGPGExchExtCommands::DoCommand (
                  LPEXCHEXTCALLBACK pEECB, // The Exchange Callback Interface.
                  UINT nCommandID)         // The command id.
{
  HRESULT hr;

  log_debug ("%s:%s: commandID=%u (%#x)\n",
             SRCNAME, __func__, nCommandID, nCommandID);
  if (nCommandID == SC_CLOSE && m_lContext == EECONTEXT_READNOTEMESSAGE)
    {
      /* This is the system close command. Replace it with our own to
         avoid the "save changes" query, apparently induced by OL
         internal syncronisation of our SetWindowText message with its
         own OOM (in this case Body). */
      LPDISPATCH pDisp;
      DISPID dispid;
      DISPPARAMS dispparams;
      VARIANT aVariant;
      
      pDisp = find_outlook_property (pEECB, "Close", &dispid);
      if (pDisp)
        {
          dispparams.rgvarg = &aVariant;
          dispparams.rgvarg[0].vt = VT_INT;
          dispparams.rgvarg[0].intVal = 1; /* olDiscard */
          dispparams.cArgs = 1;
          dispparams.cNamedArgs = 0;
          hr = pDisp->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                              DISPATCH_METHOD, &dispparams,
                              NULL, NULL, NULL);
          pDisp->Release();
          pDisp = NULL;
          if (hr == S_OK)
            {
              log_debug ("%s:%s: invoking Close succeeded", SRCNAME,__func__);
              return S_OK; /* We handled the close command. */
            }

          log_debug ("%s:%s: invoking Close failed: %#lx",
                     SRCNAME, __func__, hr);
        }

      /* We are not interested in the close command - pass it on. */
      return S_FALSE; 
    }
  else if (nCommandID == 154)
    {
      log_debug ("%s:%s: command Reply called\n", SRCNAME, __func__);
      /* What we might want to do is to call Reply, then GetInspector
         and then Activate - this allows us to get full control over
         the quoted message and avoids the ugly msgcache. */
    }
  else if (nCommandID == 155)
    {
      log_debug ("%s:%s: command ReplyAll called\n", SRCNAME, __func__);
    }
  else if (nCommandID == 156)
    {
      log_debug ("%s:%s: command Forward called\n", SRCNAME, __func__);
    }
  

  if ((nCommandID != m_nCmdEncrypt) 
      && (nCommandID != m_nCmdSign))
    return S_FALSE; 

  if (m_lContext == EECONTEXT_READNOTEMESSAGE) 
    {
      HWND hWnd = NULL;
      LPMESSAGE pMessage = NULL;
      LPMDB pMDB = NULL;

      if (FAILED (pEECB->GetWindow (&hWnd)))
        hWnd = NULL;
//       else
//         show_window_hierarchy (hWnd, 0);

      hr = pEECB->GetObject (&pMDB, (LPMAPIPROP *)&pMessage);
      if (SUCCEEDED (hr))
        {
          if (nCommandID == m_nCmdEncrypt)
            {
              GpgMsg *m = CreateGpgMsg (pMessage);
              m->setExchangeCallback ((void*)pEECB);
              m->decrypt (hWnd, 0);
              delete m;
	    }
	}
      ul_release (pMessage);
      ul_release (pMDB);
    }
  else if (m_lContext == EECONTEXT_SENDNOTEMESSAGE) 
    {
      if (nCommandID == m_nCmdEncrypt)
        m_pExchExt->m_gpgEncrypt = !m_pExchExt->m_gpgEncrypt;
      if (nCommandID == m_nCmdSign)
        m_pExchExt->m_gpgSign = !m_pExchExt->m_gpgSign;
    }
  else if (m_lContext == EECONTEXT_VIEWER)
    {
      if (start_key_manager ())
        MessageBox (NULL, _("Could not start Key-Manager"),
                    "GPGol", MB_ICONERROR|MB_OK);
    }

  return S_OK; 
}

/* Called by Exchange when it receives a WM_INITMENU message, allowing
   the extension object to enable, disable, or update its menu
   commands before the user sees them. This method is called
   frequently and should be written in a very efficient manner. */
STDMETHODIMP_(VOID) 
CGPGExchExtCommands::InitMenu(LPEXCHEXTCALLBACK pEECB) 
{
#if 0
  log_debug ("%s:%s: context=0x%lx (%s)\n", SRCNAME, __func__, 
             m_lContext,
             (m_lContext == EECONTEXT_SESSION?           "Session"          :
              m_lContext == EECONTEXT_VIEWER?            "Viewer"           :
              m_lContext == EECONTEXT_REMOTEVIEWER?      "RemoteViewer"     :
              m_lContext == EECONTEXT_SEARCHVIEWER?      "SearchViewer"     :
              m_lContext == EECONTEXT_ADDRBOOK?          "AddrBook"         :
              m_lContext == EECONTEXT_SENDNOTEMESSAGE?   "SendNoteMessage"  :
              m_lContext == EECONTEXT_READNOTEMESSAGE?   "ReadNoteMessage"  :
              m_lContext == EECONTEXT_SENDPOSTMESSAGE?   "SendPostMessage"  :
              m_lContext == EECONTEXT_READPOSTMESSAGE?   "ReadPostMessage"  :
              m_lContext == EECONTEXT_READREPORTMESSAGE? "ReadReportMessage":
              m_lContext == EECONTEXT_SENDRESENDMESSAGE? "SendResendMessage":
              m_lContext == EECONTEXT_PROPERTYSHEETS?    "PropertySheets"   :
              m_lContext == EECONTEXT_ADVANCEDCRITERIA?  "AdvancedCriteria" :
              m_lContext == EECONTEXT_TASK?              "Task" : ""));
#endif
}


/* Called by Exhange when the user requests help for a menu item.
   Return value: S_OK when it is a menu item of this plugin and the
   help was shown; otherwise S_FALSE.  */
STDMETHODIMP 
CGPGExchExtCommands::Help (
	LPEXCHEXTCALLBACK pEECB, // The pointer to Exchange Callback Interface.
	UINT nCommandID)         // The command id.
{
    if (m_lContext == EECONTEXT_READNOTEMESSAGE) {
    	if (nCommandID == m_nCmdEncrypt) {
	    MessageBox (m_hWnd,
                        _("Decrypt and verify the message."),
                        "GPGol", MB_OK);
	    return S_OK;
	}
    }
    if (m_lContext == EECONTEXT_SENDNOTEMESSAGE) {
	if (nCommandID == m_nCmdEncrypt) {
	    MessageBox(m_hWnd,
                       _("Select this option to encrypt the message."),
                       "GPGol", MB_OK);	
	    return S_OK;
	} 
	else if (nCommandID == m_nCmdSign) {
	    MessageBox(m_hWnd,
                       _("Select this option to sign the message."),
                       "GPGol", MB_OK);	
	    return S_OK;
	} 
    }

    if (m_lContext == EECONTEXT_VIEWER) {
    	if (nCommandID == m_nCmdEncrypt) {
		MessageBox(m_hWnd, 
                           _("Open GPG Key Manager"),
                           "GPGol", MB_OK);
		return S_OK;
	} 
    }

    return S_FALSE;
}


/* Called by Exhange to get the status bar text or the tooltip of a
   menu item.  Returns S_OK when it is a menu item of this plugin and
   the text was set; otherwise S_FALSE. */
STDMETHODIMP 
CGPGExchExtCommands::QueryHelpText(
          UINT nCommandID,  // The command id corresponding to the
                            //  menu item activated.
	  ULONG lFlags,     // Identifies either EECQHT_STATUS
                            //  or EECQHT_TOOLTIP.
          LPTSTR pszText,   // A pointer to buffer to be populated 
                            //  with text to display.
	  UINT nCharCnt)    // The count of characters available in psz buffer.
{
	
    if (m_lContext == EECONTEXT_READNOTEMESSAGE) {
	if (nCommandID == m_nCmdEncrypt) {
	    if (lFlags == EECQHT_STATUS)
		lstrcpyn (pszText, ".", nCharCnt);
  	    if (lFlags == EECQHT_TOOLTIP)
		lstrcpyn (pszText,
                          _("Decrypt message and verify signature"),
                          nCharCnt);
	    return S_OK;
	}
    }
    if (m_lContext == EECONTEXT_SENDNOTEMESSAGE) {
	if (nCommandID == m_nCmdEncrypt) {
	    if (lFlags == EECQHT_STATUS)
		lstrcpyn (pszText, ".", nCharCnt);
	    if (lFlags == EECQHT_TOOLTIP)
		lstrcpyn (pszText,
                          _("Encrypt message with GPG"),
                          nCharCnt);
	    return S_OK;
	}
	if (nCommandID == m_nCmdSign) {
	    if (lFlags == EECQHT_STATUS)
		lstrcpyn (pszText, ".", nCharCnt);
  	    if (lFlags == EECQHT_TOOLTIP)
		lstrcpyn (pszText,
                          _("Sign message with GPG"),
                          nCharCnt);
	    return S_OK;
	}
    }
    if (m_lContext == EECONTEXT_VIEWER) {
	if (nCommandID == m_nCmdEncrypt) {
	    if (lFlags == EECQHT_STATUS)
		lstrcpyn (pszText, ".", nCharCnt);
	    if (lFlags == EECQHT_TOOLTIP)
		lstrcpyn (pszText,
                          _("Open GPG Key Manager"),
                          nCharCnt);
	    return S_OK;
	}	
    }
    return S_FALSE;
}


/* Called by Exchange to get toolbar button infos.  Returns S_OK when
   it is a button of this plugin and the requested info was delivered;
   otherwise S_FALSE. */
STDMETHODIMP 
CGPGExchExtCommands::QueryButtonInfo (
	ULONG lToolbarID,       // The toolbar identifier.
	UINT nToolbarButtonID,  // The toolbar button index.
        LPTBBUTTON pTBB,        // A pointer to toolbar button structure
                                //  (see TBBUTTON structure).
	LPTSTR lpszDescription, // A pointer to string describing button.
	UINT nCharCnt,          // The maximum size of lpsz buffer.
        ULONG lFlags)           // EXCHEXT_UNICODE may be specified
{
	if (m_lContext == EECONTEXT_READNOTEMESSAGE)
	{
		if (nToolbarButtonID == m_nToolbarButtonID1)
		{
			pTBB->iBitmap = m_nToolbarBitmap1;             
			pTBB->idCommand = m_nCmdEncrypt;
			pTBB->fsState = TBSTATE_ENABLED;
			pTBB->fsStyle = TBSTYLE_BUTTON;
			pTBB->dwData = 0;
			pTBB->iString = -1;
			lstrcpyn (lpszDescription,
                                  _("Decrypt message and verify signature"),
                                  nCharCnt);
			return S_OK;
		}
	}
	if (m_lContext == EECONTEXT_SENDNOTEMESSAGE)
	{
		if (nToolbarButtonID == m_nToolbarButtonID1)
		{
			pTBB->iBitmap = m_nToolbarBitmap1;             
			pTBB->idCommand = m_nCmdEncrypt;
			pTBB->fsState = TBSTATE_ENABLED;
			if (m_pExchExt->m_gpgEncrypt)
				pTBB->fsState |= TBSTATE_CHECKED;
			pTBB->fsStyle = TBSTYLE_BUTTON | TBSTYLE_CHECK;
			pTBB->dwData = 0;
			pTBB->iString = -1;
			lstrcpyn (lpszDescription,
                                  _("Encrypt message with GPG"),
                                  nCharCnt);
			return S_OK;
		}
		if (nToolbarButtonID == m_nToolbarButtonID2)
		{
			pTBB->iBitmap = m_nToolbarBitmap2;             
			pTBB->idCommand = m_nCmdSign;
			pTBB->fsState = TBSTATE_ENABLED;
			if (m_pExchExt->m_gpgSign)
				pTBB->fsState |= TBSTATE_CHECKED;
			pTBB->fsStyle = TBSTYLE_BUTTON | TBSTYLE_CHECK;
			pTBB->dwData = 0;
			pTBB->iString = -1;
			lstrcpyn (lpszDescription,
                                  _("Sign message with GPG"),
                                  nCharCnt);
			return S_OK;
		}
	}
	if (m_lContext == EECONTEXT_VIEWER)
	{
		if (nToolbarButtonID == m_nToolbarButtonID1)
		{
			pTBB->iBitmap = m_nToolbarBitmap1;             
			pTBB->idCommand = m_nCmdEncrypt;
			pTBB->fsState = TBSTATE_ENABLED;
			pTBB->fsStyle = TBSTYLE_BUTTON;
			pTBB->dwData = 0;
			pTBB->iString = -1;
			lstrcpyn (lpszDescription,
                                  _("Open GPG Key Manager"),
                                  nCharCnt);
			return S_OK;
		}
	}

	return S_FALSE;
}



STDMETHODIMP 
CGPGExchExtCommands::ResetToolbar (ULONG lToolbarID, ULONG lFlags)
{	
    return S_OK;
}


CGPGExchExtPropertySheets::CGPGExchExtPropertySheets (
                                    CGPGExchExt* pParentInterface)
{ 
    m_pExchExt = pParentInterface;
    m_lRef = 0; 
}


STDMETHODIMP 
CGPGExchExtPropertySheets::QueryInterface(REFIID riid, LPVOID FAR * ppvObj)
{   
    *ppvObj = NULL;
    if (riid == IID_IExchExtPropertySheets) {
        *ppvObj = (LPVOID)this;
        AddRef();
        return S_OK;
    }
    if (riid == IID_IUnknown) {
        *ppvObj = (LPVOID)m_pExchExt;
        m_pExchExt->AddRef();
        return S_OK;
    }

    return E_NOINTERFACE;
}


/* Called by Echange to get the maximum number of property pages which
   are to be added. LFLAGS is a bitmask indicating what type of
   property sheet is being displayed. Return value: The maximum number
   of custom pages for the property sheet.  */
STDMETHODIMP_ (ULONG) 
CGPGExchExtPropertySheets::GetMaxPageCount(ULONG lFlags)
{
    if (lFlags == EEPS_TOOLSOPTIONS)
	return 1;	
    return 0;
}


/* Called by Exchange to request information about the property page.
   Return value: S_OK to signal Echange to use the pPSP
   information. */
STDMETHODIMP 
CGPGExchExtPropertySheets::GetPages(
	LPEXCHEXTCALLBACK pEECB, // A pointer to Exchange callback interface.
	ULONG lFlags,            // A  bitmask indicating what type of
                                 //  property sheet is being displayed.
	LPPROPSHEETPAGE pPSP,    // The output parm pointing to pointer
                                 //  to list of property sheets.
	ULONG FAR * plPSP)       // The output parm pointing to buffer 
                                 //  containing the number of property 
                                 //  sheets actually used.
{
  int resid ;

  if (!strncmp (gettext_localename (), "de", 2))
    resid = IDD_GPG_OPTIONS_DE;
  else
    resid = IDD_GPG_OPTIONS;

  pPSP[0].dwSize = sizeof (PROPSHEETPAGE);
  pPSP[0].dwFlags = PSP_DEFAULT | PSP_HASHELP;
  pPSP[0].hInstance = glob_hinst;
  pPSP[0].pszTemplate = MAKEINTRESOURCE (resid);
  pPSP[0].hIcon = NULL;     
  pPSP[0].pszTitle = NULL;  
  pPSP[0].pfnDlgProc = (DLGPROC) GPGOptionsDlgProc;
  pPSP[0].lParam = 0;     
  pPSP[0].pfnCallback = NULL;
  pPSP[0].pcRefParent = NULL; 

  *plPSP = 1;

  return S_OK;
}


STDMETHODIMP_ (VOID) 
CGPGExchExtPropertySheets::FreePages (LPPROPSHEETPAGE pPSP,
                                      ULONG lFlags, ULONG lPSP)
{
}




