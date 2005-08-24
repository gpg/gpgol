/* olflange.cpp - Flange between Outlook and the MapiGPGME class
 *	Copyright (C) 2001 G Data Software AG, http://www.gdata.de
 *	Copyright (C) 2004, 2005 g10 Code GmbH
 * 
 * This file is part of OutlGPG.
 * 
 * OutlGPG is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2 of the License, or (at your option) any later version.
 * 
 * OutlGPG is distributed in the hope that it will be useful,
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
#include "MapiGPGME.h"
#include "intern.h"
#include "gpgmsg.hh"
#include "msgcache.h"

#include "olflange-ids.h"
#include "olflange-def.h"
#include "olflange.h"


#define TRACEPOINT() do { ExchLogInfo ("%s:%s:%d: tracepoint\n", \
                                       __FILE__, __func__, __LINE__); \
                        } while (0)


bool g_bInitDll = FALSE;


/* FIXME!!!! Huh?  We only have one m_gpg object??? This is strange,
   AFAICS, Exchange may create several contexts and thus we may be
   required to run several instances of mapiGPGME concurrently. */
MapiGPGME *m_gpg = NULL;




/* Registers this module as an Exchange extension. This basically updates
   some Registry entries. */
STDAPI 
DllRegisterServer (void)
{    
    HKEY hkey;
    CHAR szKeyBuf[1024];
    CHAR szEntry[512];
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
      10 	EECONTEXT_READREPORTMESSAGE
      11 	EECONTEXT_SENDRESENDMESSAGE
      12 	EECONTEXT_PROPERTYSHEETS
      13 	EECONTEXT_ADVANCEDCRITERIA
      14 	EECONTEXT_TASK
    */
    lstrcat (szEntry, ";11000111111100"); 
    ec = RegCreateKeyEx (HKEY_LOCAL_MACHINE, szKeyBuf, 0, NULL, 
				   REG_OPTION_NON_VOLATILE,
				   KEY_ALL_ACCESS, NULL, &hkey, NULL);
    if (ec != ERROR_SUCCESS) {
	log_debug ("DllRegisterServer failed\n");
	return E_ACCESSDENIED;
    }
    
    dwTemp = lstrlen (szEntry) + 1;
    RegSetValueEx (hkey, "OutlGPG", 0, REG_SZ, (BYTE*) szEntry, dwTemp);

    /* To avoid conflicts with the old G-DATA plugin and older vesions
       of this Plugin, we remove the key used by these versions. */
    RegDeleteValue (hkey, "GPG Exchange");

    /* Set outlook update flag. */
    strcpy (szEntry, "4.0;Outxxx.dll;7;000000000000000;0000000000;OutXXX");
    dwTemp = lstrlen (szEntry) + 1;
    RegSetValueEx (hkey, "Outlook Setup Extension", 0, REG_SZ, (BYTE*) szEntry, dwTemp);
    RegCloseKey (hkey);
    
    hkey = NULL;
    lstrcpy (szKeyBuf, "Software\\GNU\\OutlGPG");
    RegCreateKeyEx (HKEY_CURRENT_USER, szKeyBuf, 0, NULL,
		    REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hkey, NULL);
    if (hkey != NULL)
	RegCloseKey (hkey);

    log_debug ("DllRegisterServer succeeded\n");
    return S_OK;
}


/* Unregisters this module as an Exchange extension. */
STDAPI 
DllUnregisterServer (void)
{
    HKEY hkey;
    CHAR szKeyBuf[1024];

    lstrcpy(szKeyBuf, "Software\\Microsoft\\Exchange\\Client\\Extensions");
    /* create and open key and subkey */
    long lResult = RegCreateKeyEx(HKEY_LOCAL_MACHINE, szKeyBuf, 0, NULL, 
				    REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 
				    NULL, &hkey, NULL);
    if (lResult != ERROR_SUCCESS) {
	log_debug ("DllUnregisterServer: access denied.\n");
	return E_ACCESSDENIED;
    }
    RegDeleteValue (hkey, "OutlGPG");
    /* set outlook update flag */
    CHAR szEntry[512];
    strcpy (szEntry, "4.0;Outxxx.dll;7;000000000000000;0000000000;OutXXX");
    DWORD dwTemp = lstrlen (szEntry) + 1;
    RegSetValueEx (hkey, "Outlook Setup Extension", 0, REG_SZ, (BYTE*) szEntry, dwTemp);
    RegCloseKey (hkey);

    return S_OK;
}


/* DISPLAY a MAPI property. */
static void
show_mapi_property (LPMESSAGE message, ULONG prop, const char *propname)
{
  HRESULT hr;
  LPSPropValue lpspvFEID = NULL;
  cache_item_t item;
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

  log_debug ("%s:%s: looking for `%s'\n", __FILE__, __func__, name);

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
      //           __FILE__, __func__, (int)(s-name), name, pObj, hr);
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
      
  hr = pDisp->GetIDsOfNames(IID_NULL, &wname, 1,
                            LOCALE_SYSTEM_DEFAULT, &dispid);
  xfree (wname);
  //log_debug ("   dispid(%s)=%d  (hr=0x%x)\n", name, dispid, hr);
  if (r_dispid)
    *r_dispid = dispid;

  log_debug ("%s:%s:    got IDispatch=%p dispid=%d\n",
               __FILE__, __func__, pDisp, dispid);
  return pDisp;
}


/* Return Outlook's Application object. */
/* FIXME: We should be able to fold most of the code into
   find_outlook_property. */
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
                   __FILE__, __func__, vtResult.intVal);
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
      //             __FILE__, __func__, pUnk);
      pDisp->Release();
      pDisp = NULL;
    }
  return pUnk;
}



/* The entry point which Exchange calls.  This is called for each
   context entry. Creates a new CGPGExchExt object every time so each
   context will get its own CGPGExchExt interface. */
EXTERN_C LPEXCHEXT __stdcall
ExchEntryPoint (void)
{
  log_debug ("%s:%s: creating new CGPGExchExt object\n", __FILE__, __func__);
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
  if (!m_pExchExtMessageEvents || !m_pExchExtCommands
      || !m_pExchExtPropertySheets)
    out_of_core ();
  
  if (!g_bInitDll)
    {
      if (!m_gpg)
        {
          m_gpg = CreateMapiGPGME ();
          m_gpg->readOptions ();
        }
      g_bInitDll = TRUE;
      log_debug ("%s:%s: one time init done\n", __FILE__, __func__);
    }
}


/*  Uninitializes the dll in the session context. 
   FIXME:  One instance only????  For what the hell do we use g_bInitDll then?
*/
CGPGExchExt::~CGPGExchExt (void) 
{
  log_debug ("%s:%s: cleaning up CGPGExchExt object\n", __FILE__, __func__);

    if (m_lContext == EECONTEXT_SESSION) {
	if (g_bInitDll) {
	    if (m_gpg != NULL) {
		m_gpg->writeOptions ();
		delete m_gpg;
		m_gpg = NULL;
	    }
	    g_bInitDll = FALSE;
            log_debug ("%s:%s: one time deinit done\n", __FILE__, __func__);
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

    if ((riid == IID_IUnknown) || (riid == IID_IExchExt)) {
        *ppvObj = (LPUNKNOWN) this;
    }
    else if (riid == IID_IExchExtMessageEvents) {
        *ppvObj = (LPUNKNOWN) m_pExchExtMessageEvents;
        m_pExchExtMessageEvents->SetContext (m_lContext);
    }
    else if (riid == IID_IExchExtCommands) {
        *ppvObj = (LPUNKNOWN)m_pExchExtCommands;
        m_pExchExtCommands->SetContext (m_lContext);
    }
    else if (riid == IID_IExchExtPropertySheets) {
	if (m_lContext != EECONTEXT_PROPERTYSHEETS)
	    return E_NOINTERFACE;
        *ppvObj = (LPUNKNOWN) m_pExchExtPropertySheets;
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

  /* Save the context in an instance variable. */
  m_lContext = lContext;

  log_debug ("%s:%s: context=0x%lx (%s) flags=0x%lx\n", __FILE__, __func__, 
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
  pEECB->GetVersion (&lBuildVersion, EECBGV_GETBUILDVERSION);
  if (EECBGV_BUILDVERSION_MAJOR
      != (lBuildVersion & EECBGV_BUILDVERSION_MAJOR_MASK))
    {
      log_debug ("%s:%s: invalid version 0x%lx\n",
                   __FILE__, __func__, lBuildVersion);
      return S_FALSE;
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
//                    __FILE__, __func__, pApplication);
      return S_OK;
    }
  
  log_debug ("%s:%s: can't handle this context\n", __FILE__, __func__);
  return S_FALSE;
}



CGPGExchExtMessageEvents::CGPGExchExtMessageEvents 
                                              (CGPGExchExt *pParentInterface)
{ 
  m_pExchExt = pParentInterface;
  m_lRef = 0; 
  m_bOnSubmitActive = FALSE;
};


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

  log_debug ("%s:%s: received\n", __FILE__, __func__);

  pEECB->GetObject (&pMDB, (LPMAPIPROP *)&pMessage);
  show_mapi_property (pMessage, PR_SEARCH_KEY, "SEARCH_KEY");
  msgcache_set_active (pMessage);
  if (pMessage)
    UlRelease(pMessage);
  if (pMDB)
    UlRelease(pMDB);
  return S_FALSE;
}


/* Called by Exchange after a message has been read.  Returns: S_FALSE
   to signal Exchange to continue calling extensions.  PEECB is a
   pointer to the IExchExtCallback interface. LFLAGS are some flags. */
STDMETHODIMP 
CGPGExchExtMessageEvents::OnReadComplete (LPEXCHEXTCALLBACK pEECB,
                                          ULONG lFlags)
{
  log_debug ("%s:%s: received\n", __FILE__, __func__);
  return S_FALSE;
}


/* Called by Exchange when a message will be written. Returns: S_FALSE
   to signal Exchange to continue calling extensions.  PEECB is a
   pointer to the IExchExtCallback interface. */
STDMETHODIMP 
CGPGExchExtMessageEvents::OnWrite (LPEXCHEXTCALLBACK pEECB)
{
    log_debug ("%s:%s: received\n", __FILE__, __func__);
    return S_FALSE;
}


/* Called by Exchange when the data has been written to the message.
   Encrypts and signs the message if the options are set.  PEECB is a
   pointer to the IExchExtCallback interface.  Returns: S_FALSE to
   signal Exchange to continue calling extensions.  E_FAIL to signals
   Exchange an error; the message will not be sent */
STDMETHODIMP 
CGPGExchExtMessageEvents::OnWriteComplete (LPEXCHEXTCALLBACK pEECB,
                                           ULONG lFlags)
{
  log_debug ("%s:%s: received\n", __FILE__, __func__);

  HRESULT hrReturn = S_FALSE;
  LPMESSAGE msg = NULL;
  LPMDB pMDB = NULL;
  HWND hWnd = NULL;
  int rc;
          
  if (FAILED(pEECB->GetWindow (&hWnd)))
    hWnd = NULL;

  if (!m_bOnSubmitActive) /* the user is just saving the message */
    return S_FALSE;
  
  if (m_bWriteFailed)     /* operation failed already */
    return S_FALSE;
  
  HRESULT hr = pEECB->GetObject (&pMDB, (LPMAPIPROP *)&msg);
  if (SUCCEEDED (hr))
    {
      log_debug ("%s:%s:%d: here\n", __FILE__, __func__, __LINE__);
      
      GpgMsg *m = CreateGpgMsg (msg);
      log_debug ("%s:%s:%d: here\n", __FILE__, __func__, __LINE__);
      if (m_pExchExt->m_gpgEncrypt && m_pExchExt->m_gpgSign)
        rc = m_gpg->signEncrypt (hWnd, m);
      if (m_pExchExt->m_gpgEncrypt && !m_pExchExt->m_gpgSign)
        rc = m_gpg->encrypt (hWnd, m);
      if (!m_pExchExt->m_gpgEncrypt && m_pExchExt->m_gpgSign)
        rc = m_gpg->sign (hWnd, m);
      else
        rc = 0;
      log_debug ("%s:%s:%d: here\n", __FILE__, __func__, __LINE__);
      delete m;
      
      if (rc)
        {
          hrReturn = E_FAIL;
          m_bWriteFailed = TRUE;	
        }
    }

  if (msg)
    UlRelease(msg);
  if (pMDB) 
    UlRelease(pMDB);

  return hrReturn;
}

/* Called by Exchange when the user selects the "check names" command.
   PEECB is a pointer to the IExchExtCallback interface.  Returns
   S_FALSE to signal Exchange to continue calling extensions. */
STDMETHODIMP 
CGPGExchExtMessageEvents::OnCheckNames(LPEXCHEXTCALLBACK pEECB)
{
  log_debug ("%s:%s: received\n", __FILE__, __func__);
  return S_FALSE;
}


/* Called by Exchange when "check names" command is complete.
   PEECB is a pointer to the IExchExtCallback interface.  Returns
   S_FALSE to signal Exchange to continue calling extensions. */
STDMETHODIMP 
CGPGExchExtMessageEvents::OnCheckNamesComplete (LPEXCHEXTCALLBACK pEECB,
                                                ULONG lFlags)
{
  log_debug ("%s:%s: received\n", __FILE__, __func__);
  return S_FALSE;
}


/* Called by Exchange before the message data will be written and
   submitted to MAPI.  PEECB is a pointer to the IExchExtCallback
   interface.  Returns S_FALSE to signal Exchange to continue calling
   extensions. */
STDMETHODIMP 
CGPGExchExtMessageEvents::OnSubmit (LPEXCHEXTCALLBACK pEECB)
{
  log_debug ("%s:%s: received\n", __FILE__, __func__);
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
  log_debug ("%s:%s: received\n", __FILE__, __func__);
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
};



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



// We can't read the Body object because it would fire up the
// security pop-up.  Writing is okay.
//       vtResult.pdispVal = NULL;
//       pDisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
//                     DISPATCH_PROPERTYGET, &dispparamsNoArgs,
//                     &vtResult, NULL, NULL);

//       log_debug ("%s:%s: Body=%p (%s)\n", __FILE__, __func__, 
//                    vtResult.pbVal,
//                    (tmp = wchar_to_utf8 ((wchar_t*)vtResult.pbVal)));



// XXX IExchExtSessionEvents::OnDelivery: could be used to automatically decrypt new mails
// when they arrive

/* Called by Echange to install commands and toolbar buttons.  Returns
    S_FALSE to signal Exchange to continue calling extensions. */
STDMETHODIMP 
CGPGExchExtCommands::InstallCommands (
	LPEXCHEXTCALLBACK pEECB, // The Exchange Callback Interface.
	HWND hWnd,               // The window handle to the main window
                                 // of context.
	HMENU hMenu,             // The menu handle to main menu of context.
	UINT FAR * pnCommandIDBase,  // The base conmmand id.
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
  
  log_debug ("%s:%s: context=0x%lx (%s) flags=0x%lx\n", __FILE__, __func__, 
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


  if (m_lContext == EECONTEXT_READNOTEMESSAGE
      || m_lContext == EECONTEXT_SENDNOTEMESSAGE)
    {
      LPMDB pMDB = NULL;
      LPMESSAGE pMessage = NULL;

      pEECB->GetObject (&pMDB, (LPMAPIPROP *)&pMessage);
      show_mapi_property (pMessage, PR_ENTRYID, "ENTRYID");
      show_mapi_property (pMessage, PR_SEARCH_KEY, "SEARCH_KEY");
      if (pMessage)
        UlRelease(pMessage);
      if (pMDB)
        UlRelease(pMDB);
    }

  /* Outlook 2003 sometimes displays the plaintext sometimes the
     orginal undecrypted text when doing a Reply.  This seems to
     depend on the sieze of the message; my guess it that only short
     messages are locally saved in the process and larger ones are
     fetyched again from the backend - or the other way around.
     Anyway, we can't rely on that and thus me make sure to update the
     Body object right here with our own copy of the plaintext.  To
     match the text we use the Storage ID Property of MAPI.  

     Unfortunately there seems to be no way of resetting the Saved
     property after updating the body, thus even without entering a
     single byte the user will be asked when cancelling a reply
     whether he really wants to do that.  

     Note, that we can't optimize the code here by first reading the
     body becuase this would pop up the securiy window, telling tghe
     user that someone is trying to read these data.
  */
  if (m_lContext == EECONTEXT_SENDNOTEMESSAGE)
    {
      LPMDB pMDB = NULL;
      LPMESSAGE pMessage = NULL;
      const char *body;
      void *refhandle = NULL;
      
      /*  Note that for read and send the object returned by the
          outlook extension callback is of class 43 (MailItem) so we
          only need to ask for Body then. */
      hr = pEECB->GetObject (&pMDB, (LPMAPIPROP *)&pMessage);
      if (FAILED(hr))
        log_debug ("%s:%s: getObject failed: hr=%#x\n", hr);
      else if ( (body = msgcache_get (pMessage, &refhandle)) 
                && (pDisp = find_outlook_property (pEECB, "Body", &dispid)))
        {
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
          log_debug ("%s:%s: PROPERTYPUT(body) result -> %d\n",
                       __FILE__, __func__, hr);

          pDisp->Release();
          pDisp = NULL;

          /* Because we found the plaintext in the cache we can assume
             that the orginal message has been encrypted and thus we
             now set a flag to make sure that by default the reply
             gets encrypted too. */
          force_encrypt = 1;
        }
      msgcache_unref (refhandle);
      if (pMessage)
        UlRelease(pMessage);
      if (pMDB)
        UlRelease(pMDB);
    }



  /* XXX: factor out common code */
  if (m_lContext == EECONTEXT_READNOTEMESSAGE) {
    int nTBIndex;
    HWND hwndToolbar = NULL;
    CHAR szBuffer[128];

    pEECB->GetMenuPos (EECMDID_ToolsCustomizeToolbar, &hMenuTools,
                       NULL, NULL, 0);
    AppendMenu (hMenuTools, MF_SEPARATOR, 0, NULL);
	
    LoadString (glob_hinst, IDS_DECRYPT_MENU_ITEM, szBuffer, 128);
    AppendMenu (hMenuTools, MF_BYPOSITION | MF_STRING,
                *pnCommandIDBase, szBuffer);

    m_nCmdEncrypt = *pnCommandIDBase;
    (*pnCommandIDBase)++;
	
    for (nTBIndex = nTBECnt-1; nTBIndex > -1; --nTBIndex) {	
      if (EETBID_STANDARD == pTBEArray[nTBIndex].tbid) {
        hwndToolbar = pTBEArray[nTBIndex].hwnd;		
        m_nToolbarButtonID1 = pTBEArray[nTBIndex].itbbBase;
        pTBEArray[nTBIndex].itbbBase++;
        break;		
      }	
    }

    if (hwndToolbar) {
      TBADDBITMAP tbab;
      tbab.hInst = glob_hinst;
      tbab.nID = IDB_DECRYPT;
      m_nToolbarBitmap1 = SendMessage(hwndToolbar, TB_ADDBITMAP,
                                      1, (LPARAM)&tbab);
      m_nToolbarButtonID2 = pTBEArray[nTBIndex].itbbBase;
      pTBEArray[nTBIndex].itbbBase++;
    }
  }

  if (m_lContext == EECONTEXT_SENDNOTEMESSAGE) {
    CHAR szBuffer[128];
    int nTBIndex;
    HWND hwndToolbar = NULL;

    pEECB->GetMenuPos(EECMDID_ToolsCustomizeToolbar, &hMenuTools,
                      NULL, NULL, 0);
    AppendMenu(hMenuTools, MF_SEPARATOR, 0, NULL);
	
    LoadString(glob_hinst, IDS_ENCRYPT_MENU_ITEM, szBuffer, 128);
    AppendMenu(hMenuTools, MF_BYPOSITION | MF_STRING,
               *pnCommandIDBase, szBuffer);

    m_nCmdEncrypt = *pnCommandIDBase;
    (*pnCommandIDBase)++;

    LoadString(glob_hinst, IDS_SIGN_MENU_ITEM, szBuffer, 128);
    AppendMenu(hMenuTools, MF_BYPOSITION | MF_STRING,
               *pnCommandIDBase, szBuffer);

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

    if (hwndToolbar) {
      TBADDBITMAP tbab;
      tbab.hInst = glob_hinst;
      tbab.nID = IDB_ENCRYPT;
      m_nToolbarBitmap1 = SendMessage(hwndToolbar, TB_ADDBITMAP,
                                      1, (LPARAM)&tbab);
      m_nToolbarButtonID2 = pTBEArray[nTBIndex].itbbBase;
      pTBEArray[nTBIndex].itbbBase++;
      tbab.nID = IDB_SIGN;
      m_nToolbarBitmap2 = SendMessage(hwndToolbar, TB_ADDBITMAP,
                                      1, (LPARAM)&tbab);
    }
    m_pExchExt->m_gpgEncrypt = m_gpg->getEncryptDefault ();
    m_pExchExt->m_gpgSign = m_gpg->getSignDefault ();
    if (force_encrypt)
      m_pExchExt->m_gpgEncrypt = true;
  }

  if (m_lContext == EECONTEXT_VIEWER) {
    CHAR szBuffer[128];
    int nTBIndex;
    HWND hwndToolbar = NULL;

    pEECB->GetMenuPos (EECMDID_ToolsCustomizeToolbar, &hMenuTools,
                       NULL, NULL, 0);
    AppendMenu (hMenuTools, MF_SEPARATOR, 0, NULL);
	
    LoadString (glob_hinst, IDS_KEY_MANAGER, szBuffer, 128);
    AppendMenu (hMenuTools, MF_BYPOSITION | MF_STRING,
                *pnCommandIDBase, szBuffer);

    m_nCmdEncrypt = *pnCommandIDBase;
    (*pnCommandIDBase)++;	

    for (nTBIndex = nTBECnt-1; nTBIndex > -1; --nTBIndex) {
      if (EETBID_STANDARD == pTBEArray[nTBIndex].tbid) {
        hwndToolbar = pTBEArray[nTBIndex].hwnd;
        m_nToolbarButtonID1 = pTBEArray[nTBIndex].itbbBase;
        pTBEArray[nTBIndex].itbbBase++;
        break;	
      }
    }
    if (hwndToolbar) {
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
      hr = pEECB->GetObject (&pMDB, (LPMAPIPROP *)&pMessage);
      if (SUCCEEDED (hr))
        {
          show_mapi_property (pMessage, PR_SEARCH_KEY, "SEARCH_KEY");
          if (nCommandID == m_nCmdEncrypt)
            {
              GpgMsg *m = CreateGpgMsg (pMessage);
              m_gpg->decrypt (hWnd, m);
              delete m;
	    }
	}
      if (pMessage)
        UlRelease(pMessage);
      if (pMDB)
        UlRelease(pMDB);
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
      if (m_gpg->startKeyManager ())
        MessageBox (NULL, "Could not start Key-Manager",
                    "OutlGPG", MB_ICONERROR|MB_OK);
    }

  return S_OK; 
}


STDMETHODIMP_(VOID) 
CGPGExchExtCommands::InitMenu(LPEXCHEXTCALLBACK pEECB) 
{
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
	    CHAR szBuffer[512];
	    CHAR szAppName[128];

	    LoadString (glob_hinst, IDS_DECRYPT_HELP, szBuffer, 511);
	    LoadString (glob_hinst, IDS_APP_NAME, szAppName, 127);
	    MessageBox (m_hWnd, szBuffer, szAppName, MB_OK);
	    return S_OK;
	}
    }
    if (m_lContext == EECONTEXT_SENDNOTEMESSAGE) {
	if (nCommandID == m_nCmdEncrypt) {
	    CHAR szBuffer[512];
	    CHAR szAppName[128];
	    LoadString(glob_hinst, IDS_ENCRYPT_HELP, szBuffer, 511);
	    LoadString(glob_hinst, IDS_APP_NAME, szAppName, 127);
	    MessageBox(m_hWnd, szBuffer, szAppName, MB_OK);	
	    return S_OK;
	} 
	if (nCommandID == m_nCmdSign) {
	    CHAR szBuffer[512];	
	    CHAR szAppName[128];	
	    LoadString(glob_hinst, IDS_SIGN_HELP, szBuffer, 511);	
	    LoadString(glob_hinst, IDS_APP_NAME, szAppName, 127);	
	    MessageBox(m_hWnd, szBuffer, szAppName, MB_OK);	
	    return S_OK;
	} 
    }

    if (m_lContext == EECONTEXT_VIEWER) {
    	if (nCommandID == m_nCmdEncrypt) {
		CHAR szBuffer[512];
		CHAR szAppName[128];
		LoadString(glob_hinst, IDS_KEY_MANAGER_HELP, szBuffer, 511);
		LoadString(glob_hinst, IDS_APP_NAME, szAppName, 127);
		MessageBox(m_hWnd, szBuffer, szAppName, MB_OK);
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
		LoadString (glob_hinst, IDS_DECRYPT_STATUSBAR,
                            pszText, nCharCnt);
  	    if (lFlags == EECQHT_TOOLTIP)
		LoadString (glob_hinst, IDS_DECRYPT_TOOLTIP,
                            pszText, nCharCnt);
	    return S_OK;
	}
    }
    if (m_lContext == EECONTEXT_SENDNOTEMESSAGE) {
	if (nCommandID == m_nCmdEncrypt) {
	    if (lFlags == EECQHT_STATUS)
		LoadString (glob_hinst, IDS_ENCRYPT_STATUSBAR,
                            pszText, nCharCnt);
	    if (lFlags == EECQHT_TOOLTIP)
		LoadString (glob_hinst, IDS_ENCRYPT_TOOLTIP,
                            pszText, nCharCnt);
	    return S_OK;
	}
	if (nCommandID == m_nCmdSign) {
	    if (lFlags == EECQHT_STATUS)
		LoadString (glob_hinst, IDS_SIGN_STATUSBAR, pszText, nCharCnt);
  	    if (lFlags == EECQHT_TOOLTIP)
	        LoadString (glob_hinst, IDS_SIGN_TOOLTIP, pszText, nCharCnt);
	    return S_OK;
	}
    }
    if (m_lContext == EECONTEXT_VIEWER) {
	if (nCommandID == m_nCmdEncrypt) {
	    if (lFlags == EECQHT_STATUS)
		LoadString (glob_hinst, IDS_KEY_MANAGER_STATUSBAR,
                            pszText, nCharCnt);
	    if (lFlags == EECQHT_TOOLTIP)
		LoadString (glob_hinst, IDS_KEY_MANAGER_TOOLTIP,
                            pszText, nCharCnt);
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
			LoadString(glob_hinst, IDS_DECRYPT_TOOLTIP,
                                   lpszDescription, nCharCnt);
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
			LoadString(glob_hinst, IDS_ENCRYPT_TOOLTIP,
                                   lpszDescription, nCharCnt);
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
			LoadString(glob_hinst, IDS_SIGN_TOOLTIP,
                                   lpszDescription, nCharCnt);
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
			LoadString(glob_hinst, IDS_KEY_MANAGER_TOOLTIP,
                                   lpszDescription, nCharCnt);
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
    int resid = 0;

    switch (GetUserDefaultLangID ()) {
    case 0x0407:    resid = IDD_GPG_OPTIONS_DE;break;
    default:	    resid = IDD_GPG_OPTIONS; break;
    }

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




