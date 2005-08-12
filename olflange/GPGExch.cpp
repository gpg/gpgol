/* GPGExch.cpp - exchange extension classes
 *	Copyright (C) 2001 G Data Software AG, http://www.gdata.de
 *	Copyright (C) 2004, 2005 g10 Code GmbH
 * 
 * This file is part of the G DATA Outlook Plugin for GnuPG.
 * 
 * This plugin is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2 of the License, or (at your option) any later version.
 * 
 * This plugin is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Lesser General Public License for more details.
 * 
 * You should have received a copy of the GNU Lesser General
 * Public License along with this plugin; if not, write to the 
 * Free Software Foundation, Inc., 59 Temple Place - Suite 330, 
 * Boston, MA 02111-1307, USA.
 */

#include "stdafx.h"

#ifndef INITGUID
#define INITGUID
#endif
#include <INITGUID.H>
#include <MAPIGUID.H>
#include <EXCHEXT.H>

#include "GPGExchange.h"
#include "GPGExch.h"
#include "../src/MapiGPGME.h"

BOOL g_bInitDll = FALSE;
MapiGPGME *m_gpg = NULL;

static void 
ExchLogInfo (const char * fmt, ...)
{
    if (m_gpg) {
        va_list a;
  
        va_start (a, fmt);
        m_gpg->logDebug (fmt, a);
        va_end (a);
    }
}

/* The one and only CGPGExchApp object */
BEGIN_MESSAGE_MAP(CGPGExchApp, CWinApp)
END_MESSAGE_MAP()

CGPGExchApp::CGPGExchApp (void)
{
    ExchLogInfo("GPGExch\n");
}

CGPGExchApp theApp;



/* ExchEntryPoint -
 The entry point which Exchange calls.
 This is called for each context entry. Creates a new CAvkExchExt object
 every time so each context will get its own CAvkExchExt interface.
 
 Return value: Pointer to Exchange Extension Object */
LPEXCHEXT CALLBACK 
ExchEntryPoint (void)
{
    ExchLogInfo ("extension entry point...\n");
    return new CGPGExchExt;
}


/* DllRegisterServer
 Registers this object as exchange extension. Sets the contextes which are 
 implemented by this object. */
STDAPI 
DllRegisterServer (void)
{    
    HKEY hkey;
    CHAR szKeyBuf[1024];
    CHAR szEntry[512];
    TCHAR szModuleFileName[MAX_PATH];
    DWORD dwTemp = 0;
    long ec;

    /* get server location */
    DWORD dwResult = ::GetModuleFileName(theApp.m_hInstance, szModuleFileName, MAX_PATH);
    if (dwResult == 0)	
	return E_FAIL;

    lstrcpy (szKeyBuf, "Software\\Microsoft\\Exchange\\Client\\Extensions");
    lstrcpy (szEntry, "4.0;");
    lstrcat (szEntry, szModuleFileName);
    lstrcat (szEntry, ";1;11000111111100");  /* context information */
    ec = RegCreateKeyEx (HKEY_LOCAL_MACHINE, szKeyBuf, 0, NULL, 
				   REG_OPTION_NON_VOLATILE,
				   KEY_ALL_ACCESS, NULL, &hkey, NULL);
    if (ec != ERROR_SUCCESS) {
	ExchLogInfo ("DllRegisterServer failed\n");
	return E_ACCESSDENIED;
    }
    
    dwTemp = lstrlen (szEntry) + 1;
    RegSetValueEx (hkey, "OutlGPG", 0, REG_SZ, (BYTE*) szEntry, dwTemp);

    /* set outlook update flag */
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

    ExchLogInfo ("DllRegisterServer succeeded\n");
    return S_OK;
}


/* DllUnregisterServer - Unregisters this object in the exchange extension list. */
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
	ExchLogInfo ("DllUnregisterServer: access denied.\n");
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


/* constructor of CGPGExchExt
 Initializes members and creates the interface objects for the new context.
 Does the dll initialization if it was not done before. */
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
    if (!g_bInitDll) {
	if (m_gpg == NULL) {
	    m_gpg = CreateMapiGPGME (NULL);
	    m_gpg->readOptions ();
	}
	g_bInitDll = TRUE;
	ExchLogInfo("CGPGExchExt load\n");
    }
    else
	ExchLogInfo("CGPGExchExt exists\n");
}


/* constructor of CGPGExchExt - Uninitializes the dll in the dession context. */
CGPGExchExt::~CGPGExchExt (void) 
{
    if (m_lContext == EECONTEXT_SESSION) {
	if (g_bInitDll) {
	    if (m_gpg != NULL) {
		m_gpg->writeOptions ();
		delete m_gpg;
		m_gpg = NULL;
	    }
	    g_bInitDll = FALSE;
	}	
    }
}


/* CGPGExchExt::QueryInterface
 Called by Exchage to request for interfaces.

 Return value: S_OK if the interface is supported, otherwise E_NOINTERFACE: */
STDMETHODIMP 
CGPGExchExt::QueryInterface(
	REFIID riid,      // The interface ID.
	LPVOID * ppvObj)  // The address of interface object pointer.
{
    HRESULT hr = S_OK;

    *ppvObj = NULL;

    if ((riid == IID_IUnknown) || (riid == IID_IExchExt)) {
        *ppvObj = (LPUNKNOWN) this;
    }
    else if (riid == IID_IExchExtMessageEvents) {
        *ppvObj = (LPUNKNOWN) m_pExchExtMessageEvents;
    }
    else if (riid == IID_IExchExtCommands) {
        *ppvObj = (LPUNKNOWN)m_pExchExtCommands;
        m_pExchExtCommands->SetContext(m_lContext);
    }
    else if (riid == IID_IExchExtPropertySheets) {
	if (m_lContext != EECONTEXT_PROPERTYSHEETS)
	    return E_NOINTERFACE;
        *ppvObj = (LPUNKNOWN) m_pExchExtPropertySheets;
    }
    else
        hr = E_NOINTERFACE;

    if (*ppvObj != NULL)
        ((LPUNKNOWN)*ppvObj)->AddRef();

    /*ExchLogInfo("QueryInterface %d\n", __LINE__);*/
    return hr;
}


/* CGPGExchExt::Install
 Called once for each new context. Checks the exchange extension version 
 number and the context.
 
 Return value: S_OK if the extension should used in the requested context, 
               otherwise S_FALSE. */
STDMETHODIMP CGPGExchExt::Install(
	LPEXCHEXTCALLBACK pEECB, // The pointer to Exchange Extension callback function.
	ULONG lContext,          // The context code at time of being called.
	ULONG lFlags)            // The flag to say if install is for modal or not.
{
    ULONG lBuildVersion;

    m_lContext = lContext;

    /*ExchLogInfo("Install %d\n", __LINE__);*/
    // check the version 
    pEECB->GetVersion (&lBuildVersion, EECBGV_GETBUILDVERSION);
    if (EECBGV_BUILDVERSION_MAJOR != (lBuildVersion & EECBGV_BUILDVERSION_MAJOR_MASK))
        return S_FALSE;

    // and the context
    if ((lContext == EECONTEXT_PROPERTYSHEETS) ||
	(lContext == EECONTEXT_SENDNOTEMESSAGE) ||
	(lContext == EECONTEXT_SENDPOSTMESSAGE) ||
	(lContext == EECONTEXT_SENDRESENDMESSAGE) ||
	(lContext == EECONTEXT_READNOTEMESSAGE) ||
	(lContext == EECONTEXT_READPOSTMESSAGE) ||
	(lContext == EECONTEXT_READREPORTMESSAGE) ||
	(lContext == EECONTEXT_VIEWER))
	return S_OK;
    return S_FALSE;
}


CGPGExchExtMessageEvents::CGPGExchExtMessageEvents (CGPGExchExt *pParentInterface)
{ 
    m_pExchExt = pParentInterface;
    m_lRef = 0; 
    m_bOnSubmitCalled = FALSE;
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


/* CGPGExchExtMessageEvents::OnRead - Called from Exchange on reading a message.
 Return value: S_FALSE to signal Exchange to continue calling extensions. */
STDMETHODIMP CGPGExchExtMessageEvents::OnRead(
	LPEXCHEXTCALLBACK pEECB) // A pointer to IExchExtCallback interface.
{
    ExchLogInfo ("OnRead\n");
    return S_FALSE;
}


/* CGPGExchExtMessageEvents::OnReadComplete
 Called by Exchange after a message has been read.

 Return value: S_FALSE to signal Exchange to continue calling extensions. */
STDMETHODIMP CGPGExchExtMessageEvents::OnReadComplete(
	LPEXCHEXTCALLBACK pEECB, // A pointer to IExchExtCallback interface.
	ULONG lFlags)
{
    ExchLogInfo ("OnReadComplete\n");
    return S_FALSE;
}


/* CGPGExchExtMessageEvents::OnWrite - Called by Exchange when a message will be written.
 @rdesc S_FALSE to signal Exchange to continue calling extensions. */
STDMETHODIMP CGPGExchExtMessageEvents::OnWrite(
	LPEXCHEXTCALLBACK pEECB) // A pointer to IExchExtCallback interface.
{
    ExchLogInfo ("OnWrite\n");
    return S_FALSE;
}


/* CGPGExchExtMessageEvents::OnWriteComplete
 Called by Exchange when the data has been written to the message.
 Encrypts and signs the message if the options are set.
 @pEECB - A pointer to IExchExtCallback interface.

 Return value: S_FALSE: signals Exchange to continue calling extensions
               E_FAIL:  signals Exchange an error; the message will not be sent */
STDMETHODIMP CGPGExchExtMessageEvents::OnWriteComplete (
		  LPEXCHEXTCALLBACK pEECB, ULONG lFlags)
{
    HRESULT hrReturn = S_FALSE;
    LPMESSAGE pMessage = NULL;
    LPMDB pMDB = NULL;
    HWND hWnd = NULL;

    if (FAILED(pEECB->GetWindow (&hWnd)))
	hWnd = NULL;

    if (!m_bOnSubmitCalled) /* the user is just saving the message */
	return S_FALSE;

    if (m_bWriteFailed)     /* operation failed already */
	return S_FALSE;

    HRESULT hr = pEECB->GetObject (&pMDB, (LPMAPIPROP *)&pMessage);
    if (SUCCEEDED (hr)) {
	if (m_pExchExt->m_gpgEncrypt || m_pExchExt->m_gpgSign) {
	    m_gpg->setMessage (pMessage);
	    if (m_gpg->doCmd (m_pExchExt->m_gpgEncrypt,
			      m_pExchExt->m_gpgSign)) {
		hrReturn = E_FAIL;
		m_bWriteFailed = TRUE;	
	    }
	}
    }
    if (pMessage != NULL)
	UlRelease(pMessage);
    if (pMDB != NULL)
	UlRelease(pMDB);

    return hrReturn;
}

/* CGPGExchExtMessageEvents::OnCheckNames

 Called by Exchange when the user selects the "check names" command.

 @pEECB - A pointer to IExchExtCallback interface.

 Return value: S_FALSE to signal Exchange to continue calling extensions.
*/
STDMETHODIMP CGPGExchExtMessageEvents::OnCheckNames(LPEXCHEXTCALLBACK pEECB)
{
    return S_FALSE;
}


/* CGPGExchExtMessageEvents::OnCheckNamesComplete

 Called by Exchange when "check names" command is complete.

  @pEECB - A pointer to IExchExtCallback interface.

 Return value: S_FALSE to signal Exchange to continue calling extensions.
*/
STDMETHODIMP CGPGExchExtMessageEvents::OnCheckNamesComplete(
			LPEXCHEXTCALLBACK pEECB, ULONG lFlags)
{
    return S_FALSE;
}


/* CGPGExchExtMessageEvents::OnSubmit

 Called by Exchange before the message data will be written and submitted.
 to MAPI.

 @pEECB - A pointer to IExchExtCallback interface.

 Return value: S_FALSE to signal Exchange to continue calling extensions.
*/
STDMETHODIMP CGPGExchExtMessageEvents::OnSubmit(
			    LPEXCHEXTCALLBACK pEECB)
{
    m_bOnSubmitCalled = TRUE;
    m_bWriteFailed = FALSE;
    return S_FALSE;
}


/* CGPGExchExtMessageEvents::OnSubmitComplete

  @pEECB - A pointer to IExchExtCallback interface.

 Called by Echange after the message has been submitted to MAPI.
*/
STDMETHODIMP_ (VOID) CGPGExchExtMessageEvents::OnSubmitComplete (
			    LPEXCHEXTCALLBACK pEECB, ULONG lFlags)
{
    m_bOnSubmitCalled = FALSE; 
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


// XXX IExchExtSessionEvents::OnDelivery: could be used to automatically decrypt new mails
// when they arrive

/* CGPGExchExtCommands::InstallCommands

 Called by Echange to install commands and toolbar buttons.

 Return value: S_FALSE to signal Exchange to continue calling extensions.
*/
STDMETHODIMP CGPGExchExtCommands::InstallCommands(
	LPEXCHEXTCALLBACK pEECB, // The Exchange Callback Interface.
	HWND hWnd,               // The window handle to main window of context.
	HMENU hMenu,             // The menu handle to main menu of context.
	UINT FAR * pnCommandIDBase,  // The base conmmand id.
	LPTBENTRY pTBEArray,     // The array of toolbar button entries.
	UINT nTBECnt,            // The count of button entries in array.
	ULONG lFlags)            // reserved
{
    HMENU hMenuTools;
    m_hWnd = hWnd;

    /* XXX: factor out common code */
    if (m_lContext == EECONTEXT_READNOTEMESSAGE) {
	int nTBIndex;
	HWND hwndToolbar = NULL;
	CHAR szBuffer[128];

        pEECB->GetMenuPos (EECMDID_ToolsCustomizeToolbar, &hMenuTools, NULL, NULL, 0);
        AppendMenu (hMenuTools, MF_SEPARATOR, 0, NULL);
	
	LoadString (theApp.m_hInstance, IDS_DECRYPT_MENU_ITEM, szBuffer, 128);
        AppendMenu (hMenuTools, MF_BYPOSITION | MF_STRING, *pnCommandIDBase, szBuffer);

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
	    tbab.hInst = theApp.m_hInstance;
	    tbab.nID = IDB_DECRYPT;
	    m_nToolbarBitmap1 = SendMessage(hwndToolbar, TB_ADDBITMAP, 1, (LPARAM)&tbab);
	    m_nToolbarButtonID2 = pTBEArray[nTBIndex].itbbBase;
	    pTBEArray[nTBIndex].itbbBase++;
	}
    }

    if (m_lContext == EECONTEXT_SENDNOTEMESSAGE) {
	CHAR szBuffer[128];
	int nTBIndex;
	HWND hwndToolbar = NULL;

        pEECB->GetMenuPos(EECMDID_ToolsCustomizeToolbar, &hMenuTools, NULL, NULL, 0);
        AppendMenu(hMenuTools, MF_SEPARATOR, 0, NULL);
	
	LoadString(theApp.m_hInstance, IDS_ENCRYPT_MENU_ITEM, szBuffer, 128);
        AppendMenu(hMenuTools, MF_BYPOSITION | MF_STRING, *pnCommandIDBase, szBuffer);

        m_nCmdEncrypt = *pnCommandIDBase;
        (*pnCommandIDBase)++;

	LoadString(theApp.m_hInstance, IDS_SIGN_MENU_ITEM, szBuffer, 128);
        AppendMenu(hMenuTools, MF_BYPOSITION | MF_STRING, *pnCommandIDBase, szBuffer);

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
	    tbab.hInst = theApp.m_hInstance;
	    tbab.nID = IDB_ENCRYPT;
	    m_nToolbarBitmap1 = SendMessage(hwndToolbar, TB_ADDBITMAP, 1, (LPARAM)&tbab);
	    m_nToolbarButtonID2 = pTBEArray[nTBIndex].itbbBase;
	    pTBEArray[nTBIndex].itbbBase++;
	    tbab.nID = IDB_SIGN;
	    m_nToolbarBitmap2 = SendMessage(hwndToolbar, TB_ADDBITMAP, 1, (LPARAM)&tbab);
	}
	m_pExchExt->m_gpgEncrypt = m_gpg->getEncryptDefault ();
	m_pExchExt->m_gpgSign = m_gpg->getSignDefault ();
    }

    if (m_lContext == EECONTEXT_VIEWER) {
	CHAR szBuffer[128];
	int nTBIndex;
	HWND hwndToolbar = NULL;

        pEECB->GetMenuPos (EECMDID_ToolsCustomizeToolbar, &hMenuTools, NULL, NULL, 0);
        AppendMenu (hMenuTools, MF_SEPARATOR, 0, NULL);
	
	LoadString (theApp.m_hInstance, IDS_KEY_MANAGER, szBuffer, 128);
        AppendMenu (hMenuTools, MF_BYPOSITION | MF_STRING, *pnCommandIDBase, szBuffer);

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
	    tbab.hInst = theApp.m_hInstance;
	    tbab.nID = IDB_KEY_MANAGER;
	    m_nToolbarBitmap1 = SendMessage(hwndToolbar, TB_ADDBITMAP, 1, (LPARAM)&tbab);
	}	
    }
    return S_FALSE;
}


/* CGPGExchExtCommands::DoCommand - Called by Exchange when a user selects a command. 

 Return value: S_OK if command is handled, otherwise S_FALSE.
*/
STDMETHODIMP CGPGExchExtCommands::DoCommand(
	LPEXCHEXTCALLBACK pEECB, // The Exchange Callback Interface.
	UINT nCommandID)         // The command id.
{

    if ((nCommandID != m_nCmdEncrypt) && 
	(nCommandID != m_nCmdSign))
	return S_FALSE; 

    if (m_lContext == EECONTEXT_READNOTEMESSAGE) {
	LPMESSAGE pMessage = NULL;
	LPMDB pMDB = NULL;
	HWND hWnd = NULL;

	if (FAILED (pEECB->GetWindow (&hWnd)))
	    hWnd = NULL;
	HRESULT hr = pEECB->GetObject (&pMDB, (LPMAPIPROP *)&pMessage);
	if (SUCCEEDED (hr)) {
	    if (nCommandID == m_nCmdEncrypt) {
		m_gpg->setWindow (hWnd);
		m_gpg->setMessage (pMessage);
		m_gpg->decrypt ();
	    }
	}
	if (pMessage != NULL)
	    UlRelease(pMessage);
	if (pMDB != NULL)
	    UlRelease(pMDB);
    }
    if (m_lContext == EECONTEXT_SENDNOTEMESSAGE) {
	HWND hWnd = NULL;
	if (FAILED(pEECB->GetWindow (&hWnd)))
	    hWnd = NULL;
	if (nCommandID == m_nCmdEncrypt)
	    m_pExchExt->m_gpgEncrypt = !m_pExchExt->m_gpgEncrypt;
	if (nCommandID == m_nCmdSign)
	    m_pExchExt->m_gpgSign = !m_pExchExt->m_gpgSign;
    }
    if (m_lContext == EECONTEXT_VIEWER) {
	if (m_gpg->startKeyManager ())
	    MessageBox (NULL, "Could not start Key-Manager", "GPGExch", MB_ICONERROR|MB_OK);
    }
    return S_OK; 
}

STDMETHODIMP_(VOID) CGPGExchExtCommands::InitMenu(
	LPEXCHEXTCALLBACK pEECB) // The pointer to Exchange Callback Interface.
{
}

/* CGPGExchExtCommands::Help

 Called by Exhange when the user requests help for a menu item.

 Return value: S_OK when it is a menu item of this plugin and the help was shown;
               otherwise S_FALSE
*/
STDMETHODIMP CGPGExchExtCommands::Help (
	LPEXCHEXTCALLBACK pEECB, // The pointer to Exchange Callback Interface.
	UINT nCommandID)         // The command id.
{
    if (m_lContext == EECONTEXT_READNOTEMESSAGE) {
    	if (nCommandID == m_nCmdEncrypt) {
	    CHAR szBuffer[512];
	    CHAR szAppName[128];

	    LoadString (theApp.m_hInstance, IDS_DECRYPT_HELP, szBuffer, 511);
	    LoadString (theApp.m_hInstance, IDS_APP_NAME, szAppName, 127);
	    MessageBox (m_hWnd, szBuffer, szAppName, MB_OK);
	    return S_OK;
	}
    }
    if (m_lContext == EECONTEXT_SENDNOTEMESSAGE) {
	if (nCommandID == m_nCmdEncrypt) {
	    CHAR szBuffer[512];
	    CHAR szAppName[128];
	    LoadString(theApp.m_hInstance, IDS_ENCRYPT_HELP, szBuffer, 511);
	    LoadString(theApp.m_hInstance, IDS_APP_NAME, szAppName, 127);
	    MessageBox(m_hWnd, szBuffer, szAppName, MB_OK);	
	    return S_OK;
	} 
	if (nCommandID == m_nCmdSign) {
	    CHAR szBuffer[512];	
	    CHAR szAppName[128];	
	    LoadString(theApp.m_hInstance, IDS_SIGN_HELP, szBuffer, 511);	
	    LoadString(theApp.m_hInstance, IDS_APP_NAME, szAppName, 127);	
	    MessageBox(m_hWnd, szBuffer, szAppName, MB_OK);	
	    return S_OK;
	} 
    }

    if (m_lContext == EECONTEXT_VIEWER) {
    	if (nCommandID == m_nCmdEncrypt) {
		CHAR szBuffer[512];
		CHAR szAppName[128];
		LoadString(theApp.m_hInstance, IDS_KEY_MANAGER_HELP, szBuffer, 511);
		LoadString(theApp.m_hInstance, IDS_APP_NAME, szAppName, 127);
		MessageBox(m_hWnd, szBuffer, szAppName, MB_OK);
		return S_OK;
	} 
    }

    return S_FALSE;
}

/* CGPGExchExtCommands::QueryHelpText

 Called by Exhange to get the status bar text or the tooltip of a menu item.

 @rdesc S_OK when it is a menu item of this plugin and the text was set;
        otherwise S_FALSE.
*/
STDMETHODIMP CGPGExchExtCommands::QueryHelpText(
	UINT nCommandID,  // The command id corresponding to menu item activated.
	ULONG lFlags,     // Identifies either EECQHT_STATUS or EECQHT_TOOLTIP.
    LPTSTR pszText,   // A pointer to buffer to be populated with text to display.
	UINT nCharCnt)    // The count of characters available in psz buffer.
{
	
    if (m_lContext == EECONTEXT_READNOTEMESSAGE) {
	if (nCommandID == m_nCmdEncrypt) {
	    if (lFlags == EECQHT_STATUS)
		LoadString (theApp.m_hInstance, IDS_DECRYPT_STATUSBAR, pszText, nCharCnt);
  	    if (lFlags == EECQHT_TOOLTIP)
		LoadString (theApp.m_hInstance, IDS_DECRYPT_TOOLTIP, pszText, nCharCnt);
	    return S_OK;
	}
    }
    if (m_lContext == EECONTEXT_SENDNOTEMESSAGE) {
	if (nCommandID == m_nCmdEncrypt) {
	    if (lFlags == EECQHT_STATUS)
		LoadString (theApp.m_hInstance, IDS_ENCRYPT_STATUSBAR, pszText, nCharCnt);
	    if (lFlags == EECQHT_TOOLTIP)
		LoadString (theApp.m_hInstance, IDS_ENCRYPT_TOOLTIP, pszText, nCharCnt);
	    return S_OK;
	}
	if (nCommandID == m_nCmdSign) {
	    if (lFlags == EECQHT_STATUS)
		LoadString (theApp.m_hInstance, IDS_SIGN_STATUSBAR, pszText, nCharCnt);
  	    if (lFlags == EECQHT_TOOLTIP)
	        LoadString (theApp.m_hInstance, IDS_SIGN_TOOLTIP, pszText, nCharCnt);
	    return S_OK;
	}
    }
    if (m_lContext == EECONTEXT_VIEWER) {
	if (nCommandID == m_nCmdEncrypt) {
	    if (lFlags == EECQHT_STATUS)
		LoadString (theApp.m_hInstance, IDS_KEY_MANAGER_STATUSBAR, pszText, nCharCnt);
	    if (lFlags == EECQHT_TOOLTIP)
		LoadString (theApp.m_hInstance, IDS_KEY_MANAGER_TOOLTIP, pszText, nCharCnt);
	    return S_OK;
	}	
    }
    return S_FALSE;
}

/////////////////////////////////////////////////////////////////////////////
// CGPGExchExtCommands::QueryButtonInfo 
//
// Called by Exchange to get toolbar button infos,
//
// @rdesc S_OK when it is a button of this plugin and the requested info was delivered;
//        otherwise S_FALSE
//
STDMETHODIMP CGPGExchExtCommands::QueryButtonInfo (
	ULONG lToolbarID,       // The toolbar identifier.
	UINT nToolbarButtonID,  // The toolbar button index.
    LPTBBUTTON pTBB,        // A pointer to toolbar button structure (see TBBUTTON structure).
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
			LoadString(theApp.m_hInstance, IDS_DECRYPT_TOOLTIP, lpszDescription, nCharCnt);
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
			LoadString(theApp.m_hInstance, IDS_ENCRYPT_TOOLTIP, lpszDescription, nCharCnt);
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
			LoadString(theApp.m_hInstance, IDS_SIGN_TOOLTIP, lpszDescription, nCharCnt);
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
			LoadString(theApp.m_hInstance, IDS_KEY_MANAGER_TOOLTIP, lpszDescription, nCharCnt);
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


CGPGExchExtPropertySheets::CGPGExchExtPropertySheets(CGPGExchExt* pParentInterface)
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


/* CGPGExchExtPropertySheets::GetMaxPageCount

 Called by Echange to get the maximum number of property pages which are
 to be added.

 Return value: The maximum number of custom pages for the property sheet.
*/
STDMETHODIMP_ (ULONG) 
CGPGExchExtPropertySheets::GetMaxPageCount(
	ULONG lFlags) // A bitmask indicating what type of property sheet is being displayed.
{
    if (lFlags == EEPS_TOOLSOPTIONS)
	return 1;	
    return 0;
}


BOOL CALLBACK GPGOptionsDlgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam);


/* CGPGExchExtPropertySheets::GetPages

 Called by Exchange to request information about the property page.

 Return value: S_OK to signal Echange to use the pPSP information.
*/
STDMETHODIMP 
CGPGExchExtPropertySheets::GetPages(
	LPEXCHEXTCALLBACK pEECB, // A pointer to Exchange callback interface.
	ULONG lFlags,            // A  bitmask indicating what type of property sheet is being displayed.
	LPPROPSHEETPAGE pPSP,    // The output parm pointing to pointer to list of property sheets.
	ULONG FAR * plPSP)       // The output parm pointing to buffer contaiing number of property sheets actually used.
{
    int resid = 0;

    switch (GetUserDefaultLangID ()) {
    case 0x0407:    resid = IDD_GPG_OPTIONS_DE;break;
    default:	    resid = IDD_GPG_OPTIONS; break;
    }

    pPSP[0].dwSize = sizeof (PROPSHEETPAGE);
    pPSP[0].dwFlags = PSP_DEFAULT | PSP_HASHELP;
    pPSP[0].hInstance = theApp.m_hInstance;
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
CGPGExchExtPropertySheets::FreePages (LPPROPSHEETPAGE pPSP, ULONG lFlags, ULONG lPSP)
{
}
