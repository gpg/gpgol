/* GPGExch.cpp - exchange extension classes
 *	Copyright (C) 2001 G Data Software AG, http://www.gdata.de
 *	Copyright (C) 2004 g10 Code GmbH
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
#include "..\GDGPG\Wrapper\GDGPGWrapper.h"
#include "GPG.h"

BOOL g_bInitDll = FALSE;
BOOL g_bInitCom = FALSE;
CGPG g_gpg;


/* The one and only CGPGExchApp object */
BEGIN_MESSAGE_MAP(CGPGExchApp, CWinApp)
END_MESSAGE_MAP()

CGPGExchApp::CGPGExchApp ()
{
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
    return new CGPGExchExt;
}


/* DllRegisterServer
 Registers this object as exchange extension. Sets the contextes which are 
 implemented by this object. */
STDAPI 
DllRegisterServer (void)
{
    /* get server location */
    TCHAR szModuleFileName[MAX_PATH];
    DWORD dwResult = ::GetModuleFileName(theApp.m_hInstance, szModuleFileName, MAX_PATH);
    if(dwResult == 0)	return E_FAIL;
    HKEY hKey;
    CHAR szKeyBuf[1024];
    CHAR szEntry[512];

    lstrcpy (szKeyBuf, "Software\\Microsoft\\Exchange\\Client\\Extensions");
    lstrcpy (szEntry, "4.0;");
    lstrcat (szEntry, szModuleFileName);
    lstrcat (szEntry, ";1;11000111111100");  // context information
    long lResult = RegCreateKeyEx (HKEY_LOCAL_MACHINE, szKeyBuf, 0, NULL, 
				   REG_OPTION_NON_VOLATILE,
				   KEY_ALL_ACCESS, NULL, &hKey, NULL);
    if (lResult != ERROR_SUCCESS) 
	return E_ACCESSDENIED;
    DWORD dwTemp = 0;
    dwTemp = lstrlen (szEntry) + 1;
    RegSetValueEx (hKey, "GPG Exchange", 0, REG_SZ, (BYTE*) szEntry, dwTemp);

    /* set outlook update flag */
    strcpy(szEntry, "4.0;Outxxx.dll;7;000000000000000;0000000000;OutXXX");
    dwTemp = lstrlen (szEntry) + 1;
    RegSetValueEx (hKey, "Outlook Setup Extension", 0, REG_SZ, (BYTE*) szEntry, dwTemp);
    RegCloseKey (hKey);
    
    return S_OK;
}


/* DllUnregisterServer - Unregisters this object in the exchange extension list. */
STDAPI 
DllUnregisterServer (void)
{
    HKEY hKey;
    CHAR szKeyBuf[1024];

    lstrcpy(szKeyBuf, "Software\\Microsoft\\Exchange\\Client\\Extensions");
    /* create and open key and subkey */
    long lResult = RegCreateKeyEx(HKEY_LOCAL_MACHINE, szKeyBuf, 0, NULL, 
				    REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 
				    NULL, &hKey, NULL);
    if (lResult != ERROR_SUCCESS) 
	return E_ACCESSDENIED;
    RegDeleteValue (hKey, "GPG Exchange");
    /* set outlook update flag */
    CHAR szEntry[512];
    strcpy (szEntry, "4.0;Outxxx.dll;7;000000000000000;0000000000;OutXXX");
    DWORD dwTemp = lstrlen (szEntry) + 1;
    RegSetValueEx (hKey, "Outlook Setup Extension", 0, REG_SZ, (BYTE*) szEntry, dwTemp);
    RegCloseKey (hKey);

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
    m_bEncryptWhenSending = FALSE;
    m_bSignWhenSending = FALSE;
    m_pExchExtMessageEvents = new CGPGExchExtMessageEvents (this);
    m_pExchExtCommands = new CGPGExchExtCommands (this);
    m_pExchExtPropertySheets = new CGPGExchExtPropertySheets (this);
    if (!g_bInitDll)
    {
	if (CoInitialize(NULL) == S_OK)
	    g_bInitCom = TRUE;
	g_gpg.Init();
	g_bInitDll = TRUE;
    }
}


/* constructor of CGPGExchExt - Uninitializes the dll in the dession context. */
CGPGExchExt::~CGPGExchExt () 
{
    if (m_lContext == EECONTEXT_SESSION)
    {
	if (g_bInitDll)
	{
	    if (g_bInitCom)
	    {
		CoUninitialize();
		g_bInitCom = FALSE;	
	    }
	    g_gpg.UnInit();
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

    if ((riid == IID_IUnknown) || (riid == IID_IExchExt))
    {
        *ppvObj = (LPUNKNOWN) this;
    }
    else if (riid == IID_IExchExtMessageEvents)
    {
        *ppvObj = (LPUNKNOWN) m_pExchExtMessageEvents;
    }
    else if (riid == IID_IExchExtCommands)
    {
        *ppvObj = (LPUNKNOWN)m_pExchExtCommands;
        m_pExchExtCommands->SetContext(m_lContext);
    }
    else if (riid == IID_IExchExtPropertySheets)
    {
		 if (m_lContext != EECONTEXT_PROPERTYSHEETS)
                 return E_NOINTERFACE;
        *ppvObj = (LPUNKNOWN) m_pExchExtPropertySheets;
    }
    else
        hr = E_NOINTERFACE;

    if (*ppvObj != NULL)
        ((LPUNKNOWN)*ppvObj)->AddRef();

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

    // check the version 
    pEECB->GetVersion(&lBuildVersion, EECBGV_GETBUILDVERSION);
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


CGPGExchExtMessageEvents::CGPGExchExtMessageEvents (CGPGExchExt* pParentInterface) 
{ 
    m_pExchExt = pParentInterface;
    m_lRef = 0; 
    m_bOnSubmitCalled = FALSE;
};


STDMETHODIMP 
CGPGExchExtMessageEvents::QueryInterface (REFIID riid, LPVOID FAR * ppvObj)
{   
    *ppvObj = NULL;
    if (riid == IID_IExchExtMessageEvents)
    {
        *ppvObj = (LPVOID)this;
        AddRef();
        return S_OK;
    }
    if (riid == IID_IUnknown)
    {
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
    USES_CONVERSION;

    HRESULT hr = S_OK;
    LPMDB pMDB = NULL;
    LPMESSAGE pMessage = NULL;
    LPSTREAM pStreamBody = NULL;
    LARGE_INTEGER off;
    ULARGE_INTEGER off_ret;    
    char buf[512], in_tmpname[512], out_tmpname[512];
    char  * asct;
    ULONG nr=0;
    string sMsg = "", sOutMsg = "";
    SPropValue sProp;
    LPSPropValue sPropVal=NULL;
    int ret = 0, mimeType = 0;
    FILE * fp;
    time_t t;
    struct tm * tm;    

    t = time (NULL);
    tm = localtime (&t);
    asct = asctime (tm);
    asct[strlen (asct)-1] = 0;
    
    /* XXX: clean up the mess!!! */

    hr = pEECB->GetObject (&pMDB, (LPMAPIPROP *)&pMessage);
    if (FAILED (hr) || pMessage == NULL) 
    {
	hr = S_FALSE;
	goto failed;
    }

    /* Do not process outgoing messages */
    HrGetOneProp (pMessage, PR_MESSAGE_FLAGS, &sPropVal);
    if (sPropVal->Value.l & MSGFLAG_UNSENT)
    {
	hr = S_OK;
	goto failed;
    }

    /* 
    if (g_gpg.CheckPGPMime (NULL, pMessage, mimeType) == TRUE)
    {
	g_gpg.ProcessPGPMime (NULL, pMessage, mimeType);
	goto failed;
    }
    */

    hr = pMessage->OpenProperty (PR_BODY, &IID_IStream, STGM_DIRECT|STGM_READ,
				 0, (LPUNKNOWN *)&pStreamBody);
    if (GetScode (hr) == MAPI_E_NOT_FOUND) 
    {
	hr = S_OK;
	goto failed;
    }
    if (FAILED (hr))
    {
	hr = S_FALSE;
	goto failed;
    }

    off.LowPart = 0;
    off.HighPart = 0;
    hr = pStreamBody->Seek (off, STREAM_SEEK_SET, &off_ret);
    if (FAILED (hr))
    {
	hr = S_FALSE;
	goto failed;
    }

    while ((hr = pStreamBody->Read (buf, 512, &nr)) == S_OK
	    && nr > 0)
    {
	buf[nr] = 0;
	sMsg = sMsg + buf;
    }

    if (strstr (sMsg.c_str (), "-----BEGIN PGP SIGNED MESSAGE-----") && 
	strstr (sMsg.c_str (), "-----END PGP SIGNATURE-----"))
    {
	GetTempPath (sizeof in_tmpname-1, in_tmpname);
	strcat (in_tmpname, "_gpg_in.tmp.asc");

	fp = fopen (in_tmpname, "wb");
	if (!fp)
	{
	    hr = S_FALSE;
	    goto failed;
	}
	fwrite (sMsg.c_str (), 1, sMsg.length (), fp);
	fclose (fp);
	
	sOutMsg = "[-- Begin GPG Output (";
	sOutMsg += asct;
	sOutMsg += ") --]\n";
	sOutMsg += OLE2A (g_gpg.GetGPGInfo (A2OLE (in_tmpname)));
	sOutMsg += "\n[--End GPG Output --]\n\n";
	sOutMsg += sMsg;

	sProp.ulPropTag = PR_BODY;
	sProp.Value.lpszA = (char *)sOutMsg.c_str ();
	hr = HrSetOneProp (pMessage, &sProp);

	goto failed;
    }

    if (strstr (sMsg.c_str (), "-----BEGIN PGP PUBLIC KEY BLOCK-----")
	&& strstr (sMsg.c_str (), "-----END PGP PUBLIC KEY BLOCK-----"))
    {
	GetTempPath (sizeof in_tmpname-1, in_tmpname);
	strcat (in_tmpname, "_gpg_in.tmp");

	fp = fopen (in_tmpname, "wb");
	if (!fp)
	{
	    hr = S_FALSE;
	    goto failed;
	}
	fwrite (sMsg.c_str (), 1, sMsg.length (), fp);
	fclose (fp);

	sOutMsg = "[-- Begin GPG Output (";
	sOutMsg += asct;
	sOutMsg += ") --]\n\n";
	sOutMsg += OLE2A (g_gpg.GetGPGInfo (A2OLE (in_tmpname)));
	sOutMsg += "\n[--End GPG Output --]\n\n";
	sOutMsg += sMsg;

	sProp.ulPropTag = PR_BODY;
	sProp.Value.lpszA = (char *)sOutMsg.c_str ();
	hr = HrSetOneProp (pMessage, &sProp);

	goto failed;
    }

    if (strstr (sMsg.c_str (), "-----BEGIN PGP MESSAGE-----") &&
	strstr (sMsg.c_str (), "-----END PGP MESSAGE-----"))
    {
	GetTempPath (sizeof in_tmpname-1, in_tmpname);
	strcat (in_tmpname, "_gpg_in.tmp");

	GetTempPath (sizeof out_tmpname-1, out_tmpname);
	strcat (out_tmpname, "_gpg_out.tmp");

	fp = fopen (in_tmpname, "wb");
	if (!fp)
	{
	    hr = S_FALSE;
	    goto failed;
	}
	fwrite (sMsg.c_str (), 1, sMsg.length (), fp);
	fclose (fp);

	g_gpg.DecryptFile (NULL, A2OLE (in_tmpname), A2OLE (out_tmpname), ret);
	if (ret)
	{
	    if (ret == 4 || ret == 8)
	    {
		sOutMsg = "[-- Begin GPG Output (";
		sOutMsg += asct;
		sOutMsg += ") --]\n";
		sOutMsg += OLE2A (g_gpg.GetGPGOutput ());
		sOutMsg += "\n[--End GPG Output --]\n\n";

		sProp.ulPropTag = PR_BODY;
		sProp.Value.lpszA = (char *)sOutMsg.c_str ();
		hr = HrSetOneProp (pMessage, &sProp);
		goto failed;
	    }
	}
	fp = fopen (out_tmpname, "rb");
	if (!fp)
	    sOutMsg = "";
	else
	{
	    while (!feof (fp))
	    {
		nr = fread (buf, 1, 512, fp);
		buf[nr] = 0;
		if (!nr)
		    break;
		sOutMsg = sOutMsg + buf;
	    }
	    fclose (fp);
	}

	sProp.ulPropTag = PR_BODY;
	sProp.Value.lpszA = (char *)sOutMsg.c_str ();
	hr = HrSetOneProp (pMessage, &sProp);

	sProp.ulPropTag = PR_ACCESS;
	sProp.Value.l = MAPI_ACCESS_MODIFY;
	HrSetOneProp(pMessage, &sProp);	
    }

failed:
    MAPIFreeBuffer (sPropVal);
    DeleteFile (in_tmpname);
    DeleteFile (out_tmpname);
    UlRelease (pMDB);
    UlRelease (pMessage);
    if (pStreamBody != NULL)
	pStreamBody->Release ();
    return hr;
}


/* CGPGExchExtMessageEvents::OnReadComplete
 Called by Exchange after a message has been read.

 Return value: S_FALSE to signal Exchange to continue calling extensions. */
STDMETHODIMP CGPGExchExtMessageEvents::OnReadComplete(
	LPEXCHEXTCALLBACK pEECB, // A pointer to IExchExtCallback interface.
	ULONG lFlags)
{	
    return S_FALSE;
}


/* CGPGExchExtMessageEvents::OnWrite - Called by Exchange when a message will be written.
 @rdesc S_FALSE to signal Exchange to continue calling extensions. */
STDMETHODIMP CGPGExchExtMessageEvents::OnWrite(
	LPEXCHEXTCALLBACK pEECB) // A pointer to IExchExtCallback interface.
{
    return S_FALSE;
}


/* CGPGExchExtMessageEvents::OnWriteComplete
 Called by Exchange when the data has been written to the message.
 Encrypts and signs the message if the options are set.

 Return value: S_FALSE: signals Exchange to continue calling extensions
               E_FAIL:  signals Exchange an error; the message will not be sent */
STDMETHODIMP CGPGExchExtMessageEvents::OnWriteComplete(
	LPEXCHEXTCALLBACK pEECB, // A pointer to IExchExtCallback interface.
	ULONG lFlags)
{
    HRESULT hrReturn = S_FALSE;
    LPMESSAGE pMessage = NULL;
    LPMDB pMDB = NULL;
    HWND hWnd = NULL;

    if (FAILED(pEECB->GetWindow(&hWnd)))
	hWnd = NULL;

    if (!m_bOnSubmitCalled) /* the user is just saving the message */
	return S_FALSE;

    if (m_bWriteFailed)     /* operation failed already */
	return S_FALSE;

    HRESULT hr = pEECB->GetObject(&pMDB, (LPMAPIPROP *)&pMessage);
    if (SUCCEEDED(hr))	
    {
	USES_CONVERSION;
	if (m_pExchExt->m_bEncryptWhenSending || m_pExchExt->m_bSignWhenSending)
	{
	    if (!g_gpg.EncryptAndSignMessage(hWnd, pMessage, 
					    m_pExchExt->m_bEncryptWhenSending, 
					    m_pExchExt->m_bSignWhenSending, TRUE))
	    {
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

 Return value: S_FALSE to signal Exchange to continue calling extensions.
*/
STDMETHODIMP CGPGExchExtMessageEvents::OnCheckNames(
	LPEXCHEXTCALLBACK pEECB) // A pointer to IExchExtCallback interface.
{
    return S_FALSE;
}


/* CGPGExchExtMessageEvents::OnCheckNamesComplete

 Called by Exchange when "check names" command is complete.

 Return value: S_FALSE to signal Exchange to continue calling extensions.
*/
STDMETHODIMP CGPGExchExtMessageEvents::OnCheckNamesComplete(
	LPEXCHEXTCALLBACK pEECB, // A pointer to IExchExtCallback interface.
	ULONG lFlags)
{
    return S_FALSE;
}


/* CGPGExchExtMessageEvents::OnSubmit

 Called by Exchange before the message data will be written and submitted.
 to MAPI.

 Return value: S_FALSE to signal Exchange to continue calling extensions.
*/
STDMETHODIMP CGPGExchExtMessageEvents::OnSubmit(
	LPEXCHEXTCALLBACK pEECB) // A pointer to IExchExtCallback interface.
{
    m_bOnSubmitCalled = TRUE;
    m_bWriteFailed = FALSE;
    return S_FALSE;
}


/* CGPGExchExtMessageEvents::OnSubmitComplete

 Called by Echange after the message has been submitted to MAPI.
*/
STDMETHODIMP_ (VOID) CGPGExchExtMessageEvents::OnSubmitComplete (
	LPEXCHEXTCALLBACK pEECB, // A pointer to IExchExtCallback interface-
	ULONG lFlags)
{
    m_bOnSubmitCalled = FALSE; 
}


CGPGExchExtCommands::CGPGExchExtCommands (CGPGExchExt* pParentInterface)
{ 
    m_pExchExt = pParentInterface; 
    m_lRef = 0; 
    m_lContext = 0; 
    m_nCommandID1 = 0;  
    m_nCommandID2 = 0; 
    m_nCommandID3 = 0; 
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
    if ((riid == IID_IExchExtCommands) || (riid == IID_IUnknown))
    {
        *ppvObj = (LPVOID)this;
        AddRef();
        return S_OK;
    }
    return E_NOINTERFACE;
}


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

    if (m_lContext == EECONTEXT_READNOTEMESSAGE)
    {
        pEECB->GetMenuPos(EECMDID_ToolsCustomizeToolbar, &hMenuTools, NULL, NULL, 0);
        AppendMenu(hMenuTools, MF_SEPARATOR, 0, NULL);

	CHAR szBuffer[128];
	LoadString(theApp.m_hInstance, IDS_DECRYPT_MENU_ITEM, szBuffer, 128);
        AppendMenu(hMenuTools, MF_BYPOSITION | MF_STRING, *pnCommandIDBase, szBuffer);

        m_nCommandID1 = *pnCommandIDBase;
        (*pnCommandIDBase)++;

	LoadString(theApp.m_hInstance, IDS_ADD_KEYS_MENU_ITEM, szBuffer, 128);
        AppendMenu(hMenuTools, MF_BYPOSITION | MF_STRING, *pnCommandIDBase, szBuffer);

        m_nCommandID2 = *pnCommandIDBase;
        (*pnCommandIDBase)++;

	int nTBIndex;
	HWND hwndToolbar = NULL;
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
	    tbab.hInst = theApp.m_hInstance;
	    tbab.nID = IDB_DECRYPT;
	    m_nToolbarBitmap1 = SendMessage(hwndToolbar, TB_ADDBITMAP, 1, (LPARAM)&tbab);
	    m_nToolbarButtonID2 = pTBEArray[nTBIndex].itbbBase;
	    pTBEArray[nTBIndex].itbbBase++;
	    tbab.nID = IDB_ADD_KEYS;
	    m_nToolbarBitmap2 = SendMessage(hwndToolbar, TB_ADDBITMAP, 1, (LPARAM)&tbab);
	}	
    }

    if (m_lContext == EECONTEXT_SENDNOTEMESSAGE)	
    {
        pEECB->GetMenuPos(EECMDID_ToolsCustomizeToolbar, &hMenuTools, NULL, NULL, 0);
        AppendMenu(hMenuTools, MF_SEPARATOR, 0, NULL);

	CHAR szBuffer[128];
	LoadString(theApp.m_hInstance, IDS_ENCRYPT_MENU_ITEM, szBuffer, 128);
        AppendMenu(hMenuTools, MF_BYPOSITION | MF_STRING, *pnCommandIDBase, szBuffer);

        m_nCommandID1 = *pnCommandIDBase;
        (*pnCommandIDBase)++;

	LoadString(theApp.m_hInstance, IDS_SIGN_MENU_ITEM, szBuffer, 128);
        AppendMenu(hMenuTools, MF_BYPOSITION | MF_STRING, *pnCommandIDBase, szBuffer);

        m_nCommandID2 = *pnCommandIDBase;
        (*pnCommandIDBase)++;

	LoadString(theApp.m_hInstance, IDS_ADD_STANDARD_KEY, szBuffer, 128);
        AppendMenu(hMenuTools, MF_BYPOSITION | MF_STRING, *pnCommandIDBase, szBuffer);

        m_nCommandID3 = *pnCommandIDBase;
        (*pnCommandIDBase)++;

	int nTBIndex;
	HWND hwndToolbar = NULL;

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
	    tbab.hInst = theApp.m_hInstance;
	    tbab.nID = IDB_ENCRYPT;
	    m_nToolbarBitmap1 = SendMessage(hwndToolbar, TB_ADDBITMAP, 1, (LPARAM)&tbab);
	    m_nToolbarButtonID2 = pTBEArray[nTBIndex].itbbBase;
	    pTBEArray[nTBIndex].itbbBase++;
	    tbab.nID = IDB_SIGN;
	    m_nToolbarBitmap2 = SendMessage(hwndToolbar, TB_ADDBITMAP, 1, (LPARAM)&tbab);
	}
	m_pExchExt->m_bEncryptWhenSending = g_gpg.GetEncryptDefault();
	m_pExchExt->m_bSignWhenSending = g_gpg.GetSignDefault();	
    }

    if (EECONTEXT_VIEWER == m_lContext)	
    {
        pEECB->GetMenuPos(EECMDID_ToolsCustomizeToolbar, &hMenuTools, NULL, NULL, 0);
        AppendMenu(hMenuTools, MF_SEPARATOR, 0, NULL);

	CHAR szBuffer[128];
	LoadString(theApp.m_hInstance, IDS_KEY_MANAGER, szBuffer, 128);
        AppendMenu(hMenuTools, MF_BYPOSITION | MF_STRING, *pnCommandIDBase, szBuffer);

        m_nCommandID1 = *pnCommandIDBase;
        (*pnCommandIDBase)++;

	int nTBIndex;
	HWND hwndToolbar = NULL;

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
	 
    if ((nCommandID != m_nCommandID1) && 
	(nCommandID != m_nCommandID2) && 
	(nCommandID != m_nCommandID3))
	return S_FALSE; 

    if (m_lContext == EECONTEXT_READNOTEMESSAGE)
    {
	LPMESSAGE pMessage = NULL;
	LPMDB pMDB = NULL;
	HWND hWnd = NULL;

	if (FAILED(pEECB->GetWindow(&hWnd)))
	    hWnd = NULL;
	HRESULT hr = pEECB->GetObject(&pMDB, (LPMAPIPROP *)&pMessage);
	if (SUCCEEDED(hr))
	{
	    if (nCommandID == m_nCommandID1)
		g_gpg.DecryptMessage(hWnd, pMessage, TRUE);
	    if (nCommandID == m_nCommandID2)
		g_gpg.ImportKeys(hWnd, pMessage);	
	}
	if (pMessage != NULL)
	    UlRelease(pMessage);
	if (pMDB != NULL)
	    UlRelease(pMDB);	 
    }
    if (m_lContext == EECONTEXT_SENDNOTEMESSAGE)
    {
	HWND hWnd = NULL;
	if (FAILED(pEECB->GetWindow(&hWnd)))
	    hWnd = NULL;
	if (nCommandID == m_nCommandID1)
	    m_pExchExt->m_bEncryptWhenSending = !m_pExchExt->m_bEncryptWhenSending;
	if (nCommandID == m_nCommandID2)
	    m_pExchExt->m_bSignWhenSending = !m_pExchExt->m_bSignWhenSending;
	if (nCommandID == m_nCommandID3)
	    g_gpg.AddStandardKey(hWnd);	 
    }
    if (m_lContext == EECONTEXT_VIEWER)
    {
	g_gpg.OpenKeyManager();
	g_gpg.InvalidateKeyLists();	
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
STDMETHODIMP CGPGExchExtCommands::Help(
	LPEXCHEXTCALLBACK pEECB, // The pointer to Exchange Callback Interface.
	UINT nCommandID)         // The command id.
{
	if (m_lContext == EECONTEXT_READNOTEMESSAGE)
	{
		if (nCommandID == m_nCommandID1)
		{
			CHAR szBuffer[512];
			CHAR szAppName[128];
			LoadString(theApp.m_hInstance, IDS_DECRYPT_HELP, szBuffer, 512);
			LoadString(theApp.m_hInstance, IDS_APP_NAME, szAppName, 512);
			MessageBox(m_hWnd, szBuffer, szAppName, MB_OK);
			return S_OK;
	} 
		if (nCommandID == m_nCommandID2)
		{
			CHAR szBuffer[512];
			CHAR szAppName[128];
			LoadString(theApp.m_hInstance, IDS_ADD_KEYS_HELP, szBuffer, 512);
			LoadString(theApp.m_hInstance, IDS_APP_NAME, szAppName, 512);
			MessageBox(m_hWnd, szBuffer, szAppName, MB_OK);
			return S_OK;
		} 
	}
	if (m_lContext == EECONTEXT_SENDNOTEMESSAGE)
	{
		if (nCommandID == m_nCommandID1)
		{
			CHAR szBuffer[512];
			CHAR szAppName[128];
			LoadString(theApp.m_hInstance, IDS_ENCRYPT_HELP, szBuffer, 512);
			LoadString(theApp.m_hInstance, IDS_APP_NAME, szAppName, 512);
			MessageBox(m_hWnd, szBuffer, szAppName, MB_OK);
			return S_OK;
		} 
		if (nCommandID == m_nCommandID2)
		{
			CHAR szBuffer[512];
			CHAR szAppName[128];
			LoadString(theApp.m_hInstance, IDS_SIGN_HELP, szBuffer, 512);
			LoadString(theApp.m_hInstance, IDS_APP_NAME, szAppName, 512);
			MessageBox(m_hWnd, szBuffer, szAppName, MB_OK);
			return S_OK;
		} 
		if (nCommandID == m_nCommandID3)
		{
			CHAR szBuffer[512];
			CHAR szAppName[128];
			LoadString(theApp.m_hInstance, IDS_ADD_STANDARD_KEY_HELP, szBuffer, 512);
			LoadString(theApp.m_hInstance, IDS_APP_NAME, szAppName, 512);
			MessageBox(m_hWnd, szBuffer, szAppName, MB_OK);
			return S_OK;
		} 
	}

	if (m_lContext == EECONTEXT_VIEWER)
	{
		if (nCommandID == m_nCommandID1)
		{
			CHAR szBuffer[512];
			CHAR szAppName[128];
			LoadString(theApp.m_hInstance, IDS_KEY_MANAGER_HELP, szBuffer, 512);
			LoadString(theApp.m_hInstance, IDS_APP_NAME, szAppName, 512);
			MessageBox(m_hWnd, szBuffer, szAppName, MB_OK);
			return S_OK;
		} 
	}

	return S_FALSE;
}

/////////////////////////////////////////////////////////////////////////////
// CGPGExchExtCommands::QueryHelpText
//
// Called by Exhange to get the status bar text or the tooltip of a menu item.
//
// @rdesc S_OK when it is a menu item of this plugin and the text was set;
//        otherwise S_FALSE
//
STDMETHODIMP CGPGExchExtCommands::QueryHelpText(
	UINT nCommandID,  // The command id corresponding to menu item activated.
	ULONG lFlags,     // Identifies either EECQHT_STATUS or EECQHT_TOOLTIP.
    LPTSTR pszText,   // A pointer to buffer to be populated with text to display.
	UINT nCharCnt)    // The count of characters available in psz buffer.
{
	if (m_lContext == EECONTEXT_READNOTEMESSAGE)
	{
		if (nCommandID == m_nCommandID1)
		{
			if (lFlags == EECQHT_STATUS)
				LoadString(theApp.m_hInstance, IDS_DECRYPT_STATUSBAR, pszText, nCharCnt);
  			if (lFlags == EECQHT_TOOLTIP)
				LoadString(theApp.m_hInstance, IDS_DECRYPT_TOOLTIP, pszText, nCharCnt);
			return S_OK;
		}
		if (nCommandID == m_nCommandID2)
		{
			if (lFlags == EECQHT_STATUS)
				LoadString(theApp.m_hInstance, IDS_ADD_KEYS_STATUSBAR, pszText, nCharCnt);
  			if (lFlags == EECQHT_TOOLTIP)
				LoadString(theApp.m_hInstance, IDS_ADD_KEYS_TOOLTIP, pszText, nCharCnt);
			return S_OK;
		}
	}
	if (m_lContext == EECONTEXT_SENDNOTEMESSAGE)
	{
		if (nCommandID == m_nCommandID1)
		{
			if (lFlags == EECQHT_STATUS)
				LoadString(theApp.m_hInstance, IDS_ENCRYPT_STATUSBAR, pszText, nCharCnt);
  			if (lFlags == EECQHT_TOOLTIP)
				LoadString(theApp.m_hInstance, IDS_ENCRYPT_TOOLTIP, pszText, nCharCnt);
			return S_OK;
		}
		if (nCommandID == m_nCommandID2)
		{
			if (lFlags == EECQHT_STATUS)
				LoadString(theApp.m_hInstance, IDS_SIGN_STATUSBAR, pszText, nCharCnt);
  			if (lFlags == EECQHT_TOOLTIP)
				LoadString(theApp.m_hInstance, IDS_SIGN_TOOLTIP, pszText, nCharCnt);
			return S_OK;
		}
		if (nCommandID == m_nCommandID3)
		{
			if (lFlags == EECQHT_STATUS)
				LoadString(theApp.m_hInstance, IDS_ADD_STANDARD_KEY_STATUSBAR, pszText, nCharCnt);
			return S_OK;
		}
	}

	if (m_lContext == EECONTEXT_VIEWER)
	{
		if (nCommandID == m_nCommandID1)
		{
			if (lFlags == EECQHT_STATUS)
				LoadString(theApp.m_hInstance, IDS_KEY_MANAGER_STATUSBAR, pszText, nCharCnt);
  			if (lFlags == EECQHT_TOOLTIP)
				LoadString(theApp.m_hInstance, IDS_KEY_MANAGER_TOOLTIP, pszText, nCharCnt);
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
			pTBB->idCommand = m_nCommandID1;
			pTBB->fsState = TBSTATE_ENABLED;
			pTBB->fsStyle = TBSTYLE_BUTTON;
			pTBB->dwData = 0;
			pTBB->iString = -1;
			LoadString(theApp.m_hInstance, IDS_DECRYPT_TOOLTIP, lpszDescription, nCharCnt);
			return S_OK;
		}
		if (nToolbarButtonID == m_nToolbarButtonID2)
		{
			pTBB->iBitmap = m_nToolbarBitmap2;             
			pTBB->idCommand = m_nCommandID2;
			pTBB->fsState = TBSTATE_ENABLED;
			pTBB->fsStyle = TBSTYLE_BUTTON;
			pTBB->dwData = 0;
			pTBB->iString = -1;
			LoadString(theApp.m_hInstance, IDS_ADD_KEYS_TOOLTIP, lpszDescription, nCharCnt);
			return S_OK;
		}
	}
	if (m_lContext == EECONTEXT_SENDNOTEMESSAGE)
	{
		if (nToolbarButtonID == m_nToolbarButtonID1)
		{
			pTBB->iBitmap = m_nToolbarBitmap1;             
			pTBB->idCommand = m_nCommandID1;
			pTBB->fsState = TBSTATE_ENABLED;
			if (m_pExchExt->m_bEncryptWhenSending)
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
			pTBB->idCommand = m_nCommandID2;
			pTBB->fsState = TBSTATE_ENABLED;
			if (m_pExchExt->m_bSignWhenSending)
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
			pTBB->idCommand = m_nCommandID1;
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
    if (riid == IID_IExchExtPropertySheets)
    {
        *ppvObj = (LPVOID)this;
        AddRef();
        return S_OK;
    }
    if (riid == IID_IUnknown)
    {
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
    pPSP[0].dwSize = sizeof (PROPSHEETPAGE);
    pPSP[0].dwFlags = PSP_DEFAULT | PSP_HASHELP;
    pPSP[0].hInstance = theApp.m_hInstance;
    pPSP[0].pszTemplate = MAKEINTRESOURCE(IDD_GPG_OPTIONS);
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
