/* GDGPGWrapper.cpp - GDGPG wrapper implementation
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
#include "GDGPGWrapper.h"

// GDGPGs class ID
const CLSID _CLSID_GDGPG = {0x321F09FC,0xE2FD,0x409B,{0xB8,0xD1,0x60,0xFA,0x7D,0xCD,0xA5,0x31}};

CGDGPGWrapper::CGDGPGWrapper()
{	
    m_pGDGPG = NULL;

    VariantInit (&m_vtReturn);
    for (int i=0; i < PARAM_MAX; i++)
	VariantInit (&m_vtParam[i]);
}


CGDGPGWrapper::~CGDGPGWrapper()
{
    if (m_pGDGPG != NULL)
	Release();
}


BOOL
CGDGPGWrapper::CallMethod (LPDISPATCH pDisp, DISPID dwDispID, UINT cArgs, LPVARIANT pvtReturn)
{
    if (NULL == pDisp) return FALSE;

    if (cArgs > PARAM_MAX)
	return FALSE;

    DISPPARAMS dispparams = {NULL, NULL, 0, 0};
    dispparams.rgvarg = (VARIANTARG*) &m_vtParam;
    dispparams.cArgs = cArgs;

    LPVARIANT pvt = &m_vtReturn;
    if (pvtReturn != NULL)
	pvt = pvtReturn;

    HRESULT hr = pDisp->Invoke (dwDispID, IID_NULL, LOCALE_USER_DEFAULT,
			        DISPATCH_METHOD, &dispparams, pvt, NULL, NULL);
    if (FAILED (hr)) 
    {
        hr = pDisp->Invoke (dwDispID, IID_NULL, LOCALE_USER_DEFAULT,
			    DISPATCH_METHOD | DISPATCH_PROPERTYGET, &dispparams, pvt, NULL, NULL);
    }
	
    return SUCCEEDED (hr);
}


BOOL 
CGDGPGWrapper::GetProperty(DISPID dwDispID)
{
    if (NULL == m_pGDGPG) 
	return E_POINTER;

    DISPPARAMS dispparamsNoArgs = {NULL, NULL, 0, 0};

    HRESULT hr = m_pGDGPG->Invoke (dwDispID, IID_NULL, LOCALE_USER_DEFAULT,
				  DISPATCH_METHOD | DISPATCH_PROPERTYGET, &dispparamsNoArgs,
				  &m_vtReturn, NULL, NULL);

    return SUCCEEDED (hr);
}


BOOL 
CGDGPGWrapper::CreateInstance (void)
{
    if (m_pGDGPG == NULL)
    {
	HRESULT hr = CoCreateInstance (_CLSID_GDGPG, NULL, CLSCTX_ALL, 
				       IID_IDispatch, (void**) &m_pGDGPG);
	if (FAILED(hr))
	{
	    m_pGDGPG = NULL;
	    return FALSE;
	}	
    }
    return TRUE;
}


void 
CGDGPGWrapper::Release (void)
{
    if (m_pGDGPG != NULL)
    {
	m_pGDGPG->Release();
	m_pGDGPG = NULL;	
    }
}


BOOL 
CGDGPGWrapper::OpenKeyManager (int &nReturn)
{
    BOOL bSuccess = CallMethod(m_pGDGPG, 1, 0);
    if (bSuccess)
	nReturn = m_vtReturn.lVal;	
    return bSuccess;
}


BOOL 
CGDGPGWrapper::EncryptAndSignFile(HWND hWndParent, BOOL bEncrypt, BOOL bSign, 
				  BSTR strFilenameSource, BSTR strFilenameDest, 
				  BSTR strRecipient, BOOL bArmor, 
				  BOOL bEncryptWithStandardKey, int &nReturn)
{
    m_vtParam[7].vt = VT_UI4; 
    m_vtParam[7].lVal = (ULONG) hWndParent;
    m_vtParam[6].vt = VT_I4; 
    m_vtParam[6].lVal = bEncrypt ? 1 : 0;
    m_vtParam[5].vt = VT_I4; 
    m_vtParam[5].lVal = bSign ? 1 : 0;
    m_vtParam[4].vt = VT_BSTR; 
    m_vtParam[4].bstrVal = strFilenameSource;
    m_vtParam[3].vt = VT_BSTR; 
    m_vtParam[3].bstrVal = strFilenameDest;
    m_vtParam[2].vt = VT_BSTR; 
    m_vtParam[2].bstrVal = strRecipient;
    m_vtParam[1].vt = VT_I4; 
    m_vtParam[1].lVal = bArmor ? 1 : 0;
    m_vtParam[0].vt = VT_I4; 
    m_vtParam[0].lVal = bEncryptWithStandardKey ? 1 : 0;
    BOOL bSuccess = CallMethod(m_pGDGPG, 2, 8);
    if (bSuccess)
	nReturn = m_vtReturn.lVal;	
    return bSuccess;
}


BOOL 
CGDGPGWrapper::DecryptFile(HWND hWndParent, BSTR strFilenameSource, 
			   BSTR strFilenameDest, int &nReturn)
{
    m_vtParam[2].vt = VT_UI4; 
    m_vtParam[2].lVal = (ULONG) hWndParent;
    m_vtParam[1].vt = VT_BSTR; 
    m_vtParam[1].bstrVal = strFilenameSource;
    m_vtParam[0].vt = VT_BSTR; 
    m_vtParam[0].bstrVal = strFilenameDest;
    BOOL bSuccess = CallMethod(m_pGDGPG, 3, 3);
    if (bSuccess)
	nReturn = m_vtReturn.lVal;	
    return bSuccess;
}


BOOL
CGDGPGWrapper::GetGPGOutput (BSTR * hStderr)
{
    m_vtParam[0].vt = VT_BYREF|VT_BSTR;
    m_vtParam[0].pbstrVal = hStderr;
    BOOL bSuccess = CallMethod (m_pGDGPG, 11, 1);
    return bSuccess;
}


BOOL 
CGDGPGWrapper::GetGPGInfo (BSTR strFilename, BSTR *hInfo)
{
    m_vtParam[1].vt = VT_BSTR;
    m_vtParam[1].bstrVal = strFilename;
    m_vtParam[0].vt = VT_BYREF|VT_BSTR;
    m_vtParam[0].pbstrVal = hInfo;
    BOOL bSuccess = CallMethod (m_pGDGPG, 12, 2);
    return bSuccess;
}


BOOL CGDGPGWrapper::SetLogLevel (int nLevel)
{
    m_vtParam[0].vt = VT_UI4;
    m_vtParam[1].lVal = (ULONG)nLevel;
    BOOL bSuccess = CallMethod (m_pGDGPG, 14, 1);
    return bSuccess;
}


BOOL CGDGPGWrapper::SetLogFile (BSTR strLogFilename)
{
    m_vtParam[0].vt = VT_BSTR;
    m_vtParam[0].bstrVal = strLogFilename;
    BOOL bSuccess = CallMethod (m_pGDGPG, 15, 1);
    return bSuccess;
}


BOOL 
CGDGPGWrapper::ExportStandardKey (HWND hWndParent, BSTR strExportFileName, int &nReturn)
{
    m_vtParam[1].vt = VT_UI4; 
    m_vtParam[1].lVal = (ULONG) hWndParent;
    m_vtParam[0].vt = VT_BSTR; 
    m_vtParam[0].bstrVal = strExportFileName;
    BOOL bSuccess = CallMethod(m_pGDGPG, 4, 2);
    if (bSuccess)
	nReturn = m_vtReturn.lVal;	
    return bSuccess;
}


BOOL 
CGDGPGWrapper::VerifyDetachedSignature (HWND hParent, BSTR strFilenameText, BSTR strFilenameSig, int &nReturn)
{
    m_vtParam[2].vt = VT_UI4;
    m_vtParam[2].lVal = (ULONG)hParent;
    m_vtParam[1].vt = VT_BSTR;
    m_vtParam[1].bstrVal = strFilenameText;
    m_vtParam[0].vt = VT_BSTR;
    m_vtParam[0].bstrVal = strFilenameSig;
    
    BOOL bSuccess = CallMethod (m_pGDGPG, 13, 3);
    if (bSuccess)
	nReturn = m_vtReturn.lVal;
    return bSuccess;
}


BOOL 
CGDGPGWrapper::ImportKeys(HWND hWndParent, BSTR strImportFilename, BOOL bShowMessage, 
			  int &nEditCount, int &nImportCount, int &nUnchangeCount, int &nReturn)
{
    m_vtParam[5].vt = VT_UI4; 
    m_vtParam[5].lVal = (ULONG) hWndParent;
    m_vtParam[4].vt = VT_BSTR; 
    m_vtParam[4].bstrVal = strImportFilename;
    m_vtParam[3].vt = VT_I4; 
    m_vtParam[3].lVal = bShowMessage ? 1 : 0;
    m_vtParam[2].vt = VT_BYREF | VT_I4; 
    m_vtParam[2].plVal = (long*) &nEditCount;
    m_vtParam[1].vt = VT_BYREF | VT_I4; 
    m_vtParam[1].plVal = (long*) &nImportCount;
    m_vtParam[0].vt = VT_BYREF | VT_I4; 
    m_vtParam[0].plVal = (long*) &nUnchangeCount;
    BOOL bSuccess = CallMethod(m_pGDGPG, 5, 6);
    if (bSuccess)
	nReturn = m_vtReturn.lVal;	
    return bSuccess;
}


BOOL 
CGDGPGWrapper::SetStorePassphraseTime(int nSeconds)
{
    m_vtParam[0].vt = VT_I4; 
    m_vtParam[0].lVal = nSeconds;
    return CallMethod(m_pGDGPG, 6, 1);
}


BOOL 
CGDGPGWrapper::InvalidateKeyLists()
{
    return CallMethod(m_pGDGPG, 7, 0);
}


BOOL
CGDGPGWrapper::Options(HWND hWndParent)
{
    m_vtParam[0].vt = VT_UI4; 
    m_vtParam[0].lVal = (ULONG) hWndParent;
    return CallMethod(m_pGDGPG, 8, 1);
}


BOOL
CGDGPGWrapper::EncryptAndSignNextFile(HWND hWndParent, BSTR strFilenameSource, 
				BSTR strFilenameDest, BOOL bArmor, int &nReturn)
{
    m_vtParam[3].vt = VT_UI4; 
    m_vtParam[3].lVal = (ULONG) hWndParent;
    m_vtParam[2].vt = VT_BSTR;
    m_vtParam[2].bstrVal = strFilenameSource;
    m_vtParam[1].vt = VT_BSTR; 
    m_vtParam[1].bstrVal = strFilenameDest;
    m_vtParam[0].vt = VT_I4; 
    m_vtParam[0].lVal = bArmor ? 1 : 0;
    BOOL bSuccess = CallMethod(m_pGDGPG, 9, 4);
    if (bSuccess)
	nReturn = m_vtReturn.lVal;	
    return bSuccess;
}


BOOL 
CGDGPGWrapper::DecryptNextFile(HWND hWndParent, BSTR strFilenameSource, 
				    BSTR strFilenameDest, int &nReturn)
{
    m_vtParam[2].vt = VT_UI4; 
    m_vtParam[2].lVal = (ULONG) hWndParent;
    m_vtParam[1].vt = VT_BSTR; 
    m_vtParam[1].bstrVal = strFilenameSource;
    m_vtParam[0].vt = VT_BSTR; 
    m_vtParam[0].bstrVal = strFilenameDest;
    BOOL bSuccess = CallMethod(m_pGDGPG, 10, 3);
    if (bSuccess)
	nReturn = m_vtReturn.lVal;	
    return bSuccess;
}
