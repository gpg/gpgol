/* GDGPGWrapper.h - GDGPG wrapper definition 
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

#ifndef INC_GDGPGWRAPPER_H
#define INC_GDGPGWRAPPER_H

// max. number of parameters for CallMethod() 
#define PARAM_MAX 10

/////////////////////////////////////////////////////////////////////////////
// CGDGPGWrapper
//
// The CGDGPGWrapper class wraps the calls to the GDGPG object.
//
class CGDGPGWrapper
{
// constructor/destructor
public:
    CGDGPGWrapper();  
    ~CGDGPGWrapper();

// attributes
protected:
    LPDISPATCH   m_pGDGPG;         // Points to the GDGPG interface of the GDGPG object.
    VARIANT m_vtReturn;            // The return value; used in CallMethod().
    VARIANT m_vtParam[PARAM_MAX];  // The parameters; used in CallMethod().

// implementation
public:
    BOOL CreateInstance();         // Connects with the GDGPG object.
    void Release();                // Disconnects from the GDGPG object.
    LPDISPATCH GetDispatch() { return m_pGDGPG; };  // Returns the dispatch interface of the GDGPG object.

    // wrappers of the GDGPG methods
    BOOL OpenKeyManager(int &nReturn);
    BOOL EncryptAndSignFile(HWND hWndParent, BOOL bEncrypt, BOOL bSign, BSTR strFilenameSource, BSTR strFilenameDest, BSTR strRecipient, BOOL bArmor, BOOL bEncryptWithStandardKey, int &nReturn);
    BOOL EncryptAndSignNextFile(HWND hWndParent, BSTR strFilenameSource, BSTR strFilenameDest, BOOL bArmor, int &nReturn);
    BOOL DecryptFile(HWND hWndParent, BSTR strFilenameSource, BSTR strFilenameDest, int &nReturn);
    BOOL DecryptNextFile(HWND hWndParent, BSTR strFilenameSource, BSTR strFilenameDest, int &nReturn);
    BOOL ExportStandardKey(HWND hWndParent, BSTR strExportFileName, int &nReturn);
    BOOL ImportKeys(HWND hWndParent, BSTR strImportFilename, BOOL bShowMessage, int &nEditCount, int &nImportCount, int &nUnchangeCount, int &nReturn);
    BOOL VerifyDetachedSignature (HWND hParent, BSTR strFilenameText, BSTR strFilenameSig, int &nReturn);
    BOOL SetStorePassphraseTime(int nSeconds);
    BOOL InvalidateKeyLists();
    BOOL Options(HWND hWndParent);
    BOOL GetGPGOutput (BSTR * hStderr);
    BOOL GetGPGInfo (BSTR strFilename, BSTR *hInfo);
    BOOL SetLogLevel (int nLevel);
    BOOL SetLogFile (BSTR strLogFilename);

protected:
    // Calls the specified method in the GDGPG object.
    BOOL CallMethod(LPDISPATCH pDisp, DISPID dwDispID, UINT cArgs, LPVARIANT pvtReturn = NULL);

    // Gets the specified property from the GDGPG object.
    BOOL GetProperty(DISPID dwDispID);
};

#endif // INC_GDGPGWRAPPER_H