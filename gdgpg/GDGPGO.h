/* GDGPGO.h - declaration of the COM object class (GDGPG)
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

#ifndef __GDGPG_H_
#define __GDGPG_H_

#include "resource.h"       // main symbols
#include "key.h"


/* CGDGPG - The CGDGPG class implements the IGDGPG interface. */
class ATL_NO_VTABLE CGDGPG : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CGDGPG, &CLSID_GDGPG>,
	public IDispatchImpl<IGDGPG, &IID_IGDGPG, &LIBID_GDGPGLib>
{

public:
    CGDGPG();
    ~CGDGPG();

DECLARE_REGISTRY_RESOURCEID(IDR_GDGPG)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CGDGPG)
    COM_INTERFACE_ENTRY(IGDGPG)
    COM_INTERFACE_ENTRY(IDispatch)
END_COM_MAP()


public:
    STDMETHOD(DecryptNextFile)(/*[in]*/ ULONG hWndParent, /*[in]*/ BSTR strFilenameSource, /*[in]*/ BSTR strFilenameDest, /*[out, retval]*/ int* pvReturn);
    STDMETHOD(EncryptAndSignNextFile)(/*[in]*/ ULONG hWndParent, /*[in]*/ BSTR strFilenameSource, /*[in]*/ BSTR strFilenameDest, /*[in]*/ BOOL bArmor, /*[out, retval]*/ int* pvReturn);
    STDMETHOD(Options)(/*[in]*/ ULONG hWndParent);
    STDMETHOD(InvalidateKeyLists)();
    STDMETHOD(SetStorePassphraseTime)(/*[in]*/ int nSeconds);
    STDMETHOD(ImportKeys)(/*[in]*/ ULONG hWndParent, /*[in]*/ BSTR strImportFilename, /*[in]*/ BOOL bShowMessage, /*[out]*/ int* pvEditCount, /*[out]*/ int* pvImportCount, /*[out]*/ int* pvUnchangeCnt, /*[out, retval]*/ int* pvReturn);
    STDMETHOD(ExportStandardKey)(/*[in]*/ ULONG hWndParent, /*[in]*/ BSTR strExportFileName, /*[out, retval]*/ int* pvReturn);
    STDMETHOD(DecryptFile)(/*[in]*/ ULONG hWndParent, /*[in]*/ BSTR strFilenameSource, /*[in]*/ BSTR strFilenameDest, /*[out, retval]*/ int* pvReturn);
    STDMETHOD(VerifyDetachedSignature) (/*[in]*/ULONG hWndParent, /*[in]*/BSTR strFilenameText, /*[in]*/BSTR strFilenameSig, /*[out]*/int *pvReturn);
    STDMETHOD(EncryptAndSignFile)(/*[in]*/ ULONG hWndParent, /*[in]*/ BOOL bEncrypt, /*[in]*/ BOOL bSign, /*[in]*/ BSTR strFilenameSource, /*[in]*/ BSTR strFilenameDest, /*[in]*/ BSTR strRecipient, /*[in]*/ BOOL bArmor, /*[in]*/ BOOL bEncryptWithStandardKey, /*[out, retval]*/ int* pvReturn);
    STDMETHOD(OpenKeyManager)(/*[out, retval]*/ int* pvReturn);
    STDMETHOD(GetGPGOutput) (/*[out, retval]*/BSTR *hStdErr);
    STDMETHOD(GetGPGInfo) (/*[in]*/BSTR strFilename, /*[out]*/BSTR *hInfo);
    STDMETHOD(SetLogLevel) (/*[in]*/ULONG nLevel);
    STDMETHOD(SetLogFile)(/*[in]*/BSTR strLogFilename, /*[out, retval]*/int* pvReturn);


protected:
    string       m_sGPGExe;   
    string       m_sKeyManagerExe;   
    vector<CKey> m_keyList;
    CKey         m_keyDefault;
    int          m_nStorePassphraseTime;
    string       m_sPassphraseStore;
    DWORD        m_dwPasspharseTime;
    string       m_sPassphraseInvalid;
    FILE        *m_LogFP;
    int		 m_LogLevel;

    /* parameter for calls of EncryptAndSignNextFile() */
    string m_sEncryptCommandNextFile;;
    string m_sEncryptPassphraseNextFile;
    string m_sDecryptPassphraseNextFile;

    /* GPG error output */
    string	  m_gpgStdErr;

private:
    int UseLogging (void) { return m_LogFP != NULL && m_LogLevel > 0; }
    int LogInfo (const char * fmt, ...);

protected:
    
    BOOL CallGPG (string sCommand, string sPassphrase, string &sResult, string &sError, int nTimeoutSec = 30, BOOL bKillProcessOnTimeout = TRUE);
    BOOL GenerateKeyList (void);
    void SetDefaultRecipients (vector<CKey> &keyList, string sRecipients);
    BOOL GetDefaultKey (void);
    BOOL CheckPassphrase (string sPassphrase);
    BOOL GetDecryptInfo (string sFilename, BOOL &bIsEncrypted, BOOL &bIsSigned, string &sDecryptKeys, int &nUnknownKeys);
    void StorePassphrase (string sPassphrase);

    string GetStoredPassphrase (void);
    void ReadOptions (void);
    void WriteOptions(void);
};

#endif //__GDGPG_H_
