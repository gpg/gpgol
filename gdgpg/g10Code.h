/* g10Code.h - g10Code Interface
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

#ifndef __G10CODE_H_
#define __G10CODE_H_

#include "resource.h"

class ATL_NO_VTABLE Cg10Code : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<Cg10Code, &CLSID_g10Code>,
	public IConnectionPointContainerImpl<Cg10Code>,
	public IDispatchImpl<Ig10Code, &IID_Ig10Code, &LIBID_GDGPGLib>
{
public:
    Cg10Code (void);
    ~Cg10Code (void);

DECLARE_REGISTRY_RESOURCEID(IDR_G10CODE)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(Cg10Code)
    COM_INTERFACE_ENTRY(Ig10Code)
    COM_INTERFACE_ENTRY(IDispatch)
    COM_INTERFACE_ENTRY(IConnectionPointContainer)
END_COM_MAP()
BEGIN_CONNECTION_POINT_MAP(Cg10Code)
END_CONNECTION_POINT_MAP()

private:
    FILE *m_logFP;
    int	  m_logLevel;
    char *m_passPhrase;
    int   m_useNoPassphrase;

    void logInfo (const char * fmt, ...);
    int  useLogging (void) { return m_logFP != NULL && m_logLevel > 0; }
    BOOL spawnGPG (string sCommand, string sPassphrase, string &sStdout,
		   string &sStderr, int nTimeoutSec, BOOL bKillProcessOnTimeout);
    int strFromBstr (char **strVal, BSTR newVal);
    int parseOptions (string &strCommand, const char * outName);
    char* data2File (const char * dataBuf,const char *fileName);
    int   file2Data (char ** dataBuf, const char * fileName, int delFile);

protected:
    BOOL m_forceMDC;
    BOOL m_forceV3SIG;
    BOOL m_noVersion;
    BOOL m_alwaysTrust;
    long m_pgpMode;
    BOOL m_expertMode;
    BOOL m_textMode;
    long m_compressLevel;
    char*m_localUser;
    BOOL m_armorMode;
    char*m_plainText;
    char*m_cipherText;
    char*m_gpgExe;
    char*m_outPut;
    char*m_commentStr;
    char*m_encryptTo;
    char*m_keyServer;
    char*m_homeDir;
    string m_recipientSet;
    
/* Ig10Code */
public:	
    STDMETHOD(DecryptFile)(/*[in]*/BSTR inFile, /*[out, retval]*/long *pvReturn);
    STDMETHOD(EncryptFile)(/*[in]*/BSTR inFile, /*[out, retval]*/long *pvReturn);
    STDMETHOD(get_HomeDir)(/*[out, retval]*/ BSTR *pVal);
    STDMETHOD(put_HomeDir)(/*[in]*/ BSTR newVal);
    STDMETHOD(get_Keyserver)(/*[out, retval]*/ BSTR *pVal);
    STDMETHOD(put_Keyserver)(/*[in]*/ BSTR newVal);
    STDMETHOD(get_ForceV3Sig)(/*[out, retval]*/ BOOL *pVal);
    STDMETHOD(put_ForceV3Sig)(/*[in]*/ BOOL newVal);
    STDMETHOD(get_ForceMDC)(/*[out, retval]*/ BOOL *pVal);
    STDMETHOD(put_ForceMDC)(/*[in]*/ BOOL newVal);
    STDMETHOD(get_EncryptoTo)(/*[out, retval]*/ BSTR *pVal);
    STDMETHOD(put_EncryptoTo)(/*[in]*/ BSTR newVal);
    STDMETHOD(Export)(BSTR keyNames, /*[out,retval]*/long *pvReturn);
    STDMETHOD(Decrypt)(/*[out, retval]*/long *pvReturn);
    STDMETHOD(put_Passphrase)(/*[in]*/BSTR newVal);
    STDMETHOD(get_Comment)(/*[out, retval]*/ BSTR *pVal);
    STDMETHOD(put_Comment)(/*[in]*/ BSTR newVal);
    STDMETHOD(get_NoVersion)(/*[out, retval]*/ BOOL *pVal);
    STDMETHOD(put_NoVersion)(/*[in]*/ BOOL newVal);
    STDMETHOD(get_Output)(/*[out, retval]*/ BSTR *pVal);
    STDMETHOD(put_Output)(/*[in]*/ BSTR newVal);
    STDMETHOD(ClearRecipient)(void);
    STDMETHOD(Encrypt)(long *pvReturn);
    STDMETHOD(AddRecipient)(BSTR name, long *pvReturn);
    STDMETHOD(get_Binary)(/*[out, retval]*/ BSTR *pVal);
    STDMETHOD(put_Binary)(/*[in]*/ BSTR newVal);
    STDMETHOD(SetLogFile)(BSTR logFile, long *pvReturn);
    STDMETHOD(SetLogLevel)(long logLevel);
    STDMETHOD(get_AlwaysTrust)(/*[out, retval]*/ BOOL *pVal);
    STDMETHOD(put_AlwaysTrust)(/*[in]*/ BOOL newVal);
    STDMETHOD(get_PGPMode)(/*[out, retval]*/ long *pVal);
    STDMETHOD(put_PGPMode)(/*[in]*/ long newVal);
    STDMETHOD(get_Expert)(/*[out, retval]*/ BOOL *pVal);
    STDMETHOD(put_Expert)(/*[in]*/ BOOL newVal);
    STDMETHOD(get_TextMode)(/*[out, retval]*/ BOOL *pVal);
    STDMETHOD(put_TextMode)(/*[in]*/ BOOL newVal);
    STDMETHOD(get_CompressLevel)(/*[out, retval]*/ long *pVal);
    STDMETHOD(put_CompressLevel)(/*[in]*/ long newVal);
    STDMETHOD(get_LocalUser)(/*[out, retval]*/ BSTR *pVal);
    STDMETHOD(put_LocalUser)(/*[in]*/ BSTR newVal);
    STDMETHOD(get_Armor)(/*[out, retval]*/ BOOL *pVal);
    STDMETHOD(put_Armor)(/*[in]*/ BOOL newVal);
    STDMETHOD(get_Ciphertext)(/*[out, retval]*/ BSTR *pVal);
    STDMETHOD(put_Ciphertext)(/*[in]*/ BSTR newVal);
    STDMETHOD(get_Plaintext)(/*[out, retval]*/ BSTR *pVal);
    STDMETHOD(put_Plaintext)(/*[in]*/ BSTR newVal);
};

#endif //__G10CODE_H_
