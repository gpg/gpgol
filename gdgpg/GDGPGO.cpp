/* GDGPGO.cpp - implementation of the COM object class (GDGPG)
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
#include "GDGPG.h"
#include "GDGPGO.h"
#include "SelectKeyDlg.h"
#include "PassphraseDlg.h"
#include "PassphraseDecryptDlg.h"
#include "OptionsDlg.h"
#include "gdgpgdef.h"

/* OpenKeyManager

 Opens the key manager. The name of the key manager file was read from the 
 registry.

 Return value: The error code:
               GDGPG_SUCCESS:                    success
               GDGPG_ERR_KEY_MANAGER_NOT_EXISTS: the key manager does not exists
*/
STDMETHODIMP CGDGPG::OpenKeyManager (int *pvReturn)
{
    PROCESS_INFORMATION pInfo;
    STARTUPINFO sInfo;
    BOOL bSuccess;

    *pvReturn = GDGPG_SUCCESS;

    /* create startup info for the gpg process */
    memset(&sInfo, 0, sizeof(sInfo));
    sInfo.cb = sizeof(STARTUPINFO); 
    sInfo.dwFlags = STARTF_USESHOWWINDOW; 
    sInfo.wShowWindow = SW_SHOW;

    bSuccess = CreateProcess (NULL, (char*) m_sKeyManagerExe.c_str(),
			      NULL, NULL, TRUE, CREATE_DEFAULT_ERROR_MODE,	
			      NULL, NULL, &sInfo, &pInfo);
    if (bSuccess)
    {	
	/* XXX: why is this disabled -ts */
	/* ::WaitForSingleObject(pInfo.hProcess, INFINITE);*/
	CloseHandle(pInfo.hProcess);
	CloseHandle(pInfo.hThread);	
    }
    else
    {
	TCHAR s[256], sCaption[256];
	LoadString(_Module.m_hInstResource, IDS_ERR_OPEN_KEY_MANAGER, s, sizeof(s));
	LoadString(_Module.m_hInstResource, IDS_APP_NAME, sCaption, sizeof(s));
	::MessageBox(NULL, s, sCaption, MB_ICONEXCLAMATION);
	*pvReturn = GDGPG_ERR_KEY_MANAGER_NOT_EXISTS;	
    }
    return S_OK;
}

/* EncryptAndSignFile

 Encrypts and signs the specified file.

 Return value: The error code:
               GDGPG_SUCCESS:          success
               GDGPG_ERR_CANCEL:       canceled by the user
               GDGPG_ERR_GPG_FAILED:   call to gpg failed
               GDGPG_ERR_NO_DEST_FILE: the destination was not created by gpg (gpg failed)
*/
STDMETHODIMP CGDGPG::EncryptAndSignFile(
	ULONG hWndParent,               // The handle of the parent window for messages and dialog.
	BOOL bEncrypt,                  // Indicates whether to encrypt the message.         
	BOOL bSign,                     // Indicates whether to sign the messages.
	BSTR strFilenameSource,         // The source file name.
	BSTR strFilenameDest,           // The destination file name.
	BSTR strRecipient,              // The recipient(s).
	BOOL bArmor,                    // Indicates whether to produce ASCII output.
	BOOL bEncryptWithStandardKey,   // Indicates whether to encrypt with the standard key too.
	int *pvReturn)                  // The return value.
{
	
    USES_CONVERSION;
    m_sEncryptPassphraseNextFile = "";
    string sPassphrase;
    BOOL bPassphraseOK = FALSE;
    string sResult, sError;

    if (hWndParent == NULL)
	hWndParent = (ULONG) ::GetActiveWindow();

    *pvReturn = GDGPG_SUCCESS;
    ::DeleteFile(OLE2A(strFilenameDest));

    GetDefaultKey ();
    GenerateKeyList ();
    SetDefaultRecipients (m_keyList, OLE2A(strRecipient));

    /* get passphrase for signing */
    while (bSign && !bPassphraseOK) 	
    {
	TCHAR s[256];
	sPassphrase = GetStoredPassphrase ();
	if ((sPassphrase != "") && CheckPassphrase (sPassphrase))
	    bPassphraseOK = TRUE;    
	else 
	{
	    CPassphraseDlg dlgPassphrase;
	    string sMess;

	    LoadString(_Module.m_hInstResource, IDS_SIGN_KEY, s, sizeof(s));
	    sMess = s;
	    sMess += "\n  ";
	    sMess += m_keyDefault.m_sUser;
	    dlgPassphrase.SetMessage (sMess);
	    if (dlgPassphrase.DoModal ((HWND) hWndParent) != IDOK)
	    {
		*pvReturn = GDGPG_ERR_CANCEL;
		return S_OK;	
	    }
	    sPassphrase = dlgPassphrase.GetPassphrase ();
	    if (sPassphrase != "")
		bPassphraseOK = CheckPassphrase (sPassphrase);
	    else
		bPassphraseOK = TRUE;
	}
	if (!bPassphraseOK)
	{
	    TCHAR sCaption[256];
	    LoadString (_Module.m_hInstResource, IDS_WRONG_PASSPHRASE, s, sizeof(s));
	    LoadString (_Module.m_hInstResource, IDS_APP_NAME, sCaption, sizeof(s));
	    ::MessageBox ((HWND) hWndParent, s, sCaption, 0);   
	}
	else
	{
	    StorePassphrase(sPassphrase);	
	    m_sEncryptPassphraseNextFile = sPassphrase;   
	}
    }
	
	
    if (bEncrypt)
    {
	CSelectKeyDlg dlg;

	dlg.SetKeyList(&m_keyList);
	if (bEncryptWithStandardKey)
	    dlg.SetExceptionKeyID(m_keyDefault.m_sIDPub);
	if (dlg.DoModal((HWND) hWndParent) != IDOK)
	{
	    *pvReturn = GDGPG_ERR_CANCEL;
	    return S_OK;   
	}	
    }
    string sCommand = "--batch --yes ";

    if (bSign)
    {
	if (bEncrypt)
	    sCommand += "--sign --passphrase-fd 0 ";
	else
	    sCommand += "--clearsign --passphrase-fd 0 ";	
    }

    if (bEncrypt) 
    {
	sCommand += "--always-trust ";
	sCommand += "--encrypt ";

	for (vector<CKey>::iterator i = m_keyList.begin(); i != m_keyList.end(); i++) 
	{
	    if (i->m_bSelected) 
	    {
		sCommand += " --recipient ";
		/*sCommand += i->GetKeyID (-1);*/
		sCommand += i->m_sAddress;
	    }
	    if (bEncryptWithStandardKey && (m_keyDefault.GetKeyID (-1) != "")) 
	    {
		sCommand += " --recipient ";
		/*sCommand += m_keyDefault.GetKeyID (-1);*/
		sCommand += m_keyDefault.m_sAddress;
	    }
	    sCommand += " ";
	}
    }
	
    m_sEncryptCommandNextFile = sCommand;

    if (bArmor)	
	sCommand += "--armor ";

    sCommand += "--output \"";
    sCommand += OLE2A (strFilenameDest);
    sCommand += "\" ";

    sCommand += "\"";
    sCommand += OLE2A (strFilenameSource);
    sCommand += "\"";

    if (!CallGPG (sCommand, sPassphrase, sResult, sError))
    {
	LogInfo ("EncryptAndSignFile failed (gpg error).\n");
	*pvReturn = GDGPG_ERR_GPG_FAILED;
    }
    else 
    {
	/* check whether the destintion file exists */
	FILE* file = fopen (OLE2A (strFilenameDest), "rb");
	if (file == NULL)
	{
	    LogInfo ("error, no destination file: %s\n", OLE2A (strFilenameDest));
	    *pvReturn = GDGPG_ERR_NO_DEST_FILE;
	}
	else
	    fclose (file);	
    }

    if (*pvReturn != GDGPG_SUCCESS) 
    {
	TCHAR s[256], sCaption[256];
	LoadString(_Module.m_hInstResource, IDS_ERR_CALL_GPG_FAILED, s, sizeof(s));
	LoadString(_Module.m_hInstResource, IDS_APP_NAME, sCaption, sizeof(s));
	::MessageBox((HWND) hWndParent, s, sCaption, MB_ICONEXCLAMATION);	
    }
    return S_OK;
}


STDMETHODIMP CGDGPG::VerifyDetachedSignature (
    ULONG hWndParent,
    BSTR strFilenameText,
    BSTR strFilenameSig,
    int *pvReturn)
{
    USES_CONVERSION;

    int sig_okay = -1;
    string sCommand = "";
    string sRes = "", sErr = "";

    *pvReturn = GDGPG_SUCCESS;

    sCommand = "--status-fd 1 ";
    sCommand += "--verify ";
    sCommand += OLE2A (strFilenameSig);
    sCommand += " ";
    sCommand += OLE2A (strFilenameText);

    if (!CallGPG (sCommand, "", sRes, sErr))
    {
	LogInfo ("VerifyDetachedSignature failed (gpg error).\n");
	*pvReturn = GDGPG_ERR_GPG_FAILED;
	return S_FALSE;
    }

    if (sRes.find ("[GNUPG:] GOODSIG") != -1 ||
	sRes.find ("[GNUPG:] EXPKEYSIG") != -1)
	sig_okay = 1;
    else if (sRes.find ("[GNUPG:] BADSIG") != -1)
	sig_okay = 0;
    else
    {
	if (sRes.find ("[GNUPG:] ERRSIG") != -1)
	    MessageBox ((HWND)hWndParent, "Could not check signature", "GDGPG", MB_OK);
	*pvReturn = GDGPG_ERR_NOT_CRYPTED_OR_SIGNED;
	return S_FALSE;
    }

    ::MessageBox ((HWND)hWndParent, sig_okay==1? "Good signature" : "BAD signature",
		  "GDGPG", MB_OK);

    return S_OK;
}

/* DecryptFile
 Decrypts the specified file and checks the signature.

 Return value: The error code:
               GDGPG_SUCCESS:                   success
               GDGPG_ERR_CANCEL:                canceled by the user
               GDGPG_ERR_GPG_FAILED:            call to gpg failed
               GDGPG_ERR_NO_DEST_FILE:          the destination was not created by gpg (gpg failed)
               GDGPG_ERR_NOT_CRYPTED_OR_SIGNED: the message is neigther encrypted nor signed (error message always shown)
               GDGPG_ERR_CHECK_SIGNATURE:       error while checking the signature
               GDGPG_ERR_NO_VALID_DECRYPT_KEY   none of the possible public keys in the key ring
*/
STDMETHODIMP CGDGPG::DecryptFile (
	ULONG hWndParent,        // The handle of the parent window for messages and dialog.
	BSTR strFilenameSource,  // The source file name.
	BSTR strFilenameDest,    // The destination file name.
	int *pvReturn)           // The return value.
{
	
    USES_CONVERSION;

    if (hWndParent == NULL)
	hWndParent = (ULONG) ::GetActiveWindow ();


    *pvReturn = GDGPG_SUCCESS;
    ::DeleteFile (OLE2A(strFilenameDest));

    TCHAR sCaption[256];
    LoadString (_Module.m_hInstResource, IDS_APP_NAME, sCaption, sizeof(sCaption));

    GenerateKeyList ();

    /* check file status */
    BOOL bEncrypted = FALSE;
    BOOL bSigned = FALSE;

    string sDecryptKeys;
    int nUnknownKeys = 0;
    GetDecryptInfo (OLE2A(strFilenameSource), bEncrypted, bSigned, sDecryptKeys, nUnknownKeys);
    if (!bEncrypted && !bSigned)
    {
	TCHAR s[256];
	LoadString(_Module.m_hInstResource, IDS_NOT_CRYPTED_OR_SIGNED, s, sizeof(s));
	::MessageBox((HWND) hWndParent, s, sCaption, MB_ICONINFORMATION);
	*pvReturn = GDGPG_ERR_NOT_CRYPTED_OR_SIGNED;
	return S_OK;	
    }

    if (bEncrypted && (sDecryptKeys == ""))
    {
	TCHAR s[256];
	LoadString(_Module.m_hInstResource, IDS_NO_VALID_DECRYPT_KEY, s, sizeof(s));
	::MessageBox((HWND) hWndParent, s, sCaption, MB_ICONINFORMATION);
	*pvReturn = GDGPG_ERR_NO_VALID_DECRYPT_KEY;
	return S_OK;	
    }

    BOOL bSuccess = FALSE;
    BOOL bTryWithStorePassphrase = TRUE;
    do
    {
	string sPassphrase;
	BOOL bStoredPassphrase = FALSE;
	if (bTryWithStorePassphrase)
	{
	    sPassphrase = GetStoredPassphrase ();
	    if (sPassphrase != "")
		bStoredPassphrase = TRUE;
	    bTryWithStorePassphrase = FALSE;
	}
		
	if (bEncrypted && !bStoredPassphrase)
	{
	    CPassphraseDecryptDlg dlg;
	    dlg.SetKeys (sDecryptKeys, nUnknownKeys);
	    if (dlg.DoModal ((HWND) hWndParent) != IDOK)
	    {
		*pvReturn = GDGPG_ERR_CANCEL;
		return S_OK;	
	    }
	    sPassphrase = dlg.GetPassphrase();
	}
	m_sDecryptPassphraseNextFile = sPassphrase;

	string sCommand = "--output \"";
	sCommand += OLE2A(strFilenameDest);
	sCommand += "\" --yes --status-fd 1 ";
	sCommand += "--decrypt";
	if (bEncrypted)
	    sCommand += " --passphrase-fd 0";
	sCommand += " \"";
	sCommand += OLE2A(strFilenameSource);
	sCommand += "\"";

	string sResult, sError;
	bSuccess = CallGPG(sCommand, sPassphrase, sResult, sError);
	if (!bSuccess)
	    *pvReturn = GDGPG_ERR_GPG_FAILED;
	else
	{
	    /* check whether file exists */
	    FILE* file = fopen(OLE2A(strFilenameDest), "rb");
	    if (file == NULL)
	    {
		LogInfo ("error, no destination file: %s\n", OLE2A (strFilenameDest));
		bSuccess = FALSE;
		*pvReturn = GDGPG_ERR_NO_DEST_FILE;
	    }
	    else
		fclose(file);	
	}
	
	if (bSuccess)
	{
	    if (!bSigned) 
	    {
		if ((sResult.find("[GNUPG:] GOODSIG") != -1) ||	
		    (sResult.find("[GNUPG:] BADSIG") != -1) ||
		    (sResult.find("[GNUPG:] EXPKEYSIG") != -1) ||
		    (sResult.find("[GNUPG:] ERRSIG") != -1))
		    bSigned = TRUE;	
	    }
			
	    if (bSigned)
	    {
		string sSign;
		sResult += '\n';
		while (sResult.find('\n') != -1)
		{
		    string sLine = sResult.substr(0, sResult.find('\n'));
		    sResult = sResult.substr(sResult.find('\n')+1);
		    if (sLine.find("[GNUPG:] GOODSIG") != -1 ||
			sLine.find("[GNUPG:] EXPKEYSIG") != -1)
		    {		
			int id;
			TCHAR s[256];
			id = IDS_GOODSIG;
			if (sLine.find ("EXPKEYSIG") != -1)
			    id = IDS_EXPGOODSIG;
			sLine = sLine.substr(17);
			if (sLine.find(" ") != -1)
			    sLine = sLine.substr(sLine.find(" ")+1);
			if (sSign != "")
			    sSign += "\n";
			LoadString(_Module.m_hInstResource, id, s, sizeof(s)-1);
			sSign += s ;
			sSign += " " + sLine;	
		    }
					
		    if (sLine.find("[GNUPG:] BADSIG") != -1)
		    {
			sLine = sLine.substr(16);
			if (sLine.find(" ") != -1)
			    sLine = sLine.substr(sLine.find(" ")+1);
			if (sSign != "")
			    sSign += "\n";
			sSign += "BAD signature: ";
			TCHAR s[256];
			LoadString(_Module.m_hInstResource, IDS_BADSIG, s, sizeof(s));
			sSign += s;
			sSign += " " + sLine;	
		    }
		}

		if (sSign == "") /* ERRSIG */
		{
		    TCHAR s[256];
		    LoadString(_Module.m_hInstResource, IDS_ERR_CHECK_SIGN, s, sizeof(s));
		    sSign = s;
		}
		::MessageBox((HWND) hWndParent, sSign.c_str(), sCaption, MB_ICONINFORMATION);	
	    }
	    bSuccess = TRUE;
	    if (bEncrypted)
		StorePassphrase (sPassphrase);	
	}

		
	if (!bSuccess)
	{
	    TCHAR s[256];
	    if (bEncrypted)
	    {
		if (!bStoredPassphrase)
		    LoadString(_Module.m_hInstResource, IDS_DECRYPT_FAILED, s, sizeof(s));	
	    }
	    else
		LoadString(_Module.m_hInstResource, IDS_ERR_CHECK_SIGN, s, sizeof(s));
	    ::MessageBox((HWND) hWndParent, s, sCaption, MB_ICONINFORMATION);
	    if (!bEncrypted)
		*pvReturn = GDGPG_ERR_CHECK_SIGNATURE;	
	}
    } while (bEncrypted && !bSuccess);

    if ((*pvReturn == 2) || (*pvReturn == 3))
    {
	TCHAR s[256], sCaption[256];
	LoadString(_Module.m_hInstResource, IDS_ERR_CALL_GPG_FAILED, s, sizeof(s));	
	LoadString(_Module.m_hInstResource, IDS_APP_NAME, sCaption, sizeof(s));
	::MessageBox((HWND) hWndParent, s, sCaption, MB_ICONEXCLAMATION);	
    }

    return S_OK;
}


STDMETHODIMP CGDGPG::GetGPGOutput (BSTR *hStderr)
{
    CComBSTR a (m_gpgStdErr.c_str ());
    *hStderr = a;
    return S_OK;
}


STDMETHODIMP CGDGPG::GetGPGInfo (BSTR strFilename, BSTR *hInfo)
{    
    USES_CONVERSION;
    string s = "";
    string res = "", err = "";

    s += "--batch --logger-fd=1 ";
    s += OLE2A (strFilename);
    if (!CallGPG (s, "", res, err))
	LogInfo ("GetGPGInfo failed (gpg error).\n");

    CComBSTR a((char *)res.c_str ());
    *hInfo = a;

    return S_OK;
}


STDMETHODIMP CGDGPG::SetLogLevel (ULONG nLevel)
{ 
    m_LogLevel = nLevel; 
    return S_OK;
}


STDMETHODIMP CGDGPG::SetLogFile (BSTR strLogFilename, int *pvReturn)
{
    USES_CONVERSION;

    *pvReturn = 0;
    if (m_LogFP)
    {
	fclose (m_LogFP);
	m_LogFP = NULL;
    }

    m_LogFP = fopen (OLE2A (strLogFilename), "a+b");
    if (m_LogFP != NULL && m_LogLevel < 1)
	m_LogLevel = 1;
    return S_OK;
}


/* ExportStandardKey

 Exports the standard key to the specified file.
 
 Return value: The error code:
               GDGPG_SUCCESS:                   success
               GDGPG_ERR_CANCEL:                canceled by the user
               GDGPG_ERR_GPG_FAILED:            call to gpg failed
               GDGPG_ERR_NO_DEST_FILE:          the destination was not created by gpg (gpg failed)
               GPGPG_ERR_NO_STANDARD_KEY:       no standrd key available
*/
STDMETHODIMP CGDGPG::ExportStandardKey(
	ULONG hWndParent,        // The handle of the parent window for messages and dialog.
	BSTR strExportFileName,  // The destination file name.
	int *pvReturn)           // The return value.
{

    *pvReturn = GDGPG_SUCCESS;

    m_keyDefault.m_sUser = ""; // invalidate
    if (!GetDefaultKey())
    {
	*pvReturn = GPGPG_ERR_NO_STANDARD_KEY;
	return S_OK;	
    }

    USES_CONVERSION;
    ::DeleteFile(OLE2A(strExportFileName));
    string sCommand = "--batch --yes --output \"";
    sCommand += OLE2A(strExportFileName);
    sCommand += "\" --armor --export ";
    sCommand += m_keyDefault.m_sIDPub;
    string sResult, sError;

    if (!CallGPG(sCommand, "", sResult, sError))
    {
	LogInfo ("ExportStandardKey failed (gpg error).\n");
	*pvReturn = GDGPG_ERR_GPG_FAILED;
    }

    FILE* file = fopen (OLE2A (strExportFileName), "rb");
    if (file == NULL)
    {
	LogInfo ("error, no destination file: %s\n", OLE2A (strExportFileName));
	*pvReturn = GDGPG_ERR_NO_DEST_FILE;
    }
    else
	fclose (file);

    return S_OK;
}

/////////////////////////////////////////////////////////////////////////////
// ImportKeys
//
// Import all keys from the specified file.
//
// Return value: The error code:
//               GDGPG_SUCCESS:                   success
//               GDGPG_ERR_GPG_FAILED:            call to gpg failed
//
STDMETHODIMP CGDGPG::ImportKeys(
	ULONG hWndParent,        // The handle of the parent window for messages and dialog.
	BSTR strImportFilename,  // The file name.
	BOOL bShowMessage,       // not used
	int *pvEditCount,        // The number of edited keys (return value).
	int *pvImportCount,      // The number of imported keys (return value).
	int *pvUnchangeCnt,      // The number of unchanged keys (return value).
	int *pvReturn)           // The return value.
{
	
    *pvReturn = 0;
    *pvEditCount = 0;
    *pvImportCount = 0;
    *pvUnchangeCnt = 0;

	
    USES_CONVERSION;
    string sCommand = "--batch --yes --status-fd 1 --import \"";
    sCommand += OLE2A(strImportFilename);
    sCommand += "\"";
    string sResult, sError;

    if (!CallGPG(sCommand, "", sResult, sError))
    {
	LogInfo ("ImportKeys failed (gpg error).\n");
	*pvReturn = GDGPG_ERR_GPG_FAILED;
    }

    sResult += '\n';
    while (sResult.find('\n') != -1)
    {
	string sLine = sResult.substr(0, sResult.find('\n'));
	sResult = sResult.substr(sResult.find('\n')+1);
	if (sLine.substr(0,8) == "[GNUPG:]")		
	{
	    if (sLine.find("IMPORT_RES") != -1)
		sLine = sLine.substr(sLine.find("IMPORT_RES")+11);
	    for (int i=0; i<5; i++)
	    {
		if (sLine.find(" ") > 0)
		{
		    int n = atoi(sLine.substr(0, sLine.find(" ")).c_str());
		    sLine = sLine.substr(sLine.find(" ")+1);
		    switch (i)
		    {
		    case 0: *pvEditCount = n; break;
		    case 2: *pvImportCount = n; break;
		    case 4: *pvUnchangeCnt = n; break;
		    }
		}	
	    }
	}	
    }
    return S_OK;
}

/////////////////////////////////////////////////////////////////////////////
// SetStorePassphraseTime
//
// Sets the time to store the passphrase (in seconds).
//
STDMETHODIMP CGDGPG::SetStorePassphraseTime(
	int nSeconds)  // The number og seconds.
{
    if (nSeconds < m_nStorePassphraseTime)
	m_sPassphraseStore = m_sPassphraseInvalid;	
    m_nStorePassphraseTime = nSeconds;

    return S_OK;
}

/////////////////////////////////////////////////////////////////////////////
// InvalidateKeyLists
//
// Invalidates the key lists.
//
STDMETHODIMP CGDGPG::InvalidateKeyLists()
{
    m_keyDefault.m_sUser = "";
    m_keyList.clear ();
    m_sPassphraseStore = m_sPassphraseInvalid;

    return S_OK;
}

/////////////////////////////////////////////////////////////////////////////
// Options
//
// Opens the options dialog.
//
STDMETHODIMP CGDGPG::Options(
	ULONG hWndParent)  // The handle of the parent window.
{
    COptionsDlg dlg;

    dlg.m_sGPG = m_sGPGExe;
    dlg.m_sKeyManager = m_sKeyManagerExe;
    if (dlg.DoModal((HWND) hWndParent) == IDOK)
    {
	m_sGPGExe = dlg.m_sGPG;
	m_sKeyManagerExe = dlg.m_sKeyManager;	
	WriteOptions();	
    }

    return S_OK;
}

/////////////////////////////////////////////////////////////////////////////
// EncryptAndSignNextFile
//
// Encrypts and signs the specified file using the parameter of the last call 
// to EncryptAndSignFile(). If strFilenameDest is empty, the funtion deletes 
// only the cached parameter.
//
// Return value: The error code:
//               GDGPG_SUCCESS:          success
//               GDGPG_ERR_GPG_FAILED:   call to gpg failed
//               GDGPG_ERR_NO_DEST_FILE: the destination was not created by gpg (gpg failed)
//
STDMETHODIMP CGDGPG::EncryptAndSignNextFile(
	ULONG hWndParent,        // The handle of the parent window for messages and dialog.
	BSTR strFilenameSource,  // The source file name.
	BSTR strFilenameDest,    // The destination file name.
	BOOL bArmor,             // Indicates whether to produce ASCII output.
	int *pvReturn)           // The return value.
{
    USES_CONVERSION;

    LPSTR wstrFilenameDest;
    string sResult, sError;

    *pvReturn = GDGPG_SUCCESS;
    wstrFilenameDest = OLE2A(strFilenameDest);

    if (wstrFilenameDest == "" || *wstrFilenameDest == 0)
    {
	LogInfo ("EncryptAndSignNextFile: cleanup.\n");
	m_sEncryptCommandNextFile = "";
	m_sEncryptPassphraseNextFile = m_sPassphraseInvalid;
	return S_OK;	
    }
    
    string sCommand = m_sEncryptCommandNextFile;

    sCommand += "--output \"";
    sCommand += OLE2A(strFilenameDest);
    sCommand += "\" ";
    if (bArmor)
	sCommand += "--armor ";
    sCommand += "\"";
    sCommand += OLE2A(strFilenameSource);
    sCommand += "\"";

    if (!CallGPG(sCommand, m_sEncryptPassphraseNextFile, sResult, sError))
    {
	LogInfo ("EncryptAndSignNextFile failed (gpg error).\n");
	*pvReturn = GDGPG_ERR_GPG_FAILED;	
    }
    else // check whether the destintion file exists
    {	
	FILE* file = fopen (OLE2A (strFilenameDest), "rb");
	if (file == NULL)
	{
	    LogInfo ("error, no destination file: %s\n", OLE2A (strFilenameDest));
	    *pvReturn = GDGPG_ERR_NO_DEST_FILE;
	}
	else
	    fclose (file);
    }

    return S_OK;
}


/* DecryptNextFile

 Decrypts the specified file using the parameter of the last call to 
 DecryptFile(). If strFilenameDest is empty, the funtion deletes only the 
 cached parameter.

 Return value: The error code:
               GDGPG_SUCCESS:          success
               GDGPG_ERR_GPG_FAILED:   call to gpg failed
               GDGPG_ERR_NO_DEST_FILE: the destination was not created by gpg (gpg failed) 
*/
STDMETHODIMP CGDGPG::DecryptNextFile(
	ULONG hWndParent,       // The handle of the parent window for messages and dialog.
	BSTR strFilenameSource, // The source file name.
	BSTR strFilenameDest,   // The destination file name.
	int *pvReturn)          // The return value.
{
    string sResult, sError;
    LPSTR wstrFilenameDest;

    USES_CONVERSION;

    if (hWndParent == NULL)
	hWndParent = (ULONG) ::GetActiveWindow();

    *pvReturn = GDGPG_SUCCESS;
    wstrFilenameDest = OLE2A(strFilenameDest);
    if (wstrFilenameDest == "" || *wstrFilenameDest == 0)
    {
	LogInfo ("DecryptNextFile: cleanup\n");
	m_sDecryptPassphraseNextFile = m_sPassphraseInvalid;
	return S_OK;	
    }    
	
    ::DeleteFile(OLE2A(strFilenameDest));
    if (strstr (OLE2A (strFilenameSource),".sig"))
	return S_OK;


    string sCommand = "--output \"";
    sCommand += OLE2A(strFilenameDest);
    sCommand += "\" --yes ";
    sCommand += "--decrypt";
    if (m_sDecryptPassphraseNextFile != "")
	sCommand += " --passphrase-fd 0";
    sCommand += " \"";
    sCommand += OLE2A(strFilenameSource);
    sCommand += "\"";    

    if (!CallGPG(sCommand, m_sDecryptPassphraseNextFile, sResult, sError))
    {
	LogInfo ("DecryptNextFile failed (gpg error).\n");
	*pvReturn = GDGPG_ERR_GPG_FAILED;
    }
    else // check whether the destintion file exists
    {
	FILE* file = fopen (OLE2A (strFilenameDest), "rb");
	if (file == NULL)
	{
	    LogInfo ("error, no destination file: %s\n", OLE2A (strFilenameDest));
	    *pvReturn = GDGPG_ERR_NO_DEST_FILE;
	}
	else
	    fclose(file);	
    }
    return S_OK;
}
