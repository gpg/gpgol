/* GPG.cpp - gpg functions for mapi messages
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
#include "GPG.h"

#include <EXCHEXT.H>
#include <INITGUID.H>
#include <MAPIGUID.H>
#include <COMUTIL.h>
#include "GPGExchange.h"
#include "GPGExch.h"
#include "..\GDGPG\Wrapper\GDGPGWrapper.h"

/*DEFINE_OLEGUID(IID_IMessage, 0x00020307, 0, 0);*/

// actions for ProcessAttachments()
#define PROATT_SAVE                  1
#define PROATT_ENCRYPT_AND_SIGN      2
#define PROATT_DECRYPT               3

CGPG::CGPG()
{	
    m_bInit = FALSE;
    m_gpgStderr = NULL;
    m_gpgInfo = NULL;
    m_LogFile = NULL;

    cont_type = NULL;
    cont_trans_enc = NULL;
}


CGPG::~CGPG ()
{
    if (m_LogFile)
	delete [] m_LogFile;
}


BOOL CGPG::DecryptFile (HWND hWndParent, 
			BSTR strFilenameSource, 
			BSTR strFilenameDest, 
			int &pvReturn)
{    
    return m_gdgpg.DecryptFile (hWndParent, strFilenameSource, strFilenameDest, pvReturn);
}



void CGPG::SetLogFile (const char * strLogFilename) 
{ 
    if (m_LogFile)
    {
	delete [] m_LogFile;
	m_LogFile = NULL;
    }
    m_LogFile = new char[strlen (strLogFilename)+1];
    strcpy (m_LogFile, strLogFilename);
}


/* ReadGPGOptions - Reads the plugin options from the registry. */
void CGPG::ReadGPGOptions (void)
{
    USES_CONVERSION;
    TCHAR szKeyName[] = _T("Software\\Microsoft\\Exchange\\Client\\Extensions\\GPG Exchange");
    char Buf[200];
    HKEY hKey;

    if (RegOpenKeyEx(HKEY_LOCAL_MACHINE, szKeyName, 0, KEY_READ, &hKey) == ERROR_SUCCESS)
    {
	DWORD dwSize;
	DWORD dwResult;
	dwSize	= sizeof(dwResult);
	if (RegQueryValueEx(hKey, _T("PhraseTime"), NULL, NULL, (LPBYTE) &dwResult, &dwSize) == ERROR_SUCCESS)
	    m_nStorePassPhraseTime = dwResult;
	dwSize	= sizeof(dwResult);
	if (RegQueryValueEx(hKey, _T("EncryptDefault"), NULL, NULL, (LPBYTE) &dwResult, &dwSize) == ERROR_SUCCESS)
	    m_bEncryptDefault = (dwResult != 0);
	dwSize	= sizeof(dwResult);
	if (RegQueryValueEx(hKey, _T("SignDefault"), NULL, NULL, (LPBYTE) &dwResult, &dwSize) == ERROR_SUCCESS)
	    m_bSignDefault = (dwResult != 0);
	dwSize	= sizeof(dwResult);
	if (RegQueryValueEx(hKey, _T("EncryptWithStandardKey"), NULL, NULL, (LPBYTE) &dwResult, &dwSize) == ERROR_SUCCESS)
	    m_bEncryptWithStandardKey = (dwResult != 0);
	dwSize	= sizeof(dwResult);
	if (RegQueryValueEx(hKey, _T("SaveDecrypted"), NULL, NULL, (LPBYTE) &dwResult, &dwSize) == ERROR_SUCCESS)
	    m_bSaveDecrypted = (dwResult != 0);
	dwSize = sizeof (Buf);
	if (RegQueryValueEx (hKey, _T("LogFile"), NULL, NULL, (BYTE*)Buf, &dwSize) == ERROR_SUCCESS)
	{
	    m_LogFile = new char [strlen (Buf)+1];
	    strcpy (m_LogFile, Buf);
	}
	RegCloseKey(hKey);	
    }
}



/* WriteGPGOptions - Writes the plugin options to the registry. */
void CGPG::WriteGPGOptions (void)
{
    USES_CONVERSION;
    TCHAR szKeyName[] = _T("Software\\Microsoft\\Exchange\\Client\\Extensions\\GPG Exchange");
    HKEY hKey;

    if (RegCreateKeyEx(HKEY_LOCAL_MACHINE, szKeyName, 0, NULL, REG_OPTION_NON_VOLATILE,
	KEY_ALL_ACCESS, NULL, &hKey, NULL) == ERROR_SUCCESS)	
    {
	DWORD dw = m_nStorePassPhraseTime;
	RegSetValueEx(hKey, _T("PhraseTime"), 0, REG_DWORD, (CONST BYTE *) &dw, sizeof(DWORD));
	m_gdgpg.SetStorePassphraseTime(m_nStorePassPhraseTime);
	dw = m_bEncryptDefault ? 1 : 0;
	RegSetValueEx(hKey, _T("EncryptDefault"), 0, REG_DWORD, (CONST BYTE *) &dw, sizeof(DWORD));
	dw = m_bSignDefault ? 1 : 0;
	RegSetValueEx(hKey, _T("SignDefault"), 0, REG_DWORD, (CONST BYTE *) &dw, sizeof(DWORD));
	dw = m_bEncryptWithStandardKey ? 1 : 0;
	RegSetValueEx(hKey, _T("EncryptWithStandardKey"), 0, REG_DWORD, (CONST BYTE *) &dw, sizeof(DWORD));
	dw = m_bSaveDecrypted ? 1 : 0;
	RegSetValueEx(hKey, _T("SaveDecrypted"), 0, REG_DWORD, (CONST BYTE *) &dw, sizeof(DWORD));
	if (m_LogFile != NULL)
	    RegSetValueEx (hKey, _T("LogFile"), NULL, REG_SZ, (BYTE*)m_LogFile, strlen (m_LogFile));
	RegCloseKey(hKey);	
    }
}


/* EncryptAndSignMessage - Encrypts and signs the specified message.
   Return value: TRUE if successful. */
BOOL CGPG::EncryptAndSignMessage(
	HWND hWnd,             // The handle of the parent window for messages and dialog.
	LPMESSAGE lpMessage,   // Points to the message.
	BOOL bEncrypt,         // Indicates whether to encrypt the message.
	BOOL bSign,            // Indicates whether to sign the messages.
	BOOL bIsRootMessage)   // Indicates whether this is the root message (will be FALSE when called for embedded messages).
{
    USES_CONVERSION;
    BOOL bWriteFailed = FALSE;
    BOOL bEncryptAttachmentsFailed = FALSE;
    string sFilenameSource, sFilenameDest;
    TCHAR szTempPath[MAX_PATH];
    BOOL bBodyEmpty = FALSE;
    LPSPropValue lpspvFEID = NULL;
    string sEncryptedBody = "";
    TCHAR szTemp[1024];
    FILE * file;
    HRESULT hr;
    int nRet = 0;

    // get the recipients
    string sRecipient = "";
    if (bIsRootMessage)
    {
	LPMAPITABLE lpRecipientTable = NULL;
	LPSRowSet lpRecipientRows = NULL;
	static SizedSPropTagArray( 1L, PropRecipientNum) = { 1L, {PR_EMAIL_ADDRESS}};

	hr = lpMessage->GetRecipientTable(0, &lpRecipientTable);
	if (SUCCEEDED(hr))
	{
	    hr = HrQueryAllRows(lpRecipientTable, 
		(LPSPropTagArray) &PropRecipientNum,
		NULL, NULL, 0L, &lpRecipientRows);
	    for (int j = 0L; j < lpRecipientRows->cRows; j++)
	    {
		if (sRecipient.size() > 0)
		    sRecipient += "|";
		sRecipient += lpRecipientRows->aRow[j].lpProps[0].Value.lpszA;	
	    }
	    if (NULL != lpRecipientTable)
		lpRecipientTable->Release();
	    if (NULL != lpRecipientRows)
		FreeProws(lpRecipientRows);	
	}   
    }

    GetTempPath(MAX_PATH, szTempPath);
    sFilenameSource = szTempPath;
    sFilenameSource += _T("_gdgpg_");  
    sFilenameDest = sFilenameSource;
    sFilenameSource += "Body.txt";
    sFilenameDest += "Body.gpg";

    // write the message body to a temp file
    hr = HrGetOneProp((LPMAPIPROP) lpMessage, PR_BODY, &lpspvFEID);
    if (FAILED(hr))
	bBodyEmpty = TRUE;
    // write body to file
    file = fopen (sFilenameSource.c_str (), "wb");
    if (!file)
    {
	MAPIFreeBuffer (lpspvFEID);
	bWriteFailed = TRUE;
	goto fail;
    }
    if (!bBodyEmpty)
	fwrite (lpspvFEID->Value.lpszA, 1, strlen(lpspvFEID->Value.lpszA), file); // don't write the the 0 character at the end!!!
    fclose (file);
    MAPIFreeBuffer (lpspvFEID);
    lpspvFEID = NULL;

    // encrypt and sign	
    if (bIsRootMessage)
	m_gdgpg.EncryptAndSignFile (hWnd, bEncrypt, bSign,
					A2OLE(sFilenameSource.c_str()), A2OLE(sFilenameDest.c_str()), A2OLE(sRecipient.c_str()), 
					TRUE, m_bEncryptWithStandardKey, nRet);
    else
	m_gdgpg.EncryptAndSignNextFile (hWnd, A2OLE(sFilenameSource.c_str()), A2OLE(sFilenameDest.c_str()), TRUE, nRet);
    ::DeleteFile (sFilenameSource.c_str ());
    if (nRet != 0)
    {
	bWriteFailed = TRUE;
	goto fail;
    }

    // check whether file exists
    file = fopen (sFilenameDest.c_str (), "rb");
    if (file == NULL)
    {
	bWriteFailed = TRUE;
	goto fail;
    }
    fclose (file);

    // read new body, write the body later because crypting of the attachments may fail
    int nCharsRead;
    file = fopen (sFilenameDest.c_str (), "rb");
    if (!file)
    {
	bWriteFailed = TRUE;
	goto fail;
    }
    do
    {
	nCharsRead = fread (szTemp, sizeof(TCHAR), sizeof(szTemp)-1, file);
	szTemp[nCharsRead] = '\0';
	sEncryptedBody += szTemp;
    } while (nCharsRead > 0);
    fclose (file);
    ::DeleteFile (sFilenameDest.c_str ());

    // reset the rendering positions of the attachments
    // otherwise outlook inserts a blank at the attachment position
    {
	static SizedSPropTagArray( 1L, PropAttNum) = { 1L, {PR_ATTACH_NUM}};

	BOOL bSuccess = TRUE;
	LPMAPITABLE lpAttTable = NULL;
	LPSRowSet lpAttRows = NULL;
	BOOL bSaveAttSuccess = TRUE;
	hr = lpMessage->GetAttachmentTable(0, &lpAttTable);
	if (FAILED(hr))
	    bSuccess = FALSE;
	if (bSuccess)
	{
	    hr = HrQueryAllRows(lpAttTable, (LPSPropTagArray) &PropAttNum, NULL, NULL, 0L, &lpAttRows);
	    if (FAILED(hr))
		bSuccess = FALSE;
	}
	if (bSuccess)
	{
	    for (int j = 0; j < (int) lpAttRows->cRows; j++)
	    {
		LPATTACH pAttach = NULL;
		LPSTREAM pStreamAtt = NULL;
		hr = lpMessage->OpenAttach(lpAttRows->aRow[j].lpProps[0].Value.ul, NULL, MAPI_BEST_ACCESS, &pAttach);
		if (SUCCEEDED(hr))
		{
		    SPropValue sProp; 
		    sProp.ulPropTag = PR_RENDERING_POSITION;
		    sProp.Value.ul = -1;
		    hr = HrSetOneProp(pAttach, &sProp);
		    pAttach->SaveChanges(FORCE_SAVE);
		}	
	    }
	}	
	if (NULL != lpAttTable)	
	    lpAttTable->Release ();
	if (NULL != lpAttRows)
	    FreeProws (lpAttRows);
	// encrypt and sign attachments
	bEncryptAttachmentsFailed = !EncryptAndSignAttachments(hWnd, lpMessage);

	// replace the message body
	if (!bEncryptAttachmentsFailed)
	{
	    BOOL bChanged = FALSE;
	    SPropValue sProp; 
	    sProp.ulPropTag = PR_BODY;
	    // replaces the text first with "" and calls RTFSync
	    // otherwise outlook deletes an important whitespace
	    sProp.Value.lpszA = "";
	    hr = HrSetOneProp(lpMessage, &sProp);
	    ::RTFSync(lpMessage, RTF_SYNC_BODY_CHANGED, &bChanged);
	    sProp.Value.lpszA = (char*) sEncryptedBody.c_str();
	    hr = HrSetOneProp(lpMessage, &sProp);
	    ::RTFSync(lpMessage, RTF_SYNC_BODY_CHANGED, &bChanged);
	}
	
	/* XXX: in the above code the body is not changed */
	SPropValue sProp;
	sProp.ulPropTag = PR_BODY;
	sProp.Value.lpszA = (char*)sEncryptedBody.c_str ();
	HrSetOneProp (lpMessage, &sProp);

	// erase the temporary saved paramter	
	if (bIsRootMessage)
	    m_gdgpg.EncryptAndSignNextFile(hWnd, A2OLE(""), A2OLE(""), FALSE, nRet);

	if (bIsRootMessage && bEncryptAttachmentsFailed)
	{
	    TCHAR sCaption[256];
	    TCHAR sMess[256];

	    LoadString(theApp.m_hInstance, IDS_APP_NAME, sCaption, sizeof(sCaption));
	    if (!m_bContainsEmbeddedOLE)
		LoadString(theApp.m_hInstance, IDS_ERR_ENCRYPT_ATTACHMENTS, sMess, sizeof(sMess));
	    else
		LoadString(theApp.m_hInstance, IDS_ERR_ENCRYPT_EMBEDDED_OLE, sMess, sizeof(sMess));
	    ::MessageBox(hWnd, sMess, sCaption, MB_ICONEXCLAMATION);
	    bWriteFailed = TRUE;
	}

    }
fail:
    return !bWriteFailed;
}


/* FindMessageWindow - Walks through all child windows and searches the crypted message. 
   Return value: TRUE if successful. */
HWND CGPG::FindMessageWindow (HWND hWnd)  /* Points to the parent window. */
{
    if (hWnd == NULL)
	return NULL;
    HWND wndChild = GetWindow (hWnd, GW_CHILD);
    while (wndChild != NULL)
    {
	TCHAR s[1025];
	memset (s, 0, sizeof (s));
	GetWindowText (wndChild, s, sizeof (s)-1);
	string sText = s;
	if ((sText.find ("-----BEGIN PGP MESSAGE-----") != -1) ||
	    (sText.find ("-----BEGIN PGP SIGNED MESSAGE-----") != -1))
	    return wndChild;
	HWND wndRTF = FindMessageWindow (wndChild);
	if (wndRTF != NULL)
	    return wndRTF;
	wndChild = GetNextWindow (wndChild, GW_HWNDNEXT);	
    }
    return NULL;
}


/* DecryptMessage - Decrypts the specified message.
   Return value: TRUE if successfull. */
BOOL CGPG::DecryptMessage(
	HWND hWnd,            // The handle of the parent window for messages and dialog.
	LPMESSAGE lpMessage,  // Points to the message.
	BOOL bIsRootMessage)  // Indicates whether this is the root message (will be FALSE when called for embedded messages).
{
	BOOL bSuccess = TRUE;

	string sFilenameSource, sFilenameDest;
	TCHAR szTempPath[MAX_PATH];
	GetTempPath(MAX_PATH, szTempPath);
	sFilenameSource = szTempPath;
	sFilenameSource += _T("_gdgpg_");  
	sFilenameDest = sFilenameSource;
	sFilenameSource += "DBody.gpg";
	sFilenameDest += "DBody.txt";

	// write the message body to a temp file
	BOOL bBodyEmpty = FALSE;
	LPSPropValue lpspvFEID = NULL;
	HRESULT hr = HrGetOneProp((LPMAPIPROP) lpMessage, PR_BODY, &lpspvFEID);
	if (FAILED(hr))
		bBodyEmpty = TRUE;

	FILE* file = fopen(sFilenameSource.c_str(), "wb");
	if (!bBodyEmpty)
		fwrite(lpspvFEID->Value.lpszA, 1, strlen(lpspvFEID->Value.lpszA), file); 
	fclose(file);
	MAPIFreeBuffer(lpspvFEID);
	lpspvFEID = NULL;

	// decrypt
	USES_CONVERSION;
	int nRet;
	if (bIsRootMessage)
		m_gdgpg.DecryptFile(hWnd, A2OLE(sFilenameSource.c_str()), A2OLE(sFilenameDest.c_str()), nRet);
	else
		m_gdgpg.DecryptNextFile(hWnd, A2OLE(sFilenameSource.c_str()), A2OLE(sFilenameDest.c_str()), nRet);
	::DeleteFile(sFilenameSource.c_str());

	if (nRet != 0)
		bSuccess = FALSE;

	// check whether file exists
	file = fopen(sFilenameDest.c_str(), "rb");
	if (file == NULL)
		bSuccess = FALSE;
	else
		fclose(file);

	if (bSuccess)
	{
		// replace body
		SPropValue sProp; 
		sProp.ulPropTag = PR_BODY;
		string sBody;
		TCHAR szTemp[1024];
		FILE* file = fopen(sFilenameDest.c_str(), "rb");
		int nCharsRead;
		do
		{
			nCharsRead = fread(szTemp, sizeof(TCHAR), sizeof(szTemp)-1, file);
			szTemp[nCharsRead] = '\0';
			sBody += szTemp;
		} while (nCharsRead > 0);
		fclose(file);
		::DeleteFile(sFilenameDest.c_str());

		sProp.Value.lpszA = (char*) sBody.c_str();
		hr = HrSetOneProp(lpMessage, &sProp);

		sProp.ulPropTag = PR_ACCESS;
		sProp.Value.l = MAPI_ACCESS_MODIFY;
		HrSetOneProp(lpMessage, &sProp);


		if (bIsRootMessage && (hWnd != NULL))
		{
			HWND wndRTF = FindMessageWindow(hWnd);

			if (wndRTF != NULL)
				SetWindowText(wndRTF, sBody.c_str());
			else
			{
				TCHAR sCaption[256];
				LoadString(theApp.m_hInstance, IDS_APP_NAME, sCaption, sizeof(sCaption));
				TCHAR sMess[512];
				if (m_bSaveDecrypted)
				{
					LoadString(theApp.m_hInstance, IDS_ERR_REPLACE_TEXT, sMess, sizeof(sMess));
					::MessageBox(hWnd, sMess, sCaption, MB_ICONINFORMATION);
				}
				else
				{
					LoadString(theApp.m_hInstance, IDS_ERR_REPLACE_TEXT_ASK_SAVE, sMess, sizeof(sMess));
					int nCmx = ::MessageBox(hWnd, sMess, sCaption, MB_ICONQUESTION | MB_YESNOCANCEL);
					if (nCmx == IDYES)
						lpMessage->SaveChanges(FORCE_SAVE);
				}
			}
		}

		// decrypt attachments
		if (bIsRootMessage)
		{
			m_nDecryptedAttachments = 0;
			m_bCancelSavingDecryptedAttachments = FALSE;
		}
		DecryptAttachments(hWnd, lpMessage);

		if (m_bSaveDecrypted)
		{
			lpMessage->SaveChanges(FORCE_SAVE);
			if (bIsRootMessage && (m_nDecryptedAttachments > 0))
			{
				TCHAR sCaption[256];
				LoadString(theApp.m_hInstance, IDS_APP_NAME, sCaption, sizeof(sCaption));
				TCHAR sMess[256];
				LoadString(theApp.m_hInstance, IDS_ATT_DECRYPT_AND_SAVE, sMess, sizeof(sMess));
			    ::MessageBox(hWnd, sMess, sCaption, MB_ICONINFORMATION);
			}
		}
	}

	// erase temporary saved parameter
	if (bIsRootMessage)
		m_gdgpg.DecryptNextFile(hWnd, A2OLE(""), A2OLE(""), nRet);

	return bSuccess;
}


void
CGPG::QuotedPrintEncode (char ** enc_buf, char * buf, size_t buflen)
{
}


BOOL
CGPG::ProcessPGPMime (HWND hWnd, LPMESSAGE pMessage, int mType)
{
    return TRUE;
}


BOOL 
CGPG::CheckPGPMime (HWND hWnd, LPMESSAGE pMessage, int &mType)
{
    return TRUE;
}

/////////////////////////////////////////////////////////////////////////////
// ProcessAttachments
//
// Executes the specified action on all attachments.
//
// Return value: TRUE if successful.
//
BOOL 
CGPG::ProcessAttachments(
	HWND hWnd,               // The handle of the parent window for messages and dialog.
	int nAction,             // The action, see PROATT_xxx definies.
	LPMESSAGE pMessage,      // Points to the message.
	string sPrefix,          // The prefix used for temporary file names.
	vector<string>* pFileNameVector)  // Points to a string vector to which all saved file names will be added; may be NULL.
{
	USES_CONVERSION;
	BOOL bSuccess = TRUE;
	TCHAR szTempPath[MAX_PATH];
	GetTempPath(MAX_PATH, szTempPath);
	string sFilenameAtt = szTempPath;
	sFilenameAtt += sPrefix;

    static SizedSPropTagArray( 1L, PropAttNum) = { 1L, {PR_ATTACH_NUM}};

	LPMAPITABLE lpAttTable = NULL;
	LPSRowSet lpAttRows = NULL;
	BOOL bSaveAttSuccess = TRUE;
	HRESULT hr = pMessage->GetAttachmentTable(0, &lpAttTable);
	if (FAILED(hr))
		bSuccess = FALSE;

	if (bSuccess)
	{
		hr = HrQueryAllRows(lpAttTable, (LPSPropTagArray) &PropAttNum,
			NULL, NULL, 0L, &lpAttRows);
		if (FAILED(hr))
			bSuccess = FALSE;
	}

	if (bSuccess)
	{
		for (int j = 0; j < (int) lpAttRows->cRows; j++)
		{
			TCHAR sn[20];
			itoa(j, sn, 10);
			string sFilename = sFilenameAtt;
			sFilename += sn;
			sFilename += ".tmp";

			string sFilenameDest = sFilename + ".gpg";

			string sAttName = "attach";
			LPATTACH pAttach = NULL;
			LPSTREAM pStreamAtt = NULL;
			int nAttachment = lpAttRows->aRow[j].lpProps[0].Value.ul;
			HRESULT hr = pMessage->OpenAttach(nAttachment, NULL, MAPI_BEST_ACCESS, &pAttach);
			if (SUCCEEDED(hr))
			{
				BOOL bSaveAttSuccess = TRUE;
				LPSPropValue lpspv = NULL;
				hr = HrGetOneProp((LPMAPIPROP) pAttach, PR_ATTACH_METHOD, &lpspv);
				BOOL bEmbedded = (lpspv->Value.ul == ATTACH_EMBEDDED_MSG);
				BOOL bEmbeddedOLE = (lpspv->Value.ul == ATTACH_OLE);
				MAPIFreeBuffer(lpspv);

				if (SUCCEEDED(hr) && bEmbedded)
				{
					if ((nAction == PROATT_DECRYPT) || (nAction == PROATT_ENCRYPT_AND_SIGN))
					{
						LPMESSAGE pMessageEmb = NULL;
						hr = pAttach->OpenProperty(PR_ATTACH_DATA_OBJ, &IID_IMessage, 0, MAPI_MODIFY, (LPUNKNOWN*) &pMessageEmb);
						if (SUCCEEDED(hr))
						{
							if (nAction == PROATT_DECRYPT)
							{
								if (!DecryptMessage(hWnd, pMessageEmb, FALSE))
									bSuccess = FALSE;
							}
							if (nAction == PROATT_ENCRYPT_AND_SIGN)
							{
								if (!EncryptAndSignMessage(hWnd, pMessageEmb, FALSE, FALSE, FALSE))
									bSuccess = FALSE;
							}
							pMessageEmb->SaveChanges(FORCE_SAVE);
							pAttach->SaveChanges(FORCE_SAVE);
							pMessageEmb->Release();
							m_nDecryptedAttachments++;
						}
					}
				}

				BOOL bShortFileName = FALSE;
				if (!bEmbedded && bSaveAttSuccess)
				{
					LPSPropValue lpPropValue;
					hr = HrGetOneProp((LPMAPIPROP) pAttach, PR_ATTACH_LONG_FILENAME, &lpPropValue);
					if (FAILED(hr))
					{
						bShortFileName = TRUE;
						hr = HrGetOneProp((LPMAPIPROP) pAttach, PR_ATTACH_FILENAME, &lpPropValue);
						if (SUCCEEDED(hr))
						{
							sAttName = lpPropValue[0].Value.lpszA;
							MAPIFreeBuffer(lpPropValue);
						}
					}
					else
					{
						sAttName = lpPropValue[0].Value.lpszA;
						MAPIFreeBuffer(lpPropValue);
					}

					// use correct file extensions for temp files to allow virus checking
					if (nAction == PROATT_DECRYPT)
					{
						sFilename = sFilenameAtt + sAttName;
						sFilenameDest = sFilename;
						int nPos = sFilename.rfind('.');
						if (nPos != -1)
						{
							string sExt = sFilename.substr(nPos+1);
							if ((sExt == "gpg") || (sExt == "GPG") || (sExt == "asc") || (sExt == "ASC"))
								sFilenameDest = sFilename.substr(0, nPos);
							else
								sFilename += ".gpg";
						}
						else
							sFilename += ".gpg";
					}
					if (nAction == PROATT_ENCRYPT_AND_SIGN)
					{
						sFilename = sFilenameAtt + sAttName;
						sFilenameDest = sFilename + ".gpg";
					}
				}

				if (!bEmbedded && bSaveAttSuccess && bEmbeddedOLE)
				{
					// we can crypt embedded OLE objects as attachments, but it is
					// not possible to open them after decrypting
					// so show a error message
					if (nAction == PROATT_ENCRYPT_AND_SIGN)
					{
						bSaveAttSuccess = FALSE;
						m_bContainsEmbeddedOLE = TRUE;
					}
					else
					{
						// save ole attachment to file
						hr = pAttach->OpenProperty(PR_ATTACH_DATA_OBJ, &IID_IStream, 0, MAPI_MODIFY, (LPUNKNOWN*) &pStreamAtt);
						if (SUCCEEDED(hr))
						{
							LPSTREAM  pStreamDest;
							hr = OpenStreamOnFile(MAPIAllocateBuffer, MAPIFreeBuffer,
							   STGM_CREATE | STGM_READWRITE,
							   (char*) sFilename.c_str(), NULL,  &pStreamDest);

							if (FAILED(hr))
							{
								pStreamAtt->Release();
							}
							else
							{
								STATSTG statInfo;
								pStreamAtt->Stat(&statInfo, STATFLAG_NONAME);

   								pStreamAtt->CopyTo(pStreamDest, statInfo.cbSize, NULL, NULL);
								pStreamDest->Commit(0);
								pStreamDest->Release();
								pStreamAtt->Release();

								if (pFileNameVector != NULL)
									pFileNameVector->push_back(sFilename);

								m_nDecryptedAttachments++;
							}
						}
						else
							bEmbeddedOLE = FALSE;  // use PR_ATTACH_DATA_BIN
					}
				}

				if (!bEmbedded && !bEmbeddedOLE && bSaveAttSuccess)
				{
					hr = pAttach->OpenProperty(PR_ATTACH_DATA_BIN, &IID_IStream, 0,	0, (LPUNKNOWN*) &pStreamAtt);
					if (HR_SUCCEEDED(hr))
					{
						LPSTREAM  pStreamDest;
						hr = OpenStreamOnFile(MAPIAllocateBuffer, MAPIFreeBuffer,
						   STGM_CREATE | STGM_READWRITE,
						   (char*) sFilename.c_str(), NULL,  &pStreamDest);

						if (FAILED(hr))
						{
							pStreamAtt->Release();
						}
						else
						{
							STATSTG statInfo;
							pStreamAtt->Stat(&statInfo, STATFLAG_NONAME);

   							pStreamAtt->CopyTo(pStreamDest, statInfo.cbSize, NULL, NULL);
							pStreamDest->Commit(0);
							pStreamDest->Release();
							pStreamAtt->Release();

							if (pFileNameVector != NULL)
								pFileNameVector->push_back(sFilename);
						}
					}
				}

				// encrypt, decrypt, sign
				if (!bEmbedded && bSaveAttSuccess && (nAction >= PROATT_ENCRYPT_AND_SIGN) &&  (nAction <= PROATT_DECRYPT))
				{
					int nRet = 0;
					if (nAction == PROATT_DECRYPT)
						m_gdgpg.DecryptNextFile(hWnd, A2OLE(sFilename.c_str()), A2OLE(sFilenameDest.c_str()), nRet);
					else
						m_gdgpg.EncryptAndSignNextFile(hWnd, A2OLE(sFilename.c_str()), A2OLE(sFilenameDest.c_str()), FALSE, nRet);
					::DeleteFile(sFilename.c_str());

					if (nRet == 0)
					{
						FILE* file = fopen(sFilenameDest.c_str(), "rb");
						if (file == NULL)
							bSaveAttSuccess = FALSE;
						else
							fclose(file);
					}
					else
						bSaveAttSuccess = FALSE;

					if (!bSaveAttSuccess)
						::DeleteFile(sFilenameDest.c_str());
				}

				// replace attachment
				if (!bEmbedded && bSaveAttSuccess && (nAction >= PROATT_ENCRYPT_AND_SIGN) &&  (nAction <= PROATT_DECRYPT))
				{
					if (nAction == PROATT_ENCRYPT_AND_SIGN)
						sAttName += ".gpg";
					if (nAction == PROATT_DECRYPT)
					{
						int nPos = sAttName.rfind('.');
						if (nPos != -1)
						{
							string sExt = sAttName.substr(nPos+1);
							if ((sExt == "gpg") || (sExt == "GPG"))
								sAttName = sAttName.substr(0, nPos);
						}
						m_nDecryptedAttachments++;
						if (!m_bSaveDecrypted && !m_bCancelSavingDecryptedAttachments)
						{
							// save encrypted attachment
							if (!SaveDecryptedAttachment(hWnd, sFilenameDest, sAttName))
								m_bCancelSavingDecryptedAttachments = TRUE;
						}
					}

					pMessage->DeleteAttach(nAttachment, 0, NULL, 0);

					ULONG ulNewAttNum;
					LPATTACH lpNewAttach = NULL;
					hr = pMessage->CreateAttach(NULL, 0, &ulNewAttNum, &lpNewAttach);
					if (SUCCEEDED(hr))
					{
						SPropValue sProp; 
						sProp.ulPropTag = PR_ATTACH_METHOD;
						sProp.Value.ul = ATTACH_BY_VALUE;
						hr = HrSetOneProp(lpNewAttach, &sProp);
						if (SUCCEEDED(hr))
						{
							if (bShortFileName)
								sProp.ulPropTag = PR_ATTACH_FILENAME;
							else
								sProp.ulPropTag = PR_ATTACH_LONG_FILENAME;
							sProp.Value.lpszA = (char*) sAttName.c_str();
							hr = HrSetOneProp(lpNewAttach, &sProp);

							LPSTREAM  pNewStream;
							hr = lpNewAttach->OpenProperty(PR_ATTACH_DATA_BIN, &IID_IStream, 0,	MAPI_CREATE | MAPI_MODIFY, (LPUNKNOWN*) &pNewStream);
							if (SUCCEEDED(hr))
							{
								LPSTREAM  pStreamFile;
								hr = OpenStreamOnFile(MAPIAllocateBuffer, MAPIFreeBuffer,
								   STGM_READ, (char*) sFilenameDest.c_str(), NULL,  &pStreamFile);

								if (SUCCEEDED(hr))
								{
									STATSTG statInfo;
									pStreamFile->Stat(&statInfo, STATFLAG_NONAME);
   									pStreamFile->CopyTo(pNewStream, statInfo.cbSize, NULL, NULL);
									pNewStream->Commit(0);
									pNewStream->Release();
									pStreamFile->Release();
									lpNewAttach->SaveChanges(FORCE_SAVE);
								}
							}
						}
					}
					::DeleteFile(sFilenameDest.c_str());
				}

				if (!bSaveAttSuccess)
					bSuccess = FALSE;
				pAttach->Release();
			}
		}
	}

    if (NULL != lpAttTable)
        lpAttTable->Release();

    if (NULL != lpAttRows)
        FreeProws(lpAttRows);

	return bSuccess;
}

/////////////////////////////////////////////////////////////////////////////
// SaveAttachments
//
// Saves all attachments of the message.
//
// Return value: TRUE if successful.
//
BOOL CGPG::SaveAttachments(
	HWND hWnd,          // The handle of the parent window for messages and dialog.
	LPMESSAGE pMessage, // Points to the message.
	string sPrefix,     // The prefix used for temporary file names.
	vector<string>* pFileNameVector)  // Points to a string vector to which all saved file names will be added; may be NULL.
{
    return ProcessAttachments(hWnd, PROATT_SAVE, pMessage, sPrefix, pFileNameVector);
}

/////////////////////////////////////////////////////////////////////////////
// EncryptAndSignAttachments
//
// Encrypts and signs all attachments.
//
// Return value: TRUE if successful.
//
BOOL CGPG::EncryptAndSignAttachments(
	HWND hWnd,           // The handle of the parent window for messages and dialog.
	LPMESSAGE pMessage)  // Points to the message.
{
    m_bContainsEmbeddedOLE = FALSE;
    return ProcessAttachments(hWnd, PROATT_ENCRYPT_AND_SIGN, pMessage, "_gdgpg_", NULL);
}

/////////////////////////////////////////////////////////////////////////////
// DecryptAttachments
//
// Decrypts all attachments.
//
// Return value: TRUE if successful.
//
BOOL CGPG::DecryptAttachments(
	HWND hWnd,           // The handle of the parent window for messages and dialog.
	LPMESSAGE pMessage)  // Points to the message.
{
    return ProcessAttachments(hWnd, PROATT_DECRYPT, pMessage, "_gdgpg_", NULL);
}


BSTR
CGPG::GetGPGOutput (void)
{
    if (m_gpgStderr)
    {
	SysFreeString (m_gpgStderr);
	m_gpgStderr = NULL;
    }
    m_gdgpg.GetGPGOutput (&m_gpgStderr);
    return m_gpgStderr;
}


BSTR 
CGPG::GetGPGInfo (BSTR bstrFilename)
{
    if (m_gpgInfo)
    {
	SysFreeString (m_gpgInfo);
	m_gpgInfo = NULL;
    }
    m_gdgpg.GetGPGInfo (bstrFilename, &m_gpgInfo);
    return m_gpgInfo;
}


BOOL 
CGPG::VerifyDetachedSignature (HWND hWndParent, BSTR strFilenameText, BSTR strFilenameSig, int &pvReturn)
{
    int ret=0;

    m_gdgpg.VerifyDetachedSignature (hWndParent, strFilenameText, strFilenameSig, ret);
    return ret;
}


/////////////////////////////////////////////////////////////////////////////
// ImportKeys
//
// Imports all keys from the specified message.
//
// Return value: TRUE if successful.
//
BOOL CGPG::ImportKeys(
	HWND hWnd,           // The handle of the parent window for messages and dialog.
	LPMESSAGE lpMessage) // Points to the message.
{
	
    BOOL bSuccess = TRUE;
    string sFilenameSource, sFilenameAtt;
    TCHAR szTempPath[MAX_PATH];

    GetTempPath(MAX_PATH, szTempPath);
    sFilenameSource = szTempPath;
    sFilenameSource += _T("_gdgpg_impkey.tmp");  

    // write the message body to a temp file
    BOOL bBodyEmpty = FALSE;
    LPSPropValue lpspvFEID = NULL;
    HRESULT hr = HrGetOneProp((LPMAPIPROP) lpMessage, PR_BODY, &lpspvFEID);

    if (FAILED(hr))
	bBodyEmpty = TRUE;    
	
    // write body to file
    FILE* file = fopen(sFilenameSource.c_str(), "wb");
    if (file)
    {
	if (!bBodyEmpty)
	    fwrite(lpspvFEID->Value.lpszA, 1, strlen(lpspvFEID->Value.lpszA), file);
	fclose(file);
    }
    MAPIFreeBuffer(lpspvFEID);
    lpspvFEID = NULL;
    
    vector<string> sFilenameVector;

    if (!bBodyEmpty)
	sFilenameVector.push_back(sFilenameSource);

    // save all attachments
    SaveAttachments(hWnd, lpMessage, "_gdgpg_impkeya", &sFilenameVector);

    USES_CONVERSION;
    int nEditCnt = 0;
    int nImportCnt = 0;
    int nUnchangeCnt = 0;
    int nRet = 0;
    int nEditCntTmp = 0;
    int nImportCntTmp = 0;
    int nUnchangeCntTmp = 0;

    for (vector<string>::iterator i = sFilenameVector.begin(); i != sFilenameVector.end(); i++)
	m_gdgpg.ImportKeys(hWnd, A2OLE(i->c_str()), TRUE, nEditCntTmp, nImportCntTmp, nUnchangeCntTmp, nRet);
    nEditCnt += nEditCntTmp;
    nImportCnt += nImportCntTmp;
    nUnchangeCnt += nUnchangeCntTmp;

    TCHAR sCaption[256];
    LoadString(theApp.m_hInstance, IDS_APP_NAME, sCaption, sizeof(sCaption));
    TCHAR s[256] = "";
    TCHAR s1[256] = "";
    if (nEditCnt > 0)
    {
	if (nImportCnt > 0)
	{
	    LoadString(theApp.m_hInstance, IDS_IMPORT_X_KEYS, s1, sizeof(s1));
	    wsprintf(s, s1 , nImportCnt);
	    m_gdgpg.InvalidateKeyLists();
	}
	else
	    LoadString(theApp.m_hInstance, IDS_IMPORT_NO_NEW_OR_CHANGED_KEYS, s, sizeof(s));	
    }
    else
	LoadString(theApp.m_hInstance, IDS_IMPORT_NO_KEYS, s, sizeof(s));
    ::MessageBox(hWnd, s, sCaption, MB_ICONINFORMATION);

    /* delete the temp files */
    for (i = sFilenameVector.begin(); i != sFilenameVector.end(); i++)
	::DeleteFile(i->c_str());

    return bSuccess;
}

/////////////////////////////////////////////////////////////////////////////
// Init
//
// Initialize this object. Creates th GDGPG object and reads the options from 
// the regsitry.
//
// Return value: TRUE if successful.
//
BOOL CGPG::Init(void)
{
    if (m_bInit)
	return TRUE;
    ReadGPGOptions();
    if (m_gdgpg.CreateInstance ())
    {
	m_gdgpg.SetStorePassphraseTime (m_nStorePassPhraseTime);
	USES_CONVERSION;
	if (m_LogFile != NULL)
	    m_gdgpg.SetLogFile (A2OLE (m_LogFile));
	m_bInit = TRUE;
	return TRUE;	
    }
    return FALSE;
}

/////////////////////////////////////////////////////////////////////////////
// UnInit
//
// Uninitialize this object. Frees the GDGPG object.
//
void CGPG::UnInit(void)
{
    if (!m_bInit)
	return;
    m_bInit = FALSE;
    m_gdgpg.Release();
}

/////////////////////////////////////////////////////////////////////////////
// EditExtendedOptions
//
// Shows the dialog to edit the extended options.
//
void CGPG::EditExtendedOptions(HWND hWnd)
{
    m_gdgpg.Options(hWnd);
}

/////////////////////////////////////////////////////////////////////////////
// OpenKeyManager
//
// Opens the key manager.
//
void CGPG::OpenKeyManager()
{
    int nRet;
    m_gdgpg.OpenKeyManager(nRet);
}

/////////////////////////////////////////////////////////////////////////////
// InvalidateKeyLists
//
// Invalidates the key lists (e.g. when the keys was changed by the key manager).
//
void CGPG::InvalidateKeyLists()
{
    m_gdgpg.InvalidateKeyLists();
}

/////////////////////////////////////////////////////////////////////////////
// AddStandardKey
//
// Adds the standard key to the open message. Copies the key to the clipboard 
// and pastes the text in the specified window.
//
// Return value: TRUE if successful.
//
BOOL CGPG::AddStandardKey(
	HWND hWnd) // The handle of the parent window for messages and dialog.
{
	BOOL bSuccess = TRUE;

	string sFilenameDest;
	TCHAR szTempPath[MAX_PATH];
	GetTempPath(MAX_PATH, szTempPath);
	sFilenameDest = szTempPath;
	sFilenameDest += _T("_gdgpg_stdkey.gpg");  

	USES_CONVERSION;
	int nRet;
	m_gdgpg.ExportStandardKey(hWnd, A2OLE(sFilenameDest.c_str()), nRet);

	if (nRet == 0)
	{
		string sBody;
		TCHAR szTemp[1024];
		FILE* file = fopen(sFilenameDest.c_str(), "rb");
		int nCharsRead;
		do
		{
			nCharsRead = fread(szTemp, sizeof(TCHAR), sizeof(szTemp)-1, file);
			szTemp[nCharsRead] = '\0';
			sBody += szTemp;
		} while (nCharsRead > 0);
		fclose(file);
		::DeleteFile(sFilenameDest.c_str());

		// copy sBody to clipboard
		if (OpenClipboard(NULL))
		{
			EmptyClipboard();
			long nTextSize = sBody.size() + 1;
			HGLOBAL hText = ::GlobalAlloc(GMEM_SHARE, nTextSize);
			LPSTR pText = (LPSTR) ::GlobalLock(hText);
			ASSERT(pText);
			strcpy(pText, sBody.c_str());
			::GlobalUnlock(hText);
			::SetClipboardData(CF_TEXT, hText);
			CloseClipboard();
		}

		if (hWnd != NULL)
		{
			BOOL bInserted = FALSE;
			HWND hwndFocus = ::GetFocus();
			if (hwndFocus != NULL)
			{
				RECT rect;
				GetWindowRect(hwndFocus, &rect);
				if (labs(rect.top-rect.bottom) > 50)
				{
					::SendMessage(hWnd, WM_COMMAND, 1058, NULL);
					bInserted = TRUE;
				}
			}
			if (!bInserted)
			{
				TCHAR sCaption[256];
				LoadString(theApp.m_hInstance, IDS_APP_NAME, sCaption, sizeof(sCaption));
				TCHAR sMess[256];
				LoadString(theApp.m_hInstance, IDS_COPY_KEY_TO_CLIPBOARD, sMess, sizeof(sMess));
				::MessageBox(hWnd, sMess, sCaption, MB_ICONINFORMATION);
			}
		}
	}
	else
	{
		TCHAR sCaption[256];
		LoadString(theApp.m_hInstance, IDS_APP_NAME, sCaption, sizeof(sCaption));
		TCHAR sMess[256];
		LoadString(theApp.m_hInstance, IDS_ERR_EXPORT_KEY, sMess, sizeof(sMess));
		::MessageBox(hWnd, sMess, sCaption, MB_ICONINFORMATION);
		bSuccess = FALSE;
	}

	return bSuccess;
}

/////////////////////////////////////////////////////////////////////////////
// SaveDecryptedAttachment
//
// Calls the "save as" dialog to saves an decrypted attachment.
// 
// Return value: TRUE if successful.
//
BOOL CGPG::SaveDecryptedAttachment(
	HWND hWnd,                 // The handle of the parent window.
	string &sSourceFilename,   // The source file name.
	string &sDestFilename)     // The destination file name.
{
	
    TCHAR szTitle[256];
    LoadString(theApp.m_hInstance, IDS_SAVE_ATT_TITLE, szTitle, sizeof(szTitle));
    TCHAR szFilter[256];

    LoadString(theApp.m_hInstance, IDS_SAVE_ATT_FILTER, szFilter, sizeof(szFilter));

    int nLength = strlen(szFilter);

    for (int i=0; i<nLength; i++) {
	if (szFilter[i] == '|')
	    szFilter[i] = '\0';	
    }

    OPENFILENAME ofn;

    TCHAR szFile[MAX_PATH];
    strcpy(szFile, sDestFilename.c_str());

    memset(&ofn, 0, sizeof(ofn));
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = hWnd;
    ofn.lpstrFile = szFile;
    ofn.nMaxFile = MAX_PATH;
    ofn.lpstrFileTitle = NULL;
    ofn.nMaxFileTitle = 0;
    ofn.Flags |= OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT;
    ofn.lpstrTitle = szTitle;
    ofn.lpstrFilter = szFilter;

    if (GetSaveFileName (&ofn))
    {
	sDestFilename = szFile;
	return ::CopyFile(sSourceFilename.c_str(), sDestFilename.c_str(), FALSE);	
    }
    return FALSE;
}
