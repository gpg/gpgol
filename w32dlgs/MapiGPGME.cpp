/* MapiGPGME.cpp - Mapi support with GPGME
 *	Copyright (C) 2005 g10 Code GmbH
 *
 * This file is part of GPGME Dialogs.
 *
 * GPGME Dialogs is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public License
 * as published by the Free Software Foundation; either version 2.1 
 * of the License, or (at your option) any later version.
 *  
 * GPGME Dialogs is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
 * General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public License
 * along with GPGME Dialogs; if not, write to the Free Software Foundation, 
 * Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA 
 */
#include <windows.h>
#include <mapidefs.h>
#include <mapiutil.h>
#include <time.h>

#include "MapiGPGME.h"
#include "gpgme.h"
#include "engine.h"
#include "keycache.h"
#include "intern.h"

#define LOGFILE "c:\\mapigpgme.log"


#define delete_buf(buf) delete [] (buf)

#define fail_if_null(p) do { if (!(p)) abort (); } while (0)


MapiGPGME::MapiGPGME (LPMESSAGE msg)
{
    this->msg = msg;
    op_init ();
    log_debug (LOGFILE, "constructor %p\r\n", msg);
}


MapiGPGME::MapiGPGME ()
{
    log_debug (LOGFILE, "constructor null\r\n");
    op_init ();
}


MapiGPGME::~MapiGPGME ()
{
    log_debug (LOGFILE, "destructor %p\r\n", msg);
    op_deinit ();
}


int 
MapiGPGME::setBody (char *body)
{
    SPropValue sProp; 
    HRESULT hr;
    int rc = TRUE;
    
    if (body == NULL) {
	log_debug (LOGFILE, "setBody with empty buffer\r\n");
	return FALSE;
    }
    rtfSync (body);
    sProp.ulPropTag = PR_BODY;
    sProp.Value.lpszA = body;
    hr = HrSetOneProp (msg, &sProp);
    if (FAILED (hr))
	rc = FALSE;
    log_debug (LOGFILE, "setBody rc=%d '%s'\r\n", rc, body);
    return rc;
}


void
MapiGPGME::rtfSync (char *body)
{
    BOOL bChanged = FALSE;
    SPropValue sProp; 
    HRESULT hr;

    /* Make sure that the Plaintext and the Richtext are in sync */
    sProp.ulPropTag = PR_BODY;
    sProp.Value.lpszA = "";
    hr = HrSetOneProp(msg, &sProp);
    RTFSync(msg, RTF_SYNC_BODY_CHANGED, &bChanged);
    sProp.Value.lpszA = body;
    hr = HrSetOneProp(msg, &sProp);
    RTFSync(msg, RTF_SYNC_BODY_CHANGED, &bChanged);
}


char* 
MapiGPGME::getBody (void)
{
    HRESULT hr;
    LPSPropValue lpspvFEID = NULL;
    char *body;

    hr = HrGetOneProp ((LPMAPIPROP) msg, PR_BODY, &lpspvFEID);
    if (FAILED(hr))
	return NULL;
    
    body = new char[strlen (lpspvFEID->Value.lpszA)+1];
    fail_if_null (body);
    strcpy (body, lpspvFEID->Value.lpszA);

    MAPIFreeBuffer (lpspvFEID);
    lpspvFEID = NULL;

    return body;
}


void 
MapiGPGME::freeKeyArray (void **key)
{
    gpgme_key_t *buf = (gpgme_key_t *)key;
    int i=0;

    if (buf == NULL)
	return;
    for (i = 0; buf[i] != NULL; i++)
	gpgme_key_release (buf[i]);
    xfree (buf);
}


int 
MapiGPGME::countRecipients (char **recipients)
{
    for (int i=0; recipients[i] != NULL; i++)
	;
    return i;
}


char** 
MapiGPGME::getRecipients (bool isRootMsg)
{
    HRESULT hr;
    LPMAPITABLE lpRecipientTable = NULL;
    LPSRowSet lpRecipientRows = NULL;
    char **rset = NULL;

    if (!isRootMsg)
	return NULL;
        
    static SizedSPropTagArray (1L, PropRecipientNum) = {1L, {PR_EMAIL_ADDRESS}};

    hr = msg->GetRecipientTable (0, &lpRecipientTable);
    if (SUCCEEDED(hr)) {
	size_t j = 0;
        hr = HrQueryAllRows (lpRecipientTable, 
			     (LPSPropTagArray) &PropRecipientNum,
			     NULL, NULL, 0L, &lpRecipientRows);
	rset = new char*[lpRecipientRows->cRows+1];
	fail_if_null (rset);
        for (j = 0L; j < lpRecipientRows->cRows; j++) {
	    const char *s = lpRecipientRows->aRow[j].lpProps[0].Value.lpszA;
	    rset[j] = new char[strlen (s)+1];
	    fail_if_null (rset[j]);
	    strcpy (rset[j], s);
	    log_debug (LOGFILE, "rset %d: %s\r\n", j, rset[j]);
	}
	rset[j] = NULL;
	if (NULL != lpRecipientTable)
	    lpRecipientTable->Release();
	if (NULL != lpRecipientRows)
	    FreeProws(lpRecipientRows);	
    }

    return rset;
}


void
MapiGPGME::freeUnknownKeys (char **unknown, int n)
{
    for (int i=0; i < n; i++) {
	if (unknown[i] != NULL)
	    free (unknown[i]);
    }
    free (unknown);
}

void 
MapiGPGME::freeRecipients(char **recipients)
{
    for (int i=0; recipients[i] != NULL; i++)
	delete_buf(recipients[i]);	
    delete recipients;
}


int 
MapiGPGME::encrypt (void)
{
    char *body = getBody();
    char *newBody = NULL;
    char **recipients = getRecipients (true);
    char **unknown = NULL;
    int opts = 0;
    gpgme_key_t *keys=NULL, *keys2=NULL;
    size_t all=0;

    log_debug (LOGFILE, "encrypt\r\n");
    int n = op_lookup_keys (recipients, &keys, &unknown, &all);
    log_debug (LOGFILE, "fnd %d need %d (%p)\r\n", n, all, unknown);
    if (n != countRecipients (recipients)) {
	log_debug (LOGFILE, "recipient_dialog_box2\r\n");
	recipient_dialog_box2 (keys, unknown, all, &keys2, &opts);
	free (keys);
	keys = keys2;
    }

    int err = op_encrypt ((void*)keys, body, &newBody);
    if (err)
	MessageBox (NULL, op_strerror (err), "GPG Encryption", MB_ICONERROR|MB_OK);
    else
	setBody (newBody);

    delete_buf (body);
    free (newBody);
    freeRecipients (recipients);
    freeUnknownKeys (unknown, n);
    freeKeyArray ((void **)keys);
    return err;
}


int 
MapiGPGME::decrypt (void)
{
    char *body = getBody();
    char *newBody = NULL;

    int err = op_decrypt_start (body, &newBody);
    if (err)
	MessageBox (NULL, op_strerror (err), "GPG Decryption", MB_ICONERROR|MB_OK);
    else
	setBody (newBody);

    delete_buf (body);
    free (newBody);
    return err;
}


int
MapiGPGME::sign (void)
{
    char *body = getBody();
    char *newBody = NULL;

    int err = op_sign_start (body, &newBody);
    if (err)
	MessageBox (NULL, op_strerror (err), "GPG Sign", MB_ICONERROR|MB_OK);
    else
	setBody (newBody);

    delete_buf (body);
    free (newBody);
    return err;
}


int
MapiGPGME::doCmd(int doEncrypt, int doSign)
{
    if (doEncrypt && doSign)
	return signEncrypt ();
    if (doEncrypt && !doSign)
	return encrypt ();
    if (!doEncrypt && doSign)
	return sign ();
    return -1;
}


int
MapiGPGME::signEncrypt ()
{
    char *body = getBody();
    char *newBody = NULL;
    char **recipients = getRecipients (TRUE);
    char **unknown = NULL;
    gpgme_key_t locusr, *keys = NULL, *keys2 =NULL;
    
    locusr = find_gpg_key (defaultKey, 0);
    if (!locusr)
	locusr = find_gpg_key (defaultKey, 1);
    if (!locusr) {
	const char *s;
	signer_dialog_box (&locusr, NULL);
	s = gpgme_key_get_string_attr (locusr, GPGME_ATTR_KEYID, NULL, 0);
	defaultKey = new char[strlen (s)+1];
	fail_if_null (defaultKey);
	strcpy (defaultKey, s);
    }


    size_t all;
    int n = op_lookup_keys (recipients, &keys, &unknown, &all);
    if (n != countRecipients (recipients)) {
	recipient_dialog_box2 (keys, unknown, all, &keys2, NULL);
	free (keys);
	keys = keys2;
    }

    int err = op_sign_encrypt ((void *)keys, (void*)locusr, body, &newBody);
    if (err)
	MessageBox (NULL, op_strerror (err), "GPG Sign Encrypt", MB_ICONERROR|MB_OK);
    else
	setBody (newBody);

    delete_buf (body);
    free (newBody);
    freeUnknownKeys (unknown, n);
    freeKeyArray ((void **)keys);
    return err;
}


int 
MapiGPGME::verify ()
{
    char *body = getBody();
    
    int err = op_verify_start (body, NULL);
    if (err)
	MessageBox (NULL, op_strerror (err), "GPG Verify", MB_ICONERROR|MB_OK);

    delete_buf (body);
    return err;
}


void MapiGPGME::setDefaultKey(const char *key)
{
    if (defaultKey) {
	delete_buf (defaultKey);
	defaultKey = NULL;
    }
    defaultKey = new char[strlen (key)+1];
    fail_if_null (defaultKey);
    strcpy (defaultKey, key);
}


char* MapiGPGME::getDefaultKey (void)
{
    if (defaultKey == NULL) {
	void *ctx=NULL;
	gpgme_key_t sk=NULL;

	enum_gpg_seckeys (NULL, &ctx);
	enum_gpg_seckeys (&sk, &ctx);

	defaultKey = new char[16+1];
	fail_if_null (defaultKey);
	const char *s = gpgme_key_get_string_attr (sk, GPGME_ATTR_KEYID, NULL, 0);
	strcpy (defaultKey, s);
    }

    return defaultKey;
}


void 
MapiGPGME::setMessage (LPMESSAGE msg)
{
    this->msg = msg;
    log_debug (LOGFILE, "setMessage %p\r\n", msg);
}


void
MapiGPGME::setWindow(HWND hwnd)
{
    this->hwnd = hwnd;
}


int
MapiGPGME::processAttachments (HWND hwnd, int action, const char **pFileNameVector)
{
#if 0
    USES_CONVERSION;
    BOOL bSuccess = TRUE;
    TCHAR szTempPath[MAX_PATH];
    HRESULT hr;
    LPMAPITABLE lpAttTable = NULL;
    LPSRowSet lpAttRows = NULL;
    BOOL bSaveAttSuccess = TRUE;

    static SizedSPropTagArray (1L, PropAttNum) = {1L, {PR_ATTACH_NUM}};

    GetTempPath (MAX_PATH, szTempPath);
    string sFilenameAtt = szTempPath;
    sFilenameAtt += sPrefix;
   
    hr = pMessage->GetAttachmentTable (0, &lpAttTable);
    if (FAILED (hr))
	return FALSE;

    hr = HrQueryAllRows (lpAttTable, (LPSPropTagArray) &PropAttNum,
			 NULL, NULL, 0L, &lpAttRows);
    if (FAILED(hr)) {
	if (NULL != lpAttTable)
	    lpAttTable->Release ();
	return FALSE;
    }


    for (int j = 0; j < (int) lpAttRows->cRows; j++)
    {
	TCHAR sn[20];
	itoa(j, sn, 10);
	string sFilename = sFilenameAtt;
	sFilename += sn;
	sFilename += ".tmp";
	string sFilenameDest = sFilename + ".pgp";
	string sAttName = "attach";
	LPATTACH pAttach = NULL;
	LPSTREAM pStreamAtt = NULL;
        int nAttachment = lpAttRows->aRow[j].lpProps[0].Value.ul;
        HRESULT hr = pMessage->OpenAttach (nAttachment, NULL, MAPI_BEST_ACCESS, &pAttach);
	if (SUCCEEDED(hr)) {
	    BOOL bSaveAttSuccess = TRUE;
	    LPSPropValue lpspv = NULL;
	    hr = HrGetOneProp ((LPMAPIPROP) pAttach, PR_ATTACH_METHOD, &lpspv);
	    BOOL bEmbedded = (lpspv->Value.ul == ATTACH_EMBEDDED_MSG);
	    BOOL bEmbeddedOLE = (lpspv->Value.ul == ATTACH_OLE);
	    MAPIFreeBuffer (lpspv);
		
		if (SUCCEEDED(hr) && bEmbedded) {
		    if ((nAction == PROATT_DECRYPT) || (nAction == PROATT_ENCRYPT_AND_SIGN)) {
			LPMESSAGE pMessageEmb = NULL;
			hr = pAttach->OpenProperty(PR_ATTACH_DATA_OBJ, &IID_IMessage, 0, MAPI_MODIFY, 
						   (LPUNKNOWN*) &pMessageEmb);
			if (SUCCEEDED(hr)) {
			    if (nAction == PROATT_DECRYPT) {
				if (!DecryptMessage(hWnd, pMessageEmb, FALSE))
				    bSuccess = FALSE;
			    }
			    if (nAction == PROATT_ENCRYPT_AND_SIGN) {
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
		if (!bEmbedded && bSaveAttSuccess) {
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
			    if ((sExt == "gpg") || (sExt == "GPG") || 
				(sExt == "asc") || (sExt == "ASC") ||
				(sExt == "pgp") || (sExt == "PGP"))
				sFilenameDest = sFilename.substr (0, nPos);
			    else
				sFilename += ".pgp";
			}
			else
			    sFilename += ".pgp";
					
		    }
		    if (nAction == PROATT_ENCRYPT_AND_SIGN)
		    {
			sFilename = sFilenameAtt + sAttName;
			sFilenameDest = sFilename + ".pgp";
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
			hr = pAttach->OpenProperty (PR_ATTACH_DATA_OBJ, &IID_IStream, 
						   0, MAPI_MODIFY, (LPUNKNOWN*) &pStreamAtt);
			if (SUCCEEDED(hr))
			{
			    LPSTREAM  pStreamDest;
			    hr = OpenStreamOnFile (MAPIAllocateBuffer, MAPIFreeBuffer,
						   STGM_CREATE | STGM_READWRITE,
						   (char*) sFilename.c_str(), NULL,  &pStreamDest);
			    if (FAILED(hr))
			    {
				pStreamAtt->Release();
			    }
			    else
			    {			
				STATSTG statInfo;				
				pStreamAtt->Stat (&statInfo, STATFLAG_NONAME);   				
				pStreamAtt->CopyTo (pStreamDest, statInfo.cbSize, NULL, NULL);
				pStreamDest->Commit (0);
				pStreamDest->Release ();
				pStreamAtt->Release ();				
				if (pFileNameVector != NULL)
				    pFileNameVector->push_back (sFilename);
				m_nDecryptedAttachments++;	
			    }
			}
			else
			    bEmbeddedOLE = FALSE;  // use PR_ATTACH_DATA_BIN
		    }
		}

		if (!bEmbedded && !bEmbeddedOLE && bSaveAttSuccess)		
		{		
		    hr = pAttach->OpenProperty (PR_ATTACH_DATA_BIN, &IID_IStream, 0, 0, (LPUNKNOWN*) &pStreamAtt);
		    if (HR_SUCCEEDED(hr))
		    {		
			LPSTREAM  pStreamDest;
			hr = OpenStreamOnFile (MAPIAllocateBuffer, MAPIFreeBuffer,
					       STGM_CREATE | STGM_READWRITE,
					       (char*) sFilename.c_str(), NULL,  &pStreamDest);
			if (FAILED (hr))
			{
			    pStreamAtt->Release ();
			}
			else
			{			
			    STATSTG statInfo;
			    pStreamAtt->Stat (&statInfo, STATFLAG_NONAME);
			    pStreamAtt->CopyTo (pStreamDest, statInfo.cbSize, NULL, NULL);
			    pStreamDest->Commit (0);
			    pStreamDest->Release ();
			    pStreamAtt->Release ();
			    if (pFileNameVector != NULL)
				pFileNameVector->push_back(sFilename);
			}
		    }
		}
		// encrypt, decrypt, sign
		if (!bEmbedded && bSaveAttSuccess &&
		    (nAction >= PROATT_ENCRYPT_AND_SIGN) &&
		    (nAction <= PROATT_DECRYPT))
		{
		    int nRet = 0;
		    if (nAction == PROATT_DECRYPT)
			m_gdgpg.DecryptNextFile(hWnd, A2OLE(sFilename.c_str()), A2OLE(sFilenameDest.c_str()), nRet);
		    else
			m_gdgpg.EncryptAndSignNextFile (hWnd, A2OLE(sFilename.c_str()),
							A2OLE(sFilenameDest.c_str()), 
							FALSE, nRet);
		    ::DeleteFile(sFilename.c_str());
		    if (nRet == 0)
		    {
			FILE* file = fopen (sFilenameDest.c_str (), "rb");
			if (file == NULL)
			    bSaveAttSuccess = FALSE;
			else			
			    fclose (file);
		    }
		    else
			bSaveAttSuccess = FALSE;
		    if (!bSaveAttSuccess)
			::DeleteFile (sFilenameDest.c_str());
		}

		// replace attachment		
		if (!bEmbedded && bSaveAttSuccess && 
		    (nAction >= PROATT_ENCRYPT_AND_SIGN) &&  
		    (nAction <= PROATT_DECRYPT))
		{
		    if (nAction == PROATT_ENCRYPT_AND_SIGN) {
			if (m_bSign_Clearsign && !m_bEncrypt_Clearsign)
			    sAttName += ".sig";
			else
			    sAttName += ".pgp";
		    }
		    if (nAction == PROATT_DECRYPT)
		    {
			int nPos = sAttName.rfind('.');
			if (nPos != -1)
			{
			    string sExt = sAttName.substr(nPos+1);
			    if ((sExt == "gpg") || (sExt == "GPG") || 
				(sExt == "pgp") || (sExt == "PGP"))
				sAttName = sAttName.substr(0, nPos);
			}
			m_nDecryptedAttachments++;
			if (!m_bSaveDecrypted && !m_bCancelSavingDecryptedAttachments)
			{			
			    // save encrypted attachment
			    if (!SaveDecryptedAttachment (hWnd, sFilenameDest, sAttName))
				m_bCancelSavingDecryptedAttachments = TRUE;
			}
			#if 0
			/* Dennis Martin 2005-02-10 */
			if (!(m_bSign_Clearsign && !m_bEncrypt_Clearsign))
			    pMessage->DeleteAttach(nAttachment, 0, NULL, 0);
			#endif
		    }

		    pMessage->DeleteAttach (nAttachment, 0, NULL, 0);

		    ULONG ulNewAttNum;
		    LPATTACH lpNewAttach = NULL;
		    hr = pMessage->CreateAttach (NULL, 0, &ulNewAttNum, &lpNewAttach);
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
			    hr = lpNewAttach->OpenProperty (PR_ATTACH_DATA_BIN, &IID_IStream, 0,	
							   MAPI_CREATE | MAPI_MODIFY, (LPUNKNOWN*) &pNewStream);
			    if (SUCCEEDED (hr))
			    {			
				LPSTREAM  pStreamFile;
				hr = OpenStreamOnFile (MAPIAllocateBuffer, MAPIFreeBuffer,
						       STGM_READ, (char*) sFilenameDest.c_str(), 
						       NULL,  &pStreamFile);
				
				if (SUCCEEDED (hr))
				{
				    STATSTG statInfo;
				    pStreamFile->Stat (&statInfo, STATFLAG_NONAME);
				    pStreamFile->CopyTo (pNewStream, statInfo.cbSize, NULL, NULL);
				    pNewStream->Commit (0);
				    pNewStream->Release ();
				    pStreamFile->Release ();
				    lpNewAttach->SaveChanges (FORCE_SAVE);
				}
			    }
			}
		    }
		    ::DeleteFile (sFilenameDest.c_str());
		}
		if (!bSaveAttSuccess)
		    bSuccess = FALSE;
		pAttach->Release ();
	    }
	}
    }
    if (NULL != lpAttTable)
        lpAttTable->Release ();
    if (NULL != lpAttRows)
        FreeProws (lpAttRows);

    return bSuccess;
#endif
    return FALSE;
}


int
MapiGPGME::startKeyManager (void)
{
    return start_key_manager ();
}


void
MapiGPGME::startConfigDialog (HWND parent)
{
    config_dialog_box (parent);
}


int
MapiGPGME::readOptions (void)
{
    char *val=NULL;

    load_extension_value ("encryptDefault", &val);
    doEncrypt = val == NULL || *val != '1'? 0 : 1;
    xfree (val); val=NULL;
    load_extension_value ("signDefault", &val);
    doSign = val == NULL || *val != '1'? 0 : 1;
    xfree (val); val = NULL;
    load_extension_value ("addDefaultKey", &val);
    encryptDefault = val == NULL || *val != '1' ? 0 : 1;
    xfree (val); val = NULL;
    load_extension_value ("storePasswdTime", &val);
    nstorePasswd = val == NULL || *val == '0'? 0 : atol (val);
    xfree (val); val = NULL;

    return 0;
}

int
MapiGPGME::writeOptions (void)
{
    store_extension_value ("encryptDefault", doEncrypt? "1" : "0");
    store_extension_value ("signDefault", doSign? "1" : "0");
    store_extension_value ("addDefaultKey", encryptDefault? "1" : "0");

    char buf[32];
    sprintf (buf, "%d", nstorePasswd);
    store_extension_value ("storePasswdTime", buf);

    return 0;
}
