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
    if (defaultKey)
	delete_buf (defaultKey);
    if (logfile)
	delete_buf (logfile);
    if (passPhrase)
	delete_buf (passPhrase);
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
    if (FAILED (hr))
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


/* XXX: I would prefer to use MapiGPGME::passphraseCallback, but member 
        functions have an incompatible calling convention. */
static int 
passphraseCallback (void *opaque, const char *uid_hint, 
		    const char *passphrase_info,
		    int last_was_bad, int fd)
{
    MapiGPGME *ctx = (MapiGPGME*)opaque;
    const char *passwd = ctx->getPassphrase ();

    DWORD nwritten = 0;
    WriteFile ((HANDLE)fd, passwd, strlen (passwd), &nwritten, NULL);

    return 0;
}

int 
MapiGPGME::decrypt (void)
{
    char *body = getBody();
    char *newBody = NULL;
    int err;

    if (!passPhrase) {
	if (nstorePasswd == 0)
	    err = op_decrypt_start (body, &newBody);
	else {
	    char *pCache = NULL;
	    err = op_decrypt_start_ext (body, &newBody, &pCache);
	    if (!err)
		storePassphrase (pCache);
	    xfree (pCache);
	}
    }
    else
	err = op_decrypt_next (passphraseCallback, this, body, &newBody);
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
MapiGPGME::streamOnFile (const char *file, LPSTREAM to)
{
    HRESULT hr;
    LPSTREAM from;

    hr = OpenStreamOnFile (MAPIAllocateBuffer, MAPIFreeBuffer,
   		           STGM_READ, (char*) file, 
			   NULL, &from);
    if (SUCCEEDED (hr)) {
	STATSTG statInfo;
	from->Stat (&statInfo, STATFLAG_NONAME);
	from->CopyTo (to, statInfo.cbSize, NULL, NULL);
	to->Commit (0);
	to->Release ();
	from->Release ();
    }
    return SUCCEEDED (hr);
}


int
MapiGPGME::getMessageFlags ()
{
    HRESULT hr;
    LPSPropValue propval = NULL;
    int flags = 0;

    hr = HrGetOneProp (msg, PR_MESSAGE_FLAGS, &propval);
    if (FAILED (hr))
	return 0;
    flags = propval->Value.l;
    MAPIFreeBuffer (propval);
    return flags;
}

int
MapiGPGME::getMessageHasAttachments ()
{
    HRESULT hr;
    LPSPropValue propval = NULL;
    int nattach = 0;

    hr = HrGetOneProp (msg, PR_HASATTACH, &propval);
    if (FAILED (hr))
	return 0;
    nattach = propval->Value.b? 1 : 0;
    MAPIFreeBuffer (propval);
    return nattach;   
}


bool
MapiGPGME::setMessageAccess (int access)
{
    HRESULT hr;
    SPropValue prop;
    prop.ulPropTag = PR_ACCESS;
    prop.Value.l = access;
    hr = HrSetOneProp (msg, &prop);
    return FAILED (hr)? false: true;
}


bool
MapiGPGME::setAttachMethod (LPATTACH obj, int mode)
{
    SPropValue prop;
    HRESULT hr;
    prop.ulPropTag = PR_ATTACH_METHOD;
    prop.Value.ul = mode;
    hr = HrSetOneProp (obj, &prop);
    return FAILED (hr)? true : false;
}

int
MapiGPGME::getAttachMethod (LPATTACH obj)
{
    HRESULT hr;
    LPSPropValue propval = NULL;
    int method = 0;

    hr = HrGetOneProp ((LPMAPIPROP)obj, PR_ATTACH_METHOD, &propval);
    if (FAILED (hr))
	return 0;
    method = propval->Value.ul;
    MAPIFreeBuffer (propval);
    return method;
}

bool
MapiGPGME::setAttachFilename (LPATTACH obj, const char *name, bool islong)
{
    HRESULT hr;
    SPropValue prop;

    if (!islong)
	prop.ulPropTag = PR_ATTACH_FILENAME;
    else
	prop.ulPropTag = PR_ATTACH_LONG_FILENAME;
    prop.Value.lpszA = (char*) name;   
    hr = HrSetOneProp (obj, &prop);
    return FAILED (hr)? false: true;
}


char*
MapiGPGME::getAttachFilename (LPATTACH obj)
{
    LPSPropValue propval;
    HRESULT hr;
    char *name;

    hr = HrGetOneProp ((LPMAPIPROP)obj, PR_ATTACH_LONG_FILENAME, &propval);
    if (FAILED(hr)) {
	hr = HrGetOneProp ((LPMAPIPROP)obj, PR_ATTACH_FILENAME, &propval);
	if (SUCCEEDED(hr)) {
	    name = xstrdup (propval[0].Value.lpszA);
	    MAPIFreeBuffer (propval);
	}
	else
	    return NULL;
    }
    else {
	name = xstrdup (propval[0].Value.lpszA);
	MAPIFreeBuffer (propval);	
    }
    return name;
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

    for (int j = 0; j < (int) lpAttRows->cRows; j++) {
	TCHAR sn[20];
	itoa(j, sn, 10);
	string sFilename = sFilenameAtt;
	sFilename += sn;
	sFilename += ".tmp";
	string sFilenameDest = sFilename + ".pgp";
	tring sAttName = "attach";
	LPATTACH pAttach = NULL;
	LPSTREAM pStreamAtt = NULL;
        int nAttachment = lpAttRows->aRow[j].lpProps[0].Value.ul;

        HRESULT hr = pMessage->OpenAttach (nAttachment, NULL, MAPI_BEST_ACCESS, &pAttach);
	if (SUCCEEDED(hr)) {
	    BOOL bSaveAttSuccess = TRUE;
	    int method = getAttachMethod (pAttach);	    
	    BOOL bEmbedded = (method == ATTACH_EMBEDDED_MSG);
	    BOOL bEmbeddedOLE = (method == ATTACH_OLE);
		
	    if (SUCCEEDED(hr) && bEmbedded) {
		    if ((nAction == PROATT_DECRYPT) || (nAction == PROATT_ENCRYPT_AND_SIGN)) {
			LPMESSAGE pMessageEmb = NULL;
			hr = pAttach->OpenProperty(PR_ATTACH_DATA_OBJ, &IID_IMessage, 0, MAPI_MODIFY, 
						   (LPUNKNOWN*) &pMessageEmb);
			if (SUCCEEDED(hr)) {
			    setWindow(hWnd);
			    setMessage (pMessageEmb);
			    if (nAction == PROATT_DECRYPT) {
				if (decrypt ())
				    bSuccess = FALSE;
			    }
			    if (nAction == PROATT_ENCRYPT_AND_SIGN) {
				if (signEncrypt ())
				    bSuccess = FALSE;
			    }
			    pMessageEmb->SaveChanges (FORCE_SAVE);
			    pAttach->SaveChanges (FORCE_SAVE);
			    pMessageEmb->Release ();
			    m_nDecryptedAttachments++;
			}
		    }
		}
		BOOL bShortFileName = FALSE;
		if (!bEmbedded && bSaveAttSuccess) {
		    sAttName = getAttachFilename(pAttach);
		
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

		// XXX: ignore bEmbeddedOLE
		if (!bEmbedded && !bEmbeddedOLE && bSaveAttSuccess)		
		{		
		    hr = pAttach->OpenProperty (PR_ATTACH_DATA_BIN, &IID_IStream, 0, 0, (LPUNKNOWN*) &pStreamAtt);
		    if (HR_SUCCEEDED(hr))
		    {	
			int err = streamOnFile (sFilename.c_str(), pStreamAtt);
			if (err)
			    pStreamAtt->Release ();
			else {
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
		    if (nAction == PROATT_DECRYPT) {
			setWindow(hWnd);
			op_decrypt_file (A2OLE(sFilename.c_str()), A2OLE(sFilenameDest.c_str());
		    }
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
			if (setAttachMethod (lpNewAttach, ATTACH_BY_VALUE)) {
			    setAttachFilename (lpNewAttach, sAttName.c_str(), false);

			    LPSTREAM  pNewStream;
			    hr = lpNewAttach->OpenProperty (PR_ATTACH_DATA_BIN, &IID_IStream, 0,	
							   MAPI_CREATE | MAPI_MODIFY, (LPUNKNOWN*) &pNewStream);
			    if (SUCCEEDED (hr))
			    {	
				err = streamOnFile (sFilenameDest.c_str(), pNewStream);
				if (!err)
				    lpNewAttach->SaveChanges (FORCE_SAVE);
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

    load_extension_value ("logFile", &val);
    if (val == NULL ||*val == '"')
	logfile = NULL;
    else
	setLogFile (val);
    xfree (val); val=NULL;

    return 0;
}

int
MapiGPGME::writeOptions (void)
{
    store_extension_value ("encryptDefault", doEncrypt? "1" : "0");
    store_extension_value ("signDefault", doSign? "1" : "0");
    store_extension_value ("addDefaultKey", encryptDefault? "1" : "0");
    if (logfile != NULL)
	store_extension_value ("logFile", logfile);

    char buf[32];
    sprintf (buf, "%d", nstorePasswd);
    store_extension_value ("storePasswdTime", buf);

    return 0;
}
