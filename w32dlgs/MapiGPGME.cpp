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
#include <initguid.h>
#include <mapiguid.h>
#include <atlbase.h>

#include "gpgme.h"
#include "intern.h"
#include "HashTable.h"
#include "MapiGPGME.h"
#include "engine.h"
#include "keycache.h"

/* attachment information */
#define ATTR_SIGN(action) ((action) & GPG_ATTACH_SIGN)
#define ATTR_ENCR(action) ((action) & GPG_ATTACH_ENCRYPT)

/* default file for logging */
#define LOGFILE "c:\\mapigpgme.log"

/* default extension for attachments */
#define EXT_MSG ".pgp"
#define EXT_SIG ".sig"

/* memory macros */
#define delete_buf(buf) delete [] (buf)

#define fail_if_null(p) do { if (!(p)) abort (); } while (0)


MapiGPGME::MapiGPGME (LPMESSAGE msg)
{
    this->encFormat = GPG_FMT_CLASSIC;
    this->passCache = new HashTable();
    this->msg = msg;
    op_init ();
    log_debug (LOGFILE, "constructor %p\r\n", msg);
}


MapiGPGME::MapiGPGME ()
{
    this->encFormat = GPG_FMT_CLASSIC;
    this->passCache = new HashTable();
    op_init ();
    log_debug (LOGFILE, "constructor null\r\n");
}


MapiGPGME::~MapiGPGME ()
{
    unsigned i=0;

    log_debug (LOGFILE, "destructor %p\r\n", msg);
    op_deinit ();
    if (defaultKey)
	delete_buf (defaultKey);
    if (logfile)
	delete_buf (logfile);
    
    log_debug (LOGFILE, "hash entries %d\r\n", passCache->size ());
    for (i = 0; i < passCache->size (); i++) {
	cache_item_t t = (cache_item_t)passCache->get (i);
	if (t != NULL)
	    cache_item_free (t);
    }
    
    freeAttachments ();
    delete passCache;
}


int
MapiGPGME::setRTFBody (char *body)
{
    setMessageAccess (MAPI_ACCESS_MODIFY);
    HWND rtf = findMessageWindow (parent);
    if (rtf != NULL) {
	log_debug (LOGFILE, "setRTFBody: window handle %p", rtf);
	SetWindowText (rtf, body);
	return TRUE;
    }
    return FALSE;
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


const char*
MapiGPGME::getPassphrase (const char *keyid)
{
    cache_item_t item = (cache_item_t)passCache->get(keyid);
    if (item != NULL)
	return item->pass;
    return NULL;
}


void
MapiGPGME::storePassphrase (void *itm)
{
    cache_item_t item = (cache_item_t)itm;
    cache_item_t old;
    old = (cache_item_t)passCache->get(item->keyid);
    if (old != NULL)
	cache_item_free (old);
    passCache->put (item->keyid+8, item);
    log_debug (LOGFILE, "put keyid %s = '%s'\r\n", item->keyid+8, item->pass);
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
    xfree (newBody);
    freeRecipients (recipients);
    freeUnknownKeys (unknown, n);
    if (0 && hasAttachments ()) { /*x:test*/
	log_debug (LOGFILE, "encrypt attachments\r\n");
	recipSet = (void *)keys;
	encryptAttachments (parent);
    }
    freeKeyArray ((void **)keys);
    return err;
}


/* XXX: I would prefer to use MapiGPGME::passphraseCallback, but member 
        functions have an incompatible calling convention. */
int 
passphraseCallback (void *opaque, const char *uid_hint, 
		    const char *passphrase_info,
		    int last_was_bad, int fd)
{
    MapiGPGME *ctx = (MapiGPGME*)opaque;
    const char *passwd;
    char keyid[16+1];
    DWORD nwritten = 0;
    int i=0;

    while (uid_hint && *uid_hint != ' ')
	keyid[i++] = *uid_hint++;
    keyid[i] = '\0';
    
    passwd = ctx->getPassphrase (keyid+8);
    log_debug (LOGFILE, "get keyid %s = '%s'\r\n", keyid+8, passwd);
    if (passwd != NULL) {
	WriteFile ((HANDLE)fd, passwd, strlen (passwd), &nwritten, NULL);
	WriteFile ((HANDLE)fd, "\n", 1, &nwritten, NULL);
    }

    return 0;
}

int 
MapiGPGME::decrypt (void)
{
    gpg_type_t id;
    char *body = getBody ();
    char *newBody = NULL;
    int err;

    id = getMessageType (body);
    if (id == GPG_TYPE_CLEARSIG)
	return verify ();

    if (nstorePasswd == 0)
	err = op_decrypt_start (body, &newBody);
    else if (passCache->size () == 0) {
	/* XXX: use the callback to see if a cached passphrase is available. If not
	        call the real passphrase callback and store the passphrase. */
	cache_item_t itm = NULL;
	err = op_decrypt_start_ext (body, &newBody, &itm);
	if (!err)
	    storePassphrase (itm);
    }
    else
	err = op_decrypt_next (passphraseCallback, this, body, &newBody);
    if (err)
	MessageBox (NULL, op_strerror (err), "GPG Decryption", MB_ICONERROR|MB_OK);
    else
	setRTFBody (newBody);

    if (hasAttachments ()) {
	log_debug (LOGFILE, "decrypt attachments\r\n");
	decryptAttachments (parent);
    }
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


gpg_type_t
MapiGPGME::getMessageType (const char *body)
{
    if (strstr (body, "BEGIN PGP MESSAGE"))
	return GPG_TYPE_MSG;
    if (strstr (body, "BEGIN PGP SIGNED MESSAGE"))
	return GPG_TYPE_CLEARSIG;
    if (strstr (body, "BEGIN PGP SIGNATURE"))
	return GPG_TYPE_SIG;
    /* XXX: pubkey, seckey */
    return GPG_TYPE_NONE;
}



int
MapiGPGME::doCmdFile(int action, const char *in, const char *out)
{
    if (ATTR_SIGN (action) && ATTR_ENCR (action))
	return op_sign_encrypt_file (recipSet, in, out);
    if (ATTR_SIGN (action) && !ATTR_ENCR (action))
	return op_sign_file (OP_SIG_NORMAL, in, out);
    if (!ATTR_SIGN (action) && ATTR_ENCR (action))
	return op_encrypt_file (recipSet, in, out);
    return !op_decrypt_file (in, out);
}


int
MapiGPGME::doCmdAttach(int action)
{
    if (ATTR_SIGN (action) && ATTR_ENCR (action))
	return signEncrypt ();
    if (ATTR_SIGN (action) && !ATTR_ENCR (action))
	return sign ();
    if (!ATTR_SIGN (action) && ATTR_ENCR (action))
	return encrypt ();
    return decrypt ();
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
    if (0 && hasAttachments ()) { /*x:test*/
	log_debug (LOGFILE, "encrypt attachments");
	recipSet = (void *)keys;
	encryptAttachments (parent);
    }
    freeKeyArray ((void **)keys);
    gpgme_key_release (locusr);
    return err;
}


int 
MapiGPGME::verify ()
{
    char *body = getBody();
    char *newBody = NULL;
    
    int err = op_verify_start (body, &newBody);
    if (err)
	MessageBox (NULL, op_strerror (err), "GPG Verify", MB_ICONERROR|MB_OK);
    else
	setRTFBody (newBody);

    delete_buf (body);
    xfree (newBody);
    return err;
}


void MapiGPGME::setDefaultKey (const char *key)
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
    this->parent = hwnd;
}


/* We need this to find the mailer window because we directly change the text
   of the window instead of the MAPI object itself. */
HWND
MapiGPGME::findMessageWindow (HWND parent)
{
    HWND child;

    if (parent == NULL)
	return NULL;

    child = GetWindow (parent, GW_CHILD);
    while (child != NULL) {
	char buf[1025];
	HWND rtf;

	memset (buf, 0, sizeof (buf));
	GetWindowText (child, buf, sizeof (buf)-1);
	if (getMessageType (buf) != GPG_TYPE_NONE)
	    return child;
	rtf = findMessageWindow (child);
	if (rtf != NULL)
	    return rtf;
	child = GetNextWindow (child, GW_HWNDNEXT);	
    }
    log_debug (LOGFILE, "no message window found.\r\n");
    return NULL;
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


const char*
MapiGPGME::getAttachmentExtension (const char *fname)
{
    static char ext[4];
    char *p;

    p = strchr (fname, '.');
    if (p != NULL) {
	int pos = (p-fname);
	
	strncpy (ext, fname+pos, 4);
	if (stricmp (ext+1, "gpg") == 0 ||
	    stricmp (ext+1, "pgp") == 0 ||
	    stricmp (ext+1, "asc") == 0)
	    return ext;
    }
    return EXT_MSG;
}

const char*
MapiGPGME::getPGPExtension (int action)
{
    if (ATTR_SIGN (action))
	return EXT_SIG;
    return EXT_MSG;
}


bool 
MapiGPGME::setXHeader (const char *name, const char *val)
{  
    USES_CONVERSION;
    LPMDB lpMdb = NULL;
    HRESULT hr = NULL;  
    LPSPropTagArray pProps = NULL;
    SPropValue pv;
    MAPINAMEID mnid[1];	
    // {00020386-0000-0000-C000-000000000046}  ->  GUID For X-Headers	
    GUID guid = {0x00020386, 0x0000, 0x0000, {0xC0, 0x00, 0x00, 0x00,
		 0x00, 0x00, 0x00, 0x46} };

    mnid[0].lpguid = &guid;
    mnid[0].ulKind = MNID_STRING;
    mnid[0].Kind.lpwstrName = A2W (name);

    hr = msg->GetIDsFromNames (1, (LPMAPINAMEID*)mnid, MAPI_CREATE, &pProps);
    if (FAILED (hr))
	return false;
    
    pv.ulPropTag = (pProps->aulPropTag[0] & 0xFFFF0000) | PT_STRING8;
    pv.Value.lpszA = (char *)val;
    hr = HrSetOneProp(msg, &pv);	
    if (!SUCCEEDED (hr))
	return false;

    return true;
}


char*
MapiGPGME::getXHeader (const char *name)
{
    /* XXX: PR_TRANSPORT_HEADERS is not available in my MSDN. */
    return NULL;
}


void
MapiGPGME::freeAttachments (void)
{
    if (attachTable != NULL) {
        attachTable->Release ();
	attachTable = NULL;
    }
    if (attachRows != NULL) {
        FreeProws (attachRows);
	attachRows = NULL;
    }
}


int
MapiGPGME::getAttachments (void)
{
    static SizedSPropTagArray (1L, PropAttNum) = {1L, {PR_ATTACH_NUM}};
    HRESULT hr;    
   
    hr = msg->GetAttachmentTable (0, &attachTable);
    if (FAILED (hr))
	return FALSE;

    hr = HrQueryAllRows (attachTable, (LPSPropTagArray) &PropAttNum,
			 NULL, NULL, 0L, &attachRows);
    if (FAILED (hr)) {
	freeAttachments ();
	return FALSE;
    }
    return TRUE;
}


LPATTACH
MapiGPGME::openAttachment (int pos)
{
    HRESULT hr;
    LPATTACH att = NULL;
    
    hr = msg->OpenAttach (pos, NULL, MAPI_BEST_ACCESS, &att);	
    if (SUCCEEDED(hr))
	return att;
    return NULL;
}


void
MapiGPGME::closeAttachment (LPATTACH att)
{
    att->Release ();
}


int
MapiGPGME::processAttachment (LPATTACH att, HWND hwnd, int action)
{    
    int method = getAttachMethod (att);
    BOOL success = TRUE;
    HRESULT hr;

    if (action == GPG_ATTACH_NONE)
	return FALSE;

    switch (method) {
    case ATTACH_EMBEDDED_MSG:
	LPMESSAGE emb;

	hr = att->OpenProperty (PR_ATTACH_DATA_OBJ, &IID_IMessage, 0, 
			        MAPI_MODIFY, (LPUNKNOWN*) &emb);
	if (FAILED (hr))
	    return FALSE;
	setWindow (hwnd);
	setMessage (emb);
	if (doCmdAttach (action))
	    success = FALSE;
	emb->SaveChanges (FORCE_SAVE);
	att->SaveChanges (FORCE_SAVE);
	emb->Release ();
	break;

    case ATTACH_BY_VALUE:
	LPSTREAM stream;
	char *inname;
	char *outname;
	
	inname = getAttachFilename (att);
	if (action != GPG_ATTACH_DECRYPT) {
	    outname = (char *)xcalloc (1, strlen (inname) + 16 + 4);
	    sprintf (outname, "%s%s", inname, getPGPExtension (action));
	    log_debug (LOGFILE, "enc outname: '%s'\r\n", outname);
	}
	else {
	    outname = (char*)xcalloc (1, strlen (inname) + 4);
	    strcpy (outname, inname);
	    outname[strlen (outname)-4] = '\0';
	    log_debug (LOGFILE, "dec outname: '%s'\r\n", outname);
	}
	success = FALSE;
	hr = att->OpenProperty (PR_ATTACH_DATA_BIN, &IID_IStream, 
			        0, 0, (LPUNKNOWN*) &stream);
	if (!FAILED (hr) && streamOnFile (inname, stream)) {
	    if (doCmdFile (action, inname, outname))
		success = TRUE;
	}
	stream->Release ();
	xfree (inname);
	xfree (outname);
	break;

    case ATTACH_BY_REF_ONLY:
	break;

    }
    return success;
}


int 
MapiGPGME::decryptAttachments (HWND hwnd)
{
    return 0;
}

int
MapiGPGME::encryptAttachments (HWND hwnd)
{
    int n = countAttachments ();
    if (!n)
	return TRUE;
    if (!getAttachments ())
	return FALSE;

    for (int i=0; i < n; i++) {
	LPATTACH amsg = openAttachment (i);
	if (msg == NULL)
	    continue;
	processAttachment (amsg, hwnd, GPG_ATTACH_ENCRYPT);
	closeAttachment (amsg);	
    }
    freeAttachments ();
    return 0;
}


int
MapiGPGME::processAttachments (HWND hwnd, int action, const char **pFileNameVector)
{
#if 0
    USES_CONVERSION;
    BOOL bSuccess = TRUE;
    TCHAR szTempPath[MAX_PATH];
    HRESULT hr;       
    BOOL bSaveAttSuccess = TRUE;

    GetTempPath (MAX_PATH, szTempPath);
    string sFilenameAtt = szTempPath;
    sFilenameAtt += sPrefix;
   
    if (!getAttachments ())
	return FALSE;

    for (int j = 0; j < countAttachments(); j++) {
	TCHAR sn[20];
	itoa(j, sn, 10);
	string sFilename = sFilenameAtt;
	sFilename += sn;
	sFilename += ".tmp";
	string sFilenameDest = sFilename + EXT_MSG;
	tring sAttName = "attach";
	LPATTACH pAttach = NULL;
	LPSTREAM pStreamAtt = NULL;
        int nAttachment = lpAttRows->aRow[j].lpProps[0].Value.ul;

	pAttach = openAttachment (nAttachment);
	if (pAttach) {
	    BOOL bSaveAttSuccess = TRUE;
	    int method = getAttachMethod (pAttach);	    
	    BOOL bEmbedded = (method == ATTACH_EMBEDDED_MSG);
	    BOOL bEmbeddedOLE = (method == ATTACH_OLE);
		
	    if (SUCCEEDED(hr) && bEmbedded) {
		bSuccess = processAttachment (att, hwnd, nAction);		
	    }
		BOOL bShortFileName = FALSE;
		if (!bEmbedded && bSaveAttSuccess) {
		    sAttName = getAttachFilename(pAttach);
		
		    // use correct file extensions for temp files to allow virus checking
		    if (nAction == PROATT_DECRYPT)
		    {
			sFilename = sFilenameAtt + sAttName;
			sFilenameDest = sFilename;
			sFilename + = getAttachmentExtension (sFilename);					
		    }
		    if (nAction == PROATT_ENCRYPT_AND_SIGN)
		    {
			sFilename = sFilenameAtt + sAttName;
			sFilenameDest = sFilename + EXT_MSG;
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
			    sAttName += EXT_SIG;
			else
			    sAttName += EXT_MSG;
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
			    deleteAttachment (nAttachment);
			#endif
		    }

		    deleteAttachmment (nAttachment);

		    ULONG ulNewAttNum;
		    LPATTACH lpNewAttach = NULL;
		    lpNewAttach = createAttachment (ulNewAttNum);
		    if (lpNewAttach) {
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
		closeAttachment (pAttach);
	    }
	}
    }

    freeAttachments ();
    freeKeyArray (recipSet);
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
    load_extension_value ("encodingFormat", &val);
    encFormat = val == NULL? GPG_FMT_CLASSIC  : atol (val);
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
    
    sprintf (buf, "%d", encFormat);
    store_extension_value ("encodingFormat", buf);

    return 0;
}
