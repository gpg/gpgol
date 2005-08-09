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

#ifdef HAVE_CONFIG_H
#include <config.h>
#endif

#include <windows.h>
#include <time.h>

#ifdef __MINGW32__
# include <initguid.h>
# include "mymapi.h"
# include "mymapitags.h"
#else /* !__MINGW32__ */
# include <atlbase.h>
# include <mapidefs.h>
# include <mapiutil.h>
# include <initguid.h>
# include <mapiguid.h>
#endif /* !__MINGW32__ */

#include "gpgme.h"
#include "keycache.h"
#include "intern.h"
#include "HashTable.h"
#include "MapiGPGME.h"
#include "engine.h"

// #define  FreeProws(a)               do { ; } while(0)
// #define  HrGetOneProp(a,b,c)        0
// #define  HrSetOneProp(a,b)          0
// #define  FPropExists(a,b)           0
// #define  FreePadrlist(a)            0
// #define  HrQueryAllRows(a,b,c,d,e,f)  0
// #define  PpropFindProp(a,b,c)       0
// #define  RTFSync(a,b,c)             0

/* These were omitted from the standard headers */
#ifndef PR_BODY_HTML
#define PR_BODY_HTML (PROP_TAG(PT_TSTRING, 0x1013))
#endif

/* Attachment information. */
#define ATT_SIGN(action) ((action) & GPG_ATTACH_SIGN)
#define ATT_ENCR(action) ((action) & GPG_ATTACH_ENCRYPT)
#define ATT_PREFIX ".pgpenc"

#define DEFAULT_ATTACHMENT_FORMAT GPG_FMT_CLASSIC


/* default extension for attachments */
#define EXT_MSG "pgp"
#define EXT_SIG "sig"

/* memory macros */
#define delete_buf(buf) delete [] (buf)

#define fail_if_null(p) do { if (!(p)) abort (); } while (0)


class MapiGPGMEImpl : public MapiGPGME
{
public:    
  MapiGPGMEImpl () 
  {
    clearConfig ();
    clearObject ();
    this->passCache = new HashTable ();
    op_init ();
    prepareLogging ();
    logDebug ("constructor null\r\n");
  }

  MapiGPGMEImpl (LPMESSAGE msg)
  {
    clearConfig ();
    clearObject ();
    this->msg = msg;
    this->passCache = new HashTable ();
    op_init ();
    prepareLogging ();
    logDebug ("constructor %p\r\n", msg);
  }
  
  ~MapiGPGMEImpl ()
  {
    unsigned int i=0;

    logDebug ("destructor %p\r\n", msg);
    op_deinit ();
    if (defaultKey)
      {
	delete_buf (defaultKey);
	defaultKey = NULL;
      }
    if (logfile) 
      {
	delete_buf (logfile);
	logfile = NULL;
      }
    logDebug ("hash entries %d\r\n", passCache->size ());
    for (i = 0; i < passCache->size (); i++) 
      {
	cache_item_t t = (cache_item_t)passCache->get (i);
	if (t != NULL)
          cache_item_free (t);
      }
    delete passCache; 
    passCache = NULL;
    freeAttachments ();
    cleanupTempFiles ();
  }

  void __stdcall destroy ()
  {
    delete this;
  }

  void operator delete (void *p) 
  {
    ::operator delete (p);
  }  

  
public:
  int __stdcall encrypt (void);
  int __stdcall decrypt (void);
  int __stdcall sign (void);
  int __stdcall verify (void);
  int __stdcall signEncrypt (void);

  int __stdcall doCmd(int doEncrypt, int doSign);
  int __stdcall doCmdAttach(int action);
  int __stdcall doCmdFile(int action, const char *in, const char *out);

  const char* __stdcall getLogFile (void) { return logfile; }
  void __stdcall setLogFile (const char *logfile)
  { 
    if (this->logfile) 
      {
        delete []this->logfile;
        this->logfile = NULL;
      }
    this->logfile = new char [strlen (logfile)+1];
    if (this->logfile != NULL)
      strcpy (this->logfile, logfile);
  }

  int __stdcall getStorePasswdTime (void)
  {
    return nstorePasswd;
  }

  void __stdcall setStorePasswdTime (int nCacheTime)
  {
    this->nstorePasswd = nCacheTime; 
  }

  bool __stdcall getEncryptDefault (void)
  {
    return doEncrypt;
  }

  void __stdcall setEncryptDefault (bool doEncrypt)
  {
    this->doEncrypt = doEncrypt; 
  }

  bool __stdcall getSignDefault (void)
  { 
    return doSign; 
  }

  void __stdcall setSignDefault (bool doSign)
  {
    this->doSign = doSign;
  }

  bool __stdcall getEncryptWithDefaultKey (void)
  {
    return encryptDefault;
  }
  
  void __stdcall setEncryptWithDefaultKey (bool encryptDefault)
  {
    this->encryptDefault = encryptDefault;
  }

  bool __stdcall getSaveDecryptedAttachments (void) 
  { 
    return saveDecryptedAtt;
  }

  void __stdcall setSaveDecryptedAttachments (bool saveDecrAtt)
  {
    this->saveDecryptedAtt = saveDecrAtt;
  }

  void __stdcall setEncodingFormat (int fmt)
  {
    encFormat = fmt; 
  }

  int __stdcall getEncodingFormat (void) 
  {
    return encFormat;
  }

  void __stdcall setSignAttachments (bool signAtt)
  {
    this->autoSignAtt = signAtt; 
  }

  bool __stdcall getSignAttachments (void)
  {
    return autoSignAtt;
  }

  void __stdcall setEnableLogging (bool val)
  {
    this->enableLogging = val;
  }

  bool __stdcall getEnableLogging (void)
  {
    return this->enableLogging;
  }

  int __stdcall readOptions (void);
  int __stdcall writeOptions (void);

  const char* __stdcall getAttachmentExtension (const char *fname);
  void __stdcall freeAttachments (void);
  int __stdcall getAttachments (void);
  
  int __stdcall countAttachments (void) 
  { 
    if (attachRows == NULL)
      return -1;
    return (int) attachRows->cRows; 
  }

  bool __stdcall hasAttachments (void)
  {
    if (attachRows == NULL)
      getAttachments ();
    bool has = attachRows->cRows > 0? true : false;
    freeAttachments ();
    return has;
  }

  bool __stdcall deleteAttachment (int pos)
  {
    if (msg->DeleteAttach (pos, 0, NULL, 0) == S_OK)
      return true;
    return false;
  }

  LPATTACH __stdcall createAttachment (int &pos)
  {
    ULONG attnum;	
    LPATTACH newatt = NULL;
    
    if (msg->CreateAttach (NULL, 0, &attnum, &newatt) == S_OK)
      {
        pos = attnum;
        return newatt;
      }
    return NULL;
  }

  int __stdcall startKeyManager ();
  void __stdcall startConfigDialog (HWND parent);

  int __stdcall attachPublicKey (const char *keyid);

  void __stdcall setDefaultKey (const char *key);
  char* __stdcall getDefaultKey (void);

  void __stdcall setMessage (LPMESSAGE msg);
  void __stdcall setWindow (HWND hwnd);

  const char* __stdcall getPassphrase (const char *keyid);
  void __stdcall storePassphrase (void *itm);
  outlgpg_type_t __stdcall getMessageType (const char *body);

  void __stdcall logDebug (const char *fmt, ...);
    
  void __stdcall clearPassphrase (void) 
  {
    if (passCache != NULL)
      passCache->clear ();
  }


private:
  HWND	      parent;
  char	      *defaultKey;
  HashTable   *passCache;
  LPMESSAGE   msg;
  LPMAPITABLE attachTable;
  LPSRowSet   attachRows;
  void	      *recipSet;
  bool        enableLogging;

  /* Options */
  char    *logfile;
  int	  nstorePasswd;  /* Time in seconds the passphrase is stored. */
  bool    doEncrypt;
  bool    doSign;
  bool    encryptDefault;
  bool    saveDecryptedAtt; /* Save decrypted attachments. */
  bool    autoSignAtt;	    /* Sign all outgoing attachments. */
  int	    encFormat;      /* Encryption format for attachments. */

  void displayError (HWND root, const char *title);
  void prepareLogging (void);
  void clearObject (void);
  void clearConfig (void)
  {
    nstorePasswd = 0;
    doEncrypt = false;
    doSign = false;
    encryptDefault = false;
    saveDecryptedAtt = false;
    autoSignAtt = false;
    encFormat = DEFAULT_ATTACHMENT_FORMAT;
    enableLogging = false;
  }

  HWND findMessageWindow (HWND parent);
  void rtfSync (char *body);
  int  setBody (char *body, bool isHtml);
  int  setRTFBody (char *body);
  bool isMessageEncrypted (void);
  bool isHtmlBody (const char *body);
  bool isHtmlMessage (void);
  char *getBody (bool isHtml);
  char *addHtmlLineEndings (char *newBody);
    
  void cleanupTempFiles ();
  int countRecipients (char **recipients);
  char **getRecipients (bool isRootMsg);
  void freeRecipients (char **recipients);
  void freeUnknownKeys (char **unknown, int n);
  void freeKeyArray (void **key);

  bool  setAttachMethod (LPATTACH obj, int mode);
  int   getAttachMethod (LPATTACH obj);
  char* getAttachFilename (LPATTACH obj);
  char* getAttachPathname (LPATTACH obj);
  bool  setAttachFilename (LPATTACH obj, const char *name, bool islong);
  int   getMessageFlags ();
  int   getMessageHasAttachments ();
  bool  setMessageAccess (int access);
  bool  setXHeader (const char *name, const char *val);
  char* getXHeader (const char *name);
  bool  checkAttachmentExtension (const char *ext);
  const char* getPGPExtension (int action);
  char* generateTempname (const char *name);
  int   streamOnFile (const char *file, LPATTACH att);
  int   streamFromFile (const char *file, LPATTACH att);
  int   encryptAttachments (HWND hwnd);
  int   decryptAttachments (HWND hwnd);
  int   signAttachments (HWND hwnd);
  LPATTACH openAttachment (int pos);
  void  releaseAttachment (LPATTACH att);
  int   processAttachment (LPATTACH *att, HWND hwnd, int pos, int action);
  bool  saveDecryptedAttachment (HWND root, const char *srcname);
  bool  signAttachment (const char *datfile);

};


/* Create an instance of the MapiGPGME class. MSG may be NULL. */
MapiGPGME *
_CreateMapiGPGME (LPMESSAGE msg)
{
  return new MapiGPGMEImpl (msg);
}
MapiGPGME *
CreateMapiGPGME (LPMESSAGE msg)
{
  return new MapiGPGMEImpl (msg);
}



void 
MapiGPGMEImpl::logDebug (const char *fmt, ...)
{
    FILE * logfp;
    va_list a;

    if (enableLogging == false || this->logfile == NULL)
	return;

    logfp = fopen (this->logfile, "a+b");
    if (logfp == NULL)
	return;
    va_start (a, fmt);
    vfprintf (logfp, fmt, a);
    va_end (a);
    fclose (logfp);
}


void
MapiGPGMEImpl::clearObject (void)
{
    this->attachRows = NULL;
    this->attachTable = NULL;
    this->defaultKey = NULL;
    this->logfile = NULL;
    this->recipSet = NULL;
    this->parent = NULL;
    this->msg = NULL;
}


void
MapiGPGMEImpl::prepareLogging (void)
{
    char *val = NULL;

    load_extension_value ("logFile", &val);
    if (val != NULL && *val != '"' && *val != 0) {
	setLogFile (val);
	setEnableLogging (true);
	xfree (val);	
    }
}


void 
MapiGPGMEImpl::cleanupTempFiles (void)
{
    HANDLE hd;
    WIN32_FIND_DATA fnd;
    char path[MAX_PATH+32], tmp[MAX_PATH+4];

    GetTempPath (sizeof (path)-4, path);
    if (path[strlen (path)-1] != '\\')
	strcat (path, "\\");
    strcpy (tmp, path);
    strcat (path, "*"ATT_PREFIX"*");
    hd = FindFirstFile (path, &fnd);
    if (hd == INVALID_HANDLE_VALUE)
	return;
    do {
	char *p = (char *)xcalloc (1, strlen (tmp) + strlen (fnd.cFileName) +2);
	sprintf (p, "%s%s", tmp, fnd.cFileName);
	logDebug ("delete tmp %s\r\n", p);
	DeleteFile (p);
	xfree (p);
    } while (FindNextFile (hd, &fnd) == TRUE);
    FindClose (hd);
}


int
MapiGPGMEImpl::setRTFBody (char *body)
{
    setMessageAccess (MAPI_ACCESS_MODIFY);
    HWND rtf = findMessageWindow (parent);
    if (rtf != NULL) {
	logDebug ("setRTFBody: window handle %p\r\n", rtf);
	SetWindowText (rtf, body);
	return TRUE;
    }
    return FALSE;
}


int 
MapiGPGMEImpl::setBody (char *body, bool isHtml)
{
    /* XXX: handle richtext/html */
    SPropValue sProp; 
    HRESULT hr;
    int rc = TRUE;
    
    if (body == NULL) {
	logDebug ("setBody with empty buffer\r\n");
	return FALSE;
    }
    rtfSync (body);
    sProp.ulPropTag = isHtml? PR_BODY_HTML : PR_BODY;
    sProp.Value.lpszA = body;
    hr = HrSetOneProp (msg, &sProp);
    if (FAILED (hr))
	rc = FALSE;
    logDebug ("setBody (html=%d) rc=%d '%s'\r\n", rc, isHtml, body);
    return rc;
}


void
MapiGPGMEImpl::rtfSync (char *body)
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


bool 
MapiGPGMEImpl::isHtmlBody (const char *body)
{
    char *p1, *p2;

    /* XXX: it is possible but unlikely that the message text
            contains the used keywords. */
    p1 = strstr (body, "<HTML>");
    p2 = strstr (body, "</HTML>");
    if (p1 && p2)
	return true;
    p1 = strstr (body, "<html>");
    p2 = strstr (body, "</html>");
    if (p1 && p2)
	return true;
    /* XXX: use case insentensive strstr version. */
    return false;
}


bool
MapiGPGMEImpl::isHtmlMessage (void)
{
    char *body = getBody (true);
    if (body != NULL && strlen (body) > 1) {
	delete_buf (body);
	return true;
    }
    if (body != NULL)
	delete_buf (body);
    return false;
}


char* 
MapiGPGMEImpl::getBody (bool isHtml)
{
    HRESULT hr;
    LPSPropValue lpspvFEID = NULL;
    char *body;
    int type = isHtml? PR_BODY_HTML: PR_BODY;

    hr = HrGetOneProp ((LPMAPIPROP) msg, type, &lpspvFEID);
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
MapiGPGMEImpl::freeKeyArray (void **key)
{
    gpgme_key_t *buf = (gpgme_key_t *)key;
    int i=0;

    if (buf == NULL)
	return;
    for (i = 0; buf[i] != NULL; i++) {
	gpgme_key_release (buf[i]);
	buf[i] = NULL;
    }
    xfree (buf);
}


int 
MapiGPGMEImpl::countRecipients (char **recipients)
{
  int i;
  
  for (i=0; recipients[i] != NULL; i++)
    ;
  return i;
}


char** 
MapiGPGMEImpl::getRecipients (bool isRootMsg)
{
    HRESULT hr;
    LPMAPITABLE lpRecipientTable = NULL;
    LPSRowSet lpRecipientRows = NULL;
    char **rset = NULL;
    size_t j=0;

    if (!isRootMsg)
	return NULL;
        
    static SizedSPropTagArray (1L, PropRecipientNum) = {1L, {PR_EMAIL_ADDRESS}};

    hr = msg->GetRecipientTable (0, &lpRecipientTable);
    if (SUCCEEDED (hr)) {
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
	    logDebug ( "rset %d: %s\r\n", j, rset[j]);
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
MapiGPGMEImpl::freeUnknownKeys (char **unknown, int n)
{    
    for (int i=0; i < n; i++) {
	if (unknown[i] != NULL) {
	    xfree (unknown[i]);
	    unknown[i] = NULL;
	}
    }
    if (n > 0)
	xfree (unknown);
}

void 
MapiGPGMEImpl::freeRecipients (char **recipients)
{
    for (int i=0; recipients[i] != NULL; i++) {
	delete_buf (recipients[i]);	
	recipients[i] = NULL;
    }
    delete recipients;
}


const char*
MapiGPGMEImpl::getPassphrase (const char *keyid)
{
    cache_item_t item = (cache_item_t)passCache->get(keyid);
    if (item != NULL)
	return item->pass;
    return NULL;
}


void
MapiGPGMEImpl::storePassphrase (void *itm)
{
    cache_item_t item = (cache_item_t)itm;
    cache_item_t old;
    old = (cache_item_t)passCache->get(item->keyid);
    if (old != NULL)
	cache_item_free (old);
    passCache->put (item->keyid+8, item);
    logDebug ( "put keyid %s = '%s'\r\n", item->keyid+8, "***");
}


char*
MapiGPGMEImpl::addHtmlLineEndings (char *newBody)
{
    char *p;
    char *snew = (char*)xcalloc (1, 2*strlen (newBody));

    p = strtok ((char *)newBody, "\n");
    while (p != NULL) {
	strcat (snew, p);
	strcat (snew, "\r\n");
	strcat (snew, "&nbsp;<br>");
	p = strtok (NULL, "\n");
    }
    
    return snew;
}

int 
MapiGPGMEImpl::encrypt (void)
{
    gpgme_key_t *keys=NULL, *keys2=NULL;
    bool isHtml = isHtmlMessage ();
    char *body = getBody (isHtml);
    char *newBody = NULL;
    char **recipients = getRecipients (true);
    char **unknown = NULL;
    int opts = 0;
    int err = 0;
    size_t all=0;

    if (body == NULL || strlen (body) == 0) {
	freeRecipients (recipients);
	if (body != NULL)
	    delete_buf (body);
	return 0;
    }

    logDebug ("encrypt\r\n");
    int n = op_lookup_keys (recipients, &keys, &unknown, &all);
    logDebug ("fnd %d need %d (%p)\r\n", n, all, unknown);
    if (n != countRecipients (recipients)) {
	logDebug ("recipient_dialog_box2\r\n");
	recipient_dialog_box2 (keys, unknown, all, &keys2, &opts);
	xfree (keys);
	keys = keys2;
	if (opts & OPT_FLAG_CANCEL) {
	    freeRecipients (recipients);
	    delete_buf (body);
	    return 0;
	}
    }

    err = op_encrypt ((void*)keys, body, &newBody);
    if (err)
	MessageBox (NULL, op_strerror (err), "GPG Encryption", MB_ICONERROR|MB_OK);
    else {
	if (isHtml) {
	    char *p = addHtmlLineEndings (newBody);
	    setBody (p, true);
	    xfree (p);
	}
	else
	    setBody (newBody, false);
    }
    delete_buf (body);
    xfree (newBody);
    freeRecipients (recipients);
    freeUnknownKeys (unknown, n);
    if (!err && hasAttachments ()) {
	logDebug ("encrypt attachments\r\n");
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
    /*logDebug ( "get keyid %s = '%s'\r\n", keyid+8, "***");*/
    if (passwd != NULL) {
	WriteFile ((HANDLE)fd, passwd, strlen (passwd), &nwritten, NULL);
	WriteFile ((HANDLE)fd, "\n", 1, &nwritten, NULL);
    }

    return 0;
}


int 
MapiGPGMEImpl::decrypt (void)
{
    outlgpg_type_t id;
    char *body = getBody (false);
    char *newBody = NULL;    
    bool hasAttach = hasAttachments ();
    bool isHtml = isHtmlMessage ();
    int err;

    id = getMessageType (body);
    if (id == GPG_TYPE_CLEARSIG) {
	delete_buf (body);
	return verify ();
    }
    else if (id == GPG_TYPE_NONE && !hasAttach) {
	MessageBox (NULL, "No valid OpenPGP data found.", "GPG Decryption", MB_ICONERROR|MB_OK);
	delete_buf (body);
	return 0;
    }

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
    if (err) {
	if (hasAttach && gpg_error ((gpg_err_code_t)err) == gpg_error (GPG_ERR_NO_DATA))
	    ;
	else
	    MessageBox (NULL, op_strerror (err), "GPG Decryption", MB_ICONERROR|MB_OK);
    }
    else if (newBody != NULL && *newBody) {	
	/* Also set PR_BODY but do not use 'SaveChanges' to make it permanently.
	   This way the user can reply with the plaintext but the ciphertext is
	   still stored. */
	logDebug ("decrypt isHtml=%d\r\n", isHtmlBody (newBody));
	setRTFBody (newBody);
	/* XXX: find a way to handle text/html message in a better way! */
	if (isHtmlBody (newBody)) {
	    const char *s = "The message text cannot be displayed.\n"
		            "You have to save the decrypted message to view it.\n"
			    "Then you need to re-open the message.\n\n"
			    "Do you want to save the decrypted message?";
	    int id = MessageBox (NULL, s, "GPG Decryption", MB_YESNO|MB_ICONWARNING);
	    if (id == IDYES) {
		logDebug ("decrypt: save plaintext message.\r\n");
		setBody (newBody, true);
		msg->SaveChanges (FORCE_SAVE);
	    }
	}
	else
	    setBody (newBody, false);
    }

    if (hasAttach) {
	logDebug ("decrypt attachments\r\n");
	decryptAttachments (parent);
    }
    delete_buf (body);
    xfree (newBody);
    return err;
}


int
MapiGPGMEImpl::sign (void)
{
    char *body = getBody (false);
    char *newBody = NULL;
    int hasAttach = hasAttachments ();
    int err = 0;

    if (body == NULL || strlen (body) == 0) {
	if (body != NULL)
	    delete_buf (body);
	return 0;
    }
    err = op_sign_start (body, &newBody);
    if (err)
	MessageBox (NULL, op_strerror (err), "GPG Sign", MB_ICONERROR|MB_OK);
    else
	setBody (newBody, isHtmlBody (newBody));

    if (hasAttach && autoSignAtt)
	signAttachments (parent);

    delete_buf (body);
    xfree (newBody);
    return err;
}


bool
MapiGPGMEImpl::isMessageEncrypted (void)
{
    char *body = getBody (false);
    bool enc = getMessageType (body) == GPG_TYPE_MSG;
    if (body != NULL)
	delete_buf (body);
    return enc;
}


outlgpg_type_t
MapiGPGMEImpl::getMessageType (const char *body)
{
    if (strstr (body, "BEGIN PGP MESSAGE"))
	return GPG_TYPE_MSG;
    if (strstr (body, "BEGIN PGP SIGNED MESSAGE"))
	return GPG_TYPE_CLEARSIG;
    if (strstr (body, "BEGIN PGP SIGNATURE"))
	return GPG_TYPE_SIG;
    if (strstr (body, "BEGIN PGP PUBLIC KEY"))
	return GPG_TYPE_PUBKEY;
    if (strstr (body, "BEGIN PGP PRIVATE KEY"))
	return GPG_TYPE_SECKEY;
    return GPG_TYPE_NONE;
}



int
MapiGPGMEImpl::doCmdFile(int action, const char *in, const char *out)
{
    logDebug ( "doCmdFile action=%d in=%s out=%s\r\n", action, in, out);
    if (ATT_SIGN (action) && ATT_ENCR (action))
	return !op_sign_encrypt_file (recipSet, in, out);
    if (ATT_SIGN (action) && !ATT_ENCR (action))
	return !op_sign_file (OP_SIG_NORMAL, in, out);
    if (!ATT_SIGN (action) && ATT_ENCR (action))
	return !op_encrypt_file (recipSet, in, out);
    return !op_decrypt_file (in, out);    
}


int
MapiGPGMEImpl::doCmdAttach (int action)
{
    logDebug ("doCmdAttach action=%d\n", action);
    if (ATT_SIGN (action) && ATT_ENCR (action))
	return signEncrypt ();
    if (ATT_SIGN (action) && !ATT_ENCR (action))
	return sign ();
    if (!ATT_SIGN (action) && ATT_ENCR (action))
	return encrypt ();
    return decrypt ();
}


int
MapiGPGMEImpl::doCmd (int doEncrypt, int doSign)
{
    logDebug ( "doCmd doEncrypt=%d doSign=%d\r\n", doEncrypt, doSign);
    if (doEncrypt && doSign)
	return signEncrypt ();
    if (doEncrypt && !doSign)
	return encrypt ();
    if (!doEncrypt && doSign)
	return sign ();
    return -1;
}


static const char *
userid_from_key (gpgme_key_t k)
{
  if (k && k->uids && k->uids->uid)
    return k->uids->uid;
  else
    return "?";
}

static const char *
keyid_from_key (gpgme_key_t k)
{
  
  if (k && k->subkeys && k->subkeys->keyid)
    return k->subkeys->keyid;
  else
    return "????????";
}


static void 
log_key_info (MapiGPGME *g, gpgme_key_t *keys, gpgme_key_t locusr)
{
    if (locusr)
	g->logDebug ("locusr:%s:%s\r\n", 
		     userid_from_key (locusr), keyid_from_key (locusr));
    else
	g->logDebug ( "locusr:null\r\n");
    gpgme_key_t n;
    int i;

    if (keys == NULL)
	return;
    i=0;
    for (n=keys[0]; keys[i] != NULL; i++)
	g->logDebug ("recp:%d:%s:%s\r\n", 
                     i, userid_from_key (keys[i]), keyid_from_key (keys[i]));
}
	    

int
MapiGPGMEImpl::signEncrypt (void)
{
    bool isHtml = isHtmlMessage ();
    char *body = getBody (isHtml);
    char *newBody = NULL;
    char **recipients = getRecipients (TRUE);
    char **unknown = NULL;
    gpgme_key_t locusr=NULL, *keys = NULL, *keys2 =NULL;

    if (body == NULL || strlen (body) == 0) {
	freeRecipients (recipients);
	if (body != NULL)
	    delete_buf (body);
	return 0;
    }
    if (signer_dialog_box (&locusr, NULL) == -1) {
	freeRecipients (recipients);
	delete_buf (body);
	return 0;
    }
    logDebug ("locusr keyid:%s\r\n", keyid_from_key (locusr));

    size_t all;
    int n = op_lookup_keys (recipients, &keys, &unknown, &all);
    if (n != countRecipients (recipients)) {
	recipient_dialog_box2 (keys, unknown, all, &keys2, NULL);
	xfree (keys);
	keys = keys2;
    }

    log_key_info (this, keys, locusr);
    int err = op_sign_encrypt ((void *)keys, (void*)locusr, body, &newBody);
    if (err)
	MessageBox (NULL, op_strerror (err), "GPG Sign Encrypt", MB_ICONERROR|MB_OK);
    else {
	if (isHtml) {
	    char *p = addHtmlLineEndings (newBody);
	    setBody (p, true);
	    xfree (p);
	}
	else
	    setBody (newBody, false);
    }

    delete_buf (body);
    xfree (newBody);
    freeUnknownKeys (unknown, n);
    if (!err && hasAttachments ()) {
	logDebug ( "encrypt attachments");
	recipSet = (void *)keys;
	encryptAttachments (parent);
    }
    freeKeyArray ((void **)keys);
    gpgme_key_release (locusr);
    freeRecipients (recipients);
    return err;
}


int 
MapiGPGMEImpl::verify (void)
{
    char *body = getBody (false);
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


void 
MapiGPGMEImpl::setDefaultKey (const char *key)
{
    if (defaultKey) {
	delete_buf (defaultKey);
	defaultKey = NULL;
    }
    defaultKey = new char[strlen (key)+1];
    fail_if_null (defaultKey);
    strcpy (defaultKey, key);
}


char* 
MapiGPGMEImpl::getDefaultKey (void)
{
  return defaultKey;
}


void 
MapiGPGMEImpl::setMessage (LPMESSAGE msg)
{
    this->msg = msg;
    logDebug ( "setMessage %p\r\n", msg);
}


void
MapiGPGMEImpl::setWindow(HWND hwnd)
{
    this->parent = hwnd;
}


/* We need this to find the mailer window because we directly change the text
   of the window instead of the MAPI object itself. */
HWND
MapiGPGMEImpl::findMessageWindow (HWND parent)
{
    HWND child;

    if (parent == NULL)
	return NULL;

    child = GetWindow (parent, GW_CHILD);
    while (child != NULL) {
	char buf[1024+1];
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
    /*logDebug ("no message window found.\r\n");*/
    return NULL;
}


int
MapiGPGMEImpl::streamFromFile (const char *file, LPATTACH att)
{
    HRESULT hr;
    LPSTREAM to = NULL, from = NULL;
    STATSTG statInfo;

    hr = att->OpenProperty (PR_ATTACH_DATA_BIN, &IID_IStream, 0,
	 		      MAPI_CREATE | MAPI_MODIFY, (LPUNKNOWN*) &to);
    if (FAILED (hr))
	return FALSE;

    hr = OpenStreamOnFile (MAPIAllocateBuffer, MAPIFreeBuffer,
			   STGM_READ, (char*)file, NULL, &from);
    if (!SUCCEEDED (hr)) {
	to->Release ();
	logDebug ( "streamFromFile %s failed.\r\n", file);
	return FALSE;
    }
    from->Stat (&statInfo, STATFLAG_NONAME);
    from->CopyTo (to, statInfo.cbSize, NULL, NULL);
    to->Commit (0);
    to->Release ();
    from->Release ();
    logDebug ( "streamFromFile %s succeeded\r\n", file);
    return TRUE;
}


int
MapiGPGMEImpl::streamOnFile (const char *file, LPATTACH att)
{
    HRESULT hr;
    LPSTREAM from = NULL, to = NULL;
    STATSTG statInfo;

    hr = att->OpenProperty (PR_ATTACH_DATA_BIN, &IID_IStream, 
			    0, 0, (LPUNKNOWN*) &from);
    if (FAILED (hr))
	return FALSE;

    hr = OpenStreamOnFile (MAPIAllocateBuffer, MAPIFreeBuffer,
   		           STGM_CREATE | STGM_READWRITE, (char*) file,
			   NULL, &to);
    if (!SUCCEEDED (hr)) {
	from->Release ();
	logDebug ( "streamOnFile %s failed with %s\r\n", file, 
		    hr == MAPI_E_NO_ACCESS? 
		    "no access" : hr == MAPI_E_NOT_FOUND? "not found" : "unknown");
	return FALSE;
    }
    from->Stat (&statInfo, STATFLAG_NONAME);
    from->CopyTo (to, statInfo.cbSize, NULL, NULL);
    to->Commit (0);
    to->Release ();
    from->Release ();
    logDebug ( "streamOnFile %s succeeded\r\n", file);
    return TRUE;
}


int
MapiGPGMEImpl::getMessageFlags (void)
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
MapiGPGMEImpl::getMessageHasAttachments (void)
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
MapiGPGMEImpl::setMessageAccess (int access)
{
    HRESULT hr;
    SPropValue prop;
    prop.ulPropTag = PR_ACCESS;
    prop.Value.l = access;
    hr = HrSetOneProp (msg, &prop);
    return FAILED (hr)? false: true;
}


bool
MapiGPGMEImpl::setAttachMethod (LPATTACH obj, int mode)
{
    SPropValue prop;
    HRESULT hr;
    prop.ulPropTag = PR_ATTACH_METHOD;
    prop.Value.ul = mode;
    hr = HrSetOneProp (obj, &prop);
    return FAILED (hr)? true : false;
}


int
MapiGPGMEImpl::getAttachMethod (LPATTACH obj)
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
MapiGPGMEImpl::setAttachFilename (LPATTACH obj, const char *name, bool islong)
{
    HRESULT hr;
    SPropValue prop;
    prop.ulPropTag = PR_ATTACH_LONG_FILENAME;

    if (!islong)
	prop.ulPropTag = PR_ATTACH_FILENAME;
    prop.Value.lpszA = (char*) name;   
    hr = HrSetOneProp (obj, &prop);
    return FAILED (hr)? false: true;
}


char*
MapiGPGMEImpl::getAttachPathname (LPATTACH obj)
{
    LPSPropValue propval;
    HRESULT hr;
    char *path;

    hr = HrGetOneProp ((LPMAPIPROP)obj, PR_ATTACH_LONG_PATHNAME, &propval);
    if (FAILED (hr)) {
	hr = HrGetOneProp ((LPMAPIPROP)obj, PR_ATTACH_PATHNAME, &propval);
	if (SUCCEEDED (hr)) {
	    path = xstrdup (propval[0].Value.lpszA);
	    MAPIFreeBuffer (propval);
	}
	else
	    return NULL;
    }
    else {
	path = xstrdup (propval[0].Value.lpszA);
	MAPIFreeBuffer (propval);
    }
    return path;
}


char*
MapiGPGMEImpl::getAttachFilename (LPATTACH obj)
{
    LPSPropValue propval;
    HRESULT hr;
    char *name;

    hr = HrGetOneProp ((LPMAPIPROP)obj, PR_ATTACH_LONG_FILENAME, &propval);
    if (FAILED(hr)) {
	hr = HrGetOneProp ((LPMAPIPROP)obj, PR_ATTACH_FILENAME, &propval);
	if (SUCCEEDED (hr)) {
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


bool
MapiGPGMEImpl::checkAttachmentExtension (const char *ext)
{
    if (ext == NULL)
	return false;
    if (*ext == '.')
	ext++;
    logDebug ( "checkAttachmentExtension: %s\r\n", ext);
    if (stricmp (ext, "gpg") == 0 ||
	stricmp (ext, "pgp") == 0 ||
	stricmp (ext, "asc") == 0)
	return true;
    return false;
}


const char*
MapiGPGMEImpl::getAttachmentExtension (const char *fname)
{
    static char ext[4];
    char *p;

    p = strrchr (fname, '.');
    if (p != NULL) {
	/* XXX: what if the extension is < 3 chars */
	strncpy (ext, p, 4);
	if (checkAttachmentExtension (ext))
	    return ext;
    }
    return EXT_MSG;
}


const char*
MapiGPGMEImpl::getPGPExtension (int action)
{
    if (ATT_SIGN (action))
	return EXT_SIG;
    return EXT_MSG;
}


bool 
MapiGPGMEImpl::setXHeader (const char *name, const char *val)
{  
#ifndef __MINGW32__
    USES_CONVERSION;
#endif
    LPMDB lpMdb = NULL;
    HRESULT hr = 0;
    LPSPropTagArray pProps = NULL;
    SPropValue pv;
    MAPINAMEID mnid[1];	
    // {00020386-0000-0000-C000-000000000046}  ->  GUID For X-Headers	
    GUID guid = {0x00020386, 0x0000, 0x0000, {0xC0, 0x00, 0x00, 0x00,
		 0x00, 0x00, 0x00, 0x46} };

    memset (&mnid[0], 0, sizeof (MAPINAMEID));
    mnid[0].lpguid = &guid;
    mnid[0].ulKind = MNID_STRING;
//     mnid[0].Kind.lpwstrName = A2W (name);
    hr = msg->GetIDsFromNames (1, (LPMAPINAMEID*)mnid, MAPI_CREATE, &pProps);
    if (FAILED (hr)) {
	logDebug ("set X-Header failed.\r\n");
	return false;
    }
    
    pv.ulPropTag = (pProps->aulPropTag[0] & 0xFFFF0000) | PT_STRING8;
    pv.Value.lpszA = (char *)val;
    hr = HrSetOneProp(msg, &pv);	
    if (!SUCCEEDED (hr)) {
	logDebug ("set X-Header failed.\r\n");
	return false;
    }

    logDebug ("set X-Header succeeded.\r\n");
    return true;
}


char*
MapiGPGMEImpl::getXHeader (const char *name)
{
    /* XXX: PR_TRANSPORT_HEADERS is not available in my MSDN. */
    return NULL;
}


void
MapiGPGMEImpl::freeAttachments (void)
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
MapiGPGMEImpl::getAttachments (void)
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
MapiGPGMEImpl::openAttachment (int pos)
{
    HRESULT hr;
    LPATTACH att = NULL;
    
    hr = msg->OpenAttach (pos, NULL, MAPI_BEST_ACCESS, &att);	
    if (SUCCEEDED (hr))
	return att;
    return NULL;
}


void
MapiGPGMEImpl::releaseAttachment (LPATTACH att)
{
    att->Release ();
}



char*
MapiGPGMEImpl::generateTempname (const char *name)
{
    char temp[MAX_PATH+2];
    char *p;

    GetTempPath (sizeof (temp)-1, temp);
    if (temp[strlen (temp)-1] != '\\')
	strcat (temp, "\\");
    p = (char *)xcalloc (1, strlen (temp) + strlen (name) + 16);
    sprintf (p, "%s%s", temp, name);
    return p;
}


bool
MapiGPGMEImpl::signAttachment (const char *datfile)
{
    char *sigfile;
    LPATTACH newatt;
    int pos=0, err=0;

    sigfile = (char *)xcalloc (1,strlen (datfile)+5);
    strcpy (sigfile, datfile);
    strcat (sigfile, ".asc");

    newatt = createAttachment (pos);
    setAttachMethod (newatt, ATTACH_BY_VALUE);
    setAttachFilename (newatt, sigfile, false);

    if (nstorePasswd == 0)
	err = op_sign_file (OP_SIG_DETACH, datfile, sigfile);
    else if (passCache->size () == 0) {
	cache_item_t itm=NULL;
	err = op_sign_file_ext (OP_SIG_DETACH, datfile, sigfile, &itm);
	if (!err)
	    storePassphrase (itm);
    }
    else
	err = op_sign_file_next (passphraseCallback, this, OP_SIG_DETACH, datfile, sigfile);

    if (streamFromFile (sigfile, newatt)) {
	logDebug ("signAttachment: commit changes.\r\n");
	newatt->SaveChanges (FORCE_SAVE);
    }
    releaseAttachment (newatt);
    xfree (sigfile);

    return (!err)? true : false;
}

/* XXX: find a way to see if the attachment is already secured. This could be
        done by watching at the extension or checking the first lines. */
int
MapiGPGMEImpl::processAttachment (LPATTACH *attm, HWND hwnd,
                                  int pos, int action)
{    
    LPATTACH att = *attm;
    int method = getAttachMethod (att);
    BOOL success = TRUE;
    HRESULT hr;

    /* XXX: sign-only code is still not very intuitive. */

    if (action == GPG_ATTACH_NONE)
	return FALSE;
    if (action == GPG_ATTACH_DECRYPT && !saveDecryptedAtt)
	return TRUE;

    switch (method) {
    case ATTACH_EMBEDDED_MSG:
	LPMESSAGE emb;

	/* we do not support to sign these kind of attachments. */
	if (action == GPG_ATTACH_SIGN)
	    return TRUE;
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
	char *inname;
	char *outname;
	char *tmp;

	tmp = getAttachFilename (att);
	inname =  generateTempname (tmp);
	logDebug ("enc inname: '%s'\r\n", inname);
	if (action != GPG_ATTACH_DECRYPT) {
	    char *tmp2 = (char *)xcalloc (1, strlen (inname) 
					     + strlen (ATT_PREFIX) + 4 + 1);
	    sprintf (tmp2, "%s"ATT_PREFIX".%s", tmp, getPGPExtension (action));
	    outname = generateTempname (tmp2);
	    xfree (tmp2);
	    logDebug ( "enc outname: '%s'\r\n", outname);
	}
	else {
	    if (checkAttachmentExtension (strrchr (tmp, '.')) == false) {
		logDebug ( "%s: no pgp extension found.\r\n", tmp);
		xfree (tmp);
		xfree (inname);
		return TRUE;
	    }
	    char *tmp2 = (char*)xcalloc (1, strlen (tmp) + 4);
	    strcpy (tmp2, tmp);
	    tmp2[strlen (tmp2) - 4] = '\0';
	    outname = generateTempname (tmp2);
	    xfree (tmp2);
	    logDebug ("dec outname: '%s'\r\n", outname);
	}
	success = FALSE;
	/* if we are in sign-only mode, just create a detached signature
	   for each attachment but do not alter the attachment data itself. */
	if (action != GPG_ATTACH_SIGN && streamOnFile (inname, att)) {
	    if (doCmdFile (action, inname, outname))
		success = TRUE;
	    else
		logDebug ( "doCmdFile failed\r\n");
	}
	if ((action == GPG_ATTACH_ENCRYPT || action == GPG_ATTACH_SIGN) 
	    && autoSignAtt)
	    signAttachment (inname);

	/*DeleteFile (inname);*/
	/* XXX: the file does not seemed to be closed. */
	xfree (inname);
	xfree (tmp);
	
	if (action != GPG_ATTACH_SIGN)
	    deleteAttachment (pos);

	if (action == GPG_ATTACH_ENCRYPT) {
	    LPATTACH newatt;
	    *attm = newatt = createAttachment (pos);
	    setAttachMethod (newatt, ATTACH_BY_VALUE);
	    setAttachFilename (newatt, outname, false);

	    if (streamFromFile (outname, newatt)) {
		logDebug ( "commit changes.\r\n");	    
		newatt->SaveChanges (FORCE_SAVE);
	    }
	}
	else if (success && action == GPG_ATTACH_DECRYPT) {
	    success = saveDecryptedAttachment (NULL, outname);
	    logDebug ("saveDecryptedAttachment ec=%d\r\n", success);
	}
	DeleteFile (outname);
	xfree (outname);
	releaseAttachment (att);
	break;

    case ATTACH_BY_REF_ONLY:
	break;

    case ATTACH_OLE:
	break;

    }

    return success;
}


int 
MapiGPGMEImpl::decryptAttachments (HWND hwnd)
{
    int n;

    if (!getAttachments ())
	return FALSE;
    n = countAttachments ();
    logDebug ( "dec: mail has %d attachments\r\n", n);
    if (!n) {
	freeAttachments ();
	return TRUE;
    }
    for (int i=0; i < n; i++) {
	LPATTACH amsg = openAttachment (i);
	if (!amsg)
	    continue;
	processAttachment (&amsg, hwnd, i, GPG_ATTACH_DECRYPT);
    }
    freeAttachments ();
    return 0;
}


int
MapiGPGMEImpl::signAttachments (HWND hwnd)
{
    if (!getAttachments ())
	return FALSE;
    int n = countAttachments ();
    logDebug ("sig: mail has %d attachments\r\n", n);
    if (!n) {
	freeAttachments ();
	return TRUE;
    }
    for (int i=0; i < n; i++) {
	LPATTACH amsg = openAttachment (i);
	if (!amsg)
	    continue;
	processAttachment (&amsg, hwnd, i, GPG_ATTACH_SIGN);
	releaseAttachment (amsg);
    }
    freeAttachments ();
    return 0;
}


int
MapiGPGMEImpl::encryptAttachments (HWND hwnd)
{    
    int n;

    if (!getAttachments ())
	return FALSE;
    n = countAttachments ();
    logDebug ("enc: mail has %d attachments\r\n", n);
    if (!n) {
	freeAttachments ();
	return TRUE;
    }
    for (int i=0; i < n; i++) {
	LPATTACH amsg = openAttachment (i);
	if (amsg == NULL)
	    continue;
	processAttachment (&amsg, hwnd, i, GPG_ATTACH_ENCRYPT);
	releaseAttachment (amsg);	
    }
    freeAttachments ();
    return 0;
}


bool 
MapiGPGMEImpl::saveDecryptedAttachment (HWND root, const char *srcname)
				     
{
    char filter[] = "All Files (*.*)|*.*||";
    char fname[MAX_PATH+1];
    char *p;
    OPENFILENAME ofn;

    for (size_t i=0; i< strlen (filter); i++)  {
	if (filter[i] == '|')
	    filter[i] = '\0';
    }

    memset (fname, 0, sizeof (fname));
    p = strstr (srcname, ATT_PREFIX);
    if (!p)
	strncpy (fname, srcname, MAX_PATH);
    else {
	strncpy (fname, srcname, (p-srcname));
	strcat (fname, srcname+(p-srcname)+strlen (ATT_PREFIX));	
    }

    memset (&ofn, 0, sizeof (ofn));
    ofn.lStructSize = sizeof (ofn);
    ofn.hwndOwner = root;
    ofn.lpstrFile = fname;
    ofn.nMaxFile = MAX_PATH;
    ofn.lpstrFileTitle = NULL;
    ofn.nMaxFileTitle = 0;
    ofn.Flags |= OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT;
    ofn.lpstrTitle = "GPG - Save decrypted attachments";
    ofn.lpstrFilter = filter;

    if (GetSaveFileName (&ofn)) {
	logDebug ("copy %s -> %s\r\n", srcname, fname);
	return CopyFile (srcname, fname, FALSE) == 0? false : true;
    }
    return true;
}


int
MapiGPGMEImpl::startKeyManager (void)
{
    return start_key_manager ();
}


void
MapiGPGMEImpl::startConfigDialog (HWND parent)
{
    config_dialog_box (parent);
}


int
MapiGPGMEImpl::readOptions (void)
{
    char *val=NULL;

    load_extension_value ("autoSignAttachments", &val);
    autoSignAtt = val == NULL || *val != '1' ? 0 : 1;
    xfree (val); val =NULL;

    load_extension_value ("saveDecryptedAttachments", &val);
    saveDecryptedAtt = val == NULL || *val != '1'? 0 : 1;
    xfree (val); val =NULL;

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
    if (val == NULL ||*val == '"' || *val == 0)
	logfile = NULL;
    else {
	setLogFile (val);
	setEnableLogging (true);
    }
    xfree (val); val=NULL;

    load_extension_value ("defaultKey", &val);
    if (val == NULL || *val == '"') {
	encryptDefault = 0;
	defaultKey = NULL;
    }
    else {
	setDefaultKey (val);
	encryptDefault = 1;
    }

    xfree (val); val=NULL;

    return 0;
}


void
MapiGPGMEImpl::displayError (HWND root, const char *title)
{	
    char buf[256];
    
    FormatMessage (FORMAT_MESSAGE_FROM_SYSTEM, NULL, GetLastError (), 
		   MAKELANGID (LANG_NEUTRAL, SUBLANG_DEFAULT), 
		   buf, sizeof (buf)-1, NULL);
    MessageBox (root, buf, title, MB_OK|MB_ICONERROR);
}


int
MapiGPGMEImpl::writeOptions (void)
{
    struct conf {
	const char *name;
	bool value;
    };
    struct conf opt[] = {
	{"encryptDefault", doEncrypt},
	{"signDefault", doSign},
	{"addDefaultKey", encryptDefault},
	{"saveDecryptedAttachments", saveDecryptedAtt},
	{"autoSignAttachments", autoSignAtt},
	{NULL, 0}
    };
    char buf[32];

    for (int i=0; opt[i].name != NULL; i++) {
	int rc = store_extension_value (opt[i].name, opt[i].value? "1": "0");
	if (rc)
	    displayError (NULL, "Save options in the registry");
	/* XXX: also show the name of the value */
    }

    if (logfile != NULL)
	store_extension_value ("logFile", logfile);
    if (defaultKey != NULL)
	store_extension_value ("defaultKey", defaultKey);
    
    sprintf (buf, "%d", nstorePasswd);
    store_extension_value ("storePasswdTime", buf);
    
    sprintf (buf, "%d", encFormat);
    store_extension_value ("encodingFormat", buf);

    return 0;
}


int 
MapiGPGMEImpl::attachPublicKey (const char *keyid)
{
    /* @untested@ */
    const char *patt[1];
    char *keyfile;
    int err, pos = 0;
    LPATTACH newatt;

    keyfile = generateTempname (keyid);
    patt[0] = xstrdup (keyid);
    err = op_export_keys (patt, keyfile);

    newatt = createAttachment (pos);
    setAttachMethod (newatt, ATTACH_BY_VALUE);
    setAttachFilename (newatt, keyfile, false);
    /* XXX: set proper RFC3156 MIME types. */

    if (streamFromFile (keyfile, newatt)) {
	logDebug ("attachPublicKey: commit changes.\r\n");
	newatt->SaveChanges (FORCE_SAVE);
    }
    releaseAttachment (newatt);
    xfree (keyfile);
    xfree ((void *)patt[0]);
    return err;
}
