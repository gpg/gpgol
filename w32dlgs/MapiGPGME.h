/* MapiGPGME.h - Mapi support with GPGME
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
#ifndef MAPI_GPGME_H
#define MAPI_GPGME_H

#define DLL_EXPORT __declspec(dllexport)

typedef enum {
    GPG_TYPE_NONE = 0,
    GPG_TYPE_MSG,
    GPG_TYPE_SIG,
    GPG_TYPE_CLEARSIG,
    GPG_TYPE_PUBKEY,	/* the key types are recognized but nut supported */
    GPG_TYPE_SECKEY     /* by this implementaiton. */
} outlgpg_type_t;

typedef enum {
    GPG_ATTACH_NONE = 0,
    GPG_ATTACH_DECRYPT = 1,
    GPG_ATTACH_ENCRYPT = 2,
    GPG_ATTACH_SIGN = 4,
    GPG_ATTACH_SIGNENCRYPT = GPG_ATTACH_SIGN|GPG_ATTACH_ENCRYPT,
} outlgpg_attachment_action_t;

typedef enum {
    GPG_FMT_NONE = 0,	    /* do not encrypt attachments */
    GPG_FMT_CLASSIC = 1,    /* encrypt attachments without any encoding */
    GPG_FMT_PGP_PEF = 2	    /* use the PGP partioned encoding format (PEF) */
} outlgpg_format_t;
#define DEFAULT_ATTACHMENT_FORMAT GPG_FMT_CLASSIC

class MapiGPGME
{
private:
    HWND	parent;
    char	*defaultKey;
    HashTable	*passCache;
    LPMESSAGE	msg;
    LPMAPITABLE attachTable;
    LPSRowSet	attachRows;
    void	*recipSet;
    bool        enableLogging;

    /* Options */
    char    *logfile;
    int	    nstorePasswd;	/* time in seconds the passphrase is stored. */
    bool    doEncrypt;
    bool    doSign;
    bool    encryptDefault;
    bool    saveDecryptedAtt;	/* save decrypted attachments */
    bool    autoSignAtt;	/* sign all outgoing attachments */
    int	    encFormat;		/* encryption format for attachments. */

public:    
    DLL_EXPORT MapiGPGME ();
    DLL_EXPORT MapiGPGME (LPMESSAGE msg);
    DLL_EXPORT ~MapiGPGME ();

private:
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

public:
    DLL_EXPORT int encrypt (void);
    DLL_EXPORT int decrypt (void);
    DLL_EXPORT int sign (void);
    DLL_EXPORT int verify (void);
    DLL_EXPORT int signEncrypt (void);

    DLL_EXPORT int doCmd(int doEncrypt, int doSign);
    DLL_EXPORT int doCmdAttach(int action);
    DLL_EXPORT int doCmdFile(int action, const char *in, const char *out);

    DLL_EXPORT const char* getLogFile (void) { return logfile; }
    DLL_EXPORT void setLogFile (const char *logfile)
    { 
	if (this->logfile) {
	    delete []this->logfile;
	    this->logfile = NULL;
	}
	this->logfile = new char [strlen (logfile)+1];
	if (this->logfile != NULL)
	    strcpy (this->logfile, logfile);
    }

    DLL_EXPORT int  getStorePasswdTime (void) { return nstorePasswd; }
    DLL_EXPORT void setStorePasswdTime (int nCacheTime) { this->nstorePasswd = nCacheTime; }
    DLL_EXPORT bool getEncryptDefault (void) { return doEncrypt; }
    DLL_EXPORT void setEncryptDefault (bool doEncrypt) { this->doEncrypt = doEncrypt; }
    DLL_EXPORT bool getSignDefault (void) { return doSign; }
    DLL_EXPORT void setSignDefault (bool doSign) { this->doSign = doSign; }
    DLL_EXPORT bool getEncryptWithDefaultKey (void) { return encryptDefault; }
    DLL_EXPORT void setEncryptWithDefaultKey (bool encryptDefault) { this->encryptDefault = encryptDefault; }
    DLL_EXPORT bool getSaveDecryptedAttachments (void) { return saveDecryptedAtt; }
    DLL_EXPORT void setSaveDecryptedAttachments (bool saveDecrAtt) { this->saveDecryptedAtt = saveDecrAtt; }
    DLL_EXPORT void setEncodingFormat (int fmt) { encFormat = fmt; }
    DLL_EXPORT int  getEncodingFormat (void) { return encFormat; }
    DLL_EXPORT void setSignAttachments (bool signAtt) { this->autoSignAtt = signAtt; }
    DLL_EXPORT bool getSignAttachments (void) { return autoSignAtt; }
    DLL_EXPORT void setEnableLogging (bool val) { this->enableLogging = val; }
    DLL_EXPORT bool getEnableLogging (void) { return this->enableLogging; }

    DLL_EXPORT int readOptions (void);
    DLL_EXPORT int writeOptions (void);

public:
    const char* getAttachmentExtension (const char *fname);
    DLL_EXPORT void freeAttachments (void);
    DLL_EXPORT int getAttachments (void);
    DLL_EXPORT int countAttachments (void) 
    { 
	if (attachRows == NULL)
	    return -1;
	return (int) attachRows->cRows; 
    }

    DLL_EXPORT bool hasAttachments (void)
    {
	if (attachRows == NULL)
	    getAttachments ();
	bool has = attachRows->cRows > 0? true : false;
	freeAttachments ();
	return has;
    }

    DLL_EXPORT bool deleteAttachment (int pos)
    {
	if (msg->DeleteAttach (pos, 0, NULL, 0) == S_OK)
	    return true;
	return false;
    }

    DLL_EXPORT LPATTACH createAttachment (int &pos)
    {
	ULONG attnum;	
	LPATTACH newatt = NULL;

	if (msg->CreateAttach (NULL, 0, &attnum, &newatt) == S_OK) {
	    pos = attnum;
	    return newatt;
	}
	return NULL;
    }

    DLL_EXPORT int startKeyManager ();
    DLL_EXPORT void startConfigDialog (HWND parent);

private:
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

public:
    DLL_EXPORT int attachPublicKey (const char *keyid);

    DLL_EXPORT void setDefaultKey (const char *key);
    DLL_EXPORT char* getDefaultKey (void);

    DLL_EXPORT void setMessage (LPMESSAGE msg);
    DLL_EXPORT void setWindow (HWND hwnd);

    const char* getPassphrase (const char *keyid);
    void storePassphrase (void *itm);
    outlgpg_type_t getMessageType (const char *body);

    void logDebug (const char *fmt, ...);
    
    DLL_EXPORT void clearPassphrase (void) 
    {
	if (passCache != NULL)
	    passCache->clear ();
    }
};

#endif /*MAPI_GPGME_H*/
