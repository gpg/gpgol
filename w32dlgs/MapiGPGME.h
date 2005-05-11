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
    GPG_TYPE_PUBKEY,
    GPG_TYPE_SECKEY
} gpg_type_t;

class MapiGPGME
{
private:
    HWND parent;
    LPMESSAGE msg;
    char* defaultKey;
    char* passPhrase;

    /* Options */
    char *logfile;
    int nstorePasswd;
    bool doEncrypt;
    bool doSign;
    bool encryptDefault;
    bool saveDecrAttr;

public:
    DLL_EXPORT MapiGPGME ();
    DLL_EXPORT MapiGPGME (LPMESSAGE msg);
    DLL_EXPORT ~MapiGPGME ();

private:
    HWND findMessageWindow (HWND parent);
    void rtfSync (char *body);
    int setBody (char *body);
    int setRTFBody (char *body);
    char *getBody ();

private:
    int countRecipients (char **recipients);
    char **getRecipients (bool isRootMsg);
    void freeRecipients (char **recipients);
    void freeUnknownKeys (char **unknown, int n);
    void freeKeyArray (void **key);

public:
    DLL_EXPORT int encrypt ();
    DLL_EXPORT int decrypt ();
    DLL_EXPORT int sign ();
    DLL_EXPORT int verify ();
    DLL_EXPORT int signEncrypt ();
    DLL_EXPORT int doCmd(int doEncrypt, int doSign);

public:
    DLL_EXPORT const char* getLogFile (void) { return logfile; }
    DLL_EXPORT void setLogFile (const char *logfile) 
    { 
	if (this->logfile)
	    delete []this->logfile;
	this->logfile = new char [strlen (logfile)+1];
	strcpy (this->logfile, logfile);
    }

    DLL_EXPORT int getStorePasswdTime () { return nstorePasswd; }
    DLL_EXPORT void setStorePasswdTime (int nCacheTime) { this->nstorePasswd = nCacheTime; }
    DLL_EXPORT bool getEncryptDefault (void) { return doEncrypt; }
    DLL_EXPORT void setEncryptDefault (bool doEncrypt) { this->doEncrypt = doEncrypt; }
    DLL_EXPORT bool getSignDefault (void) { return doSign; }
    DLL_EXPORT void setSignDefault (bool doSign) { this->doSign = doSign; }
    DLL_EXPORT bool getEncryptWithDefaultKey (void) { return encryptDefault; }
    DLL_EXPORT void setEncryptWithDefaultKey (bool encryptDefault) { this->encryptDefault = encryptDefault; }
    DLL_EXPORT bool getSaveDecryptedAttachments (void) { return saveDecrAttr; }
    DLL_EXPORT void setSaveDecryptedAttachments (bool saveDecrAttr) { this->saveDecrAttr = saveDecrAttr; }

    DLL_EXPORT int readOptions ();
    DLL_EXPORT int writeOptions ();

public:
    const char* getAttachmentExtension (const char *fname);
    DLL_EXPORT int decryptAttachments ();
    DLL_EXPORT int encryptAttachments ();

public:
    DLL_EXPORT int startKeyManager ();
    DLL_EXPORT void startConfigDialog (HWND parent);

private:
    bool setAttachMethod (LPATTACH obj, int mode);
    int getAttachMethod (LPATTACH obj);
    char* getAttachFilename (LPATTACH obj);
    bool setAttachFilename (LPATTACH obj, const char *name, bool islong);
    int getMessageFlags ();
    int getMessageHasAttachments ();
    bool setMessageAccess (int access);    

private:    
    int streamOnFile (const char *file, LPSTREAM to);
    int processAttachments (HWND hwnd, int action, const char **pFileNameVector);

public:
    DLL_EXPORT void setDefaultKey (const char *key);
    DLL_EXPORT char* getDefaultKey ();

public:
    DLL_EXPORT void setMessage (LPMESSAGE msg);
    DLL_EXPORT void setWindow (HWND hwnd);

public:
    gpg_type_t getMessageType (const char *body);
    const char *getPassphrase () { return passPhrase; }
    DLL_EXPORT void storePassphrase (const char *passPhrase)
    {
	if (this->passPhrase)
	    delete []this->passPhrase;
	this->passPhrase = new char[strlen (passPhrase)+1];
	strcpy (this->passPhrase, passPhrase);
    }
    DLL_EXPORT void clearPassphrase () 
    { 
	if (this->passPhrase) {
	    delete []this->passPhrase;
	    passPhrase = NULL;
	}
    }
};

#endif /*MAPI_GPGME_H*/