    /* GPG.h - gpg functions for mapi messages
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

#ifndef INC_GPG_H
#define INC_GPG_H

#include "..\GDGPG\Wrapper\GDGPGWrapper.h"

#define PGP_MIME_NONE 0
#define PGP_MIME_ENCR 1
#define PGP_MIME_SIG  2
#define PGP_MIME_KEY  4
#define PGP_MIME_IN   8

/* The CGPG class implements the gpg functions for MAPI messages. */
class CGPG  
{
public:
    CGPG();
    ~CGPG ();

protected:
    // options
    int  m_nStorePassPhraseTime;      // The time to store the passpharse (in seconds).
    BOOL m_bEncryptDefault;           // Indicates whether to encrypt messages on default.
    BOOL m_bSignDefault;              // Indicates whether to sign messages on default.
    BOOL m_bEncryptWithStandardKey;   // Indicates whether to encrypt the messages with the standard key too.
    BOOL m_bSaveDecrypted;            // Indicates whether to save decrypted messages.
    CGDGPGWrapper m_gdgpg;            // Wraps the calls to the GDGPG com object.
    BOOL m_bInit;                     // Indicates whether this object was initialized.
    int  m_nDecryptedAttachments;     // The number of decrypted attachments (helper variable for decrypting messages).
    BOOL m_bCancelSavingDecryptedAttachments; // Indicates whether the saving decrypted attachments was canceled (helper variable for decrypting messages).
    BOOL m_bContainsEmbeddedOLE;      // Indicates whether there was an embedded ole object.
    BSTR m_gpgStderr;
    BSTR m_gpgInfo;

    char * m_LogFile;
    char * cont_type;
    char * cont_trans_enc;

public:    
	
    BOOL Init();              // Initialize this object.
    void UnInit();            // Uninitialize this object.

    void ReadGPGOptions();    // Reads the plugin options from the registry.
    void WriteGPGOptions();   // Writes the plugin options to the registry.
    void EditExtendedOptions(HWND hWnd);  // Shows the dialog to edit the extended options.

    // Decrypts the specified message.
    BOOL DecryptMessage (HWND hWnd, LPMESSAGE lpMessage, BOOL bIsRootMessage); 
	
    // Encrypts and signs the specified message.
    BOOL EncryptAndSignMessage (HWND hWnd, LPMESSAGE lpMessage, BOOL bEncrypt, BOOL bSign, BOOL bIsRootMessage);

	
    // Imports all keys from the specified message.
    BOOL ImportKeys (HWND hWnd, LPMESSAGE lpMessage);
	
    void OpenKeyManager ();            // Opens the key manager.
    void InvalidateKeyLists ();        // Invalidates the key lists (e.g. when the keys was changed by the key manager).
    BOOL AddStandardKey (HWND hWnd);   // Adds the standard key to the open message.

	
    // get parameters
    const char * GetLogFile (void) { return m_LogFile; }
    int GetStorePassPhraseTime() { return m_nStorePassPhraseTime; };
    BOOL GetEncryptDefault() { return m_bEncryptDefault; };
    BOOL GetSignDefault() { return m_bSignDefault; };
    BOOL GetEncryptWithStandardKey() { return m_bEncryptWithStandardKey; };
    BOOL GetSaveDecrypted() { return m_bSaveDecrypted; };

    BSTR GetGPGOutput (void);
    BSTR GetGPGInfo (BSTR strFilename);
	
    // set parameters
    void SetLogFile (const char * strLogFilename);
    void SetStorePassPhraseTime(int nTime) { m_nStorePassPhraseTime = nTime; };
    void SetEncryptDefault(BOOL b) { m_bEncryptDefault = b; };
    void SetSignDefault(BOOL b) { m_bSignDefault = b; };
    void SetEncryptWithStandardKey(BOOL b) { m_bEncryptWithStandardKey = b; };
    void SetSaveDecrypted(BOOL b) { m_bSaveDecrypted = b; };

    // Internal function for decryption 
    BOOL DecryptFile (HWND hWndParent, BSTR strFilenameSource, BSTR strFilenameDest, int &pvReturn);
    BOOL VerifyDetachedSignature (HWND hWndParent, BSTR strFilenameText, BSTR strFilenameSig, int &pvReturn);

    BOOL CheckPGPMime (HWND hWnd, LPMESSAGE pMessage, int &mType);
    BOOL ProcessPGPMime (HWND hWnd, LPMESSAGE pMessage, int mType);
    BOOL ConvertOldPGP (HWND hWnd, LPMESSAGE pMessage);

public:
    // Decrypts all attachments.
    BOOL DecryptAttachments(HWND hWnd, LPMESSAGE pMessage);

protected:
    // Saves all attachments.
    BOOL SaveAttachments(HWND hWnd, LPMESSAGE pMessage, string sPrefix, vector<string>* pFileNameVector);
    // Encrypts and signs all attachments.
    BOOL EncryptAndSignAttachments(HWND hWnd, LPMESSAGE pMessage);
    // Finds the window which contains the crypted message.
    HWND FindMessageWindow(HWND hWnd);
    // Executes the specified action on all attachments.
    BOOL ProcessAttachments(HWND hWnd, int nAction, LPMESSAGE pMessage, string sPrefix, vector<string>* pFileNameVector);
    // Calls the "save as" dialog to saves an decrypted attachment.
    BOOL SaveDecryptedAttachment(HWND hWnd, string &sSourceFilename, string &sDestFilename);


private:
    BOOL MsgFromFile (LPMESSAGE pMessage, const char *strFile);
    void QuotedPrintEncode (char ** enc_buf, char * buf, size_t buflen);
};

#endif // INC_GPG_H
