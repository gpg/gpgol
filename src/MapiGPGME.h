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


class MapiGPGME
{
public:    
  virtual void __stdcall destroy () = 0;

  virtual int __stdcall encrypt (void) = 0;
  virtual int __stdcall decrypt (void) = 0;
  virtual int __stdcall sign (void) = 0;
  virtual int __stdcall verify (void) = 0;
  virtual int __stdcall signEncrypt (void) = 0;
  virtual int __stdcall attachPublicKey (const char *keyid) = 0;

  virtual int __stdcall doCmd (int doEncrypt, int doSign) = 0;
  virtual int __stdcall doCmdAttach (int action) = 0;
  virtual int __stdcall doCmdFile (int action,
                                   const char *in, const char *out) = 0;

  virtual const char* __stdcall getAttachmentExtension (const char *fname) = 0;
  virtual void __stdcall freeAttachments (void) = 0;
  virtual int __stdcall getAttachments (void) = 0;
  virtual int __stdcall countAttachments (void) = 0;
  virtual bool __stdcall hasAttachments (void) = 0;
  virtual bool __stdcall deleteAttachment (int pos) = 0;
  virtual LPATTACH __stdcall createAttachment (int &pos) = 0;

  virtual outlgpg_type_t __stdcall getMessageType (const char *body) = 0;
  virtual void __stdcall setMessage (LPMESSAGE msg) = 0;
  virtual void __stdcall setWindow (HWND hwnd) = 0;

  virtual const char * __stdcall getPassphrase (const char *keyid) = 0;
  virtual void __stdcall storePassphrase (void *itm) = 0;
  virtual void __stdcall clearPassphrase (void) = 0;

  virtual void __stdcall logDebug (const char *fmt, ...) = 0;

  virtual int __stdcall readOptions (void) = 0;
  virtual int __stdcall writeOptions (void) = 0;

  virtual int  __stdcall getStorePasswdTime (void) = 0;
  virtual void __stdcall setStorePasswdTime (int nCacheTime) = 0;
  virtual bool __stdcall getEncryptDefault (void) = 0;
  virtual void __stdcall setEncryptDefault (bool doEncrypt) = 0;
  virtual bool __stdcall getSignDefault (void) = 0;
  virtual void __stdcall setSignDefault (bool doSign) = 0;
  virtual bool __stdcall getEncryptWithDefaultKey (void) = 0;
  virtual void __stdcall setEncryptWithDefaultKey (bool encryptDefault) = 0;
  virtual bool __stdcall getSaveDecryptedAttachments (void) = 0;
  virtual void __stdcall setSaveDecryptedAttachments (bool saveDecrAtt) = 0;
  virtual void __stdcall setEncodingFormat (int fmt) = 0;
  virtual int  __stdcall getEncodingFormat (void) = 0;
  virtual void __stdcall setSignAttachments (bool signAtt) = 0;
  virtual bool __stdcall getSignAttachments (void) = 0;
  virtual void __stdcall setEnableLogging (bool val) = 0;
  virtual bool __stdcall getEnableLogging (void) = 0;
  virtual const char * __stdcall getLogFile (void) = 0;
  virtual void __stdcall setLogFile (const char *logfile) = 0;
  virtual void __stdcall setDefaultKey (const char *key) = 0;
  virtual char * __stdcall getDefaultKey (void) = 0;

  virtual int __stdcall startKeyManager () = 0;
  virtual void __stdcall startConfigDialog (HWND parent) = 0;

  void operator delete (void *p)
    {
      if (p)
        {
          MapiGPGME *m = (MapiGPGME*)(p);
          m->destroy();
        }
    }
};


extern "C" {

MapiGPGME * __stdcall CreateMapiGPGME (LPMESSAGE msg);
}

#endif /*MAPI_GPGME_H*/
