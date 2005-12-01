/* gpgmsg.hh - The GpgMsg class
 *	Copyright (C) 2005 g10 Code GmbH
 *
 * This file is part of GPGol.
 * 
 * GPGol is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2 of the License, or (at your option) any later version.
 * 
 * GPGol is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Lesser General Public License for more details.
 * 
 * You should have received a copy of the GNU Lesser General Public
 * License along with this program; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA
 * 02110-1301, USA.
 */
#ifndef GPGMSG_HH
#define GPGMSG_HH

#include <gpgme.h>

#include "intern.h"

/* To manage a message we use our own class to keep track about all
   the information we known on the content of a message.  This is
   useful to remember the state of conversion (sometimes we need to
   copy between utf8 and the native character set) and to parse the
   message down into the MIME structure. */

class GpgMsg
{
public:    
  virtual void destroy () = 0;
  void operator delete (void *p)
    {
      if (p)
        {
          GpgMsg *m = (GpgMsg*)(p);
          m->destroy();
        }
    }

  /* Set a new MAPI message into the object. */
  virtual void setMapiMessage (LPMESSAGE msg) = 0;

  /* Set the callback for Exchange. */
  virtual void setExchangeCallback (void *cb) = 0;

  /* Don't pop up any message boxes and run the decryption only on the body. */
  virtual void setPreview (bool value) = 0;

  /* Return the type of the message. */
  virtual openpgp_t getMessageType (const char *text) = 0;

  /* Returns whether the message has any attachments. */
  virtual bool hasAttachments (void) = 0;

  /* Return a malloced array of malloced strings with the recipients
     of the message. Caller is responsible for freeing this array and
     the strings.  On failure NULL is returned.  */
  virtual char **getRecipients (void) = 0;

  /* Decrypt and verify the message and all attachments.  */
  virtual int decrypt (HWND hwnd) = 0;

  /* Sign the message and optionally the attachments. */
  virtual int sign (HWND hwnd) = 0;

  /* Encrypt the entire message including any attachments. Returns 0
     on success. */
  virtual int encrypt (HWND hwnd, bool want_html) = 0;

  /* Encrypt and sign the entire message including any
     attachments. Return 0 on success. */
  virtual int signEncrypt (HWND hwnd, bool want_html) = 0;

  /* Attach the key identified by KEYID to the message. */
  virtual int attachPublicKey (const char *keyid) = 0;

  /* Return the number of attachments. */
  virtual unsigned int getAttachments (void) = 0;

  /* Decrypt the attachment with the internal number POS.
     SAVE_PLAINTEXT must be true to save the attachemnt; displaying a
     attachemnt is not yet supported. */
  virtual void decryptAttachment (HWND hwnd, int pos, bool save_plaintext,
                                  int ttl, const char *filename) = 0;

  virtual void signAttachment (HWND hwnd, int pos,
                               gpgme_key_t sign_key, int ttl) = 0;

  virtual int encryptAttachment (HWND hwnd, int pos, gpgme_key_t *keys,
                                 gpgme_key_t sign_key, int ttl) = 0;

};


/* Create a new instance and initialize with the MAPI message object
   MSG. */
GpgMsg *CreateGpgMsg (LPMESSAGE msg);

#endif /*GPGMSG_HH*/
