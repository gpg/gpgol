/* gpgmsg.cpp - Implementation ofthe GpgMsg class
 *	Copyright (C) 2005 g10 Code GmbH
 *
 * This file is part of OutlGPG.
 * 
 * OutlGPG is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2 of the License, or (at your option) any later version.
 * 
 * OutlGPG is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Lesser General Public License for more details.
 * 
 * You should have received a copy of the GNU Lesser General Public
 * License along with this program; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA
 * 02110-1301, USA.
 */

#ifdef HAVE_CONFIG_H
#include <config.h>
#endif

#include <windows.h>
#include <assert.h>
#include <string.h>

#include "mymapi.h"
#include "mymapitags.h"

#include "intern.h"
#include "gpgmsg.hh"
#include "util.h"
#include "msgcache.h"
#include "engine.h"
#include "display.h"

static const char oid_mimetag[] =
  {0x2A, 0x86, 0x48, 0x86, 0xf7, 0x14, 0x03, 0x0a, 0x04};


/* The string used as the standard XXXXX of decrypted attachments. */
#define ATT_FILE_PREFIX ".pgpenc"

#define TRACEPOINT() do { log_debug ("%s:%s:%d: tracepoint\n", \
                                       __FILE__, __func__, __LINE__); \
                        } while (0)

struct attach_info
{
  int end_of_table;  /* True if this is the last plus one entry of the
                        table. */
  int is_encrypted;  /* This is an encrypted attchment. */
  int is_signed;     /* This is a signed attachment. */
  unsigned int sig_pos; /* For signed attachments the index of the
                           attachment with the detached signature. */
  
  int method;        /* MAPI attachmend method. */
  char *filename;    /* Malloced filename of this attachment or NULL. */
  char *content_type;/* Malloced string with the mime attrib or NULL.
                        Parameters are stripped off thus a compare
                        against "type/subtype" is sufficient. */
  const char *content_type_parms; /* If not NULL the parameters of the
                                     content_type. */

};
typedef struct attach_info *attach_info_t;


/*
   The implementation class of MapiGPGME.  
 */
class GpgMsgImpl : public GpgMsg
{
public:    
  GpgMsgImpl () 
  {
    message = NULL;
    body = NULL;
    body_plain = NULL;
    body_cipher = NULL;
    body_signed = NULL;
    body_cipher_is_html = false;

    attach.att_table = NULL;
    attach.rows = NULL;
  }

  ~GpgMsgImpl ()
  {
    if (message)
      message->Release ();
    xfree (body);
    xfree (body_plain);
    xfree (body_cipher);
    xfree (body_signed);

    if (attach.att_table)
      {
        attach.att_table->Release ();
        attach.att_table = NULL;
      }
    if (attach.rows)
      {
        FreeProws (attach.rows);
        attach.rows = NULL;
      }
  }

  void destroy ()
  {
    delete this;
  }

  void operator delete (void *p) 
  {
    ::operator delete (p);
  }

  void setMapiMessage (LPMESSAGE msg)
  {
    if (message)
      {
        message->Release ();
        message = NULL;
      }
    if (msg)
      {
        msg->AddRef ();
        message = msg;
      }
  }
  
  openpgp_t getMessageType (void);
  bool hasAttachments (void);
  const char *getOrigText (void);
  const char *GpgMsgImpl::getDisplayText (void);
  const char *getPlainText (void);
  void setPlainText (char *string);
  void setSignedText (char *string);
  bool matchesString (const char *string);

  int decrypt (HWND hwnd);
  int verify (HWND hwnd);
  int sign (HWND hwnd);
  int encrypt (HWND hwnd)
  {
    return encrypt_and_sign (hwnd, false);
  }
  int signEncrypt (HWND hwnd)
  {
    return encrypt_and_sign (hwnd, true);
  }
  int attachPublicKey (const char *keyid);

  char **getRecipients (void);
  unsigned int getAttachments (void);
  void verifyAttachment (HWND hwnd, attach_info_t table,
                         unsigned int pos_data,
                         unsigned int pos_sig);
  void decryptAttachment (HWND hwnd, int pos, bool save_plaintext, int ttl,
                          const char *filename);
  void signAttachment (HWND hwnd, int pos, gpgme_key_t sign_key, int ttl);
  int encryptAttachment (HWND hwnd, int pos, gpgme_key_t *keys,
                         gpgme_key_t sign_key, int ttl);


private:
  LPMESSAGE message;  /* Pointer to the message. */
  char *body;         /* utf-8 encoded body string or NULL. */
  char *body_plain;   /* Plaintext version of BODY or NULL. */
  char *body_cipher;  /* Enciphered version of BODY or NULL. */
  char *body_signed;  /* Signed version of BODY or NULL. */
  bool body_cipher_is_html; /* Indicating that BODY_CIPHER holds HTML. */

  /* This structure collects the information about attachments. */
  struct 
  {
    LPMAPITABLE att_table;/* The loaded attachment table or NULL. */
    LPSRowSet   rows;     /* The retrieved set of rows from the table. */
  } attach;
  
  void loadBody (void);
  attach_info_t gatherAttachmentInfo (void);
  int encrypt_and_sign (HWND hwnd, bool sign);
};


/* Return a new instance and initialize with the MAPI message object
   MSG. */
GpgMsg *
CreateGpgMsg (LPMESSAGE msg)
{
  GpgMsg *m = new GpgMsgImpl ();
  if (!m)
    out_of_core ();
  m->setMapiMessage (msg);
  return m;
}


/* Release an array of GPGME keys. */
static void 
free_key_array (gpgme_key_t *keys)
{
  if (keys)
    {
      for (int i = 0; keys[i]; i++) 
	gpgme_key_release (keys[i]);
      xfree (keys);
    }
}

/* Release an array of strings with recipient names. */
static void
free_recipient_array (char **recipients)
{
  int i;

  if (recipients)
    {
      for (i=0; recipients[i]; i++) 
	xfree (recipients[i]);	
      xfree (recipients);
    }
}

/* Release a table with attachments infos. */
static void
release_attach_info (attach_info_t table)
{
  int i;

  if (!table)
    return;
  for (i=0; !table[i].end_of_table; i++)
    {
      xfree (table[i].filename);
      xfree (table[i].content_type);
    }
  xfree (table);
}


/* Return the number of recipients in the array RECIPIENTS. */
static int 
count_recipients (char **recipients)
{
  int i;
  
  for (i=0; recipients[i] != NULL; i++)
    ;
  return i;
}


/* Return a string suitable for displaying in a message box.  The
   function takes FORMAT and replaces the string "@LIST@" with the
   names of the attachmets. Depending on the set bits in WHAT only
   certain attachments are inserted. 

   Defined bits in MODE are:
      0 = Any attachment
      1 = signed attachments
      2 = encrypted attachments

   Caller must free the returned value.  Routine is guaranteed to
   return a string.
*/
static char *
text_from_attach_info (attach_info_t table, const char *format,
                       unsigned int what)
{
  int pos;
  size_t length;
  char *buffer, *p;
  const char *marker;

  marker = strstr (format, "@LIST@");
  if (!marker)
    return xstrdup (format);

#define CONDITION  (table[pos].filename \
                    && ( (what&1) \
                         || ((what & 2) && table[pos].is_signed) \
                         || ((what & 4) && table[pos].is_encrypted)))

  for (length=0, pos=0; !table[pos].end_of_table; pos++)
    if (CONDITION)
      length += 2 + strlen (table[pos].filename) + 1;

  length += strlen (format);
  buffer = p = (char*)xmalloc (length+1);

  strncpy (p, format, marker - format);
  p += marker - format;

  for (pos=0; !table[pos].end_of_table; pos++)
    if (CONDITION)
      {
        if (table[pos].is_signed)
          p = stpcpy (p, "S ");
        else if (table[pos].is_encrypted)
          p = stpcpy (p, "E ");
        else
          p = stpcpy (p, "* ");
        p = stpcpy (p, table[pos].filename);
        p = stpcpy (p, "\n");
      }
  strcpy (p, marker+6);
#undef CONDITION

  return buffer;
}



/* Load the body and make it available as an UTF8 string in the
   instance variable BODY.  */
void
GpgMsgImpl::loadBody (void)
{
  HRESULT hr;
  LPSPropValue lpspvFEID = NULL;
  LPSTREAM stream;
  SPropValue prop;
  STATSTG statInfo;
  ULONG nread;

  if (body || !message)
    return;

#if 1
  hr = message->OpenProperty (PR_BODY, &IID_IStream,
                              0, 0, (LPUNKNOWN*)&stream);
  if ( hr != S_OK )
    {
      log_debug ("%s:%s: OpenProperty failed: hr=%#lx",
                 __FILE__, __func__, hr);
      return;
    }

  hr = stream->Stat (&statInfo, STATFLAG_NONAME);
  if ( hr != S_OK )
    {
      log_debug ("%s:%s: Stat failed: hr=%#lx", __FILE__, __func__, hr);
      stream->Release ();
      return;
    }
  
  /* FIXME: We might want to read only the first 1k to decide whetehr
     this is actually an OpenPGP message and only then continue
     reading.  This requires some changes in this module. */
  body = (char*)xmalloc ((size_t)statInfo.cbSize.QuadPart + 2);
  hr = stream->Read (body, (size_t)statInfo.cbSize.QuadPart, &nread);
  if ( hr != S_OK )
    {
      log_debug ("%s:%s: Read failed: hr=%#lx", __FILE__, __func__, hr);
      xfree (body);
      body = NULL;
      stream->Release ();
      return;
    }
  body[nread] = 0;
  body[nread+1] = 0;
  if (nread != statInfo.cbSize.QuadPart)
    {
      log_debug ("%s:%s: not enough bytes returned\n", __FILE__, __func__);
      xfree (body);
      body = NULL;
      stream->Release ();
      return;
    }
  stream->Release ();

  /* Fixme: We need to optimize this. */
  {
    char *tmp;
    tmp = wchar_to_utf8 ((wchar_t*)body);
    if (!tmp)
      log_debug ("%s: error converting to utf8\n", __func__);
    else
      {
        xfree (body);
        body = tmp;
      }
  }

#else /* Old method. */
  hr = HrGetOneProp ((LPMAPIPROP)message, PR_BODY, &lpspvFEID);
  if (FAILED (hr))
    {
      log_debug ("%s: HrGetOneProp failed\n", __func__);
      return;
    }
    
  switch ( PROP_TYPE (lpspvFEID->ulPropTag) )
    {
    case PT_UNICODE:
      body = wchar_to_utf8 (lpspvFEID->Value.lpszW);
      if (!body)
        log_debug ("%s: error converting to utf8\n", __func__);
      break;
      
    case PT_STRING8:
      body = xstrdup (lpspvFEID->Value.lpszA);
      break;
      
    default:
      log_debug ("%s: proptag=0x%08lx not supported\n",
                 __func__, lpspvFEID->ulPropTag);
      break;
    }
  MAPIFreeBuffer (lpspvFEID);
#endif  

  if (body)
    log_debug ("%s:%s: loaded body `%s' at %p\n",
               __FILE__, __func__, body, body);
  

//   prop.ulPropTag = PR_ACCESS;
//   prop.Value.l = MAPI_ACCESS_MODIFY;
//   hr = HrSetOneProp (message, &prop);
//   if (FAILED (hr))
//     log_debug_w32 (-1,"%s:%s: updating access to 0x%08lx failed",
//                    __FILE__, __func__, prop.Value.l);
}


/* Return the type of a message. */
openpgp_t
GpgMsgImpl::getMessageType (void)
{
  const char *s;
  
  loadBody ();
  
  if (!body || !(s = strstr (body, "BEGIN PGP ")))
    return OPENPGP_NONE;

  /* (The extra strstr() above is just a simple optimization.) */
  if (strstr (body, "BEGIN PGP MESSAGE"))
    return OPENPGP_MSG;
  else if (strstr (body, "BEGIN PGP SIGNED MESSAGE"))
    return OPENPGP_CLEARSIG;
  else if (strstr (body, "BEGIN PGP SIGNATURE"))
    return OPENPGP_SIG;
  else if (strstr (body, "BEGIN PGP PUBLIC KEY"))
    return OPENPGP_PUBKEY;
  else if (strstr (body, "BEGIN PGP PRIVATE KEY"))
    return OPENPGP_SECKEY;
  else
    return OPENPGP_NONE;
}


/* Return the body text as received or composed.  This is guaranteed
   to never return NULL.  */
const char *
GpgMsgImpl::getOrigText ()
{
  loadBody ();
  
  return body? body : "";
}


/* Return the text of the message to be used for the display.  The
   message objects has intrinsic knowledge about the correct text.  */
const char *
GpgMsgImpl::getDisplayText (void)
{
  loadBody ();

  if (body_plain)
    return body_plain;
  else if (body)
    return body;
  else
    return "";
}



/* Save STRING as the plaintext version of the message.  WARNING:
   ownership of STRING is transferred to this object. */
void
GpgMsgImpl::setPlainText (char *string)
{
  xfree (body_plain);
  body_plain = string;
  msgcache_put (body_plain, 0, message);
}


/* Save STRING as the signed version of the message.  WARNING:
   ownership of STRING is transferred to this object. */
void
GpgMsgImpl::setSignedText (char *string)
{
  xfree (body_signed);
  body_signed = string;
}


/* Returns true if STRING matches the actual message. */ 
bool
GpgMsgImpl::matchesString (const char *string)
{
  /* FIXME:  This is a too simple implementation. */
  if (string && strstr (string, "BEGIN PGP ") )
    return true;
  return false;
}



/* Return an array of strings with the recipients of the message. On
   success a malloced array is returned containing allocated strings
   for each recipient.  The end of the array is marked by NULL.
   Caller is responsible for releasing the array.  On failure NULL is
   returned.  */
char ** 
GpgMsgImpl::getRecipients ()
{
  static SizedSPropTagArray (1L, PropRecipientNum) = {1L, {PR_EMAIL_ADDRESS}};
  HRESULT hr;
  LPMAPITABLE lpRecipientTable = NULL;
  LPSRowSet lpRecipientRows = NULL;
  char **rset;
  const char *s;
  int i, j;

  if (!message)
    return NULL;

  hr = message->GetRecipientTable (0, &lpRecipientTable);
  if (FAILED (hr)) 
    {
      log_debug_w32 (-1, "%s:%s: GetRecipientTable failed",
                     __FILE__, __func__);
      return NULL;
    }

  hr = HrQueryAllRows (lpRecipientTable, (LPSPropTagArray) &PropRecipientNum,
                       NULL, NULL, 0L, &lpRecipientRows);
  if (FAILED (hr)) 
    {
      log_debug_w32 (-1, "%s:%s: GHrQueryAllRows failed", __FILE__, __func__);
      if (lpRecipientTable)
        lpRecipientTable->Release();
      return NULL;
    }

  rset = (char**)xcalloc (lpRecipientRows->cRows+1, sizeof *rset);

  for (i = j = 0; i < lpRecipientRows->cRows; i++)
    {
      LPSPropValue row;

      if (!lpRecipientRows->aRow[j].cValues)
        continue;
      row = lpRecipientRows->aRow[j].lpProps;

      switch ( PROP_TYPE (row->ulPropTag) )
        {
        case PT_UNICODE:
          rset[j] = wchar_to_utf8 (row->Value.lpszW);
          if (rset[j])
            j++;
          else
            log_debug ("%s:%s: error converting recipient to utf8\n",
                       __FILE__, __func__);
          break;
      
        case PT_STRING8: /* Assume Ascii. */
          rset[j++] = xstrdup (row->Value.lpszA);
          break;
          
        default:
          log_debug ("%s:%s: proptag=0x%08lx not supported\n",
                     __FILE__, __func__, row->ulPropTag);
          break;
        }
    }
  rset[j] = NULL;

  if (lpRecipientTable)
    lpRecipientTable->Release();
  if (lpRecipientRows)
    FreeProws(lpRecipientRows);	
  
  log_debug ("%s:%s: got %d recipients:\n",
             __FILE__, __func__, j);
  for (i=0; rset[i]; i++)
    log_debug ("%s:%s: \t`%s'\n", __FILE__, __func__, rset[i]);

  return rset;
}




/* Decrypt the message MSG and update the window.  HWND identifies the
   current window. */
int 
GpgMsgImpl::decrypt (HWND hwnd)
{
  log_debug ("%s:%s: enter\n", __FILE__, __func__);
  openpgp_t mtype;
  char *plaintext = NULL;
  attach_info_t table = NULL;
  int err;
  unsigned int pos;
  unsigned int n_attach = 0;
  unsigned int n_encrypted = 0;
  unsigned int n_signed = 0;
  HRESULT hr;

  mtype = getMessageType ();
  if (mtype == OPENPGP_CLEARSIG)
    {
      log_debug ("%s:%s: leave (passing on to verify)\n", __FILE__, __func__);
      return verify (hwnd);
    }

  /* Check whether this possibly encrypted message has encrypted
     attachments.  We check right now because we need to get into the
     decryption code even if the body is not encrypted but attachments
     are available. */
  table = gatherAttachmentInfo ();
  if (table)
    {
      for (pos=0; !table[pos].end_of_table; pos++)
        if (table[pos].is_encrypted)
          n_encrypted++;
        else if (table[pos].is_signed)
          n_signed++;
      n_attach = pos;
    }
  log_debug ("%s:%s: message has %u attachments with "
             "%u signed and %d encrypted\n",
             __FILE__, __func__, n_attach, n_signed, n_encrypted);
  if (mtype == OPENPGP_NONE && !n_encrypted && !n_signed) 
    {
      /* Fixme: we should display the messsage box only if decryption
         has explicity be requested. */
      MessageBox (hwnd, "No valid OpenPGP data found.",
                  "GPG Decryption", MB_ICONWARNING|MB_OK);
      log_debug ("%s:%s: leave (no OpenPGP data)\n", __FILE__, __func__);
      release_attach_info (table);
      return 0;
    }


  err = *getOrigText()? op_decrypt (getOrigText (), &plaintext,
                                    opt.passwd_ttl, NULL)
                      : gpg_error (GPG_ERR_NO_DATA);
  if (err)
    {
      if (n_attach && gpg_err_code (err) == GPG_ERR_NO_DATA)
        ;
      else
        MessageBox (hwnd, op_strerror (err),
                    "GPG Decryption", MB_ICONERROR|MB_OK);
    }
  else if (plaintext && *plaintext)
    {	
      int is_html = is_html_body (plaintext);

      set_message_body (message, plaintext);
      setPlainText (plaintext);
      plaintext = NULL;

      /* Also set PR_BODY but do not use 'SaveChanges' to make it
         permanently.  This way the user can reply with the
         plaintext but the ciphertext is still stored. 

         NOTE: That does not work for OL2003!  This is the reason we
         implemented the (not fully working) msgcache system. */
      log_debug ("decrypt isHtml=%d\n", is_html);

      /* XXX: find a way to handle text/html message in a better way! */
      /* I have disabled the kludge to see what happens to a html
         message. */
      if (/*is_html ||*/ update_display (hwnd, this))
        {
          const char s[] = 
            "The message text cannot be displayed.\n"
            "You have to save the decrypted message to view it.\n"
            "Then you need to re-open the message.\n\n"
            "Do you want to save the decrypted message?";
          int what;
          
          what = MessageBox (hwnd, s, "GPG Decryption",
                             MB_YESNO|MB_ICONWARNING);
          if (what == IDYES) 
            {
              log_debug ("decrypt: saving plaintext message.\n");
              hr = message->SaveChanges (KEEP_OPEN_READWRITE|FORCE_SAVE);
              if (FAILED (hr))
                log_debug ("%s:%s: SaveChanges failed: hr=%#lx",
                           __FILE__, __func__, hr);
            }
	}
    }


  /* If we have signed attachments.  Ask whether the signatures should
     be verified; we do this is case of large attachments where
     verification might take long. */
  if (n_signed)
    {
      const char s[] = 
        "Signed attachments found.\n\n"
        "@LIST@\n"
        "Do you want to verify the signatures?";
      int what;
      char *text;

      text = text_from_attach_info (table, s, 2);
      
      what = MessageBox (hwnd, text, "Attachment Verification",
                         MB_YESNO|MB_ICONINFORMATION);
      xfree (text);
      if (what == IDYES) 
        {
          for (pos=0; !table[pos].end_of_table; pos++)
            if (table[pos].is_signed)
              {
                assert (table[pos].sig_pos < n_attach);
                verifyAttachment (hwnd, table, pos, table[pos].sig_pos);
              }
        }
    }

  if (n_encrypted)
    {
      const char s[] = 
        "Encrypted attachments found.\n\n"
        "@LIST@\n"
        "Do you want to decrypt and save them?";
      int what;
      char *text;

      text = text_from_attach_info (table, s, 4);
      what = MessageBox (hwnd, text, "Attachment Decryption",
                         MB_YESNO|MB_ICONINFORMATION);
      xfree (text);
      if (what == IDYES) 
        {
          for (pos=0; !table[pos].end_of_table; pos++)
            if (table[pos].is_encrypted)
              decryptAttachment (hwnd, pos, true, opt.passwd_ttl,
                                 table[pos].filename);
        }
    }

  log_debug ("%s:%s: leave (rc=%d)\n", __FILE__, __func__, err);
  return err;
}


/* Verify the message and display the verification result. */
int 
GpgMsgImpl::verify (HWND hwnd)
{
  log_debug ("%s:%s: enter\n", __FILE__, __func__);
  openpgp_t mtype;
  int err, has_attach;
  
  mtype = getMessageType ();
  has_attach = hasAttachments ();
  if (mtype == OPENPGP_NONE && !has_attach ) 
    {
      log_debug ("%s:%s: leave (no OpenPGP data)\n", __FILE__, __func__);
      return 0;
    }

  err = op_verify (getOrigText (), NULL, NULL);
  if (err)
    MessageBox (hwnd, op_strerror (err), "GPG Verify", MB_ICONERROR|MB_OK);
  else
    update_display (hwnd, this);

  log_debug ("%s:%s: leave (rc=%d)\n", __FILE__, __func__, err);
  return err;
}





/* Sign the current message. Returns 0 on success. */
int
GpgMsgImpl::sign (HWND hwnd)
{
  const char *plaintext;
  char *signedtext = NULL;
  int err = 0;
  gpgme_key_t sign_key = NULL;

  log_debug ("%s:%s: enter\n", __FILE__, __func__);
  
  /* We don't sign an empty body - a signature on a zero length string
     is pretty much useless. */
  if (!*(plaintext = getOrigText ()) && !hasAttachments ()) 
    {
      log_debug ("%s:%s: leave (empty)", __FILE__, __func__);
      return 0; 
    }

  /* Pop up a dialog box to ask for the signer of the message. */
  if (signer_dialog_box (&sign_key, NULL) == -1)
    {
      log_debug ("%s.%s: leave (dialog failed)\n", __FILE__, __func__);
      return gpg_error (GPG_ERR_CANCELED);  
    }

  if (*plaintext)
    {
      err = op_sign (plaintext, &signedtext, 
                     OP_SIG_CLEAR, sign_key, opt.passwd_ttl);
      if (err)
        {
          MessageBox (hwnd, op_strerror (err),
                      "GPG Sign", MB_ICONERROR|MB_OK);
          goto leave;
        }
      setSignedText (signedtext);
      signedtext = NULL;
    }

  if (opt.auto_sign_attach && hasAttachments ())
    {
      unsigned int n;
      
      n = getAttachments ();
      log_debug ("%s:%s: message has %u attachments\n", __FILE__, __func__, n);
      for (int i=0; i < n; i++) 
        signAttachment (hwnd, i, sign_key, opt.passwd_ttl);
      /* FIXME: we should throw an error if signing of any attachment
         failed. */
    }

  /* Now that we successfully processed the attachments, we can save
     the changes to the body.  For unknown reasons we need to set it
     to empty first. */
  if (*plaintext)
    {
      err = set_message_body (message, "");
      if (!err)
        set_message_body (message, signedtext);
      if (err)
        goto leave;
    }


 leave:
  xfree (signedtext);
  gpgme_key_release (sign_key);
  log_debug ("%s:%s: leave (err=%s)\n", __FILE__, __func__, op_strerror (err));
  return err;
}



/* Encrypt and optionally sign (if SIGN is true) the entire message
   including all attachments. Returns 0 on success. */
int 
GpgMsgImpl::encrypt_and_sign (HWND hwnd, bool sign)
{
  log_debug ("%s:%s: enter\n", __FILE__, __func__);
  gpgme_key_t *keys = NULL;
  gpgme_key_t sign_key = NULL;
  bool is_html;
  const char *plaintext;
  char *ciphertext = NULL;
  char **recipients = NULL;
  char **unknown = NULL;
  int err = 0;
  size_t all = 0;
  int n_keys;
    
  
  if (!*(plaintext = getOrigText ()) && !hasAttachments ()) 
    {
      log_debug ("%s:%s: leave (empty)", __FILE__, __func__);
      return 0; 
    }

  /* Pop up a dialog box to ask for the signer of the message. */
  if (sign)
    {
      if (signer_dialog_box (&sign_key, NULL) == -1)
        {
          log_debug ("%s.%s: leave (dialog failed)\n", __FILE__, __func__);
          return gpg_error (GPG_ERR_CANCELED);  
        }
    }

  /* Gather the keys for the recipients. */
  recipients = getRecipients ();
  n_keys = op_lookup_keys (recipients, &keys, &unknown, &all);

  log_debug ("%s:%s: found %d recipients, need %d, unknown=%p\n",
             __FILE__, __func__, n_keys, all, unknown);
  
  if (n_keys != count_recipients (recipients))
    {
      int opts = 0;
      gpgme_key_t *keys2 = NULL;

      /* FIXME: The implementation is not correct: op_lookup_keys
         returns the number of missing keys but we compare against the
         total number of keys; thus the box would pop up even when all
         have been found. */
      log_debug ("%s:%s: calling recipient_dialog_box2", __FILE__, __func__);
      recipient_dialog_box2 (keys, unknown, all, &keys2, &opts);
      xfree (keys);
      keys = keys2;
      if (opts & OPT_FLAG_CANCEL) 
        {
          err = gpg_error (GPG_ERR_CANCELED);
          goto leave;
	}
    }

  if (sign_key)
    log_debug ("%s:%s: signer: 0x%s %s\n",  __FILE__, __func__,
               keyid_from_key (sign_key), userid_from_key (sign_key));
  else
    log_debug ("%s:%s: no signer\n", __FILE__, __func__);
  if (keys)
    {
      for (int i=0; keys[i] != NULL; i++)
        log_debug ("%s.%s: recp.%d 0x%s %s\n", __FILE__, __func__,
                   i, keyid_from_key (keys[i]), userid_from_key (keys[i]));
    }


  if (*plaintext)
    {
      is_html = is_html_body (plaintext);

      err = op_encrypt (plaintext, &ciphertext, keys, NULL, 0);
      if (err)
        {
          MessageBox (hwnd, op_strerror (err),
                      "GPG Encryption", MB_ICONERROR|MB_OK);
          goto leave;
        }

      if (is_html) 
        {
          char *tmp = add_html_line_endings (ciphertext);
          xfree (ciphertext);
          ciphertext = tmp;
        }

//       {
//         HRESULT hr;
//         SPropValue prop;

//         prop.ulPropTag=PR_MESSAGE_CLASS_A;
//         prop.Value.lpszA="IPM.Note.OPENPGP";
//         hr = HrSetOneProp (message, &prop);
//         if (hr != S_OK)
//           {
//             log_error ("%s:%s: can't set message class: hr=%#lx\n",
//                        __FILE__, __func__, hr); 
//           }

//         prop.ulPropTag=PR_CONTENT_TYPE_A;
//         prop.Value.lpszA="application/encrypted;foo=bar;type=mytype";
//         hr = HrSetOneProp (message, &prop);
//         if (hr != S_OK)
//           {
//             log_error ("%s:%s: can't set content type: hr=%#lx\n",
//                        __FILE__, __func__, hr); 
//           }
//       }

    }

  if (hasAttachments ())
    {
      unsigned int n;
      
      n = getAttachments ();
      log_debug ("%s:%s: message has %u attachments\n", __FILE__, __func__, n);
      for (int i=0; !err && i < n; i++) 
        err = encryptAttachment (hwnd, i, keys, NULL, 0);
      if (err)
        {
          MessageBox (hwnd, op_strerror (err),
                      "GPG Attachment Encryption", MB_ICONERROR|MB_OK);
          goto leave;
        }
    }

  /* Now that we successfully processed the attachments, we can save
     the changes to the body.  For unknown reasons we need to set it
     to empty first. */
  if (*plaintext)
    {
      err = set_message_body (message, "");
      if (!err)
        set_message_body (message, ciphertext);
      if (err)
        goto leave;
    }


 leave:
  /* FIXME: What to do with already encrypted attachments if some of
     the encrypted (or other operations) failed? */

  for (int i=0; i < n_keys; i++)
    xfree (unknown[i]);
  if (n_keys)
    xfree (unknown);

  free_key_array (keys);
  free_recipient_array (recipients);
  xfree (ciphertext);
  log_debug ("%s:%s: leave (err=%s)\n", __FILE__, __func__, op_strerror (err));
  return err;
}




/* Attach a public key to a message. */
int 
GpgMsgImpl::attachPublicKey (const char *keyid)
{
    /* @untested@ */
#if 0
    const char *patt[1];
    char *keyfile;
    int err, pos = 0;
    LPATTACH newatt;

    keyfile = generateTempname (keyid);
    patt[0] = xstrdup (keyid);
    err = op_export_keys (patt, keyfile);

    newatt = createAttachment (NULL/*FIXME*/,pos);
    setAttachMethod (newatt, ATTACH_BY_VALUE);
    setAttachFilename (newatt, keyfile, false);
    /* XXX: set proper RFC3156 MIME types. */

    if (streamFromFile (keyfile, newatt)) {
	log_debug ("attachPublicKey: commit changes.\n");
	newatt->SaveChanges (FORCE_SAVE);
    }
    releaseAttachment (newatt);
    xfree (keyfile);
    xfree ((void *)patt[0]);
    return err;
#endif
    return -1;
}





/* Returns whether the message has any attachments. */
bool
GpgMsgImpl::hasAttachments (void)
{
  return !!getAttachments ();
}


/* Reads the attachment information and returns the number of
   attachments. */
unsigned int
GpgMsgImpl::getAttachments (void)
{
  SizedSPropTagArray (1L, propAttNum) = {
    1L, {PR_ATTACH_NUM}
  };
  HRESULT hr;    
  LPMAPITABLE table;
  LPSRowSet   rows;

  if (!message)
    return 0;

  if (!attach.att_table)
    {
      hr = message->GetAttachmentTable (0, &table);
      if (FAILED (hr))
        {
          log_debug ("%s:%s: GetAttachmentTable failed: hr=%#lx",
                     __FILE__, __func__, hr);
          return 0;
        }
      
      hr = HrQueryAllRows (table, (LPSPropTagArray)&propAttNum,
                           NULL, NULL, 0, &rows);
      if (FAILED (hr))
        {
          log_debug ("%s:%s: HrQueryAllRows failed: hr=%#lx",
                     __FILE__, __func__, hr);
          table->Release ();
          return 0;
        }
      attach.att_table = table;
      attach.rows = rows;
    }

  return attach.rows->cRows > 0? attach.rows->cRows : 0;
}

/* Return the attachment method for attachment OBJ. In case of error we
   return 0 which happens to be not defined. */
static int
get_attach_method (LPATTACH obj)
{
  HRESULT hr;
  LPSPropValue propval = NULL;
  int method ;

  hr = HrGetOneProp ((LPMAPIPROP)obj, PR_ATTACH_METHOD, &propval);
  if (FAILED (hr))
    {
      log_error ("%s:%s: error getting attachment method: hr=%#lx",
                 __FILE__, __func__, hr);
      return 0; 
    }
  /* We don't bother checking whether we really get a PT_LONG ulong
     back; if not the system is seriously damaged and we can't do
     further harm by returning a possible random value. */
  method = propval->Value.l;
  MAPIFreeBuffer (propval);
  return method;
}


/* Return the content-type of the attachment OBJ or NULL if it dow not
   exists.  Caller must free. */
static char *
get_attach_mime_tag (LPATTACH obj)
{
  HRESULT hr;
  LPSPropValue propval = NULL;
  int method ;
  char *name;

  hr = HrGetOneProp ((LPMAPIPROP)obj, PR_ATTACH_MIME_TAG_A, &propval);
  if (FAILED (hr))
    {
      log_error ("%s:%s: error getting attachment's MIME tag: hr=%#lx",
                 __FILE__, __func__, hr);
      return NULL; 
    }
  switch ( PROP_TYPE (propval->ulPropTag) )
    {
    case PT_UNICODE:
      name = wchar_to_utf8 (propval->Value.lpszW);
      if (!name)
        log_debug ("%s:%s: error converting to utf8\n", __FILE__, __func__);
      break;
      
    case PT_STRING8:
      name = xstrdup (propval->Value.lpszA);
      break;
      
    default:
      log_debug ("%s:%s: proptag=%xlx not supported\n",
                 __FILE__, __func__, propval->ulPropTag);
      name = NULL;
      break;
    }
  MAPIFreeBuffer (propval);
  return name;
}


/* Return the data property of an attachments or NULL in case of an
   error.  Caller must free.  Note, that this routine should only be
   used for short data objects like detached signatures. */
static char *
get_short_attach_data (LPATTACH obj)
{
  HRESULT hr;
  LPSPropValue propval = NULL;
  int method ;
  char *data;

  hr = HrGetOneProp ((LPMAPIPROP)obj, PR_ATTACH_DATA_BIN, &propval);
  if (FAILED (hr))
    {
      log_error ("%s:%s: error getting attachment's data: hr=%#lx",
                 __FILE__, __func__, hr);
      return NULL; 
    }
  switch ( PROP_TYPE (propval->ulPropTag) )
    {
    case PT_BINARY:
      /* This is a binary obnject but we know that it must be plain
         ASCII due to the armoed format.  */
      data = (char*)xmalloc (propval->Value.bin.cb + 1);
      memcpy (data, propval->Value.bin.lpb, propval->Value.bin.cb);
      data[propval->Value.bin.cb] = 0;
      break;
      
    default:
      log_debug ("%s:%s: proptag=%#lx not supported\n",
                 __FILE__, __func__, propval->ulPropTag);
      data = NULL;
      break;
    }
  MAPIFreeBuffer (propval);
  return data;
}



/* Example Code. */
static bool 
set_x_header (GpgMsg * msg, const char *name, const char *val)
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
// FIXME
//    hr = msg->GetIDsFromNames (1, (LPMAPINAMEID*)mnid, MAPI_CREATE, &pProps);
    if (FAILED (hr)) {
	log_debug ("set X-Header failed.\n");
	return false;
    }
    
    pv.ulPropTag = (pProps->aulPropTag[0] & 0xFFFF0000) | PT_STRING8;
    pv.Value.lpszA = (char *)val;
//FIXME     hr = HrSetOneProp(msg, &pv);	
    if (!SUCCEEDED (hr)) {
	log_debug ("set X-Header failed.\n");
	return false;
    }

    log_debug ("set X-Header succeeded.\n");
    return true;
}




/* Return the filename from the attachment as a malloced string.  The
   encoding we return will be utf8, however the MAPI docs declare that
   MAPI does only handle plain ANSI and thus we don't really care
   later on.  In fact we would need to convert the filename back to
   wchar and use the Unicode versions of the file API.  Returns NULL
   on error or if no filename is available. */
static char *
get_attach_filename (LPATTACH obj)
{
  HRESULT hr;
  LPSPropValue propval;
  char *name = NULL;

  hr = HrGetOneProp ((LPMAPIPROP)obj, PR_ATTACH_LONG_FILENAME, &propval);
  if (FAILED(hr)) 
    hr = HrGetOneProp ((LPMAPIPROP)obj, PR_ATTACH_FILENAME, &propval);
  if (FAILED(hr))
    {
      log_debug ("%s:%s: no filename property found", __FILE__, __func__);
      return NULL;
    }

  switch ( PROP_TYPE (propval->ulPropTag) )
    {
    case PT_UNICODE:
      name = wchar_to_utf8 (propval->Value.lpszW);
      if (!name)
        log_debug ("%s:%s: error converting to utf8\n", __FILE__, __func__);
      break;
      
    case PT_STRING8:
      name = xstrdup (propval->Value.lpszA);
      break;
      
    default:
      log_debug ("%s:%s: proptag=%xlx not supported\n",
                 __FILE__, __func__, propval->ulPropTag);
      name = NULL;
      break;
    }
  MAPIFreeBuffer (propval);
  return name;
}




/* Return a filename to be used for saving an attachment. Returns an
   malloced string on success. HWND is the current Window and SRCNAME
   the filename to be used as suggestion.  On error; i.e. cancel NULL
   is returned. */
static char *
get_save_filename (HWND root, const char *srcname)
				     
{
  char filter[] = "All Files (*.*)\0*.*\0\0";
  char fname[MAX_PATH+1];
  const char *s;
  OPENFILENAME ofn;

  memset (fname, 0, sizeof (fname));
  strncpy (fname, srcname, MAX_PATH-1);
  fname[MAX_PATH] = 0;  
  

  memset (&ofn, 0, sizeof (ofn));
  ofn.lStructSize = sizeof (ofn);
  ofn.hwndOwner = root;
  ofn.lpstrFile = fname;
  ofn.nMaxFile = MAX_PATH;
  ofn.lpstrFileTitle = NULL;
  ofn.nMaxFileTitle = 0;
  ofn.Flags |= OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT;
  ofn.lpstrTitle = "GPG - Save decrypted attachment";
  ofn.lpstrFilter = filter;

  if (GetSaveFileName (&ofn))
    return xstrdup (fname);
  return NULL;
}


/* Gather information about attachments and return a new object with
   these information.  Caller must release the returned information.
   The routine will return NULL in case of an error or if no
   attachments are available. */
attach_info_t
GpgMsgImpl::gatherAttachmentInfo (void)
{    
  HRESULT hr;
  attach_info_t table;
  unsigned int pos, n_attach;
  int method, err;
  const char *s;

  n_attach = getAttachments ();
  log_debug ("%s:%s: message has %u attachments\n",
             __FILE__, __func__, n_attach);
  if (!n_attach)
      return NULL;

  table = (attach_info_t)xcalloc (n_attach+1, sizeof *table);
  for (pos=0; pos < n_attach; pos++) 
    {
      LPATTACH att;

      hr = message->OpenAttach (pos, NULL, MAPI_BEST_ACCESS, &att);	
      if (FAILED (hr))
        {
          log_error ("%s:%s: can't open attachment %d: hr=%#lx",
                     __FILE__, __func__, pos, hr);
          continue;
        }

      table[pos].method = get_attach_method (att);
      table[pos].filename = get_attach_filename (att);
      table[pos].content_type = get_attach_mime_tag (att);
      if (table[pos].content_type)
        {
          char *p = strchr (table[pos].content_type, ';');
          if (p)
            {
              *p++ = 0;
              trim_trailing_spaces (table[pos].content_type);
              while (strchr (" \t\r\n", *p))
                p++;
              trim_trailing_spaces (p);
              table[pos].content_type_parms = p;
            }
        }
      att->Release ();
    }
  table[pos].end_of_table = 1;

  /* Figure out whether there are encrypted attachments. */
  for (pos=0; !table[pos].end_of_table; pos++)
    {
      if (table[pos].filename && (s = strrchr (table[pos].filename, '.'))
          &&  (!stricmp (s, ".pgp") || !stricmp (s, ".gpg")))
        table[pos].is_encrypted = 1;
      else if (table[pos].content_type  
               && ( !stricmp (table[pos].content_type,
                              "application/pgp-encrypted")
                   || (!stricmp (table[pos].content_type,
                                 "multipart/encrypted")
                       && table[pos].content_type_parms
                       && strstr (table[pos].content_type_parms,
                                  "application/pgp-encrypted"))))
        table[pos].is_encrypted = 1;
    }
     
  /* Figure out what attachments are signed. */
  for (pos=0; !table[pos].end_of_table; pos++)
    {
      if (table[pos].filename && (s = strrchr (table[pos].filename, '.'))
          &&  !stricmp (s, ".asc")
          && table[pos].content_type  
          && !stricmp (table[pos].content_type, "application/pgp-signature"))
        {
          size_t len = (s - table[pos].filename);

          /* We mark the actual file, assuming that the .asc is a
             detached signature.  To correlate the data file and the
             signature we keep track of the POS. */
          for (int i=0; !table[i].end_of_table; i++)
            if (i != pos && table[i].filename 
                && strlen (table[i].filename) == len
                && !strncmp (table[i].filename, table[pos].filename, len))
              {
                table[i].is_signed = 1;
                table[i].sig_pos = pos;
              }
        }
    }

  log_debug ("%s:%s: attachment info:\n", __FILE__, __func__);
  for (int i=0; !table[i].end_of_table; i++)
    {
      log_debug ("\t%d %d %d %u `%s' `%s' `%s'\n",
                 i, table[i].is_encrypted,
                 table[i].is_signed, table[i].sig_pos,
                 table[i].filename, table[i].content_type,
                 table[i].content_type_parms);
    }
  

  return table;
}




/* Verify the ATTachment at attachments and table position POS_DATA
   agains the signature at position POS_SIG.  Display the status for
   each signature. */
void
GpgMsgImpl::verifyAttachment (HWND hwnd, attach_info_t table,
                              unsigned int pos_data,
                              unsigned int pos_sig)

{    
  HRESULT hr;
  LPATTACH att;
  int err;
  char *sig_data;

  log_debug ("%s:%s: verifying attachment %d/%d",
             __FILE__, __func__, pos_data, pos_sig);

  assert (table);
  assert (message);

  /* First we copy the actual signature into a memory buffer.  Such a
     signature is expected to be samll enough to be readable directly
     (i.e.less that 16k as suggested by the MS MAPI docs). */
  hr = message->OpenAttach (pos_sig, NULL, MAPI_BEST_ACCESS, &att);	
  if (FAILED (hr))
    {
      log_error ("%s:%s: can't open attachment %d (sig): hr=%#lx",
                 __FILE__, __func__, pos_sig, hr);
      return;
    }

  if ( table[pos_sig].method == ATTACH_BY_VALUE )
    sig_data = get_short_attach_data (att);
  else
    {
      log_error ("%s:%s: attachment %d (sig): method %d not supported",
                 __FILE__, __func__, pos_sig, table[pos_sig].method);
      att->Release ();
      return;
    }
  att->Release ();
  if (!sig_data)
    return; /* Problem getting signature; error has already been
               logged. */

  /* Now get on with the actual signed data. */
  hr = message->OpenAttach (pos_data, NULL, MAPI_BEST_ACCESS, &att);	
  if (FAILED (hr))
    {
      log_error ("%s:%s: can't open attachment %d (data): hr=%#lx",
                 __FILE__, __func__, pos_data, hr);
      xfree (sig_data);
      return;
    }

  if ( table[pos_data].method == ATTACH_BY_VALUE )
    {
      LPSTREAM stream;

      hr = att->OpenProperty (PR_ATTACH_DATA_BIN, &IID_IStream, 
                              0, 0, (LPUNKNOWN*) &stream);
      if (FAILED (hr))
        {
          log_error ("%s:%s: can't open data of attachment %d: hr=%#lx",
                     __FILE__, __func__, pos_data, hr);
          goto leave;
        }
      err = op_verify_detached_sig (stream, sig_data,
                                    table[pos_data].filename);
      if (err)
        {
          log_debug ("%s:%s: verify detached signature failed: %s",
                     __FILE__, __func__, op_strerror (err)); 
          MessageBox (hwnd, op_strerror (err),
                      "GPG Attachment Verification", MB_ICONERROR|MB_OK);
        }
      stream->Release ();
    }
  else
    {
      log_error ("%s:%s: attachment %d (data): method %d not supported",
                 __FILE__, __func__, pos_data, table[pos_data].method);
    }

 leave:
  /* Close this attachment. */
  xfree (sig_data);
  att->Release ();
}


/* Decrypt the attachment with the internal number POS.
   SAVE_PLAINTEXT must be true to save the attachemnt; displaying a
   attachment is not yet supported.  If FILENAME is not NULL it will
   be displayed along with status outputs. */
void
GpgMsgImpl::decryptAttachment (HWND hwnd, int pos, bool save_plaintext,
                               int ttl, const char *filename)
{    
  HRESULT hr;
  LPATTACH att;
  int method, err;

  log_debug ("%s:%s: processing attachment %d", __FILE__, __func__, pos);

  /* Make sure that we can access the attachment table. */
  if (!message || !getAttachments ())
    {
      log_debug ("%s:%s: no attachemnts at all", __FILE__, __func__);
      return;
    }

  if (!save_plaintext)
    {
      log_error ("%s:%s: save_plaintext not requested", __FILE__, __func__);
      return;
    }

  hr = message->OpenAttach (pos, NULL, MAPI_BEST_ACCESS, &att);	
  if (FAILED (hr))
    {
      log_debug ("%s:%s: can't open attachment %d: hr=%#lx",
                 __FILE__, __func__, pos, hr);
      return;
    }

  method = get_attach_method (att);
  if ( method == ATTACH_EMBEDDED_MSG)
    {
      /* This is an embedded message.  The orginal G-DATA plugin
         decrypted the message and then updated the attachemnt;
         i.e. stored the plaintext.  This seemed to ensure that the
         attachemnt message was properly displayed.  I am not sure
         what we should do - it might be necessary to have a callback
         to allow displaying the attachment.  Needs further
         experiments. */
      LPMESSAGE emb;
      
      hr = att->OpenProperty (PR_ATTACH_DATA_OBJ, &IID_IMessage, 0, 
                              MAPI_MODIFY, (LPUNKNOWN*)&emb);
      if (FAILED (hr))
        {
          log_error ("%s:%s: can't open data obj of attachment %d: hr=%#lx",
                     __FILE__, __func__, pos, hr);
          goto leave;
        }

      //FIXME  Not sure what to do here.  Did it ever work?
      // 	setWindow (hwnd);
      // 	setMessage (emb);
      //if (doCmdAttach (action))
      //  success = FALSE;
      //XXX;
      //emb->SaveChanges (FORCE_SAVE);
      //att->SaveChanges (FORCE_SAVE);
      emb->Release ();
    }
  else if (method == ATTACH_BY_VALUE)
    {
      char *s;
      char *outname;
      char *suggested_name;
      LPSTREAM from, to;

      suggested_name = get_attach_filename (att);
      if (suggested_name)
        log_debug ("%s:%s: attachment %d, filename `%s'", 
                  __FILE__, __func__, pos, suggested_name);
      /* Strip of know extensions or use a default name. */
      if (!suggested_name)
        {
          xfree (suggested_name);
          suggested_name = (char*)xmalloc (50);
          snprintf (suggested_name, 49, "unnamed-%d.dat", pos);
        }
      else if ((s = strrchr (suggested_name, '.'))
               && (!stricmp (s, ".pgp") 
                   || !stricmp (s, ".gpg") 
                   || !stricmp (s, ".asc")) )
        {
          *s = 0;
        }
      outname = get_save_filename (hwnd, suggested_name);
      xfree (suggested_name);

      hr = att->OpenProperty (PR_ATTACH_DATA_BIN, &IID_IStream, 
                              0, 0, (LPUNKNOWN*) &from);
      if (FAILED (hr))
        {
          log_error ("%s:%s: can't open data of attachment %d: hr=%#lx",
                     __FILE__, __func__, pos, hr);
          xfree (outname);
          goto leave;
        }

      hr = OpenStreamOnFile (MAPIAllocateBuffer, MAPIFreeBuffer,
                             (STGM_CREATE | STGM_READWRITE),
                             outname, NULL, &to); 
      if (FAILED (hr)) 
        {
          log_error ("%s:%s: can't create stream for `%s': hr=%#lx\n",
                     __FILE__, __func__, outname, hr); 
          from->Release ();
          xfree (outname);
          goto leave;
        }
      
      err = op_decrypt_stream (from, to, ttl, filename);
      if (err)
        {
          log_debug ("%s:%s: decrypt stream failed: %s",
                     __FILE__, __func__, op_strerror (err)); 
          to->Revert ();
          to->Release ();
          from->Release ();
          MessageBox (hwnd, op_strerror (err),
                      "GPG Attachment Decryption", MB_ICONERROR|MB_OK);
          /* FIXME: We might need to delete outname now.  However a
             sensible implementation of the stream object should have
             done it trhough the Revert call. */
          xfree (outname);
          goto leave;
        }
        
      to->Commit (0);
      to->Release ();
      from->Release ();

      /*  Hmmm: Why are we deleting the attachment now????? 
          Disabled until clarified.   FIXME */
      //if (message->DeleteAttach (pos, 0, NULL, 0) == S_OK)
      //   show error;

      xfree (outname);
    }
  else
    {
      log_error ("%s:%s: attachment %d: method %d not supported",
                 __FILE__, __func__, pos, method);
    }

 leave:
  /* Close this attachment. */
  att->Release ();
}


/* Sign the attachment with the internal number POS.  TTL is caching
   time for a required passphrase. */
void
GpgMsgImpl::signAttachment (HWND hwnd, int pos, gpgme_key_t sign_key, int ttl)
{    
  HRESULT hr;
  LPATTACH att;
  int method, err;
  LPSTREAM from = NULL;
  LPSTREAM to = NULL;
  char *signame = NULL;
  LPATTACH newatt = NULL;

  /* Make sure that we can access the attachment table. */
  if (!message || !getAttachments ())
    {
      log_debug ("%s:%s: no attachemnts at all", __FILE__, __func__);
      return;
    }

  hr = message->OpenAttach (pos, NULL, MAPI_BEST_ACCESS, &att);	
  if (FAILED (hr))
    {
      log_debug ("%s:%s: can't open attachment %d: hr=%#lx",
                 __FILE__, __func__, pos, hr);
      return;
    }

  /* Construct a filename for the new attachment. */
  {
    char *tmpname = get_attach_filename (att);
    if (!tmpname)
      {
        signame = (char*)xmalloc (70);
        snprintf (signame, 70, "gpg-signature-%d.asc", pos);
      }
    else
      {
        signame = (char*)xmalloc (strlen (tmpname) + 4 + 1);
        strcpy (stpcpy (signame, tmpname), ".asc");
        xfree (tmpname);
      }
  }

  method = get_attach_method (att);
  if (method == ATTACH_EMBEDDED_MSG)
    {
      log_debug ("%s:%s: signing embedded attachments is not supported",
                 __FILE__, __func__);
    }
  else if (method == ATTACH_BY_VALUE)
    {
      ULONG newpos;
      SPropValue prop;

      hr = att->OpenProperty (PR_ATTACH_DATA_BIN, &IID_IStream, 
                              0, 0, (LPUNKNOWN*)&from);
      if (FAILED (hr))
        {
          log_error ("%s:%s: can't open data of attachment %d: hr=%#lx",
                     __FILE__, __func__, pos, hr);
          goto leave;
        }

      hr = message->CreateAttach (NULL, 0, &newpos, &newatt);
      if (hr != S_OK)
        {
          log_error ("%s:%s: can't create attachment: hr=%#lx\n",
                     __FILE__, __func__, hr); 
          goto leave;
        }

      prop.ulPropTag = PR_ATTACH_METHOD;
      prop.Value.ul = ATTACH_BY_VALUE;
      hr = HrSetOneProp (newatt, &prop);
      if (hr != S_OK)
        {
          log_error ("%s:%s: can't set attach method: hr=%#lx\n",
                     __FILE__, __func__, hr); 
          goto leave;
        }
      
      prop.ulPropTag = PR_ATTACH_LONG_FILENAME_A;
      prop.Value.lpszA = signame;   
      hr = HrSetOneProp (newatt, &prop);
      if (hr != S_OK)
        {
          log_error ("%s:%s: can't set attach filename: hr=%#lx\n",
                     __FILE__, __func__, hr); 
          goto leave;
        }
      log_debug ("%s:%s: setting filename of attachment %d/%ld to `%s'",
                 __FILE__, __func__, pos, newpos, signame);

      prop.ulPropTag = PR_ATTACH_EXTENSION_A;
      prop.Value.lpszA = ".pgpsig";   
      hr = HrSetOneProp (newatt, &prop);
      if (hr != S_OK)
        {
          log_error ("%s:%s: can't set attach extension: hr=%#lx\n",
                     __FILE__, __func__, hr); 
          goto leave;
        }

      prop.ulPropTag = PR_ATTACH_TAG;
      prop.Value.bin.cb  = sizeof oid_mimetag;
      prop.Value.bin.lpb = (LPBYTE)oid_mimetag;
      hr = HrSetOneProp (newatt, &prop);
      if (hr != S_OK)
        {
          log_error ("%s:%s: can't set attach tag: hr=%#lx\n",
                     __FILE__, __func__, hr); 
          goto leave;
        }

      prop.ulPropTag = PR_ATTACH_MIME_TAG_A;
      prop.Value.lpszA = "application/pgp-signature";
      hr = HrSetOneProp (newatt, &prop);
      if (hr != S_OK)
        {
          log_error ("%s:%s: can't set attach mime tag: hr=%#lx\n",
                     __FILE__, __func__, hr); 
          goto leave;
        }

      hr = newatt->OpenProperty (PR_ATTACH_DATA_BIN, &IID_IStream, 0,
                                 MAPI_CREATE | MAPI_MODIFY, (LPUNKNOWN*)&to);
      if (FAILED (hr)) 
        {
          log_error ("%s:%s: can't create output stream: hr=%#lx\n",
                     __FILE__, __func__, hr); 
          goto leave;
        }
      
      err = op_sign_stream (from, to, OP_SIG_DETACH, sign_key, ttl);
      if (err)
        {
          log_debug ("%s:%s: sign stream failed: %s",
                     __FILE__, __func__, op_strerror (err)); 
          to->Revert ();
          MessageBox (hwnd, op_strerror (err),
                      "GPG Attachment Signing", MB_ICONERROR|MB_OK);
          goto leave;
        }
      from->Release ();
      from = NULL;
      to->Commit (0);
      to->Release ();
      to = NULL;

      hr = newatt->SaveChanges (0);
      if (hr != S_OK)
        {
          log_error ("%s:%s: SaveChanges failed: hr=%#lx\n",
                     __FILE__, __func__, hr); 
          goto leave;
        }

    }
  else
    {
      log_error ("%s:%s: attachment %d: method %d not supported",
                 __FILE__, __func__, pos, method);
    }

 leave:
  if (from)
    from->Release ();
  if (to)
    to->Release ();
  xfree (signame);
  if (newatt)
    newatt->Release ();

  att->Release ();
}

/* Encrypt the attachment with the internal number POS.  KEYS is a
   NULL terminates array with recipients to whom the message should be
   encrypted.  If SIGN_KEY is not NULL the attachment will also get
   signed. TTL is the passphrase caching time and only used if
   SIGN_KEY is not NULL. Returns 0 on success. */
int
GpgMsgImpl::encryptAttachment (HWND hwnd, int pos, gpgme_key_t *keys,
                               gpgme_key_t sign_key, int ttl)
{    
  HRESULT hr;
  LPATTACH att;
  int method, err;
  LPSTREAM from = NULL;
  LPSTREAM to = NULL;
  char *filename = NULL;
  LPATTACH newatt = NULL;

  /* Make sure that we can access the attachment table. */
  if (!message || !getAttachments ())
    {
      log_debug ("%s:%s: no attachemnts at all", __FILE__, __func__);
      return 0;
    }

  hr = message->OpenAttach (pos, NULL, MAPI_BEST_ACCESS, &att);	
  if (FAILED (hr))
    {
      log_debug ("%s:%s: can't open attachment %d: hr=%#lx",
                 __FILE__, __func__, pos, hr);
      err = gpg_error (GPG_ERR_GENERAL);
      return err;
    }

  /* Construct a filename for the new attachment. */
  {
    char *tmpname = get_attach_filename (att);
    if (!tmpname)
      {
        filename = (char*)xmalloc (70);
        snprintf (filename, 70, "gpg-encrypted-%d.pgp", pos);
      }
    else
      {
        filename = (char*)xmalloc (strlen (tmpname) + 4 + 1);
        strcpy (stpcpy (filename, tmpname), ".pgp");
        xfree (tmpname);
      }
  }

  method = get_attach_method (att);
  if (method == ATTACH_EMBEDDED_MSG)
    {
      log_debug ("%s:%s: encrypting embedded attachments is not supported",
                 __FILE__, __func__);
      err = gpg_error (GPG_ERR_NOT_SUPPORTED);
    }
  else if (method == ATTACH_BY_VALUE)
    {
      ULONG newpos;
      SPropValue prop;

      hr = att->OpenProperty (PR_ATTACH_DATA_BIN, &IID_IStream, 
                              0, 0, (LPUNKNOWN*)&from);
      if (FAILED (hr))
        {
          log_error ("%s:%s: can't open data of attachment %d: hr=%#lx",
                     __FILE__, __func__, pos, hr);
          err = gpg_error (GPG_ERR_GENERAL);
          goto leave;
        }

      hr = message->CreateAttach (NULL, 0, &newpos, &newatt);
      if (hr != S_OK)
        {
          log_error ("%s:%s: can't create attachment: hr=%#lx\n",
                     __FILE__, __func__, hr); 
          err = gpg_error (GPG_ERR_GENERAL);
          goto leave;
        }

      prop.ulPropTag = PR_ATTACH_METHOD;
      prop.Value.ul = ATTACH_BY_VALUE;
      hr = HrSetOneProp (newatt, &prop);
      if (hr != S_OK)
        {
          log_error ("%s:%s: can't set attach method: hr=%#lx\n",
                     __FILE__, __func__, hr); 
          err = gpg_error (GPG_ERR_GENERAL);
          goto leave;
        }
      
      prop.ulPropTag = PR_ATTACH_LONG_FILENAME_A;
      prop.Value.lpszA = filename;   
      hr = HrSetOneProp (newatt, &prop);
      if (hr != S_OK)
        {
          log_error ("%s:%s: can't set attach filename: hr=%#lx\n",
                     __FILE__, __func__, hr); 
          err = gpg_error (GPG_ERR_GENERAL);
          goto leave;
        }
      log_debug ("%s:%s: setting filename of attachment %d/%ld to `%s'",
                 __FILE__, __func__, pos, newpos, filename);

      prop.ulPropTag = PR_ATTACH_EXTENSION_A;
      prop.Value.lpszA = ".pgpenc";   
      hr = HrSetOneProp (newatt, &prop);
      if (hr != S_OK)
        {
          log_error ("%s:%s: can't set attach extension: hr=%#lx\n",
                     __FILE__, __func__, hr); 
          err = gpg_error (GPG_ERR_GENERAL);
          goto leave;
        }

      prop.ulPropTag = PR_ATTACH_TAG;
      prop.Value.bin.cb  = sizeof oid_mimetag;
      prop.Value.bin.lpb = (LPBYTE)oid_mimetag;
      hr = HrSetOneProp (newatt, &prop);
      if (hr != S_OK)
        {
          log_error ("%s:%s: can't set attach tag: hr=%#lx\n",
                     __FILE__, __func__, hr); 
          err = gpg_error (GPG_ERR_GENERAL);
          goto leave;
        }

      prop.ulPropTag = PR_ATTACH_MIME_TAG_A;
      prop.Value.lpszA = "application/pgp-encrypted;foomode=baz";
      hr = HrSetOneProp (newatt, &prop);
      if (hr != S_OK)
        {
          log_error ("%s:%s: can't set attach mime tag: hr=%#lx\n",
                     __FILE__, __func__, hr); 
          err = gpg_error (GPG_ERR_GENERAL);
          goto leave;
        }

      hr = newatt->OpenProperty (PR_ATTACH_DATA_BIN, &IID_IStream, 0,
                                 MAPI_CREATE | MAPI_MODIFY, (LPUNKNOWN*)&to);
      if (FAILED (hr)) 
        {
          log_error ("%s:%s: can't create output stream: hr=%#lx\n",
                     __FILE__, __func__, hr); 
          err = gpg_error (GPG_ERR_GENERAL);
          goto leave;
        }
      
      err = op_encrypt_stream (from, to, keys, sign_key, ttl);
      if (err)
        {
          log_debug ("%s:%s: encrypt stream failed: %s",
                     __FILE__, __func__, op_strerror (err)); 
          to->Revert ();
          MessageBox (hwnd, op_strerror (err),
                      "GPG Attachment Encryption", MB_ICONERROR|MB_OK);
          goto leave;
        }
      from->Release ();
      from = NULL;
      to->Commit (0);
      to->Release ();
      to = NULL;

      hr = newatt->SaveChanges (0);
      if (hr != S_OK)
        {
          log_error ("%s:%s: SaveChanges failed: hr=%#lx\n",
                     __FILE__, __func__, hr); 
          err = gpg_error (GPG_ERR_GENERAL);
          goto leave;
        }

      hr = message->DeleteAttach (pos, 0, NULL, 0);
      if (hr != S_OK)
        {
          log_error ("%s:%s: DeleteAtatch failed: hr=%#lx\n",
                     __FILE__, __func__, hr); 
          err = gpg_error (GPG_ERR_GENERAL);
          goto leave;
        }

    }
  else
    {
      log_error ("%s:%s: attachment %d: method %d not supported",
                 __FILE__, __func__, pos, method);
      err = gpg_error (GPG_ERR_NOT_SUPPORTED);
    }

 leave:
  if (from)
    from->Release ();
  if (to)
    to->Release ();
  xfree (filename);
  if (newatt)
    newatt->Release ();

  att->Release ();
  return err;
}
