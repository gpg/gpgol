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


#include "gpgmsg.hh"
#include "util.h"
#include "msgcache.h"
#include "engine.h"

static const char oid_mimetag[] =
  {0x2A, 0x86, 0x48, 0x86, 0xf7, 0x14, 0x03, 0x0a, 0x04};


/* The string used as the standard XXXXX of decrypted attachments. */
#define ATT_FILE_PREFIX ".pgpenc"

#define TRACEPOINT() do { log_debug ("%s:%s:%d: tracepoint\n", \
                                       __FILE__, __func__, __LINE__); \
                        } while (0)



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

    attach.table = NULL;
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

    if (attach.table)
      {
        attach.table->Release ();
        attach.table = NULL;
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
      log_debug ("%s:%s:%d: here\n", __FILE__, __func__, __LINE__);
        msg->AddRef ();
      log_debug ("%s:%s:%d: here\n", __FILE__, __func__, __LINE__);
        message = msg;
      }
  }
  
  openpgp_t getMessageType (void);
  bool hasAttachments (void);
  const char *getOrigText (void);
  const char *GpgMsgImpl::getDisplayText (void);
  const char *getPlainText (void);
  void setPlainText (char *string);
  void setCipherText (char *string, bool html);
  void setSignedText (char *string);
  void saveChanges (bool permanent);
  bool matchesString (const char *string);
  char **getRecipients (void);
  unsigned int getAttachments (void);
  void decryptAttachment (HWND hwnd, int pos, bool save_plaintext, int ttl);
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
    LPMAPITABLE table;/* The loaded attachment table or NULL. */
    LPSRowSet   rows; /* The retrieved set of rows from the table. */
  } attach;
  
  void loadBody (void);
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
      log_debug_w32 (hr, "%s:%s: OpenProperty failed", __FILE__, __func__);
      return;
    }

  hr = stream->Stat (&statInfo, STATFLAG_NONAME);
  if ( hr != S_OK )
    {
      log_debug_w32 (hr, "%s:%s: Stat failed", __FILE__, __func__);
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
      log_debug_w32 (hr, "%s:%s: Read failed", __FILE__, __func__);
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

/* Save STRING as the ciphertext version of the message.  WARNING:
   ownership of STRING is transferred to this object. HTML indicates
   whether the ciphertext was originally HTML. */
void
GpgMsgImpl::setCipherText (char *string, bool html)
{
  xfree (body_cipher);
  body_cipher = string;
  body_cipher_is_html = html;
}

/* Save STRING as the signed version of the message.  WARNING:
   ownership of STRING is transferred to this object. */
void
GpgMsgImpl::setSignedText (char *string)
{
  xfree (body_signed);
  body_signed = string;
}

/* Save the changes made to the message.  With PERMANENT set to true
   they are really stored, when not set they are only saved
   temporary. */
void
GpgMsgImpl::saveChanges (bool permanent)
{
  SPropValue sProp; 
  HRESULT hr;
  int rc = TRUE;

  if (!body_plain)
    return; /* Nothing to save. */

  if (!permanent)
    return;
  
  /* Make sure that the Plaintext and the Richtext are in sync. */
//   if (message)
//     {
//       BOOL changed;

//       sProp.ulPropTag = PR_BODY_A;
//       sProp.Value.lpszA = "";
//       hr = HrSetOneProp(message, &sProp);
//       changed = false;
//       RTFSync(message, RTF_SYNC_BODY_CHANGED, &changed);
//       sProp.Value.lpszA = body_plain;
//       hr = HrSetOneProp(message, &sProp);
//       RTFSync(message, RTF_SYNC_BODY_CHANGED, &changed);
//     }

  sProp.ulPropTag = PR_BODY_W;
  sProp.Value.lpszW = utf8_to_wchar (body_plain);
  if (!sProp.Value.lpszW)
    {
      log_debug_w32 (-1, "%s:%s: error converting from utf8\n",
                     __FILE__, __func__);
      return;
    }
  hr = HrSetOneProp (message, &sProp);
  xfree (sProp.Value.lpszW);
  if (hr < 0)
    log_debug_w32 (-1, "%s:%s: HrSetOneProp failed", __FILE__, __func__);
  else
    {
      log_debug ("%s:%s: PR_BODY set to `%s'\n",
                 __FILE__, __func__, body_plain);
      {
        GpgMsg *xmsg = CreateGpgMsg (message);
        log_debug ("%s:%s:    cross check `%s'\n",
                   __FILE__, __func__, xmsg->getOrigText ());
        delete xmsg;
      }
      if (permanent && message)
        {
          hr = message->SaveChanges (KEEP_OPEN_READWRITE|FORCE_SAVE);
          if (hr < 0)
            log_debug_w32 (-1, "%s:%s: SaveChanges failed",
                           __FILE__, __func__);
        }
    }

  log_debug ("%s:%s: leave\n", __FILE__, __func__);
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

  if (!attach.table)
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
      attach.table = table;
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

  if (GetSaveFileName (&ofn))
    return xstrdup (fname);
  return NULL;
}


/* Decrypt the attachment with the internal number POS.
   SAVE_PLAINTEXT must be true to save the attachemnt; displaying a
   attachment is not yet supported. */
void
GpgMsgImpl::decryptAttachment (HWND hwnd, int pos, bool save_plaintext,
                               int ttl)
{    
  HRESULT hr;
  LPATTACH att;
  int method, err;

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
      const char *s;
      char *outname;
      char *suggested_name;
      LPSTREAM from, to;

      suggested_name = get_attach_filename (att);
      /* We only want to automatically decrypt attachmenst with
         certain extensions.  FIXME: Also look for content-types. */
      if (!suggested_name 
          || !(s = strrchr (suggested_name, '.'))
          || (stricmp (s, ".pgp") 
              && stricmp (s, ".gpg") 
              && stricmp (s, ".asc")))
        {
          log_debug ( "%s:%s: attachment %d has no pgp extension\n", 
                      __FILE__, __func__);
          goto leave;
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
      
      err = op_decrypt_stream (from, to, ttl);
      if (err)
        {
          log_debug ("%s:%s: decrypt stream failed: %s",
                     __FILE__, __func__, op_strerror (err)); 
          to->Revert ();
          to->Release ();
          from->Release ();
          MessageBox (NULL, op_strerror (err),
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
          MessageBox (NULL, op_strerror (err),
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
          MessageBox (NULL, op_strerror (err),
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
