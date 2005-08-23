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

/*
   The implementation class of MapiGPGME.  
 */
class GpgMsgImpl : public GpgMsg
{
public:    
  GpgMsgImpl () 
  {
    this->message = NULL;
    this->body = NULL;
    this->body_plain = NULL;
    this->body_cipher = NULL;
    this->body_signed = NULL;
    this->body_cipher_is_html = false;
  }

  ~GpgMsgImpl ()
  {
    if (message)
      message->Release ();
    xfree (body);
    xfree (body_plain);
    xfree (body_cipher);
    xfree (body_signed);
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


private:
  LPMESSAGE message;  /* Pointer to the message. */
  char *body;         /* utf-8 encoded body string or NULL. */
  char *body_plain;   /* Plaintext version of BODY or NULL. */
  char *body_cipher;  /* Enciphered version of BODY or NULL. */
  char *body_signed;  /* Signed version of BODY or NULL. */
  bool body_cipher_is_html; /* Indicating that BODY_CIPHER holds HTML. */
  
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
GpgMsgImpl::hasAttachments ()
{
//   if (!attachRows)
//     getAttachments ();

//   bool has = attachRows->cRows > 0? true : false;
//     freeAttachments ();
//     return has;
  return false;
}



