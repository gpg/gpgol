/* mapihelp.cpp - Helper functions for MAPI
 *	Copyright (C) 2005, 2007, 2008 g10 Code GmbH
 * 
 * This file is part of GpgOL.
 * 
 * GpgOL is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 * 
 * GpgOL is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Lesser General Public License for more details.
 * 
 * You should have received a copy of the GNU Lesser General Public License
 * along with this program; if not, see <http://www.gnu.org/licenses/>.
 */

#ifdef HAVE_CONFIG_H
#include <config.h>
#endif

#include <ctype.h>
#include <windows.h>

#include "mymapi.h"
#include "mymapitags.h"
#include "common.h"
#include "rfc822parse.h"
#include "serpent.h"
#include "mapihelp.h"
#include "parsetlv.h"
#include "gpgolstr.h"

#ifndef CRYPT_E_STREAM_INSUFFICIENT_DATA
#define CRYPT_E_STREAM_INSUFFICIENT_DATA 0x80091011
#endif
#ifndef CRYPT_E_ASN1_BADTAG
#define CRYPT_E_ASN1_BADTAG 0x8009310B
#endif


#define TRACEPOINT() do { log_debug ("%s:%s:%d: tracepoint\n", \
                                     SRCNAME, __func__, __LINE__); \
                        } while (0)


static int get_attach_method (LPATTACH obj);
static int has_smime_filename (LPATTACH obj);
static char *get_attach_mime_tag (LPATTACH obj);




/* Print a MAPI property to the log stream. */
void
log_mapi_property (LPMESSAGE message, ULONG prop, const char *propname)
{
  HRESULT hr;
  LPSPropValue propval = NULL;
  size_t keylen;
  void *key;
  char *buf;

  if (!message)
    return; /* No message: Nop. */

  hr = HrGetOneProp ((LPMAPIPROP)message, prop, &propval);
  if (FAILED (hr))
    {
      log_debug ("%s:%s: HrGetOneProp(%s) failed: hr=%#lx\n",
                 SRCNAME, __func__, propname, hr);
      return;
    }
    
  switch ( PROP_TYPE (propval->ulPropTag) )
    {
    case PT_BINARY:
      keylen = propval->Value.bin.cb;
      key = propval->Value.bin.lpb;
      log_hexdump (key, keylen, "%s: %20s=", __func__, propname);
      break;

    case PT_UNICODE:
      buf = wchar_to_utf8 (propval->Value.lpszW);
      if (!buf)
        log_debug ("%s:%s: error converting to utf8\n", SRCNAME, __func__);
      else
        log_debug ("%s: %20s=`%s'", __func__, propname, buf);
      xfree (buf);
      break;
      
    case PT_STRING8:
      log_debug ("%s: %20s=`%s'", __func__, propname, propval->Value.lpszA);
      break;

    case PT_LONG:
      log_debug ("%s: %20s=%ld", __func__, propname, propval->Value.l);
      break;

    default:
      log_debug ("%s:%s: HrGetOneProp(%s) property type %lu not supported\n",
                 SRCNAME, __func__, propname,
                 PROP_TYPE (propval->ulPropTag) );
      return;
    }
  MAPIFreeBuffer (propval);
}


/* Helper to create a named property. */
static ULONG 
create_gpgol_tag (LPMESSAGE message, const wchar_t *name, const char *func)
{
  HRESULT hr;
  LPSPropTagArray proparr = NULL;
  MAPINAMEID mnid, *pmnid;
  GpgOLStr propname(name);
  /* {31805ab8-3e92-11dc-879c-00061b031004}: GpgOL custom properties.  */
  GUID guid = {0x31805ab8, 0x3e92, 0x11dc, {0x87, 0x9c, 0x00, 0x06,
                                            0x1b, 0x03, 0x10, 0x04}};
  ULONG result;
  
  memset (&mnid, 0, sizeof mnid);
  mnid.lpguid = &guid;
  mnid.ulKind = MNID_STRING;
  mnid.Kind.lpwstrName = propname;
  pmnid = &mnid;
  hr = message->GetIDsFromNames (1, &pmnid, MAPI_CREATE, &proparr);
  if (FAILED (hr))
    proparr = NULL;
  if (FAILED (hr) || !(proparr->aulPropTag[0] & 0xFFFF0000) ) 
    {
      log_error ("%s:%s: can't map GpgOL property: hr=%#lx\n",
                 SRCNAME, func, hr); 
      result = 0;
    }
  else
    result = (proparr->aulPropTag[0] & 0xFFFF0000);
  if (proparr)
    MAPIFreeBuffer (proparr);
    
  return result;
}


/* Return the property tag for GpgOL Msg Class. */
int 
get_gpgolmsgclass_tag (LPMESSAGE message, ULONG *r_tag)
{
  if (!(*r_tag = create_gpgol_tag (message, L"GpgOL Msg Class", __func__)))
    return -1;
  *r_tag |= PT_STRING8;
  return 0;
}

/* Return the property tag for GpgOL Old Msg Class.  The Old Msg Class
   saves the message class as seen before we changed it the first
   time. */
int 
get_gpgololdmsgclass_tag (LPMESSAGE message, ULONG *r_tag)
{
  if (!(*r_tag = create_gpgol_tag (message, L"GpgOL Old Msg Class", __func__)))
    return -1;
  *r_tag |= PT_STRING8;
  return 0;
}


/* Return the property tag for GpgOL Attach Type. */
int 
get_gpgolattachtype_tag (LPMESSAGE message, ULONG *r_tag)
{
  if (!(*r_tag = create_gpgol_tag (message, L"GpgOL Attach Type", __func__)))
    return -1;
  *r_tag |= PT_LONG;
  return 0;
}


/* Return the property tag for GpgOL Sig Status. */
int 
get_gpgolsigstatus_tag (LPMESSAGE message, ULONG *r_tag)
{
  if (!(*r_tag = create_gpgol_tag (message, L"GpgOL Sig Status", __func__)))
    return -1;
  *r_tag |= PT_STRING8;
  return 0;
}


/* Return the property tag for GpgOL Protect IV. */
int 
get_gpgolprotectiv_tag (LPMESSAGE message, ULONG *r_tag)
{
  if (!(*r_tag = create_gpgol_tag (message, L"GpgOL Protect IV", __func__)))
    return -1;
  *r_tag |= PT_BINARY;
  return 0;
}

/* Return the property tag for GpgOL Last Decrypted. */
int 
get_gpgollastdecrypted_tag (LPMESSAGE message, ULONG *r_tag)
{
  if (!(*r_tag = create_gpgol_tag (message, L"GpgOL Last Decrypted",__func__)))
    return -1;
  *r_tag |= PT_BINARY;
  return 0;
}


/* Return the property tag for GpgOL MIME structure. */
int 
get_gpgolmimeinfo_tag (LPMESSAGE message, ULONG *r_tag)
{
  if (!(*r_tag = create_gpgol_tag (message, L"GpgOL MIME Info", __func__)))
    return -1;
  *r_tag |= PT_STRING8;
  return 0;
}


/* Return the property tag for GpgOL Charset. */
int 
get_gpgolcharset_tag (LPMESSAGE message, ULONG *r_tag)
{
  if (!(*r_tag = create_gpgol_tag (message, L"GpgOL Charset", __func__)))
    return -1;
  *r_tag |= PT_STRING8;
  return 0;
}


/* Return the property tag for GpgOL Draft Info.  */
int 
get_gpgoldraftinfo_tag (LPMESSAGE message, ULONG *r_tag)
{
  if (!(*r_tag = create_gpgol_tag (message, L"GpgOL Draft Info", __func__)))
    return -1;
  *r_tag |= PT_STRING8;
  return 0;
}


/* Return the tag of the Internet Charset Body property which seems to
   hold the PR_BODY as received and thus before charset
   conversion.  */
int
get_internetcharsetbody_tag (LPMESSAGE message, ULONG *r_tag)
{
  HRESULT hr;
  LPSPropTagArray proparr = NULL;
  MAPINAMEID mnid, *pmnid;	
  /* {4E3A7680-B77A-11D0-9DA5-00C04FD65685} */
  GUID guid = {0x4E3A7680, 0xB77A, 0x11D0, {0x9D, 0xA5, 0x00, 0xC0,
                                            0x4F, 0xD6, 0x56, 0x85}};
  GpgOLStr propname (L"Internet Charset Body");
  int result;

  memset (&mnid, 0, sizeof mnid);
  mnid.lpguid = &guid;
  mnid.ulKind = MNID_STRING;
  mnid.Kind.lpwstrName = propname;
  pmnid = &mnid;
  hr = message->GetIDsFromNames (1, &pmnid, 0, &proparr);
  if (FAILED (hr))
    proparr = NULL;
  if (FAILED (hr) || !(proparr->aulPropTag[0] & 0xFFFF0000) ) 
    {
      log_error ("%s:%s: can't get the Internet Charset Body property:"
                 " hr=%#lx\n", SRCNAME, __func__, hr); 
      result = -1;
    }
  else
    {
      result = 0;
      *r_tag = ((proparr->aulPropTag[0] & 0xFFFF0000) | PT_BINARY);
    }

  if (proparr)
    MAPIFreeBuffer (proparr);
  
  return result;
}


/* A Wrapper around the SaveChanges method.  This function should be
   called indirect through the mapi_save_changes macro.  Returns 0 on
   success. */
int
mapi_do_save_changes (LPMESSAGE message, ULONG flags, int only_del_body,
                      const char *dbg_file, const char *dbg_func)
{
  HRESULT hr;
  SPropTagArray proparray;
  int any = 0;
  
  if (mapi_has_last_decrypted (message))
    {
      proparray.cValues = 1;
      proparray.aulPropTag[0] = PR_BODY;
      hr = message->DeleteProps (&proparray, NULL);
      if (hr)
        log_debug_w32 (hr, "%s:%s: deleting PR_BODY failed",
                       log_srcname (dbg_file), dbg_func);
      else
        any = 1;

      proparray.cValues = 1;
      proparray.aulPropTag[0] = PR_BODY_HTML;
      hr = message->DeleteProps (&proparray, NULL);
      if (hr)
        log_debug_w32 (hr, "%s:%s: deleting PR_BODY_HTML failed",
                       log_srcname (dbg_file), dbg_func);
      else
        any = 1;
    }

  if (!only_del_body || any)
    {
      hr = message->SaveChanges (flags);
      if (hr)
        {
          log_error ("%s:%s: SaveChanges(%lu) failed: hr=%#lx\n",
                     log_srcname (dbg_file), dbg_func,
                     (unsigned long)flags, hr); 
          return -1;
        }
    }
  
  return 0;
}


/* Set an arbitary header in the message MSG with NAME to the value
   VAL. */
int
mapi_set_header (LPMESSAGE msg, const char *name, const char *val)
{  
  HRESULT hr;
  LPSPropTagArray pProps = NULL;
  SPropValue pv;
  MAPINAMEID mnid, *pmnid;	
  /* {00020386-0000-0000-C000-000000000046}  ->  GUID For X-Headers */
  GUID guid = {0x00020386, 0x0000, 0x0000, {0xC0, 0x00, 0x00, 0x00,
                                            0x00, 0x00, 0x00, 0x46} };
  int result;

  if (!msg)
    return -1;

  memset (&mnid, 0, sizeof mnid);
  mnid.lpguid = &guid;
  mnid.ulKind = MNID_STRING;
  mnid.Kind.lpwstrName = utf8_to_wchar (name);
  pmnid = &mnid;
  hr = msg->GetIDsFromNames (1, &pmnid, MAPI_CREATE, &pProps);
  xfree (mnid.Kind.lpwstrName);
  if (FAILED (hr)) 
    {
      pProps = NULL;
      log_error ("%s:%s: can't get mapping for header `%s': hr=%#lx\n",
                 SRCNAME, __func__, name, hr); 
      result = -1;
    }
  else
    {
      pv.ulPropTag = (pProps->aulPropTag[0] & 0xFFFF0000) | PT_STRING8;
      pv.Value.lpszA = (char *)val;
      hr = HrSetOneProp(msg, &pv);	
      if (hr)
        {
          log_error ("%s:%s: can't set header `%s': hr=%#lx\n",
                     SRCNAME, __func__, name, hr); 
          result = -1;
        }
      else
        result = 0;
    }

  if (pProps)
    MAPIFreeBuffer (pProps);

  return result;
}



/* Return the body as a new IStream object.  Returns NULL on failure.
   The stream returns the body as an ASCII stream (Use mapi_get_body
   for an UTF-8 value).  */
LPSTREAM
mapi_get_body_as_stream (LPMESSAGE message)
{
  HRESULT hr;
  ULONG tag;
  LPSTREAM stream;

  if (!message)
    return NULL;

  if (!get_internetcharsetbody_tag (message, &tag) )
    {
      /* The store knows about the Internet Charset Body property,
         thus try to get the body from this property if it exists.  */
      
      hr = message->OpenProperty (tag, &IID_IStream, 0, 0, 
                                  (LPUNKNOWN*)&stream);
      if (!hr)
        return stream;

      log_debug ("%s:%s: OpenProperty tag=%lx failed: hr=%#lx",
                 SRCNAME, __func__, tag, hr);
    }

  /* We try to get it as an ASCII body.  If this fails we would either
     need to implement some kind of stream filter to translated to
     utf-8 or read everyting into a memory buffer and [provide an
     istream from that memory buffer.  */
  hr = message->OpenProperty (PR_BODY_A, &IID_IStream, 0, 0, 
                              (LPUNKNOWN*)&stream);
  if (hr)
    {
      log_debug ("%s:%s: OpenProperty failed: hr=%#lx", SRCNAME, __func__, hr);
      return NULL;
    }

  return stream;
}



/* Return the body of the message in an allocated buffer.  The buffer
   is guaranteed to be Nul terminated.  The actual length (ie. the
   strlen()) will be stored at R_NBYTES.  The body will be returned in
   UTF-8 encoding. Returns NULL if no body is available.  */
char *
mapi_get_body (LPMESSAGE message, size_t *r_nbytes)
{
  HRESULT hr;
  LPSPropValue lpspvFEID = NULL;
  LPSTREAM stream;
  STATSTG statInfo;
  ULONG nread;
  char *body = NULL;

  if (r_nbytes)
    *r_nbytes = 0;
  hr = HrGetOneProp ((LPMAPIPROP)message, PR_BODY, &lpspvFEID);
  if (SUCCEEDED (hr))  /* Message is small enough to be retrieved directly. */
    { 
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
    }
  else /* Message is large; use an IStream to read it.  */
    {
      hr = message->OpenProperty (PR_BODY, &IID_IStream, 0, 0, 
                                  (LPUNKNOWN*)&stream);
      if (hr)
        {
          log_debug ("%s:%s: OpenProperty failed: hr=%#lx",
                     SRCNAME, __func__, hr);
          return NULL;
        }
      
      hr = stream->Stat (&statInfo, STATFLAG_NONAME);
      if (hr)
        {
          log_debug ("%s:%s: Stat failed: hr=%#lx", SRCNAME, __func__, hr);
          stream->Release ();
          return NULL;
        }
      
      /* Fixme: We might want to read only the first 1k to decide
         whether this is actually an OpenPGP message and only then
         continue reading.  */
      body = (char*)xmalloc ((size_t)statInfo.cbSize.QuadPart + 2);
      hr = stream->Read (body, (size_t)statInfo.cbSize.QuadPart, &nread);
      if (hr)
        {
          log_debug ("%s:%s: Read failed: hr=%#lx", SRCNAME, __func__, hr);
          xfree (body);
          stream->Release ();
          return NULL;
        }
      body[nread] = 0;
      body[nread+1] = 0;
      if (nread != statInfo.cbSize.QuadPart)
        {
          log_debug ("%s:%s: not enough bytes returned\n", SRCNAME, __func__);
          xfree (body);
          stream->Release ();
          return NULL;
        }
      stream->Release ();
      
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
    }

  if (r_nbytes)
    *r_nbytes = strlen (body);
  return body;
}



/* Look at the body of the MESSAGE and try to figure out whether this
   is a supported PGP message.  Returns the new message class or NULL
   if it does not look like a PGP message.  */
static char *
get_msgcls_from_pgp_lines (LPMESSAGE message)
{
  HRESULT hr;
  LPSTREAM stream;
  STATSTG statInfo;
  ULONG nread;
  size_t nbytes;
  char *body = NULL;
  char *p;
  char *msgcls = NULL;
  ULONG tag;
  int   is_binary = 0;
  
  hr = 0;
  if (!get_internetcharsetbody_tag (message, &tag) )
    {
      hr = message->OpenProperty (tag, &IID_IStream, 0, 0, 
                                  (LPUNKNOWN*)&stream);
      if (!hr)
        is_binary = 1;
    }
  if (hr)
    {
      tag = PR_BODY;
      hr = message->OpenProperty (tag, &IID_IStream, 0, 0, 
                                  (LPUNKNOWN*)&stream);
    }
  if (hr)
    {
      log_debug ("%s:%s: OpenProperty(%lx) failed: hr=%#lx",
                 SRCNAME, __func__, tag, hr);
      return NULL;
    }
  
  hr = stream->Stat (&statInfo, STATFLAG_NONAME);
  if (hr)
    {
      log_debug ("%s:%s: Stat failed: hr=%#lx", SRCNAME, __func__, hr);
      stream->Release ();
      return NULL;
    }
  
  /* We read only the first 1k to decide whether this is actually an
     OpenPGP armored message .  */
  nbytes = (size_t)statInfo.cbSize.QuadPart;
  if (nbytes > 1024*2)
    nbytes = 1024*2;
  body = (char*)xmalloc (nbytes + 2);
  hr = stream->Read (body, nbytes, &nread);
  if (hr)
    {
      log_debug ("%s:%s: Read failed: hr=%#lx", SRCNAME, __func__, hr);
      xfree (body);
      stream->Release ();
      return NULL;
    }
  body[nread] = 0;
  body[nread+1] = 0;
  if (nread != nbytes)
    {
      log_debug ("%s:%s: not enough bytes returned\n", SRCNAME, __func__);
      
      xfree (body);
      stream->Release ();
      return NULL;
    }
  stream->Release ();

  if (!is_binary)
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


  /* The first ~1k of the body of the message is now available in the
     utf-8 string BODY.  Walk over it to figure out its type.  */
  for (p=body; p && *p; p = (p=strchr (p+1, '\n')? (p+1):NULL))
    {
      if (!strncmp (p, "-----BEGIN PGP ", 15))
        {
          if (!strncmp (p+15, "SIGNED MESSAGE-----", 19)
              && trailing_ws_p (p+15+19))
            msgcls = xstrdup ("IPM.Note.GpgOL.ClearSigned");
          else if (!strncmp (p+15, "MESSAGE-----", 12)
                   && trailing_ws_p (p+15+12))
            msgcls = xstrdup ("IPM.Note.GpgOL.PGPMessage");
          break;
        }
      else if (!trailing_ws_p (p))
        break;  /* Text before the PGP message - don't take this as a
                   proper message.  */
    }
  

  xfree (body);
  return msgcls;
}


/* Check whether the message is really a CMS encrypted message.  
   We check here whether the message is really encrypted by looking at
   the object identifier inside the CMS data.  Returns:
    -1 := Unknown message type,
     0 := The message is signed,
     1 := The message is encrypted.

   This function is required for two reasons: 

   1. Due to a bug in CryptoEx which sometimes assignes the *.CexEnc
      message class to signed messages and only updates the message
      class after accessing them.  Thus in old stores there may be a
      lot of *.CexEnc message which are actually just signed.
 
   2. If the smime-type parameter is missing we need another way to
      decide whether to decrypt or to verify.

   3. Some messages lack a PR_TRANSPORT_MESSAGE_HEADERS and thus it is
      not possible to deduce the message type from the mail headers.
      This function may be used to identify the message anyway.
 */
static int
is_really_cms_encrypted (LPMESSAGE message)
{    
  HRESULT hr;
  SizedSPropTagArray (1L, propAttNum) = { 1L, {PR_ATTACH_NUM} };
  LPMAPITABLE mapitable;
  LPSRowSet   mapirows;
  unsigned int pos, n_attach;
  int result = -1; /* Unknown.  */
  LPATTACH att = NULL;
  LPSTREAM stream = NULL;
  char buffer[24];  /* 24 bytes are more than enough to peek at.
                       Cf. ksba_cms_identify() from the libksba
                       package.  */
  const char *p;
  ULONG nread;
  size_t n;
  tlvinfo_t ti;

  hr = message->GetAttachmentTable (0, &mapitable);
  if (FAILED (hr))
    {
      log_debug ("%s:%s: GetAttachmentTable failed: hr=%#lx",
                 SRCNAME, __func__, hr);
      return -1;
    }
      
  hr = HrQueryAllRows (mapitable, (LPSPropTagArray)&propAttNum,
                       NULL, NULL, 0, &mapirows);
  if (FAILED (hr))
    {
      log_debug ("%s:%s: HrQueryAllRows failed: hr=%#lx",
                 SRCNAME, __func__, hr);
      mapitable->Release ();
      return -1;
    }
  n_attach = mapirows->cRows > 0? mapirows->cRows : 0;
  if (n_attach != 1)
    {
      FreeProws (mapirows);
      mapitable->Release ();
      log_debug ("%s:%s: not just one attachment", SRCNAME, __func__);
      return -1;
    }
  pos = 0;

  if (mapirows->aRow[pos].cValues < 1)
    {
      log_error ("%s:%s: invalid row at pos %d", SRCNAME, __func__, pos);
      goto leave;
    }
  if (mapirows->aRow[pos].lpProps[0].ulPropTag != PR_ATTACH_NUM)
    {
      log_error ("%s:%s: invalid prop at pos %d", SRCNAME, __func__, pos);
      goto leave;
    }
  hr = message->OpenAttach (mapirows->aRow[pos].lpProps[0].Value.l,
                            NULL, MAPI_BEST_ACCESS, &att);	
  if (FAILED (hr))
    {
      log_error ("%s:%s: can't open attachment %d (%ld): hr=%#lx",
                 SRCNAME, __func__, pos, 
                 mapirows->aRow[pos].lpProps[0].Value.l, hr);
      goto leave;
    }
  if (!has_smime_filename (att))
    {
      log_debug ("%s:%s: no smime filename", SRCNAME, __func__);
      goto leave;
    }
  if (get_attach_method (att) != ATTACH_BY_VALUE)
    {
      log_debug ("%s:%s: wrong attach method", SRCNAME, __func__);
      goto leave;
    }
  
  hr = att->OpenProperty (PR_ATTACH_DATA_BIN, &IID_IStream, 
                          0, 0, (LPUNKNOWN*) &stream);
  if (FAILED (hr))
    {
      log_error ("%s:%s: can't open data stream of attachment: hr=%#lx",
                 SRCNAME, __func__, hr);
      goto leave;
    }

  hr = stream->Read (buffer, sizeof buffer, &nread);
  if ( hr != S_OK )
    {
      log_error ("%s:%s: Read failed: hr=%#lx", SRCNAME, __func__, hr);
      goto leave;
    }
  if (nread < sizeof buffer)
    {
      log_error ("%s:%s: not enough bytes returned\n", SRCNAME, __func__);
      goto leave;
    }

  p = buffer;
  n = nread;
  if (parse_tlv (&p, &n, &ti))
    goto leave;
  if (!(ti.cls == ASN1_CLASS_UNIVERSAL && ti.tag == ASN1_TAG_SEQUENCE
        && ti.is_cons) )
    goto leave;
  if (parse_tlv (&p, &n, &ti))
    goto leave;
  if (!(ti.cls == ASN1_CLASS_UNIVERSAL && ti.tag == ASN1_TAG_OBJECT_ID
        && !ti.is_cons && ti.length) || ti.length > n)
    goto leave;
  /* Now is this enveloped data (1.2.840.113549.1.7.3)
                 or signed data (1.2.840.113549.1.7.2) ? */
  if (ti.length == 9)
    {
      if (!memcmp (p, "\x2A\x86\x48\x86\xF7\x0D\x01\x07\x03", 9))
        result = 1; /* Encrypted.  */
      else if (!memcmp (p, "\x2A\x86\x48\x86\xF7\x0D\x01\x07\x02", 9))
        result = 0; /* Signed.  */
    }
  
 leave:
  if (stream)
    stream->Release ();
  if (att)
    att->Release ();
  FreeProws (mapirows);
  mapitable->Release ();
  return result;
}



/* Return the content-type of the first and only attachment of MESSAGE
   or NULL if it does not exists.  Caller must free. */
static char *
get_first_attach_mime_tag (LPMESSAGE message)
{    
  HRESULT hr;
  SizedSPropTagArray (1L, propAttNum) = { 1L, {PR_ATTACH_NUM} };
  LPMAPITABLE mapitable;
  LPSRowSet   mapirows;
  unsigned int pos, n_attach;
  LPATTACH att = NULL;
  char *result = NULL;

  hr = message->GetAttachmentTable (0, &mapitable);
  if (FAILED (hr))
    {
      log_debug ("%s:%s: GetAttachmentTable failed: hr=%#lx",
                 SRCNAME, __func__, hr);
      return NULL;
    }
      
  hr = HrQueryAllRows (mapitable, (LPSPropTagArray)&propAttNum,
                       NULL, NULL, 0, &mapirows);
  if (FAILED (hr))
    {
      log_debug ("%s:%s: HrQueryAllRows failed: hr=%#lx",
                 SRCNAME, __func__, hr);
      mapitable->Release ();
      return NULL;
    }
  n_attach = mapirows->cRows > 0? mapirows->cRows : 0;
  if (n_attach != 1)
    {
      FreeProws (mapirows);
      mapitable->Release ();
      log_debug ("%s:%s: not just one attachment", SRCNAME, __func__);
      return NULL;
    }
  pos = 0;

  if (mapirows->aRow[pos].cValues < 1)
    {
      log_error ("%s:%s: invalid row at pos %d", SRCNAME, __func__, pos);
      goto leave;
    }
  if (mapirows->aRow[pos].lpProps[0].ulPropTag != PR_ATTACH_NUM)
    {
      log_error ("%s:%s: invalid prop at pos %d", SRCNAME, __func__, pos);
      goto leave;
    }
  hr = message->OpenAttach (mapirows->aRow[pos].lpProps[0].Value.l,
                            NULL, MAPI_BEST_ACCESS, &att);	
  if (FAILED (hr))
    {
      log_error ("%s:%s: can't open attachment %d (%ld): hr=%#lx",
                 SRCNAME, __func__, pos, 
                 mapirows->aRow[pos].lpProps[0].Value.l, hr);
      goto leave;
    }

  /* Note: We do not expect a filename.  */

  if (get_attach_method (att) != ATTACH_BY_VALUE)
    {
      log_debug ("%s:%s: wrong attach method", SRCNAME, __func__);
      goto leave;
    }

  result = get_attach_mime_tag (att);
  
 leave:
  if (att)
    att->Release ();
  FreeProws (mapirows);
  mapitable->Release ();
  return result;
}


/* Helper for mapi_change_message_class.  Returns the new message
   class as an allocated string.

   Most message today are of the message class "IPM.Note".  However a
   PGP/MIME encrypted message also has this class.  We need to see
   whether we can detect such a mail right here and change the message
   class accordingly. */
static char *
change_message_class_ipm_note (LPMESSAGE message)
{
  char *newvalue = NULL;
  char *ct, *proto;

  ct = mapi_get_message_content_type (message, &proto, NULL);
  if (ct)
    {
      log_debug ("%s:%s: content type is '%s'", SRCNAME, __func__, ct);
      if (proto)
        {
          log_debug ("%s:%s:     protocol is '%s'", SRCNAME, __func__, proto);
          
          if (!strcmp (ct, "multipart/encrypted")
              && !strcmp (proto, "application/pgp-encrypted"))
            {
              newvalue = xstrdup ("IPM.Note.GpgOL.MultipartEncrypted");
            }
          else if (!strcmp (ct, "multipart/signed")
                   && !strcmp (proto, "application/pgp-signature"))
            {
              /* Sometimes we receive a PGP/MIME signed message with a
                 class IPM.Note.  */
              newvalue = xstrdup ("IPM.Note.GpgOL.MultipartSigned");
            }
          xfree (proto);
        }
      else if (!strcmp (ct, "text/plain"))
        {
          newvalue = get_msgcls_from_pgp_lines (message);
        }
      else if (!strcmp (ct, "multipart/mixed"))
        {
          /* It is quite common to have a multipart/mixed mail with
             separate encrypted PGP parts.  Look at the body to
             decide.  */
          newvalue = get_msgcls_from_pgp_lines (message);
        }
      
      xfree (ct);
    }
  else
    log_debug ("%s:%s: message has no content type", SRCNAME, __func__);

  return newvalue;
}

/* Helper for mapi_change_message_class.  Returns the new message
   class as an allocated string.

   This function is used for the message class "IPM.Note.SMIME".  It
   indicates an S/MIME opaque encrypted or signed message.  This may
   also be an PGP/MIME mail. */
static char *
change_message_class_ipm_note_smime (LPMESSAGE message)
{
  char *newvalue = NULL;
  char *ct, *proto, *smtype;
  
  ct = mapi_get_message_content_type (message, &proto, &smtype);
  if (ct)
    {
      log_debug ("%s:%s: content type is '%s'", SRCNAME, __func__, ct);
      if (proto 
          && !strcmp (ct, "multipart/signed")
          && !strcmp (proto, "application/pgp-signature"))
        {
          newvalue = xstrdup ("IPM.Note.GpgOL.MultipartSigned");
        }
      else if (!opt.enable_smime)
        ; /* S/MIME not enabled; thus no further checks.  */
      else if (smtype)
        {
          log_debug ("%s:%s:   smime-type is '%s'", SRCNAME, __func__, smtype);
          
          if (!strcmp (ct, "application/pkcs7-mime")
              || !strcmp (ct, "application/x-pkcs7-mime"))
            {
              if (!strcmp (smtype, "signed-data"))
                newvalue = xstrdup ("IPM.Note.GpgOL.OpaqueSigned");
              else if (!strcmp (smtype, "enveloped-data"))
                newvalue = xstrdup ("IPM.Note.GpgOL.OpaqueEncrypted");
            }
        }
      else
        {
          /* No smime type.  The filename parameter is often not
             reliable, thus we better look into the message to see if
             it is encrypted and assume an opaque signed one if this
             is not the case.  */
          switch (is_really_cms_encrypted (message))
            {
            case 0:
              newvalue = xstrdup ("IPM.Note.GpgOL.OpaqueSigned");
              break;
            case 1:
              newvalue = xstrdup ("IPM.Note.GpgOL.OpaqueEncrypted");
              break;
            }

        }
      xfree (smtype);
      xfree (proto);
      xfree (ct);
    }
  else
    {
      log_debug ("%s:%s: message has no content type", SRCNAME, __func__);

      /* CryptoEx (or the Toltec Connector) create messages without
         the transport headers property and thus we don't know the
         content type.  We try to detect the message type anyway by
         looking into the first and only attachments.  */
      switch (is_really_cms_encrypted (message))
        {
        case 0:
          newvalue = xstrdup ("IPM.Note.GpgOL.OpaqueSigned");
          break;
        case 1:
          newvalue = xstrdup ("IPM.Note.GpgOL.OpaqueEncrypted");
          break;
        default: /* Unknown.  */
          break;
        }
    }

  /* If we did not found anything but let's change the class anyway.  */
  if (!newvalue && opt.enable_smime)
    newvalue = xstrdup ("IPM.Note.GpgOL");

  return newvalue;
}

/* Helper for mapi_change_message_class.  Returns the new message
   class as an allocated string.

   This function is used for the message class
   "IPM.Note.SMIME.MultipartSigned".  This is an S/MIME message class
   but smime support is not enabled.  We need to check whether this is
   actually a PGP/MIME message.  */
static char *
change_message_class_ipm_note_smime_multipartsigned (LPMESSAGE message)
{
  char *newvalue = NULL;
  char *ct, *proto;

  ct = mapi_get_message_content_type (message, &proto, NULL);
  if (ct)
    {
      log_debug ("%s:%s: content type is '%s'", SRCNAME, __func__, ct);
      if (proto 
          && !strcmp (ct, "multipart/signed")
          && !strcmp (proto, "application/pgp-signature"))
        {
          newvalue = xstrdup ("IPM.Note.GpgOL.MultipartSigned");
        }
      xfree (proto);
      xfree (ct);
    }
  else
    log_debug ("%s:%s: message has no content type", SRCNAME, __func__);
  
  return newvalue;
}

/* Helper for mapi_change_message_class.  Returns the new message
   class as an allocated string.

   This function is used for the message classes
   "IPM.Note.Secure.CexSig" and "IPM.Note.Secure.Cexenc" (in the
   latter case IS_CEXSIG is true).  These are CryptoEx generated
   signature or encryption messages.  */
static char *
change_message_class_ipm_note_secure_cex (LPMESSAGE message, int is_cexenc)
{
  char *newvalue = NULL;
  char *ct, *smtype, *proto;
  
  ct = mapi_get_message_content_type (message, &proto, &smtype);
  if (ct)
    {
      log_debug ("%s:%s: content type is '%s'", SRCNAME, __func__, ct);
      if (smtype)
        log_debug ("%s:%s:   smime-type is '%s'", SRCNAME, __func__, smtype);
      if (proto)
        log_debug ("%s:%s:     protocol is '%s'", SRCNAME, __func__, proto);

      if (smtype)
        {
          if (!strcmp (ct, "application/pkcs7-mime")
              || !strcmp (ct, "application/x-pkcs7-mime"))
            {
              if (!strcmp (smtype, "signed-data"))
                newvalue = xstrdup ("IPM.Note.GpgOL.OpaqueSigned");
              else if (!strcmp (smtype, "enveloped-data"))
                newvalue = xstrdup ("IPM.Note.GpgOL.OpaqueEncrypted");
            }
        }

      if (!newvalue && proto)
        {
          if (!strcmp (ct, "multipart/signed")
              && (!strcmp (proto, "application/pkcs7-signature")
                  || !strcmp (proto, "application/x-pkcs7-signature")))
            {
              newvalue = xstrdup ("IPM.Note.GpgOL.MultipartSigned");
            }
          else if (!strcmp (ct, "multipart/signed")
                   && (!strcmp (proto, "application/pgp-signature")))
            {
              newvalue = xstrdup ("IPM.Note.GpgOL.MultipartSigned");
            }
        }
      
      if (!newvalue && !strcmp (ct, "text/plain"))
        {
          newvalue = get_msgcls_from_pgp_lines (message);
        }
      
      if (!newvalue)
        {
          switch (is_really_cms_encrypted (message))
            {
            case 0:
              newvalue = xstrdup ("IPM.Note.GpgOL.OpaqueSigned");
              break;
            case 1:
              newvalue = xstrdup ("IPM.Note.GpgOL.OpaqueEncrypted");
              break;
            }
        }
      
      xfree (smtype);
      xfree (proto);
      xfree (ct);
    }
  else
    {
      log_debug ("%s:%s: message has no content type", SRCNAME, __func__);
      if (is_cexenc)
        {
          switch (is_really_cms_encrypted (message))
            {
            case 0:
              newvalue = xstrdup ("IPM.Note.GpgOL.OpaqueSigned");
              break;
            case 1:
              newvalue = xstrdup ("IPM.Note.GpgOL.OpaqueEncrypted");
              break;
            }
        }
      else
        {
          char *mimetag;

          mimetag = get_first_attach_mime_tag (message);
          if (mimetag && !strcmp (mimetag, "multipart/signed"))
            newvalue = xstrdup ("IPM.Note.GpgOL.MultipartSigned");
          xfree (mimetag);
        }

      if (!newvalue)
        {
          newvalue = get_msgcls_from_pgp_lines (message);
        }
    }

  if (!newvalue)
    newvalue = xstrdup ("IPM.Note.GpgOL");

  return newvalue;
}


/* This function checks whether MESSAGE requires processing by us and
   adjusts the message class to our own.  By passing true for
   SYNC_OVERRIDE the actual MAPI message class will be updated to our
   own message class overide.  Return true if the message was
   changed. */
int
mapi_change_message_class (LPMESSAGE message, int sync_override)
{
  HRESULT hr;
  ULONG tag;
  SPropValue prop;
  LPSPropValue propval = NULL;
  char *newvalue = NULL;
  int need_save = 0;
  int have_override = 0;

  if (!message)
    return 0; /* No message: Nop. */

  if (get_gpgolmsgclass_tag (message, &tag) )
    return 0; /* Ooops. */

  hr = HrGetOneProp ((LPMAPIPROP)message, tag, &propval);
  if (FAILED (hr))
    {
      hr = HrGetOneProp ((LPMAPIPROP)message, PR_MESSAGE_CLASS_A, &propval);
      if (FAILED (hr))
        {
          log_error ("%s:%s: HrGetOneProp() failed: hr=%#lx\n",
                     SRCNAME, __func__, hr);
          return 0;
        }
    }
  else
    {
      have_override = 1;
      log_debug ("%s:%s: have override message class\n", SRCNAME, __func__);
    }
    
  if ( PROP_TYPE (propval->ulPropTag) == PT_STRING8 )
    {
      const char *s = propval->Value.lpszA;
      int cexenc = 0;
      
      log_debug ("%s:%s: checking message class `%s'", 
                       SRCNAME, __func__, s);
      if (!strcmp (s, "IPM.Note"))
        {
          newvalue = change_message_class_ipm_note (message);
        }
      else if (!strcmp (s, "IPM.Note.SMIME"))
        {
          newvalue = change_message_class_ipm_note_smime (message);
        }
      else if (opt.enable_smime
               && !strncmp (s, "IPM.Note.SMIME", 14) && (!s[14]||s[14] =='.'))
        {
          /* This is "IPM.Note.SMIME.foo" (where ".foo" is optional
             but the previous condition has already taken care of
             this).  Note that we can't just insert a new part and
             keep the SMIME; we need to change the SMIME part of the
             class name so that Outlook does not process it as an
             SMIME message. */
          newvalue = (char*)xmalloc (strlen (s) + 1);
          strcpy (stpcpy (newvalue, "IPM.Note.GpgOL"), s+14);
        }
      else if (!strcmp (s, "IPM.Note.SMIME.MultipartSigned"))
        {
          /* This is an S/MIME message class but smime support is not
             enabled.  We need to check whether this is actually a
             PGP/MIME message.  */
          newvalue = change_message_class_ipm_note_smime_multipartsigned
            (message);
        }
      else if (opt.enable_smime && sync_override && have_override
               && !strncmp (s, "IPM.Note.GpgOL", 14) && (!s[14]||s[14] =='.'))
        {
          /* In case the original message class is not yet an GpgOL
             class we set it here.  This is needed to convince Outlook
             not to do any special processing for IPM.Note.SMIME etc.  */
          LPSPropValue propval2 = NULL;

          hr = HrGetOneProp ((LPMAPIPROP)message, PR_MESSAGE_CLASS_A,
                             &propval2);
          if (SUCCEEDED (hr) && PROP_TYPE (propval2->ulPropTag) == PT_STRING8
              && propval2->Value.lpszA && strcmp (propval2->Value.lpszA, s))
            newvalue = (char*)xstrdup (s);
          MAPIFreeBuffer (propval2);
        }
      else if (opt.enable_smime 
               && (!strcmp (s, "IPM.Note.Secure.CexSig")
                   || (cexenc = !strcmp (s, "IPM.Note.Secure.CexEnc"))))
        {
          newvalue = change_message_class_ipm_note_secure_cex
            (message, cexenc);
        }
    }

  if (!newvalue)
    {
      /* We use our Sig-Status property to mark messages which passed
         this function.  This helps us to avoid later tests.  */
      if (!mapi_has_sig_status (message))
        {
          mapi_set_sig_status (message, "#");
          need_save = 1;
        }
    }
  else
    {
      /* Save old message class if not yet done.  (The second
         condition is just a failsafe check). */
      if (!get_gpgololdmsgclass_tag (message, &tag)
          && PROP_TYPE (propval->ulPropTag) == PT_STRING8)
        {
          LPSPropValue propval2 = NULL;

          hr = HrGetOneProp ((LPMAPIPROP)message, tag, &propval2);
          if (!FAILED (hr))
            MAPIFreeBuffer (propval2);
          else
            {
              /* No such property - save it.  */
              log_debug ("%s:%s: saving old message class\n",
                         SRCNAME, __func__);
              prop.ulPropTag = tag;
              prop.Value.lpszA = propval->Value.lpszA; 
              hr = message->SetProps (1, &prop, NULL);
              if (hr)
                {
                  log_error ("%s:%s: can't save old message class: hr=%#lx\n",
                             SRCNAME, __func__, hr);
                  MAPIFreeBuffer (propval);
                  return 0;
                }
              need_save = 1;
            }
        }
      
      /* Change message class.  */
      log_debug ("%s:%s: setting message class to `%s'\n",
                 SRCNAME, __func__, newvalue);
      prop.ulPropTag = PR_MESSAGE_CLASS_A;
      prop.Value.lpszA = newvalue; 
      hr = message->SetProps (1, &prop, NULL);
      xfree (newvalue);
      if (hr)
        {
          log_error ("%s:%s: can't set message class: hr=%#lx\n",
                     SRCNAME, __func__, hr);
          MAPIFreeBuffer (propval);
          return 0;
        }
      need_save = 1;
    }
  MAPIFreeBuffer (propval);

  if (need_save)
    {
      if (mapi_save_changes (message, KEEP_OPEN_READWRITE|FORCE_SAVE))
        return 0;
    }

  return 1;
}


/* Return the message class.  This function will never return NULL so
   it is mostly useful for debugging.  Caller needs to release the
   returned string.  */
char *
mapi_get_message_class (LPMESSAGE message)
{
  HRESULT hr;
  LPSPropValue propval = NULL;
  char *retstr;

  if (!message)
    return xstrdup ("[No message]");
  
  hr = HrGetOneProp ((LPMAPIPROP)message, PR_MESSAGE_CLASS_A, &propval);
  if (FAILED (hr))
    {
      log_error ("%s:%s: HrGetOneProp() failed: hr=%#lx\n",
                 SRCNAME, __func__, hr);
      return xstrdup (hr == MAPI_E_NOT_FOUND?
                        "[No message class property]":
                        "[Error getting message class property]");
    }

  if ( PROP_TYPE (propval->ulPropTag) == PT_STRING8 )
    retstr = xstrdup (propval->Value.lpszA);
  else
    retstr = xstrdup ("[Invalid message class property]");
    
  MAPIFreeBuffer (propval);
  return retstr;
}

/* Return the old message class.  This function returns NULL if no old
   message class has been saved.  Caller needs to release the returned
   string.  */
char *
mapi_get_old_message_class (LPMESSAGE message)
{
  HRESULT hr;
  ULONG tag;
  LPSPropValue propval = NULL;
  char *retstr;

  if (!message)
    return NULL;
  
  if (get_gpgololdmsgclass_tag (message, &tag))
    return NULL;

  hr = HrGetOneProp ((LPMAPIPROP)message, tag, &propval);
  if (FAILED (hr))
    {
      log_error ("%s:%s: HrGetOneProp() failed: hr=%#lx\n",
                 SRCNAME, __func__, hr);
      return NULL;
    }

  if ( PROP_TYPE (propval->ulPropTag) == PT_STRING8 )
    retstr = xstrdup (propval->Value.lpszA);
  else
    retstr = NULL;
    
  MAPIFreeBuffer (propval);
  return retstr;
}



/* Return the sender of the message.  According to the specs this is
   an UTF-8 string; we rely on that the UI server handles
   internationalized domain names.  */ 
char *
mapi_get_sender (LPMESSAGE message)
{
  HRESULT hr;
  LPSPropValue propval = NULL;
  char *buf;
  char *p0, *p;
  
  if (!message)
    return NULL; /* No message: Nop. */

  hr = HrGetOneProp ((LPMAPIPROP)message, PR_PRIMARY_SEND_ACCT, &propval);
  if (FAILED (hr))
    {
      log_debug ("%s:%s: HrGetOneProp failed: hr=%#lx\n",
                 SRCNAME, __func__, hr);
      return NULL;
    }
    
  if (PROP_TYPE (propval->ulPropTag) != PT_UNICODE) 
    {
      log_debug ("%s:%s: HrGetOneProp returns invalid type %lu\n",
                 SRCNAME, __func__, PROP_TYPE (propval->ulPropTag) );
      MAPIFreeBuffer (propval);
      return NULL;
    }
  
  buf = wchar_to_utf8 (propval->Value.lpszW);
  MAPIFreeBuffer (propval);
  if (!buf)
    {
      log_error ("%s:%s: error converting to utf8\n", SRCNAME, __func__);
      return NULL;
    }
  /* The PR_PRIMARY_SEND_ACCT property seems to be divided into fields
     using Ctrl-A as delimiter.  The first field looks like the ascii
     formatted number of fields to follow, the second field like the
     email account and the third seems to be a textual description of
     that account.  We return the second field. */
  p = strchr (buf, '\x01');
  if (!p)
    {
      log_error ("%s:%s: unknown format of the value `%s'\n",
                 SRCNAME, __func__, buf);
      xfree (buf);
      return NULL;
    }
  for (p0=buf, p++; *p && *p != '\x01';)
    *p0++ = *p++;
  *p0 = 0;

  /* When using an Exchange account this is an X.509 address and not
     an SMTP address.  We try to detect this here and extract only the
     CN RDN.  Note that there are two CNs.  This is just a simple
     approach and not a real parser.  A better way to do this would be
     to ask MAPI to resolve the X.500 name to an SMTP name.  */
  if (strstr (buf, "/o=") && strstr (buf, "/ou=") &&
      (p = strstr (buf, "/cn=Recipients")) && (p = strstr (p+1, "/cn=")))
    {
      log_debug ("%s:%s: orig address is `%s'\n", SRCNAME, __func__, buf);
      memmove (buf, p+4, strlen (p+4)+1);
      if (!strchr (buf, '@'))
        {
          /* Some Exchange accounts return only the accoutn name and
             no rfc821 mail address.  Kleopatra chokes on that, thus
             we append a domain name.  Thisis a bad hack.  */
          char *newbuf = (char *)xmalloc (strlen (buf) + 6 + 1);
          strcpy (stpcpy (newbuf, buf), "@local");
          xfree (buf);
          buf = newbuf;
        }
      
    }
  log_debug ("%s:%s: address is `%s'\n", SRCNAME, __func__, buf);
  return buf;
}

/* Return the from address of the message as a malloced UTF-8 string.
   Returns NULL if that address is not available.  */
char *
mapi_get_from_address (LPMESSAGE message)
{
  HRESULT hr;
  LPSPropValue propval = NULL;
  char *buf;
  
  if (!message)
    return xstrdup ("[no message]"); /* Ooops.  */

  hr = HrGetOneProp ((LPMAPIPROP)message, PR_SENDER_EMAIL_ADDRESS_W, &propval);
  if (FAILED (hr))
    {
      log_debug ("%s:%s: HrGetOneProp failed: hr=%#lx\n",
                 SRCNAME, __func__, hr);
      return NULL;
    }
    
  if (PROP_TYPE (propval->ulPropTag) != PT_UNICODE) 
    {
      log_debug ("%s:%s: HrGetOneProp returns invalid type %lu\n",
                 SRCNAME, __func__, PROP_TYPE (propval->ulPropTag) );
      MAPIFreeBuffer (propval);
      return NULL;
    }
  
  buf = wchar_to_utf8 (propval->Value.lpszW);
  MAPIFreeBuffer (propval);
  if (!buf)
    {
      log_error ("%s:%s: error converting to utf8\n", SRCNAME, __func__);
      return NULL;
    }

  return buf;
}


/* Return the subject of the message as a malloced UTF-8 string.
   Returns a replacement string if a subject is missing.  */
char *
mapi_get_subject (LPMESSAGE message)
{
  HRESULT hr;
  LPSPropValue propval = NULL;
  char *buf;
  
  if (!message)
    return xstrdup ("[no message]"); /* Ooops.  */

  hr = HrGetOneProp ((LPMAPIPROP)message, PR_SUBJECT_W, &propval);
  if (FAILED (hr))
    {
      log_debug ("%s:%s: HrGetOneProp failed: hr=%#lx\n",
                 SRCNAME, __func__, hr);
      return xstrdup (_("[no subject]"));
    }
    
  if (PROP_TYPE (propval->ulPropTag) != PT_UNICODE) 
    {
      log_debug ("%s:%s: HrGetOneProp returns invalid type %lu\n",
                 SRCNAME, __func__, PROP_TYPE (propval->ulPropTag) );
      MAPIFreeBuffer (propval);
      return xstrdup (_("[no subject]"));
    }
  
  buf = wchar_to_utf8 (propval->Value.lpszW);
  MAPIFreeBuffer (propval);
  if (!buf)
    {
      log_error ("%s:%s: error converting to utf8\n", SRCNAME, __func__);
      return xstrdup (_("[no subject]"));
    }

  return buf;
}




/* Return the message type.  This function knows only about our own
   message types.  Returns MSGTYPE_UNKNOWN for any MESSAGE we have
   no special support for.  */
msgtype_t
mapi_get_message_type (LPMESSAGE message)
{
  HRESULT hr;
  ULONG tag;
  LPSPropValue propval = NULL;
  msgtype_t msgtype = MSGTYPE_UNKNOWN;

  if (!message)
    return msgtype; 

  if (get_gpgolmsgclass_tag (message, &tag) )
    return msgtype; /* Ooops */

  hr = HrGetOneProp ((LPMAPIPROP)message, tag, &propval);
  if (FAILED (hr))
    {
      hr = HrGetOneProp ((LPMAPIPROP)message, PR_MESSAGE_CLASS_A, &propval);
      if (FAILED (hr))
        {
          log_error ("%s:%s: HrGetOneProp(PR_MESSAGE_CLASS) failed: hr=%#lx\n",
                     SRCNAME, __func__, hr);
          return msgtype;
        }
    }
  else
    log_debug ("%s:%s: have override message class\n", SRCNAME, __func__);
    
  if ( PROP_TYPE (propval->ulPropTag) == PT_STRING8 )
    {
      const char *s = propval->Value.lpszA;

      if (!strncmp (s, "IPM.Note.GpgOL", 14) && (!s[14] || s[14] =='.'))
        {
          s += 14;
          if (!*s)
            msgtype = MSGTYPE_GPGOL;
          else if (!strcmp (s, ".MultipartSigned"))
            msgtype = MSGTYPE_GPGOL_MULTIPART_SIGNED;
          else if (!strcmp (s, ".MultipartEncrypted"))
            msgtype = MSGTYPE_GPGOL_MULTIPART_ENCRYPTED;
          else if (!strcmp (s, ".OpaqueSigned"))
            msgtype = MSGTYPE_GPGOL_OPAQUE_SIGNED;
          else if (!strcmp (s, ".OpaqueEncrypted"))
            msgtype = MSGTYPE_GPGOL_OPAQUE_ENCRYPTED;
          else if (!strcmp (s, ".ClearSigned"))
            msgtype = MSGTYPE_GPGOL_CLEAR_SIGNED;
          else if (!strcmp (s, ".PGPMessage"))
            msgtype = MSGTYPE_GPGOL_PGP_MESSAGE;
          else
            log_debug ("%s:%s: message class `%s' not supported",
                       SRCNAME, __func__, s-14);
        }
      else if (!strncmp (s, "IPM.Note.SMIME", 14) && (!s[14] || s[14] =='.'))
        msgtype = MSGTYPE_SMIME;
    }
  MAPIFreeBuffer (propval);
  return msgtype;
}


/* This function is pretty useless because IConverterSession won't
   take attachments into account.  Need to write our own version.  */
int
mapi_to_mime (LPMESSAGE message, const char *filename)
{
  HRESULT hr;
  LPCONVERTERSESSION session;
  LPSTREAM stream;

  hr = CoCreateInstance (CLSID_IConverterSession, NULL, CLSCTX_INPROC_SERVER,
                         IID_IConverterSession, (void **) &session);
  if (FAILED (hr))
    {
      log_error ("%s:%s: can't create new IConverterSession object: hr=%#lx",
                 SRCNAME, __func__, hr);
      return -1;
    }


  hr = OpenStreamOnFile (MAPIAllocateBuffer, MAPIFreeBuffer,
                         (STGM_CREATE | STGM_READWRITE),
                         (char*)filename, NULL, &stream); 
  if (FAILED (hr)) 
    {
      log_error ("%s:%s: can't create file `%s': hr=%#lx\n",
                 SRCNAME, __func__, filename, hr); 
      hr = -1;
    }
  else
    {
      hr = session->MAPIToMIMEStm (message, stream, CCSF_SMTP);
      if (FAILED (hr))
        {
          log_error ("%s:%s: MAPIToMIMEStm failed: hr=%#lx",
                     SRCNAME, __func__, hr);
          stream->Revert ();
          hr = -1;
        }
      else
        {
          stream->Commit (0);
          hr = 0;
        }

      stream->Release ();
    }

  session->Release ();
  return hr;
}


/* Return a binary property in a malloced buffer with its length stored
   at R_NBYTES.  Returns NULL on error.  */
char *
mapi_get_binary_prop (LPMESSAGE message, ULONG proptype, size_t *r_nbytes)
{
  HRESULT hr;
  LPSPropValue propval = NULL;
  char *data;

  *r_nbytes = 0;
  hr = HrGetOneProp ((LPMAPIPROP)message, proptype, &propval);
  if (FAILED (hr))
    {
      log_error ("%s:%s: error getting property %#lx: hr=%#lx",
                 SRCNAME, __func__, proptype, hr);
      return NULL; 
    }
  switch ( PROP_TYPE (propval->ulPropTag) )
    {
    case PT_BINARY:
      /* This is a binary object but we know that it must be plain
         ASCII due to the armored format.  */
      data = (char*)xmalloc (propval->Value.bin.cb + 1);
      memcpy (data, propval->Value.bin.lpb, propval->Value.bin.cb);
      data[propval->Value.bin.cb] = 0;
      *r_nbytes = propval->Value.bin.cb;
      break;
      
    default:
      log_debug ("%s:%s: requested property %#lx has unknown tag %#lx\n",
                 SRCNAME, __func__, proptype, propval->ulPropTag);
      data = NULL;
      break;
    }
  MAPIFreeBuffer (propval);
  return data;
}

/* Return an integer property at R_VALUE.  On error the function
   returns -1 and sets R_VALUE to 0, on success 0 is returned.  */
int
mapi_get_int_prop (LPMAPIPROP object, ULONG proptype, LONG *r_value)
{
  int rc = -1;
  HRESULT hr;
  LPSPropValue propval = NULL;

  *r_value = 0;
  hr = HrGetOneProp (object, proptype, &propval);
  if (FAILED (hr))
    {
      log_error ("%s:%s: error getting property %#lx: hr=%#lx",
                 SRCNAME, __func__, proptype, hr);
      return -1; 
    }
  switch ( PROP_TYPE (propval->ulPropTag) )
    {
    case PT_LONG:
      *r_value = propval->Value.l;
      rc = 0;
      
      break;
      
    default:
      log_debug ("%s:%s: requested property %#lx has unknown tag %#lx\n",
                 SRCNAME, __func__, proptype, propval->ulPropTag);
      break;
    }
  MAPIFreeBuffer (propval);
  return rc;
}


/* Return the attachment method for attachment OBJ.  In case of error
   we return 0 which happens not to be defined.  */
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
                 SRCNAME, __func__, hr);
      return 0; 
    }
  /* We don't bother checking whether we really get a PT_LONG ulong
     back; if not the system is seriously damaged and we can't do
     further harm by returning a possible random value.  */
  method = propval->Value.l;
  MAPIFreeBuffer (propval);
  return method;
}



/* Return the filename from the attachment as a malloced string.  The
   encoding we return will be UTF-8, however the MAPI docs declare
   that MAPI does only handle plain ANSI and thus we don't really care
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
      log_debug ("%s:%s: no filename property found", SRCNAME, __func__);
      return NULL;
    }

  switch ( PROP_TYPE (propval->ulPropTag) )
    {
    case PT_UNICODE:
      name = wchar_to_utf8 (propval->Value.lpszW);
      if (!name)
        log_debug ("%s:%s: error converting to utf8\n", SRCNAME, __func__);
      break;
      
    case PT_STRING8:
      name = xstrdup (propval->Value.lpszA);
      break;
      
    default:
      log_debug ("%s:%s: proptag=%#lx not supported\n",
                 SRCNAME, __func__, propval->ulPropTag);
      name = NULL;
      break;
    }
  MAPIFreeBuffer (propval);
  return name;
}


/* Return the content-type of the attachment OBJ or NULL if it does
   not exists.  Caller must free. */
static char *
get_attach_mime_tag (LPATTACH obj)
{
  HRESULT hr;
  LPSPropValue propval = NULL;
  char *name;

  hr = HrGetOneProp ((LPMAPIPROP)obj, PR_ATTACH_MIME_TAG_A, &propval);
  if (FAILED (hr))
    {
      if (hr != MAPI_E_NOT_FOUND)
        log_error ("%s:%s: error getting attachment's MIME tag: hr=%#lx",
                   SRCNAME, __func__, hr);
      return NULL; 
    }
  switch ( PROP_TYPE (propval->ulPropTag) )
    {
    case PT_UNICODE:
      name = wchar_to_utf8 (propval->Value.lpszW);
      if (!name)
        log_debug ("%s:%s: error converting to utf8\n", SRCNAME, __func__);
      break;
      
    case PT_STRING8:
      name = xstrdup (propval->Value.lpszA);
      break;
      
    default:
      log_debug ("%s:%s: proptag=%#lx not supported\n",
                 SRCNAME, __func__, propval->ulPropTag);
      name = NULL;
      break;
    }
  MAPIFreeBuffer (propval);
  return name;
}


/* Return the GpgOL Attach Type for attachment OBJ.  Tag needs to be
   the tag of that property. */
attachtype_t
get_gpgolattachtype (LPATTACH obj, ULONG tag)
{
  HRESULT hr;
  LPSPropValue propval = NULL;
  attachtype_t retval;

  hr = HrGetOneProp ((LPMAPIPROP)obj, tag, &propval);
  if (FAILED (hr))
    {
      if (hr != MAPI_E_NOT_FOUND)
        log_error ("%s:%s: error getting GpgOL Attach Type: hr=%#lx",
                   SRCNAME, __func__, hr);
      return ATTACHTYPE_UNKNOWN; 
    }
  retval = (attachtype_t)propval->Value.l;
  MAPIFreeBuffer (propval);
  return retval;
}


/* Gather information about attachments and return a new table of
   attachments.  Caller must release the returned table.s The routine
   will return NULL in case of an error or if no attachments are
   available.  With FAST set only some information gets collected. */
mapi_attach_item_t *
mapi_create_attach_table (LPMESSAGE message, int fast)
{    
  HRESULT hr;
  SizedSPropTagArray (1L, propAttNum) = { 1L, {PR_ATTACH_NUM} };
  LPMAPITABLE mapitable;
  LPSRowSet   mapirows;
  mapi_attach_item_t *table; 
  unsigned int pos, n_attach;
  ULONG moss_tag;

  if (get_gpgolattachtype_tag (message, &moss_tag) )
    return NULL;

  /* Open the attachment table.  */
  hr = message->GetAttachmentTable (0, &mapitable);
  if (FAILED (hr))
    {
      log_debug ("%s:%s: GetAttachmentTable failed: hr=%#lx",
                 SRCNAME, __func__, hr);
      return NULL;
    }
      
  hr = HrQueryAllRows (mapitable, (LPSPropTagArray)&propAttNum,
                       NULL, NULL, 0, &mapirows);
  if (FAILED (hr))
    {
      log_debug ("%s:%s: HrQueryAllRows failed: hr=%#lx",
                 SRCNAME, __func__, hr);
      mapitable->Release ();
      return NULL;
    }
  n_attach = mapirows->cRows > 0? mapirows->cRows : 0;

  log_debug ("%s:%s: message has %u attachments\n",
             SRCNAME, __func__, n_attach);
  if (!n_attach)
    {
      FreeProws (mapirows);
      mapitable->Release ();
      return NULL;
    }

  /* Allocate our own table.  */
  table = (mapi_attach_item_t *)xcalloc (n_attach+1, sizeof *table);
  for (pos=0; pos < n_attach; pos++) 
    {
      LPATTACH att;

      if (mapirows->aRow[pos].cValues < 1)
        {
          log_error ("%s:%s: invalid row at pos %d", SRCNAME, __func__, pos);
          table[pos].mapipos = -1;
          continue;
        }
      if (mapirows->aRow[pos].lpProps[0].ulPropTag != PR_ATTACH_NUM)
        {
          log_error ("%s:%s: invalid prop at pos %d", SRCNAME, __func__, pos);
          table[pos].mapipos = -1;
          continue;
        }
      table[pos].mapipos = mapirows->aRow[pos].lpProps[0].Value.l;

      hr = message->OpenAttach (table[pos].mapipos, NULL,
                                MAPI_BEST_ACCESS, &att);	
      if (FAILED (hr))
        {
          log_error ("%s:%s: can't open attachment %d (%d): hr=%#lx",
                     SRCNAME, __func__, pos, table[pos].mapipos, hr);
          table[pos].mapipos = -1;
          continue;
        }

      table[pos].method = get_attach_method (att);
      table[pos].filename = fast? NULL : get_attach_filename (att);
      table[pos].content_type = fast? NULL : get_attach_mime_tag (att);
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
      table[pos].attach_type = get_gpgolattachtype (att, moss_tag);
      att->Release ();
    }
  table[0].private_mapitable = mapitable;
  FreeProws (mapirows);
  table[pos].end_of_table = 1;
  mapitable = NULL;

  if (fast)
    {
      log_debug ("%s:%s: attachment info: not shown due to fast flag\n",
                 SRCNAME, __func__);
    }
  else
    {
      log_debug ("%s:%s: attachment info:\n", SRCNAME, __func__);
      for (pos=0; !table[pos].end_of_table; pos++)
        {
          log_debug ("\t%d mt=%d fname=`%s' ct=`%s' ct_parms=`%s'\n",
                     table[pos].mapipos,
                     table[pos].attach_type,
                     table[pos].filename, table[pos].content_type,
                     table[pos].content_type_parms);
        }
    }

  return table;
}


/* Release a table as created by mapi_create_attach_table. */
void
mapi_release_attach_table (mapi_attach_item_t *table)
{
  unsigned int pos;
  LPMAPITABLE mapitable;

  if (!table)
    return;

  mapitable = (LPMAPITABLE)table[0].private_mapitable;
  if (mapitable)
    mapitable->Release ();
  for (pos=0; !table[pos].end_of_table; pos++)
    {
      xfree (table[pos].filename);
      xfree (table[pos].content_type);
    }
  xfree (table);
}


/* Return an attachment as a new IStream object.  Returns NULL on
   failure.  If R_ATTACH is not NULL the actual attachment will not be
   released but stored at that address; the caller needs to release it
   in this case.  */
LPSTREAM
mapi_get_attach_as_stream (LPMESSAGE message, mapi_attach_item_t *item,
                           LPATTACH *r_attach)
{
  HRESULT hr;
  LPATTACH att;
  LPSTREAM stream;

  if (r_attach)
    *r_attach = NULL;

  if (!item || item->end_of_table || item->mapipos == -1)
    return NULL;

  hr = message->OpenAttach (item->mapipos, NULL, MAPI_BEST_ACCESS, &att);
  if (FAILED (hr))
    {
      log_error ("%s:%s: can't open attachment at %d: hr=%#lx",
                 SRCNAME, __func__, item->mapipos, hr);
      return NULL;
    }
  if (item->method != ATTACH_BY_VALUE)
    {
      log_error ("%s:%s: attachment: method not supported", SRCNAME, __func__);
      att->Release ();
      return NULL;
    }

  hr = att->OpenProperty (PR_ATTACH_DATA_BIN, &IID_IStream, 
                          0, 0, (LPUNKNOWN*) &stream);
  if (FAILED (hr))
    {
      log_error ("%s:%s: can't open data stream of attachment: hr=%#lx",
                 SRCNAME, __func__, hr);
      att->Release ();
      return NULL;
    }

  if (r_attach)
    *r_attach = att;
  else
    att->Release ();

  return stream;
}


/* Return a malloced buffer with the content of the attachment. If
   R_NBYTES is not NULL the number of bytes will get stored there.
   ATT must have an attachment method of ATTACH_BY_VALUE.  Returns
   NULL on error.  If UNPROTECT is set and the appropriate crypto
   attribute is available, the function returns the unprotected
   version of the atatchment. */
static char *
attach_to_buffer (LPATTACH att, size_t *r_nbytes, int unprotect, 
                  int *r_was_protected)
{
  HRESULT hr;
  LPSTREAM stream;
  STATSTG statInfo;
  ULONG nread;
  char *buffer;
  symenc_t symenc = NULL;

  if (r_was_protected)
    *r_was_protected = 0;

  if (unprotect)
    {
      ULONG tag;
      char *iv;
      size_t ivlen;

      if (!get_gpgolprotectiv_tag ((LPMESSAGE)att, &tag) 
          && (iv = mapi_get_binary_prop ((LPMESSAGE)att, tag, &ivlen)))
        {
          symenc = symenc_open (get_128bit_session_key (), 16, iv, ivlen);
          xfree (iv);
          if (!symenc)
            log_error ("%s:%s: can't open encryption context", 
                       SRCNAME, __func__);
        }
    }
  

  hr = att->OpenProperty (PR_ATTACH_DATA_BIN, &IID_IStream, 
                          0, 0, (LPUNKNOWN*) &stream);
  if (FAILED (hr))
    {
      log_error ("%s:%s: can't open data stream of attachment: hr=%#lx",
                 SRCNAME, __func__, hr);
      return NULL;
    }

  hr = stream->Stat (&statInfo, STATFLAG_NONAME);
  if ( hr != S_OK )
    {
      log_error ("%s:%s: Stat failed: hr=%#lx", SRCNAME, __func__, hr);
      stream->Release ();
      return NULL;
    }
      
  /* Allocate one byte more so that we can terminate the string.  */
  buffer = (char*)xmalloc ((size_t)statInfo.cbSize.QuadPart + 1);

  hr = stream->Read (buffer, (size_t)statInfo.cbSize.QuadPart, &nread);
  if ( hr != S_OK )
    {
      log_error ("%s:%s: Read failed: hr=%#lx", SRCNAME, __func__, hr);
      xfree (buffer);
      stream->Release ();
      return NULL;
    }
  if (nread != statInfo.cbSize.QuadPart)
    {
      log_error ("%s:%s: not enough bytes returned\n", SRCNAME, __func__);
      xfree (buffer);
      buffer = NULL;
    }
  stream->Release ();

  if (buffer && symenc)
    {
      symenc_cfb_decrypt (symenc, buffer, buffer, nread);
      if (nread < 16 || memcmp (buffer, "GpgOL attachment", 16))
        {
          xfree (buffer);
          buffer = native_to_utf8 
            (_("[The content of this message is not visible because it has "
               "been decrypted by another Outlook session.  Use the "
               "\"decrypt/verify\" command to make it visible]"));
          nread = strlen (buffer);
        }
      else
        {
          memmove (buffer, buffer+16, nread-16);
          nread -= 16;
          if (r_was_protected)
            *r_was_protected = 1;
        }
    }

  /* Make sure that the buffer is a C string.  */
  if (buffer)
    buffer[nread] = 0;

  symenc_close (symenc);
  if (r_nbytes)
    *r_nbytes = nread;
  return buffer;
}



/* Return an attachment as a malloced buffer.  The size of the buffer
   will be stored at R_NBYTES.  If unprotect is true, the atatchment
   will be unprotected.  Returns NULL on failure. */
char *
mapi_get_attach (LPMESSAGE message, int unprotect, 
                 mapi_attach_item_t *item, size_t *r_nbytes)
{
  HRESULT hr;
  LPATTACH att;
  char *buffer;

  if (!item || item->end_of_table || item->mapipos == -1)
    return NULL;

  hr = message->OpenAttach (item->mapipos, NULL, MAPI_BEST_ACCESS, &att);
  if (FAILED (hr))
    {
      log_error ("%s:%s: can't open attachment at %d: hr=%#lx",
                 SRCNAME, __func__, item->mapipos, hr);
      return NULL;
    }
  if (item->method != ATTACH_BY_VALUE)
    {
      log_error ("%s:%s: attachment: method not supported", SRCNAME, __func__);
      att->Release ();
      return NULL;
    }

  buffer = attach_to_buffer (att, r_nbytes, unprotect, NULL);
  att->Release ();

  return buffer;
}


/* Mark this attachment as the original MOSS message.  We set a custom
   property as well as the hidden flag.  */
int 
mapi_mark_moss_attach (LPMESSAGE message, mapi_attach_item_t *item)
{
  int retval = -1;
  HRESULT hr;
  LPATTACH att;
  SPropValue prop;

  if (!item || item->end_of_table || item->mapipos == -1)
    return -1;

  hr = message->OpenAttach (item->mapipos, NULL, MAPI_BEST_ACCESS, &att);
  if (FAILED (hr))
    {
      log_error ("%s:%s: can't open attachment at %d: hr=%#lx",
                 SRCNAME, __func__, item->mapipos, hr);
      return -1;
    }

  if (get_gpgolattachtype_tag (message, &prop.ulPropTag) )
    goto leave;
  prop.Value.l = ATTACHTYPE_MOSS;
  hr = HrSetOneProp (att, &prop);	
  if (hr)
    {
      log_error ("%s:%s: can't set %s property: hr=%#lx\n",
                 SRCNAME, __func__, "GpgOL Attach Type", hr); 
      return false;
    }

  prop.ulPropTag = PR_ATTACHMENT_HIDDEN;
  prop.Value.b = TRUE;
  hr = HrSetOneProp (att, &prop);
  if (hr)
    {
      log_error ("%s:%s: can't set hidden attach flag: hr=%#lx\n",
                 SRCNAME, __func__, hr); 
      goto leave;
    }
  

  hr = att->SaveChanges (KEEP_OPEN_READWRITE);
  if (hr)
    {
      log_error ("%s:%s: SaveChanges(attachment) failed: hr=%#lx\n",
                 SRCNAME, __func__, hr); 
      goto leave;
    }
  
  retval = 0;
    
 leave:
  att->Release ();
  return retval;
}


/* If the hidden property has not been set on ATTACH, set it and save
   the changes. */
int 
mapi_set_attach_hidden (LPATTACH attach)
{
  int retval = -1;
  HRESULT hr;
  LPSPropValue propval;
  SPropValue prop;

  hr = HrGetOneProp ((LPMAPIPROP)attach, PR_ATTACHMENT_HIDDEN, &propval);
  if (SUCCEEDED (hr) 
      && PROP_TYPE (propval->ulPropTag) == PT_BOOLEAN
      && propval->Value.b)
    return 0;/* Already set to hidden. */

  prop.ulPropTag = PR_ATTACHMENT_HIDDEN;
  prop.Value.b = TRUE;
  hr = HrSetOneProp (attach, &prop);
  if (hr)
    {
      log_error ("%s:%s: can't set hidden attach flag: hr=%#lx\n",
                 SRCNAME, __func__, hr); 
      goto leave;
    }
  
  hr = attach->SaveChanges (KEEP_OPEN_READWRITE);
  if (hr)
    {
      log_error ("%s:%s: SaveChanges(attachment) failed: hr=%#lx\n",
                 SRCNAME, __func__, hr); 
      goto leave;
    }
  
  retval = 0;
    
 leave:
  return retval;
}


/* Returns true if ATTACH has the hidden flag set to true.  */
int
mapi_test_attach_hidden (LPATTACH attach)
{
  HRESULT hr;
  LPSPropValue propval = NULL;
  int result = 0;
  
  hr = HrGetOneProp ((LPMAPIPROP)attach, PR_ATTACHMENT_HIDDEN, &propval);
  if (FAILED (hr))
    return result; /* No.  */  
  
  if (PROP_TYPE (propval->ulPropTag) == PT_BOOLEAN && propval->Value.b)
    result = 1; /* Yes.  */

  MAPIFreeBuffer (propval);
  return result;
}




/* Returns True if MESSAGE has the GpgOL Sig Status property.  */
int
mapi_has_sig_status (LPMESSAGE msg)
{
  HRESULT hr;
  LPSPropValue propval = NULL;
  ULONG tag;
  int yes;

  if (get_gpgolsigstatus_tag (msg, &tag) )
    return 0; /* Error:  Assume No.  */
  hr = HrGetOneProp ((LPMAPIPROP)msg, tag, &propval);
  if (FAILED (hr))
    return 0; /* No.  */  
  if (PROP_TYPE (propval->ulPropTag) == PT_STRING8)
    yes = 1;
  else
    yes = 0;

  MAPIFreeBuffer (propval);
  return yes;
}


/* Returns True if MESSAGE has a GpgOL Sig Status property and that it
   is not set to unchecked.  */
int
mapi_test_sig_status (LPMESSAGE msg)
{
  HRESULT hr;
  LPSPropValue propval = NULL;
  ULONG tag;
  int yes;

  if (get_gpgolsigstatus_tag (msg, &tag) )
    return 0; /* Error:  Assume No.  */
  hr = HrGetOneProp ((LPMAPIPROP)msg, tag, &propval);
  if (FAILED (hr))
    return 0; /* No.  */  

  /* We return False if we have an unknown signature status (?) or the
     message has been sent by us and not yet checked (@).  */
  if (PROP_TYPE (propval->ulPropTag) == PT_STRING8)
    yes = !(propval->Value.lpszA && (!strcmp (propval->Value.lpszA, "?")
                                     || !strcmp (propval->Value.lpszA, "@")));
  else
    yes = 0;

  MAPIFreeBuffer (propval);
  return yes;
}


/* Return the signature status as an allocated string.  Will never
   return NULL.  */
char *
mapi_get_sig_status (LPMESSAGE msg)
{
  HRESULT hr;
  LPSPropValue propval = NULL;
  ULONG tag;
  char *retstr;

  if (get_gpgolsigstatus_tag (msg, &tag) )
    return xstrdup ("[Error getting tag for sig status]");
  hr = HrGetOneProp ((LPMAPIPROP)msg, tag, &propval);
  if (FAILED (hr))
    return xstrdup ("");
  if (PROP_TYPE (propval->ulPropTag) == PT_STRING8)
    retstr = xstrdup (propval->Value.lpszA);
  else
    retstr = xstrdup ("[Sig status has an invalid type]");

  MAPIFreeBuffer (propval);
  return retstr;
}




/* Set the signature status property to STATUS_STRING.  There are a
   few special values:

     "#" The message is not of interest to us.
     "@" The message has been created and signed or encrypted by us.
     "?" The signature status has not been checked.
     "!" The signature verified okay 
     "~" The signature was not fully verified.
     "-" The signature is bad

   Note that this function does not call SaveChanges.  */
int 
mapi_set_sig_status (LPMESSAGE message, const char *status_string)
{
  HRESULT hr;
  SPropValue prop;

  if (get_gpgolsigstatus_tag (message, &prop.ulPropTag) )
    return -1;
  prop.Value.lpszA = xstrdup (status_string);
  hr = HrSetOneProp (message, &prop);	
  xfree (prop.Value.lpszA);
  if (hr)
    {
      log_error ("%s:%s: can't set %s property: hr=%#lx\n",
                 SRCNAME, __func__, "GpgOL Sig Status", hr); 
      return -1;
    }

  return 0;
}


/* When sending a message we need to fake the message class so that OL
   processes it according to our needs.  However, if we later try to
   get the message class from the sent message, OL still has the SMIME
   message class and tries to hide this by trying to decrypt the
   message and return the message class from the plaintext.  To
   mitigate the problem we define our own msg class override
   property.  */
int 
mapi_set_gpgol_msg_class (LPMESSAGE message, const char *name)
{
  HRESULT hr;
  SPropValue prop;

  if (get_gpgolmsgclass_tag (message, &prop.ulPropTag) )
    return -1;
  prop.Value.lpszA = xstrdup (name);
  hr = HrSetOneProp (message, &prop);	
  xfree (prop.Value.lpszA);
  if (hr)
    {
      log_error ("%s:%s: can't set %s property: hr=%#lx\n",
                 SRCNAME, __func__, "GpgOL Msg Class", hr); 
      return -1;
    }

  return 0;
}


/* Return the charset as assigned by GpgOL to an attachment.  This may
   return NULL it is has not been assigned or is the standard
   (UTF-8). */
char *
mapi_get_gpgol_charset (LPMESSAGE obj)
{
  HRESULT hr;
  LPSPropValue propval = NULL;
  ULONG tag;
  char *retstr;

  if (get_gpgolcharset_tag (obj, &tag) )
    return NULL; /* Error.  */
  hr = HrGetOneProp ((LPMAPIPROP)obj, tag, &propval);
  if (FAILED (hr))
    return NULL;
  if (PROP_TYPE (propval->ulPropTag) == PT_STRING8)
    {
      if (!strcmp (propval->Value.lpszA, "utf-8"))
        retstr = NULL;
      else
        retstr = xstrdup (propval->Value.lpszA);
    }
  else
    retstr = NULL;

  MAPIFreeBuffer (propval);
  return retstr;
}


/* Set the GpgOl charset property to an attachment. 
   Note that this function does not call SaveChanges.  */
int 
mapi_set_gpgol_charset (LPMESSAGE obj, const char *charset)
{
  HRESULT hr;
  SPropValue prop;
  char *p;

  /* Note that we lowercase the value and cut it to a max of 32
     characters.  The latter is required to make sure that
     HrSetOneProp will always work.  */
  if (get_gpgolcharset_tag (obj, &prop.ulPropTag) )
    return -1;
  prop.Value.lpszA = xstrdup (charset);
  for (p=prop.Value.lpszA; *p; p++)
    *p = tolower (*(unsigned char*)p);
  if (strlen (prop.Value.lpszA) > 32)
    prop.Value.lpszA[32] = 0;
  hr = HrSetOneProp ((LPMAPIPROP)obj, &prop);	
  xfree (prop.Value.lpszA);
  if (hr)
    {
      log_error ("%s:%s: can't set %s property: hr=%#lx\n",
                 SRCNAME, __func__, "GpgOL Charset", hr); 
      return -1;
    }

  return 0;
}



/* Return GpgOL's draft info string as an allocated string.  If no
   draft info is available, NULL is returned.  */
char *
mapi_get_gpgol_draft_info (LPMESSAGE msg)
{
  HRESULT hr;
  LPSPropValue propval = NULL;
  ULONG tag;
  char *retstr;

  if (get_gpgoldraftinfo_tag (msg, &tag) )
    return NULL;
  hr = HrGetOneProp ((LPMAPIPROP)msg, tag, &propval);
  if (FAILED (hr))
    return NULL;
  if (PROP_TYPE (propval->ulPropTag) == PT_STRING8)
    retstr = xstrdup (propval->Value.lpszA);
  else
    retstr = NULL;

  MAPIFreeBuffer (propval);
  return retstr;
}


/* Set GpgOL's draft info string to STRING.  This string is defined as:

   Character 1:  'E' = encrypt selected,
                 'e' = encrypt not selected.
                 '-' = don't care
   Character 2:  'S' = sign selected,
                 's' = sign not selected.
                 '-' = don't care
   Character 3:  'A' = Auto protocol 
                 'P' = OpenPGP protocol
                 'X' = S/MIME protocol
                 '-' = don't care
                 
   If string is NULL, the property will get deleted.

   Note that this function does not call SaveChanges.  */
int 
mapi_set_gpgol_draft_info (LPMESSAGE message, const char *string)
{
  HRESULT hr;
  SPropValue prop;
  SPropTagArray proparray;

  if (get_gpgoldraftinfo_tag (message, &prop.ulPropTag) )
    return -1;
  if (string)
    {
      prop.Value.lpszA = xstrdup (string);
      hr = HrSetOneProp (message, &prop);	
      xfree (prop.Value.lpszA);
    }
  else
    {
      proparray.cValues = 1;
      proparray.aulPropTag[0] = prop.ulPropTag;
      hr = message->DeleteProps (&proparray, NULL);
    }
  if (hr)
    {
      log_error ("%s:%s: can't %s %s property: hr=%#lx\n",
                 SRCNAME, __func__, string?"set":"delete",
                 "GpgOL Draft Info", hr); 
      return -1;
    }

  return 0;
}


/* Return the MIME info as an allocated string.  Will never return
   NULL.  */
char *
mapi_get_mime_info (LPMESSAGE msg)
{
  HRESULT hr;
  LPSPropValue propval = NULL;
  ULONG tag;
  char *retstr;

  if (get_gpgolmimeinfo_tag (msg, &tag) )
    return xstrdup ("[Error getting tag for MIME info]");
  hr = HrGetOneProp ((LPMAPIPROP)msg, tag, &propval);
  if (FAILED (hr))
    return xstrdup ("");
  if (PROP_TYPE (propval->ulPropTag) == PT_STRING8)
    retstr = xstrdup (propval->Value.lpszA);
  else
    retstr = xstrdup ("[MIME info has an invalid type]");

  MAPIFreeBuffer (propval);
  return retstr;
}


/* Helper around mapi_get_gpgol_draft_info to avoid
   the string handling.
   Return values are:
   0 -> Do nothing
   1 -> Encrypt
   2 -> Sign
   3 -> Encrypt & Sign*/
int
get_gpgol_draft_info_flags (LPMESSAGE message)
{
  char *buf = mapi_get_gpgol_draft_info (message);
  int ret = 0;
  if (!buf)
    {
      return 0;
    }
  if (buf[0] == 'E')
    {
      ret |= 1;
    }
  if (buf[1] == 'S')
    {
      ret |= 2;
    }
  xfree (buf);
  return ret;
}

/* Sets the draft info flags. Protocol is always Auto.
   flags should be the same as defined by
   get_gpgol_draft_info_flags
*/
int
set_gpgol_draft_info_flags (LPMESSAGE message, int flags)
{
  char buf[4];
  buf[3] = '\0';
  buf[2] = 'A'; /* Protocol */
  buf[1] = flags & 2 ? 'S' : 's';
  buf[0] = flags & 1 ? 'E' : 'e';

  return mapi_set_gpgol_draft_info (message, buf);
}


/* Helper for mapi_get_msg_content_type() */
static int
get_message_content_type_cb (void *dummy_arg,
                             rfc822parse_event_t event, rfc822parse_t msg)
{
  (void)dummy_arg;
  (void)msg;

  if (event == RFC822PARSE_T2BODY)
    return 42; /* Hack to stop the parsing after having read the
                  outer headers. */
  return 0;
}


/* Return Content-Type of the current message.  This one is taken
   directly from the rfc822 header.  If R_PROTOCOL is not NULL a
   string with the protocol parameter will be stored at this address,
   if no protocol is given NULL will be stored.  If R_SMTYPE is not
   NULL a string with the smime-type parameter will be stored there.
   Caller must release all returned strings.  */
char *
mapi_get_message_content_type (LPMESSAGE message,
                               char **r_protocol, char **r_smtype)
{
  HRESULT hr;
  LPSPropValue propval = NULL;
  rfc822parse_t msg;
  const char *header_lines, *s;
  rfc822parse_field_t ctx;
  size_t length;
  char *retstr = NULL;
  
  if (r_protocol)
    *r_protocol = NULL;
  if (r_smtype)
    *r_smtype = NULL;

  hr = HrGetOneProp ((LPMAPIPROP)message,
                     PR_TRANSPORT_MESSAGE_HEADERS_A, &propval);
  if (FAILED (hr))
    {
      log_error ("%s:%s: error getting the headers lines: hr=%#lx",
                 SRCNAME, __func__, hr);
      return NULL; 
    }
  if (PROP_TYPE (propval->ulPropTag) != PT_STRING8)
    {
      /* As per rfc822, header lines must be plain ascii, so no need
         to cope with unicode etc. */
      log_error ("%s:%s: proptag=%#lx not supported\n",
                 SRCNAME, __func__, propval->ulPropTag);
      MAPIFreeBuffer (propval);
      return NULL;
    }
  header_lines = propval->Value.lpszA;

  /* Read the headers into an rfc822 object. */
  msg = rfc822parse_open (get_message_content_type_cb, NULL);
  if (!msg)
    {
      log_error ("%s:%s: rfc822parse_open failed\n", SRCNAME, __func__);
      MAPIFreeBuffer (propval);
      return NULL;
    }
  
  while ((s = strchr (header_lines, '\n')))
    {
      length = (s - header_lines);
      if (length && s[-1] == '\r')
        length--;
      rfc822parse_insert (msg, (const unsigned char*)header_lines, length);
      header_lines = s+1;
    }
  
  /* Parse the content-type field. */
  ctx = rfc822parse_parse_field (msg, "Content-Type", -1);
  if (ctx)
    {
      const char *s1, *s2;
      s1 = rfc822parse_query_media_type (ctx, &s2);
      if (s1)
        {
          retstr = (char*)xmalloc (strlen (s1) + 1 + strlen (s2) + 1);
          strcpy (stpcpy (stpcpy (retstr, s1), "/"), s2);

          if (r_protocol)
            {
              s = rfc822parse_query_parameter (ctx, "protocol", 0);
              if (s)
                *r_protocol = xstrdup (s);
            }
          if (r_smtype)
            {
              s = rfc822parse_query_parameter (ctx, "smime-type", 0);
              if (s)
                *r_smtype = xstrdup (s);
            }
        }
      rfc822parse_release_field (ctx);
    }

  rfc822parse_close (msg);
  MAPIFreeBuffer (propval);
  return retstr;
}


/* Returns True if MESSAGE has a GpgOL Last Decrypted property with any value.
   This indicates that there should be no PR_BODY tag.  */
int
mapi_has_last_decrypted (LPMESSAGE message)
{
  HRESULT hr;
  LPSPropValue propval = NULL;
  ULONG tag;
  int yes = 0;
  
  if (get_gpgollastdecrypted_tag (message, &tag) )
    return 0; /* No.  */
  hr = HrGetOneProp ((LPMAPIPROP)message, tag, &propval);
  if (FAILED (hr))
    return 0; /* No.  */  
  
  if (PROP_TYPE (propval->ulPropTag) == PT_BINARY)
    yes = 1;

  MAPIFreeBuffer (propval);
  return yes;
}


/* Returns True if MESSAGE has a GpgOL Last Decrypted property and
   that matches the current session. */
int
mapi_test_last_decrypted (LPMESSAGE message)
{
  HRESULT hr;
  LPSPropValue propval = NULL;
  ULONG tag;
  int yes = 0;

  if (get_gpgollastdecrypted_tag (message, &tag) )
    goto leave; /* No.  */
  hr = HrGetOneProp ((LPMAPIPROP)message, tag, &propval);
  if (FAILED (hr))
    goto leave; /* No.  */  

  if (PROP_TYPE (propval->ulPropTag) == PT_BINARY
      && propval->Value.bin.cb == 8
      && !memcmp (propval->Value.bin.lpb, get_64bit_session_marker (), 8) )
    yes = 1;

  MAPIFreeBuffer (propval);
 leave:
  log_debug ("%s:%s: message decrypted during this session: %s\n",
             SRCNAME, __func__, yes?"yes":"no");
  return yes;
}



/* Helper for mapi_get_gpgol_body_attachment.  */
static int
has_gpgol_body_name (LPATTACH obj)
{
  HRESULT hr;
  LPSPropValue propval;
  int yes = 0;

  hr = HrGetOneProp ((LPMAPIPROP)obj, PR_ATTACH_FILENAME, &propval);
  if (FAILED(hr))
    return 0;

  if ( PROP_TYPE (propval->ulPropTag) == PT_UNICODE)
    {
      if (!wcscmp (propval->Value.lpszW, L"gpgol000.txt"))
        yes = 1;
      else if (!wcscmp (propval->Value.lpszW, L"gpgol000.htm"))
        yes = 2;
    }
  else if ( PROP_TYPE (propval->ulPropTag) == PT_STRING8)
    {
      if (!strcmp (propval->Value.lpszA, "gpgol000.txt"))
        yes = 1;
      else if (!strcmp (propval->Value.lpszA, "gpgol000.htm"))
        yes = 2;
    }
  MAPIFreeBuffer (propval);
  return yes;
}

/* Helper to check whether the file name of OBJ is "smime.p7m".
   Returns on true if so.  */
static int
has_smime_filename (LPATTACH obj)
{
  HRESULT hr;
  LPSPropValue propval;
  int yes = 0;

  hr = HrGetOneProp ((LPMAPIPROP)obj, PR_ATTACH_FILENAME, &propval);
  if (FAILED(hr))
    {
      hr = HrGetOneProp ((LPMAPIPROP)obj, PR_ATTACH_LONG_FILENAME, &propval);
      if (FAILED(hr))
        return 0;
    }

  if ( PROP_TYPE (propval->ulPropTag) == PT_UNICODE)
    {
      if (!wcscmp (propval->Value.lpszW, L"smime.p7m"))
        yes = 1;
    }
  else if ( PROP_TYPE (propval->ulPropTag) == PT_STRING8)
    {
      if (!strcmp (propval->Value.lpszA, "smime.p7m"))
        yes = 1;
    }
  MAPIFreeBuffer (propval);
  return yes;
}


/* Return the content of the body attachment of MESSAGE.  The body
   attachment is a hidden attachment created by us for later display.
   If R_NBYTES is not NULL the number of bytes in the returned buffer
   is stored there.  If R_ISHTML is not NULL a flag indicating whether
   the HTML is html formatted is stored there.  If R_PROTECTED is not
   NULL a flag indicating whether the message was protected is stored
   there.  If no body attachment can be found or on any other error an
   error codes is returned and NULL is stored at R_BODY.  Caller must
   free the returned string.  If NULL is passed for R_BODY, the
   function will only test whether a body attachment is available and
   return an error code if not.  R_IS_HTML and R_PROTECTED are not
   defined in this case.  */
int
mapi_get_gpgol_body_attachment (LPMESSAGE message, 
                                char **r_body, size_t *r_nbytes, 
                                int *r_ishtml, int *r_protected)
{    
  HRESULT hr;
  SizedSPropTagArray (1L, propAttNum) = { 1L, {PR_ATTACH_NUM} };
  LPMAPITABLE mapitable;
  LPSRowSet   mapirows;
  unsigned int pos, n_attach;
  ULONG moss_tag;
  char *body = NULL;
  int bodytype;
  int found = 0;

  if (r_body)
    *r_body = NULL;
  if (r_ishtml)
    *r_ishtml = 0;
  if (r_protected)
    *r_protected = 0;

  if (get_gpgolattachtype_tag (message, &moss_tag) )
    return -1;

  hr = message->GetAttachmentTable (0, &mapitable);
  if (FAILED (hr))
    {
      log_debug ("%s:%s: GetAttachmentTable failed: hr=%#lx",
                 SRCNAME, __func__, hr);
      return -1;
    }
      
  hr = HrQueryAllRows (mapitable, (LPSPropTagArray)&propAttNum,
                       NULL, NULL, 0, &mapirows);
  if (FAILED (hr))
    {
      log_debug ("%s:%s: HrQueryAllRows failed: hr=%#lx",
                 SRCNAME, __func__, hr);
      mapitable->Release ();
      return -1;
    }
  n_attach = mapirows->cRows > 0? mapirows->cRows : 0;
  if (!n_attach)
    {
      FreeProws (mapirows);
      mapitable->Release ();
      log_debug ("%s:%s: No attachments at all", SRCNAME, __func__);
      return -1;
    }
  log_debug ("%s:%s: message has %u attachments\n",
             SRCNAME, __func__, n_attach);

  for (pos=0; pos < n_attach; pos++) 
    {
      LPATTACH att;

      if (mapirows->aRow[pos].cValues < 1)
        {
          log_error ("%s:%s: invalid row at pos %d", SRCNAME, __func__, pos);
          continue;
        }
      if (mapirows->aRow[pos].lpProps[0].ulPropTag != PR_ATTACH_NUM)
        {
          log_error ("%s:%s: invalid prop at pos %d", SRCNAME, __func__, pos);
          continue;
        }
      hr = message->OpenAttach (mapirows->aRow[pos].lpProps[0].Value.l,
                                NULL, MAPI_BEST_ACCESS, &att);	
      if (FAILED (hr))
        {
          log_error ("%s:%s: can't open attachment %d (%ld): hr=%#lx",
                     SRCNAME, __func__, pos, 
                     mapirows->aRow[pos].lpProps[0].Value.l, hr);
          continue;
        }
      if ((bodytype=has_gpgol_body_name (att))
           && get_gpgolattachtype (att, moss_tag) == ATTACHTYPE_FROMMOSS)
        {
          found = 1;
          if (!r_body)
            ; /* Body content has not been requested. */
          else if (opt.body_as_attachment && !mapi_test_attach_hidden (att))
            {
              /* The body is to be shown as an attachment. */
              body = native_to_utf8 
                (bodytype == 2
                 ? ("[Open the attachment \"gpgol000.htm\""
                    " to view the message.]")
                 : ("[Open the attachment \"gpgol000.txt\""
                    " to view the message.]"));
              found = 1;
            }
          else
            {
              char *charset;
              
              if (get_attach_method (att) == ATTACH_BY_VALUE)
                body = attach_to_buffer (att, r_nbytes, 1, r_protected);
              if (body && (charset = mapi_get_gpgol_charset ((LPMESSAGE)att)))
                {
                  /* We only support transcoding from Latin-1 for now.  */
                  if (strcmp (charset, "iso-8859-1") 
                      && !strcmp (charset, "latin-1"))
                    log_debug ("%s:%s: Using Latin-1 instead of %s",
                               SRCNAME, __func__, charset);
                  xfree (charset);
                  charset = latin1_to_utf8 (body);
                  xfree (body);
                  body = charset;
                }
            }
          att->Release ();
          if (r_ishtml)
            *r_ishtml = (bodytype == 2);
          break;
        }
      att->Release ();
    }
  FreeProws (mapirows);
  mapitable->Release ();
  if (!found)
    {
      log_error ("%s:%s: no suitable body attachment found", SRCNAME,__func__);
      if (r_body)
        *r_body = native_to_utf8 
          (_("[The content of this message is not visible"
             " due to an processing error in GpgOL.]"));
      return -1;
    }

  if (r_body)
    *r_body = body;
  else
    xfree (body);  /* (Should not happen.)  */
  return 0;
}


/* Delete a possible body atatchment.  Returns true if an atatchment
   has been deleted.  */
int
mapi_delete_gpgol_body_attachment (LPMESSAGE message)
{    
  HRESULT hr;
  SizedSPropTagArray (1L, propAttNum) = { 1L, {PR_ATTACH_NUM} };
  LPMAPITABLE mapitable;
  LPSRowSet   mapirows;
  unsigned int pos, n_attach;
  ULONG moss_tag;
  int found = 0;

  if (get_gpgolattachtype_tag (message, &moss_tag) )
    return 0;

  hr = message->GetAttachmentTable (0, &mapitable);
  if (FAILED (hr))
    {
      log_debug ("%s:%s: GetAttachmentTable failed: hr=%#lx",
                 SRCNAME, __func__, hr);
      return 0;
    }
      
  hr = HrQueryAllRows (mapitable, (LPSPropTagArray)&propAttNum,
                       NULL, NULL, 0, &mapirows);
  if (FAILED (hr))
    {
      log_debug ("%s:%s: HrQueryAllRows failed: hr=%#lx",
                 SRCNAME, __func__, hr);
      mapitable->Release ();
      return 0;
    }
  n_attach = mapirows->cRows > 0? mapirows->cRows : 0;
  if (!n_attach)
    {
      FreeProws (mapirows);
      mapitable->Release ();
      return 0; /* No Attachments.  */
    }

  for (pos=0; pos < n_attach; pos++) 
    {
      LPATTACH att;

      if (mapirows->aRow[pos].cValues < 1)
        {
          log_error ("%s:%s: invalid row at pos %d", SRCNAME, __func__, pos);
          continue;
        }
      if (mapirows->aRow[pos].lpProps[0].ulPropTag != PR_ATTACH_NUM)
        {
          log_error ("%s:%s: invalid prop at pos %d", SRCNAME, __func__, pos);
          continue;
        }
      hr = message->OpenAttach (mapirows->aRow[pos].lpProps[0].Value.l,
                                NULL, MAPI_BEST_ACCESS, &att);	
      if (FAILED (hr))
        {
          log_error ("%s:%s: can't open attachment %d (%ld): hr=%#lx",
                     SRCNAME, __func__, pos, 
                     mapirows->aRow[pos].lpProps[0].Value.l, hr);
          continue;
        }
      if (has_gpgol_body_name (att)
          && get_gpgolattachtype (att, moss_tag) == ATTACHTYPE_FROMMOSS)
        {
          att->Release ();
          hr = message->DeleteAttach (mapirows->aRow[pos].lpProps[0].Value.l,
                                      0, NULL, 0);
          if (hr)
            log_error ("%s:%s: DeleteAttach failed: hr=%#lx\n",
                         SRCNAME, __func__, hr); 
          else
            {
              log_debug ("%s:%s: body attachment deleted\n", 
                         SRCNAME, __func__); 
              found = 1;
              
            }
          break;
        }
      att->Release ();
    }
  FreeProws (mapirows);
  mapitable->Release ();
  return found;
}


/* Copy the attachment ITEM of the message MESSAGE verbatim to the
   PR_BODY property.  Returns 0 on success.  This function does not
   call SaveChanges. */
int
mapi_attachment_to_body (LPMESSAGE message, mapi_attach_item_t *item)
{
  int result = -1;
  HRESULT hr; 
  LPATTACH att = NULL;
  LPSTREAM instream = NULL;
  LPSTREAM outstream = NULL;
  LPUNKNOWN punk;

  if (!message || !item || item->end_of_table || item->mapipos == -1)
    return -1; /* Error.  */

  hr = message->OpenAttach (item->mapipos, NULL, MAPI_BEST_ACCESS, &att);
  if (FAILED (hr))
    {
      log_error ("%s:%s: can't open attachment at %d: hr=%#lx",
                 SRCNAME, __func__, item->mapipos, hr);
      goto leave;
    }
  if (item->method != ATTACH_BY_VALUE)
    {
      log_error ("%s:%s: attachment: method not supported", SRCNAME, __func__);
      goto leave;
    }

  hr = att->OpenProperty (PR_ATTACH_DATA_BIN, &IID_IStream, 
                          0, 0, (LPUNKNOWN*) &instream);
  if (FAILED (hr))
    {
      log_error ("%s:%s: can't open data stream of attachment: hr=%#lx",
                 SRCNAME, __func__, hr);
      goto leave;
    }


  punk = (LPUNKNOWN)outstream;
  hr = message->OpenProperty (PR_BODY_A, &IID_IStream, 0,
                              MAPI_CREATE|MAPI_MODIFY, &punk);
  if (FAILED (hr))
    {
      log_error ("%s:%s: can't open body stream for update: hr=%#lx",
                 SRCNAME, __func__, hr);
      goto leave;
    }
  outstream = (LPSTREAM)punk;

  {
    ULARGE_INTEGER cb;
    cb.QuadPart = 0xffffffffffffffffll;
    hr = instream->CopyTo (outstream, cb, NULL, NULL);
  }
  if (hr)
    {
      log_error ("%s:%s: can't copy streams: hr=%#lx\n",
                 SRCNAME, __func__, hr); 
      goto leave;
    }
  hr = outstream->Commit (0);
  if (hr)
    {
      log_error ("%s:%s: commiting output stream failed: hr=%#lx",
                 SRCNAME, __func__, hr);
      goto leave;
    }
  result = 0;
  
 leave:
  if (outstream)
    {
      if (result)
        outstream->Revert ();
      outstream->Release ();
    }
  if (instream)
    instream->Release ();
  if (att)
    att->Release ();
  return result;
}


