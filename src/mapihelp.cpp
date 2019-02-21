/* mapihelp.cpp - Helper functions for MAPI
 * Copyright (C) 2005, 2007, 2008 g10 Code GmbH
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
#include "mapihelp.h"
#include "parsetlv.h"
#include "oomhelp.h"

#include <string>

#ifndef CRYPT_E_STREAM_INSUFFICIENT_DATA
#define CRYPT_E_STREAM_INSUFFICIENT_DATA 0x80091011
#endif
#ifndef CRYPT_E_ASN1_BADTAG
#define CRYPT_E_ASN1_BADTAG 0x8009310B
#endif


static int get_attach_method (LPATTACH obj);
static int has_smime_filename (LPATTACH obj);
static char *get_attach_mime_tag (LPATTACH obj);




/* Print a MAPI property to the log stream. */
void
log_mapi_property (LPMESSAGE message, ULONG prop, const char *propname)
{
  TSTART;
  HRESULT hr;
  LPSPropValue propval = NULL;
  size_t keylen;
  void *key;
  char *buf;

  if (!message)
    {
      TRETURN; /* No message: Nop. */
    }

  hr = HrGetOneProp ((LPMAPIPROP)message, prop, &propval);
  if (FAILED (hr))
    {
      log_debug ("%s:%s: HrGetOneProp(%s) failed: hr=%#lx\n",
                 SRCNAME, __func__, propname, hr);
      TRETURN;
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
      TRETURN;
    }
  MAPIFreeBuffer (propval);
  TRETURN;
}


/* Helper to create a named property. */
static ULONG 
create_gpgol_tag (LPMESSAGE message, const wchar_t *name, const char *func)
{
  TSTART;
  HRESULT hr;
  LPSPropTagArray proparr = NULL;
  MAPINAMEID mnid, *pmnid;
  wchar_t *propname = xwcsdup (name);
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
  xfree (propname);
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
    
  TRETURN result;
}


/* Return the property tag for GpgOL Msg Class. */
int 
get_gpgolmsgclass_tag (LPMESSAGE message, ULONG *r_tag)
{
  TSTART;
  if (!(*r_tag = create_gpgol_tag (message, L"GpgOL Msg Class", __func__)))
    {
      TRETURN -1;
    }
  *r_tag |= PT_STRING8;
  TRETURN 0;
}

/* Return the property tag for GpgOL Old Msg Class.  The Old Msg Class
   saves the message class as seen before we changed it the first
   time. */
int 
get_gpgololdmsgclass_tag (LPMESSAGE message, ULONG *r_tag)
{
  TSTART;
  if (!(*r_tag = create_gpgol_tag (message, L"GpgOL Old Msg Class", __func__)))
    {
      TRETURN -1;
    }
  *r_tag |= PT_STRING8;
  TRETURN 0;
}


/* Return the property tag for GpgOL Attach Type. */
int 
get_gpgolattachtype_tag (LPMESSAGE message, ULONG *r_tag)
{
  TSTART;
  if (!(*r_tag = create_gpgol_tag (message, L"GpgOL Attach Type", __func__)))
    {
      TRETURN -1;
    }
  *r_tag |= PT_LONG;
  TRETURN 0;
}


/* Return the property tag for GpgOL Protect IV. */
int 
get_gpgolprotectiv_tag (LPMESSAGE message, ULONG *r_tag)
{
  TSTART;
  if (!(*r_tag = create_gpgol_tag (message, L"GpgOL Protect IV", __func__)))
    {
      TRETURN -1;
    }
  *r_tag |= PT_BINARY;
  TRETURN 0;
}

/* Return the property tag for GpgOL Last Decrypted. */
int 
get_gpgollastdecrypted_tag (LPMESSAGE message, ULONG *r_tag)
{
  TSTART;
  if (!(*r_tag = create_gpgol_tag (message, L"GpgOL Last Decrypted",__func__)))
    {
      TRETURN -1;
    }
  *r_tag |= PT_BINARY;
  TRETURN 0;
}


/* Return the property tag for GpgOL MIME structure. */
int 
get_gpgolmimeinfo_tag (LPMESSAGE message, ULONG *r_tag)
{
  TSTART;
  if (!(*r_tag = create_gpgol_tag (message, L"GpgOL MIME Info", __func__)))
    {
      TRETURN -1;
    }
  *r_tag |= PT_STRING8;
  TRETURN 0;
}


/* Return the property tag for GpgOL Charset. */
int 
get_gpgolcharset_tag (LPMESSAGE message, ULONG *r_tag)
{
  TSTART;
  if (!(*r_tag = create_gpgol_tag (message, L"GpgOL Charset", __func__)))
    {
      TRETURN -1;
    }
  *r_tag |= PT_STRING8;
  TRETURN 0;
}


/* Return the property tag for GpgOL Draft Info.  */
int 
get_gpgoldraftinfo_tag (LPMESSAGE message, ULONG *r_tag)
{
  TSTART;
  if (!(*r_tag = create_gpgol_tag (message, L"GpgOL Draft Info", __func__)))
    {
      TRETURN -1;
    }
  *r_tag |= PT_STRING8;
  TRETURN 0;
}


/* Return the tag of the Internet Charset Body property which seems to
   hold the PR_BODY as received and thus before charset
   conversion.  */
int
get_internetcharsetbody_tag (LPMESSAGE message, ULONG *r_tag)
{
  TSTART;
  HRESULT hr;
  LPSPropTagArray proparr = NULL;
  MAPINAMEID mnid, *pmnid;	
  /* {4E3A7680-B77A-11D0-9DA5-00C04FD65685} */
  GUID guid = {0x4E3A7680, 0xB77A, 0x11D0, {0x9D, 0xA5, 0x00, 0xC0,
                                            0x4F, 0xD6, 0x56, 0x85}};
  wchar_t propname[] = L"Internet Charset Body";
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
      log_debug ("%s:%s: can't get the Internet Charset Body property:"
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
  
  TRETURN result;
}

/* Return the property tag for GpgOL UUID Info.  */
static int
get_gpgoluid_tag (LPMESSAGE message, ULONG *r_tag)
{
  TSTART;
  if (!(*r_tag = create_gpgol_tag (message, L"GpgOL UID", __func__)))
    {
      TRETURN -1;
    }
  *r_tag |= PT_UNICODE;
  TRETURN 0;
}

char *
mapi_get_uid (LPMESSAGE msg)
{
  TSTART;
  /* If the UUID is not in OOM maybe we find it in mapi. */
  if (!msg)
    {
      log_error ("%s:%s: Called without message",
                 SRCNAME, __func__);
      TRETURN NULL;
    }
  ULONG tag;
  if (get_gpgoluid_tag (msg, &tag))
    {
      log_debug ("%s:%s: Failed to get tag for '%p'",
                 SRCNAME, __func__, msg);
      TRETURN NULL;
    }
  LPSPropValue propval = NULL;
  HRESULT hr = HrGetOneProp ((LPMAPIPROP)msg, tag, &propval);
  if (hr)
    {
      log_debug ("%s:%s: Failed to get prop for '%p'",
                 SRCNAME, __func__, msg);
      TRETURN NULL;
    }
  char *ret = NULL;
  if (PROP_TYPE (propval->ulPropTag) == PT_UNICODE)
    {
      ret = wchar_to_utf8 (propval->Value.lpszW);
      log_debug ("%s:%s: Fund uuid in MAPI for %p",
                 SRCNAME, __func__, msg);
    }
  else if (PROP_TYPE (propval->ulPropTag) == PT_STRING8)
    {
      ret = xstrdup (propval->Value.lpszA);
      log_debug ("%s:%s: Fund uuid in MAPI for %p",
                 SRCNAME, __func__, msg);
    }
  MAPIFreeBuffer (propval);
  TRETURN ret;
}


/* A Wrapper around the SaveChanges method.  This function should be
   called indirect through the mapi_save_changes macro.  Returns 0 on
   success. */
int
mapi_do_save_changes (LPMESSAGE message, ULONG flags, int only_del_body,
                      const char *dbg_file, const char *dbg_func)
{
  TSTART;
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
      int i;
      for (i = 0, hr = 0; hr && i < 10; i++)
        {
          hr = message->SaveChanges (flags);
          if (hr)
            {
              log_debug ("%s:%s: Failed try to save.",
                         SRCNAME, __func__);
              Sleep (1000);
            }
        }
      if (hr)
        {
          log_error ("%s:%s: SaveChanges(%lu) failed: hr=%#lx\n",
                     log_srcname (dbg_file), dbg_func,
                     (unsigned long)flags, hr); 
          TRETURN -1;
        }
    }
  
  TRETURN 0;
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
    {
      TRETURN -1;
    }

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

  TRETURN result;
}


/* Return the headers as ASCII string. Returns empty
   string on failure. */
std::string
mapi_get_header (LPMESSAGE message)
{
  TSTART;
  HRESULT hr;
  LPSTREAM stream = nullptr;
  ULONG bRead;
  std::string ret;

  if (!message)
    {
      TRETURN ret;
    }

  hr = gpgol_openProperty (message, PR_TRANSPORT_MESSAGE_HEADERS_A, &IID_IStream, 0, 0,
                              (LPUNKNOWN*)&stream);
  if (hr)
    {
      log_debug ("%s:%s: OpenProperty failed: hr=%#lx", SRCNAME, __func__, hr);
      TRETURN ret;
    }

  char buf[8192];
  while ((hr = stream->Read (buf, 8192, &bRead)) == S_OK ||
         hr == S_FALSE)
    {
      if (!bRead)
        {
          // EOF
          break;
        }
      ret += std::string (buf, bRead);
    }
  gpgol_release (stream);
  TRETURN ret;
}


/* Return the body as a new IStream object.  Returns NULL on failure.
   The stream Returns the body as an ASCII stream (Use mapi_get_body
   for an UTF-8 value).  */
LPSTREAM
mapi_get_body_as_stream (LPMESSAGE message)
{
  TSTART;
  HRESULT hr;
  ULONG tag;
  LPSTREAM stream;

  if (!message)
    {
      TRETURN NULL;
    }

  if (!get_internetcharsetbody_tag (message, &tag) )
    {
      /* The store knows about the Internet Charset Body property,
         thus try to get the body from this property if it exists.  */

      hr = gpgol_openProperty (message, tag, &IID_IStream, 0, 0,
                                  (LPUNKNOWN*)&stream);
      if (!hr)
        {
          TRETURN stream;
        }

      log_debug ("%s:%s: OpenProperty tag=%lx failed: hr=%#lx",
                 SRCNAME, __func__, tag, hr);
    }

  /* We try to get it as an ASCII body.  If this fails we would either
     need to implement some kind of stream filter to translated to
     utf-8 or read everyting into a memory buffer and [provide an
     istream from that memory buffer.  */
  hr = gpgol_openProperty (message, PR_BODY_A, &IID_IStream, 0, 0,
                              (LPUNKNOWN*)&stream);
  if (hr)
    {
      log_debug ("%s:%s: OpenProperty failed: hr=%#lx", SRCNAME, __func__, hr);
      TRETURN NULL;
    }

  TRETURN stream;
}



/* Return the body of the message in an allocated buffer.  The buffer
   is guaranteed to be Nul terminated.  The actual length (ie. the
   strlen()) will be stored at R_NBYTES.  The body will be returned in
   UTF-8 encoding. Returns NULL if no body is available.  */
char *
mapi_get_body (LPMESSAGE message, size_t *r_nbytes)
{
  TSTART;
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
      hr = gpgol_openProperty (message, PR_BODY, &IID_IStream, 0, 0,
                                  (LPUNKNOWN*)&stream);
      if (hr)
        {
          log_debug ("%s:%s: OpenProperty failed: hr=%#lx",
                     SRCNAME, __func__, hr);
          TRETURN NULL;
        }
      
      hr = stream->Stat (&statInfo, STATFLAG_NONAME);
      if (hr)
        {
          log_debug ("%s:%s: Stat failed: hr=%#lx", SRCNAME, __func__, hr);
          gpgol_release (stream);
          TRETURN NULL;
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
          gpgol_release (stream);
          TRETURN NULL;
        }
      body[nread] = 0;
      body[nread+1] = 0;
      if (nread != statInfo.cbSize.QuadPart)
        {
          log_debug ("%s:%s: not enough bytes returned\n", SRCNAME, __func__);
          xfree (body);
          gpgol_release (stream);
          TRETURN NULL;
        }
      gpgol_release (stream);
      
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
  TRETURN body;
}



/* Look at the body of the MESSAGE and try to figure out whether this
   is a supported PGP message.  Returns the new message class or NULL
   if it does not look like a PGP message.

   If r_nobody is not null it is set to true if no body was found.
   */
static char *
get_msgcls_from_pgp_lines (LPMESSAGE message, bool *r_nobody = nullptr)
{
  TSTART;
  HRESULT hr;
  LPSTREAM stream;
  STATSTG statInfo;
  ULONG nread;
  size_t nbytes;
  char *body = NULL;
  char *p;
  char *msgcls = NULL;
  int is_wchar = 0;

  if (r_nobody)
    {
      *r_nobody = false;
    }

  stream = mapi_get_body_as_stream (message);
  if (!stream)
    {
      log_debug ("%s:%s: Failed to get body ASCII stream.",
                 SRCNAME, __func__);
      hr = gpgol_openProperty (message, PR_BODY_W, &IID_IStream, 0, 0,
                                  (LPUNKNOWN*)&stream);
      if (hr)
        {
          log_error ("%s:%s: Failed to get  w_body stream. : hr=%#lx",
                     SRCNAME, __func__, hr);
          if (r_nobody)
            {
              *r_nobody = true;
            }
          TRETURN NULL;
        }
      else
        {
          is_wchar = 1;
        }
    }

  hr = stream->Stat (&statInfo, STATFLAG_NONAME);
  if (hr)
    {
      log_debug ("%s:%s: Stat failed: hr=%#lx", SRCNAME, __func__, hr);
      gpgol_release (stream);
      TRETURN NULL;
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
      gpgol_release (stream);
      TRETURN NULL;
    }
  body[nread] = 0;
  body[nread+1] = 0;
  if (nread != nbytes)
    {
      log_debug ("%s:%s: not enough bytes returned\n", SRCNAME, __func__);

      xfree (body);
      gpgol_release (stream);
      TRETURN NULL;
    }
  gpgol_release (stream);

  if (is_wchar)
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
  for (p=body; p && *p; p = ((p=strchr (p+1, '\n')) ? (p+1) : NULL))
    {
      if (!strncmp (p, "-----BEGIN PGP ", 15))
        {
          /* Enabling clearsigned detection for Outlook 2010 and later
             would result in data loss as the signature is not reverted. */
          if (!strncmp (p+15, "SIGNED MESSAGE-----", 19)
              && trailing_ws_p (p+15+19))
            msgcls = xstrdup ("IPM.Note.GpgOL.ClearSigned");
          else if (!strncmp (p+15, "MESSAGE-----", 12)
                   && trailing_ws_p (p+15+12))
            msgcls = xstrdup ("IPM.Note.GpgOL.PGPMessage");
          break;
        }
      else if (!trailing_ws_p (p))
        {
          /* We have text before the message. In that case we need
             to break because some bad MUA's like Outlook do not insert
             quote characters before a replied to message. In that case
             the reply to an inline Mail from an Outlook without GpgOL
             enabled could cause the behavior that we would detect
             the original message.
          */
          log_debug ("%s:%s: Detected non whitespace %c before a PGP Marker",
                     SRCNAME, __func__, *p);
          break;
        }
    }


  xfree (body);
  TRETURN msgcls;
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
      TRETURN -1;
    }
  memdbg_addRef (mapitable);

  hr = HrQueryAllRows (mapitable, (LPSPropTagArray)&propAttNum,
                       NULL, NULL, 0, &mapirows);
  if (FAILED (hr))
    {
      log_debug ("%s:%s: HrQueryAllRows failed: hr=%#lx",
                 SRCNAME, __func__, hr);
      gpgol_release (mapitable);
      TRETURN -1;
    }
  n_attach = mapirows->cRows > 0? mapirows->cRows : 0;
  if (n_attach != 1)
    {
      FreeProws (mapirows);
      gpgol_release (mapitable);
      log_debug ("%s:%s: not just one attachment", SRCNAME, __func__);
      TRETURN -1;
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
  memdbg_addRef (att);
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
  
  hr = gpgol_openProperty (att, PR_ATTACH_DATA_BIN, &IID_IStream,
                          0, 0, (LPUNKNOWN*) &stream);
  if (FAILED (hr))
    {
      log_error ("%s:%s: can't open data stream of attachment: hr=%#lx",
                 SRCNAME, __func__, hr);
      goto leave;
    }

  memdbg_addRef (stream);
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
    gpgol_release (stream);
  if (att)
    gpgol_release (att);
  FreeProws (mapirows);
  gpgol_release (mapitable);
  TRETURN result;
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
      TRETURN NULL;
    }
      
  hr = HrQueryAllRows (mapitable, (LPSPropTagArray)&propAttNum,
                       NULL, NULL, 0, &mapirows);
  if (FAILED (hr))
    {
      log_debug ("%s:%s: HrQueryAllRows failed: hr=%#lx",
                 SRCNAME, __func__, hr);
      gpgol_release (mapitable);
      TRETURN NULL;
    }
  n_attach = mapirows->cRows > 0? mapirows->cRows : 0;
  if (n_attach < 1)
    {
      FreeProws (mapirows);
      gpgol_release (mapitable);
      log_debug ("%s:%s: less then one attachment", SRCNAME, __func__);
      TRETURN NULL;
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
  memdbg_addRef (att);
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
    gpgol_release (att);
  FreeProws (mapirows);
  gpgol_release (mapitable);
  TRETURN result;
}


/* Look at the first attachment's content type to determine the
   messageclass. */
static char *
get_msgcls_from_first_attachment (LPMESSAGE message)
{
  TSTART;
  char *ret = nullptr;
  char *attach_mime = get_first_attach_mime_tag (message);
  if (!attach_mime)
    {
      TRETURN nullptr;
    }
  if (!strcmp (attach_mime, "application/pgp-encrypted"))
    {
      ret = xstrdup ("IPM.Note.GpgOL.MultipartEncrypted");
      xfree (attach_mime);
    }
  else if (!strcmp (attach_mime, "application/pgp-signature"))
    {
      ret = xstrdup ("IPM.Note.GpgOL.MultipartSigned");
      xfree (attach_mime);
    }
  TRETURN ret;
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
  TSTART;
  char *newvalue = NULL;
  char *ct, *proto;

  ct = mapi_get_message_content_type (message, &proto, NULL);
  log_debug ("%s:%s: content type is '%s'", SRCNAME, __func__,
             ct ? ct : "null");
  if (ct && proto)
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
      else if (!strcmp (ct, "multipart/mixed")
               && !strcmp (proto, "application/pgp-encrypted"))
        {
          /* This case happens if an encrypted mail is moved
             by outlook local filter rules.
          */
          newvalue = xstrdup ("IPM.Note.GpgOL.MultipartEncrypted");
        }
      xfree (proto);
    }
  else if (ct && !strcmp (ct, "application/ms-tnef"))
    {
      /* ms-tnef can either be inline PGP or PGP/MIME. First check
         for inline and then look at the attachments if they look
         like PGP /MIME .*/
      newvalue = get_msgcls_from_pgp_lines (message);
      if (!newvalue)
        {
          /* So no PGP Inline. Lets look at the attachment. */
          newvalue = get_msgcls_from_first_attachment (message);
        }
    }
  else if (!ct || !strcmp (ct, "text/plain") ||
           !strcmp (ct, "multipart/mixed") ||
           !strcmp (ct, "multipart/alternative") ||
           !strcmp (ct, "multipart/related") ||
           !strcmp (ct, "text/html"))
    {
      bool has_no_body = false;
      /* It is quite common to have a multipart/mixed or alternative
         mail with separate encrypted PGP parts.  Look at the body to
         decide.  */
      newvalue = get_msgcls_from_pgp_lines (message, &has_no_body);

      if (!newvalue && has_no_body && ct && !strcmp (ct, "multipart/mixed"))
        {
          /* This is uncommon. But some Exchanges might break a PGP/MIME mail
             this way. Let's take a look at the attachments. Maybe it's
             a PGP/MIME mail. */
          log_debug ("%s:%s: Multipart mixed without body found. Looking at attachments.",
                     SRCNAME, __func__);
          newvalue = get_msgcls_from_first_attachment (message);
        }
    }

  xfree (ct);

  TRETURN newvalue;
}

/* Helper for mapi_change_message_class.  Returns the new message
   class as an allocated string.

   This function is used for the message class "IPM.Note.SMIME".  It
   indicates an S/MIME opaque encrypted or signed message.  This may
   also be an PGP/MIME mail. */
static char *
change_message_class_ipm_note_smime (LPMESSAGE message)
{
  TSTART;
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
      else if (ct && !strcmp (ct, "application/ms-tnef"))
        {
          /* So no PGP Inline. Lets look at the attachment. */
          char *attach_mime = get_first_attach_mime_tag (message);
          if (!attach_mime)
            {
              xfree (ct);
              xfree (proto);
              TRETURN nullptr;
            }
          if (!strcmp (attach_mime, "multipart/signed"))
            {
              newvalue = xstrdup ("IPM.Note.GpgOL.MultipartSigned");
              xfree (attach_mime);
            }
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
    newvalue = xstrdup ("IPM.Note.GpgOL.OpaqueEncrypted");

  TRETURN newvalue;
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
  TSTART;
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
      else if (!strcmp (ct, "wks.confirmation.mail"))
        {
          newvalue = xstrdup ("IPM.Note.GpgOL.WKSConfirmation");
        }
      else if (ct && !strcmp (ct, "application/ms-tnef"))
        {
          /* So no PGP Inline. Lets look at the attachment. */
          char *attach_mime = get_first_attach_mime_tag (message);
          if (!attach_mime)
            {
              xfree (ct);
              xfree (proto);
              TRETURN nullptr;
            }
          if (!strcmp (attach_mime, "multipart/signed"))
            {
              newvalue = xstrdup ("IPM.Note.GpgOL.MultipartSigned");
              xfree (attach_mime);
            }
        }
      xfree (proto);
      xfree (ct);
    }
  else
    log_debug ("%s:%s: message has no content type", SRCNAME, __func__);
  
  TRETURN newvalue;
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
  TSTART;
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
      
      if (!newvalue && (!strcmp (ct, "text/plain") ||
                        !strcmp (ct, "multipart/alternative") ||
                        !strcmp (ct, "multipart/mixed")))
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
    newvalue = xstrdup ("IPM.Note.GpgOL.OpaqueEncrypted");

  TRETURN newvalue;
}

static msgtype_t
string_to_type (const char *s)
{
  TSTART;
  if (!s || strlen (s) < 14)
    {
      TRETURN MSGTYPE_UNKNOWN;
    }
  if (!strncmp (s, "IPM.Note.GpgOL", 14) && (!s[14] || s[14] =='.'))
    {
      s += 14;
      if (!*s)
        {
          TRETURN MSGTYPE_GPGOL;
        }
      else if (!strcmp (s, ".MultipartSigned"))
        {
          TRETURN MSGTYPE_GPGOL_MULTIPART_SIGNED;
        }
      else if (!strcmp (s, ".MultipartEncrypted"))
        {
          TRETURN MSGTYPE_GPGOL_MULTIPART_ENCRYPTED;
        }
      else if (!strcmp (s, ".OpaqueSigned"))
        {
          TRETURN MSGTYPE_GPGOL_OPAQUE_SIGNED;
        }
      else if (!strcmp (s, ".OpaqueEncrypted"))
        {
          TRETURN MSGTYPE_GPGOL_OPAQUE_ENCRYPTED;
        }
      else if (!strcmp (s, ".ClearSigned"))
        {
          TRETURN MSGTYPE_GPGOL_CLEAR_SIGNED;
        }
      else if (!strcmp (s, ".PGPMessage"))
        {
          TRETURN MSGTYPE_GPGOL_PGP_MESSAGE;
        }
      else if (!strcmp (s, ".WKSConfirmation"))
        {
          TRETURN MSGTYPE_GPGOL_WKS_CONFIRMATION;
        }
      else
        log_debug ("%s:%s: message class `%s' not supported",
                   SRCNAME, __func__, s-14);
    }
  else if (!strncmp (s, "IPM.Note.SMIME", 14) && (!s[14] || s[14] =='.'))
    {
      TRETURN MSGTYPE_SMIME;
    }
  TRETURN MSGTYPE_UNKNOWN;
}


/* This function checks whether MESSAGE requires processing by us and
   adjusts the message class to our own.  By passing true for
   SYNC_OVERRIDE the actual MAPI message class will be updated to our
   own message class overide.  TRETURN true if the message was
   changed. */
int
mapi_change_message_class (LPMESSAGE message, int sync_override,
                           msgtype_t *r_type)
{
  TSTART;
  HRESULT hr;
  ULONG tag;
  SPropValue prop;
  LPSPropValue propval = NULL;
  char *newvalue = NULL;
  int need_save = 0;
  int have_override = 0;

  if (!message)
    {
      TRETURN 0; /* No message: Nop. */
    }

  if (get_gpgolmsgclass_tag (message, &tag) )
    {
      TRETURN 0; /* Ooops. */
    }

  hr = HrGetOneProp ((LPMAPIPROP)message, tag, &propval);
  if (FAILED (hr))
    {
      hr = HrGetOneProp ((LPMAPIPROP)message, PR_MESSAGE_CLASS_A, &propval);
      if (FAILED (hr))
        {
          log_error ("%s:%s: HrGetOneProp() failed: hr=%#lx\n",
                     SRCNAME, __func__, hr);
          TRETURN 0;
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
      else if (opt.enable_smime && (!strcmp (s, "IPM.Note.SMIME") ||
                                    !strcmp (s, "IPM.Note.GpgOL")))
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

          char *tmp = change_message_class_ipm_note_smime_multipartsigned
            (message);
          /* This case happens even for PGP/MIME mails but that is ok
             as we later fiddle out the protocol. But we have to
             check if this is a WKS Mail now so that we can do the
             special handling for that. */
          if (tmp && !strcmp (tmp, "IPM.Note.GpgOL.WKSConfirmation"))
            {
              newvalue = tmp;
            }
          else
            {
              xfree (tmp);
              newvalue = (char*)xmalloc (strlen (s) + 1);
              strcpy (stpcpy (newvalue, "IPM.Note.GpgOL"), s+14);
            }
        }
      else if (!strcmp (s, "IPM.Note.SMIME.MultipartSigned"))
        {
          /* This is an S/MIME message class but smime support is not
             enabled.  We need to check whether this is actually a
             PGP/MIME message.  */
          newvalue = change_message_class_ipm_note_smime_multipartsigned
            (message);
        }
      else if (sync_override && have_override
               && !strncmp (s, "IPM.Note.GpgOL", 14) && (!s[14]||s[14] =='.'))
        {
          /* In case the original message class is not yet an GpgOL
             class we set it here.  This is needed to convince Outlook
             not to do any special processing for IPM.Note.SMIME etc.  */
          LPSPropValue propval2 = NULL;

          hr = HrGetOneProp ((LPMAPIPROP)message, PR_MESSAGE_CLASS_A,
                             &propval2);
          if (!SUCCEEDED (hr))
            {
              log_debug ("%s:%s: Failed to get PR_MESSAGE_CLASS_A property.",
                         SRCNAME, __func__);
            }
          else if (PROP_TYPE (propval2->ulPropTag) != PT_STRING8)
            {
              log_debug ("%s:%s: PR_MESSAGE_CLASS_A is not string.",
                         SRCNAME, __func__);
            }
          else if (!propval2->Value.lpszA)
            {
              log_debug ("%s:%s: PR_MESSAGE_CLASS_A is null.",
                         SRCNAME, __func__);
            }
          else if (!strcmp (propval2->Value.lpszA, s))
            {
              log_debug ("%s:%s: PR_MESSAGE_CLASS_A is already the same.",
                         SRCNAME, __func__);
            }
          else
            {
              newvalue = (char*)xstrdup (s);
            }
          MAPIFreeBuffer (propval2);
        }
      else if (opt.enable_smime 
               && (!strcmp (s, "IPM.Note.Secure.CexSig")
                   || (cexenc = !strcmp (s, "IPM.Note.Secure.CexEnc"))))
        {
          newvalue = change_message_class_ipm_note_secure_cex
            (message, cexenc);
        }
      if (r_type && !newvalue)
        {
          *r_type = string_to_type (s);
        }
    }

  if (!newvalue)
    {
      log_debug ("%s:%s Message is not a crypto message.",
                 SRCNAME, __func__);
    }
  else
    {
      if (r_type)
        {
          *r_type = string_to_type (newvalue);
        }
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
                  TRETURN 0;
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
          TRETURN 0;
        }
      need_save = 1;
    }

  MAPIFreeBuffer (propval);

  if (need_save)
    {
      if (mapi_save_changes (message, KEEP_OPEN_READWRITE|FORCE_SAVE))
        {
          TRETURN 0;
        }
    }

  TRETURN 1;
}


/* Return the message class.  This function will never return NULL so
   it is mostly useful for debugging.  Caller needs to release the
   returned string.  */
char *
mapi_get_message_class (LPMESSAGE message)
{
  TSTART;
  HRESULT hr;
  LPSPropValue propval = NULL;
  char *retstr;

  if (!message)
    {
      TRETURN xstrdup ("[No message]");
    }
  
  hr = HrGetOneProp ((LPMAPIPROP)message, PR_MESSAGE_CLASS_A, &propval);
  if (FAILED (hr))
    {
      log_error ("%s:%s: HrGetOneProp() failed: hr=%#lx\n",
                 SRCNAME, __func__, hr);
      TRETURN xstrdup (hr == MAPI_E_NOT_FOUND?
                        "[No message class property]":
                        "[Error getting message class property]");
    }

  if ( PROP_TYPE (propval->ulPropTag) == PT_STRING8 )
    retstr = xstrdup (propval->Value.lpszA);
  else
    retstr = xstrdup ("[Invalid message class property]");
    
  MAPIFreeBuffer (propval);
  TRETURN retstr;
}

/* Return the old message class.  This function returns NULL if no old
   message class has been saved.  Caller needs to release the returned
   string.  */
char *
mapi_get_old_message_class (LPMESSAGE message)
{
  TSTART;
  HRESULT hr;
  ULONG tag;
  LPSPropValue propval = NULL;
  char *retstr;

  if (!message)
    {
      TRETURN NULL;
    }
  
  if (get_gpgololdmsgclass_tag (message, &tag))
    {
      TRETURN NULL;
    }

  hr = HrGetOneProp ((LPMAPIPROP)message, tag, &propval);
  if (FAILED (hr))
    {
      log_error ("%s:%s: HrGetOneProp() failed: hr=%#lx\n",
                 SRCNAME, __func__, hr);
      TRETURN NULL;
    }

  if ( PROP_TYPE (propval->ulPropTag) == PT_STRING8 )
    retstr = xstrdup (propval->Value.lpszA);
  else
    retstr = NULL;
    
  MAPIFreeBuffer (propval);
  TRETURN retstr;
}



/* Return the sender of the message.  According to the specs this is
   an UTF-8 string; we rely on that the UI server handles
   internationalized domain names.  */ 
char *
mapi_get_sender (LPMESSAGE message)
{
  TSTART;
  HRESULT hr;
  LPSPropValue propval = NULL;
  char *buf;
  char *p0, *p;
  
  if (!message)
    {
      TRETURN NULL; /* No message: Nop. */
    }

  hr = HrGetOneProp ((LPMAPIPROP)message, PR_PRIMARY_SEND_ACCT, &propval);
  if (FAILED (hr))
    {
      log_debug ("%s:%s: HrGetOneProp failed: hr=%#lx\n",
                 SRCNAME, __func__, hr);
      TRETURN NULL;
    }
    
  if (PROP_TYPE (propval->ulPropTag) != PT_UNICODE) 
    {
      log_debug ("%s:%s: HrGetOneProp Returns invalid type %lu\n",
                 SRCNAME, __func__, PROP_TYPE (propval->ulPropTag) );
      MAPIFreeBuffer (propval);
      TRETURN NULL;
    }
  
  buf = wchar_to_utf8 (propval->Value.lpszW);
  MAPIFreeBuffer (propval);
  if (!buf)
    {
      log_error ("%s:%s: error converting to utf8\n", SRCNAME, __func__);
      TRETURN NULL;
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
                 SRCNAME, __func__, anonstr (buf));
      xfree (buf);
      TRETURN NULL;
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
      log_debug ("%s:%s: orig address is `%s'\n", SRCNAME, __func__,
                 anonstr (buf));
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
  log_debug ("%s:%s: address is `%s'\n", SRCNAME, __func__, anonstr (buf));
  TRETURN buf;
}

static char *
resolve_ex_from_address (LPMESSAGE message)
{
  TSTART;
  HRESULT hr;
  char *sender_entryid;
  size_t entryidlen;
  LPMAPISESSION session;
  ULONG utype;
  LPUNKNOWN user;
  LPSPropValue propval = NULL;
  char *buf;

  if (g_ol_version_major < 14)
    {
      log_debug ("%s:%s: Not implemented for Ol < 14", SRCNAME, __func__);
      TRETURN NULL;
    }

  sender_entryid = mapi_get_binary_prop (message, PR_SENDER_ENTRYID,
                                         &entryidlen);
  if (!sender_entryid)
    {
      log_error ("%s:%s: Error: %i", SRCNAME, __func__, __LINE__);
      TRETURN NULL;
    }

  session = get_oom_mapi_session ();

  if (!session)
    {
      log_error ("%s:%s: Error: %i", SRCNAME, __func__, __LINE__);
      xfree (sender_entryid);
      TRETURN NULL;
    }

  hr = session->OpenEntry (entryidlen,  (LPENTRYID)sender_entryid,
                           &IID_IMailUser,
                           MAPI_BEST_ACCESS | MAPI_CACHE_ONLY,
                           &utype, (IUnknown**)&user);
  if (FAILED (hr))
    {
      log_debug ("%s:%s: Failed to open cached entry. Fallback to uncached.",
                 SRCNAME, __func__);
      hr = session->OpenEntry (entryidlen,  (LPENTRYID)sender_entryid,
                               &IID_IMailUser,
                               MAPI_BEST_ACCESS,
                               &utype, (IUnknown**)&user);
    }
  gpgol_release (session);

  if (FAILED (hr))
    {
      log_error ("%s:%s: Error: %i", SRCNAME, __func__, __LINE__);
      TRETURN NULL;
    }

  hr = HrGetOneProp ((LPMAPIPROP)user, PR_SMTP_ADDRESS_W, &propval);
  if (FAILED (hr))
    {
      log_error ("%s:%s: Error: %i", SRCNAME, __func__, __LINE__);
      TRETURN NULL;
    }

  if (PROP_TYPE (propval->ulPropTag) != PT_UNICODE)
    {
      log_debug ("%s:%s: HrGetOneProp Returns invalid type %lu\n",
                 SRCNAME, __func__, PROP_TYPE (propval->ulPropTag) );
      MAPIFreeBuffer (propval);
      TRETURN NULL;
    }
  buf = wchar_to_utf8 (propval->Value.lpszW);
  MAPIFreeBuffer (propval);

  TRETURN buf;
}

/* Return the from address of the message as a malloced UTF-8 string.
   returns NULL if that address is not available.  */
char *
mapi_get_from_address (LPMESSAGE message)
{
  TSTART;
  HRESULT hr;
  LPSPropValue propval = NULL;
  char *buf;
  ULONG try_props[3] = {PidTagSenderSmtpAddress_W,
                        PR_SENT_REPRESENTING_SMTP_ADDRESS_W,
                        PR_SENDER_EMAIL_ADDRESS_W};

  if (!message)
    {
      TRETURN xstrdup ("[no message]"); /* Ooops.  */
    }

  for (int i = 0; i < 3; i++)
    {
      /* We try to get different properties first as they contain
         the SMTP address of the sender. EMAIL address can be
         some LDAP stuff for exchange. */
      hr = HrGetOneProp ((LPMAPIPROP)message, try_props[i],
                         &propval);
      if (!FAILED (hr))
        {
          break;
        }
    }
   /* This is the last result that should always work but not necessarily
      contain an SMTP Address. */
  if (FAILED (hr))
    {
      log_debug ("%s:%s: HrGetOneProp failed: hr=%#lx\n",
                 SRCNAME, __func__, hr);
      TRETURN NULL;
    }

  if (PROP_TYPE (propval->ulPropTag) != PT_UNICODE) 
    {
      log_debug ("%s:%s: HrGetOneProp Returns invalid type %lu\n",
                 SRCNAME, __func__, PROP_TYPE (propval->ulPropTag) );
      MAPIFreeBuffer (propval);
      TRETURN NULL;
    }
  
  buf = wchar_to_utf8 (propval->Value.lpszW);
  MAPIFreeBuffer (propval);
  if (!buf)
    {
      log_error ("%s:%s: error converting to utf8\n", SRCNAME, __func__);
      TRETURN NULL;
    }

  if (strstr (buf, "/o="))
    {
      char *buf2;
      /* If both SMTP Address properties are not set
         we need to fallback to resolve the address
         through the address book */
      log_debug ("%s:%s: resolving exchange address.",
                 SRCNAME, __func__);
      buf2 = resolve_ex_from_address (message);
      if (buf2)
        {
          xfree (buf);
          TRETURN buf2;
        }
    }

  TRETURN buf;
}


/* Return the subject of the message as a malloced UTF-8 string.
   returns a replacement string if a subject is missing.  */
char *
mapi_get_subject (LPMESSAGE message)
{
  TSTART;
  HRESULT hr;
  LPSPropValue propval = NULL;
  char *buf;
  
  if (!message)
    {
      TRETURN xstrdup ("[no message]"); /* Ooops.  */
    }

  hr = HrGetOneProp ((LPMAPIPROP)message, PR_SUBJECT_W, &propval);
  if (FAILED (hr))
    {
      log_debug ("%s:%s: HrGetOneProp failed: hr=%#lx\n",
                 SRCNAME, __func__, hr);
      TRETURN xstrdup (_("[no subject]"));
    }
    
  if (PROP_TYPE (propval->ulPropTag) != PT_UNICODE) 
    {
      log_debug ("%s:%s: HrGetOneProp Returns invalid type %lu\n",
                 SRCNAME, __func__, PROP_TYPE (propval->ulPropTag) );
      MAPIFreeBuffer (propval);
      TRETURN xstrdup (_("[no subject]"));
    }
  
  buf = wchar_to_utf8 (propval->Value.lpszW);
  MAPIFreeBuffer (propval);
  if (!buf)
    {
      log_error ("%s:%s: error converting to utf8\n", SRCNAME, __func__);
      TRETURN xstrdup (_("[no subject]"));
    }

  TRETURN buf;
}




/* Return the message type.  This function knows only about our own
   message types.  Returns MSGTYPE_UNKNOWN for any MESSAGE we have
   no special support for.  */
msgtype_t
mapi_get_message_type (LPMESSAGE message)
{
  TSTART;
  HRESULT hr;
  ULONG tag;
  LPSPropValue propval = NULL;
  msgtype_t msgtype = MSGTYPE_UNKNOWN;

  if (!message)
    {
      TRETURN msgtype;
    }

  if (get_gpgolmsgclass_tag (message, &tag) )
    {
      TRETURN msgtype; /* Ooops */
    }

  hr = HrGetOneProp ((LPMAPIPROP)message, tag, &propval);
  if (FAILED (hr))
    {
      hr = HrGetOneProp ((LPMAPIPROP)message, PR_MESSAGE_CLASS_A, &propval);
      if (FAILED (hr))
        {
          log_error ("%s:%s: HrGetOneProp(PR_MESSAGE_CLASS) failed: hr=%#lx\n",
                     SRCNAME, __func__, hr);
          TRETURN msgtype;
        }
    }
  else
    log_debug ("%s:%s: have override message class\n", SRCNAME, __func__);

  if ( PROP_TYPE (propval->ulPropTag) == PT_STRING8 )
    {
      msgtype = string_to_type (propval->Value.lpszA);
    }
  MAPIFreeBuffer (propval);
  TRETURN msgtype;
}


/* This function is pretty useless because IConverterSession won't
   take attachments into account.  Need to write our own version.  */
int
mapi_to_mime (LPMESSAGE message, const char *filename)
{
  TSTART;
  HRESULT hr;
  LPCONVERTERSESSION session;
  LPSTREAM stream;

  hr = CoCreateInstance (CLSID_IConverterSession, NULL, CLSCTX_INPROC_SERVER,
                         IID_IConverterSession, (void **) &session);
  if (FAILED (hr))
    {
      log_error ("%s:%s: can't create new IConverterSession object: hr=%#lx",
                 SRCNAME, __func__, hr);
      TRETURN -1;
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

      gpgol_release (stream);
    }

  gpgol_release (session);
  TRETURN hr;
}


/* Return a binary property in a malloced buffer with its length stored
   at R_NBYTES.  Returns NULL on error.  */
char *
mapi_get_binary_prop (LPMESSAGE message, ULONG proptype, size_t *r_nbytes)
{
  TSTART;
  HRESULT hr;
  LPSPropValue propval = NULL;
  char *data;

  *r_nbytes = 0;
  hr = HrGetOneProp ((LPMAPIPROP)message, proptype, &propval);
  if (FAILED (hr))
    {
      log_error ("%s:%s: error getting property %#lx: hr=%#lx",
                 SRCNAME, __func__, proptype, hr);
      TRETURN NULL;
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
  TRETURN data;
}

/* Return an integer property at R_VALUE.  On error the function
   Returns -1 and sets R_VALUE to 0, on success 0 is returned.  */
int
mapi_get_int_prop (LPMAPIPROP object, ULONG proptype, LONG *r_value)
{
  TSTART;
  int rc = -1;
  HRESULT hr;
  LPSPropValue propval = NULL;

  *r_value = 0;
  hr = HrGetOneProp (object, proptype, &propval);
  if (FAILED (hr))
    {
      log_error ("%s:%s: error getting property %#lx: hr=%#lx",
                 SRCNAME, __func__, proptype, hr);
      TRETURN -1;
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
  TRETURN rc;
}


/* Return the attachment method for attachment OBJ.  In case of error
   we TRETURN 0 which happens not to be defined.  */
static int
get_attach_method (LPATTACH obj)
{
  TSTART;
  HRESULT hr;
  LPSPropValue propval = NULL;
  int method ;

  hr = HrGetOneProp ((LPMAPIPROP)obj, PR_ATTACH_METHOD, &propval);
  if (FAILED (hr))
    {
      log_error ("%s:%s: error getting attachment method: hr=%#lx",
                 SRCNAME, __func__, hr);
      TRETURN 0;
    }
  /* We don't bother checking whether we really get a PT_LONG ulong
     back; if not the system is seriously damaged and we can't do
     further harm by returning a possible random value.  */
  method = propval->Value.l;
  MAPIFreeBuffer (propval);
  TRETURN method;
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
  TSTART;
  HRESULT hr;
  LPSPropValue propval;
  char *name = NULL;

  hr = HrGetOneProp ((LPMAPIPROP)obj, PR_ATTACH_LONG_FILENAME, &propval);
  if (FAILED(hr)) 
    hr = HrGetOneProp ((LPMAPIPROP)obj, PR_ATTACH_FILENAME, &propval);
  if (FAILED(hr))
    {
      log_debug ("%s:%s: no filename property found", SRCNAME, __func__);
      TRETURN NULL;
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
  TRETURN name;
}

/* Return the content-id of the attachment OBJ or NULL if it does
   not exists.  Caller must free. */
static char *
get_attach_content_id (LPATTACH obj)
{
  TSTART;
  HRESULT hr;
  LPSPropValue propval = NULL;
  char *name;

  hr = HrGetOneProp ((LPMAPIPROP)obj, PR_ATTACH_CONTENT_ID, &propval);
  if (FAILED (hr))
    {
      if (hr != MAPI_E_NOT_FOUND)
        log_error ("%s:%s: error getting attachment's MIME tag: hr=%#lx",
                   SRCNAME, __func__, hr);
      TRETURN NULL;
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
  TRETURN name;
}

/* Return the content-type of the attachment OBJ or NULL if it does
   not exists.  Caller must free. */
static char *
get_attach_mime_tag (LPATTACH obj)
{
  TSTART;
  HRESULT hr;
  LPSPropValue propval = NULL;
  char *name;

  hr = HrGetOneProp ((LPMAPIPROP)obj, PR_ATTACH_MIME_TAG_A, &propval);
  if (FAILED (hr))
    {
      if (hr != MAPI_E_NOT_FOUND)
        log_error ("%s:%s: error getting attachment's MIME tag: hr=%#lx",
                   SRCNAME, __func__, hr);
      TRETURN NULL;
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
  TRETURN name;
}


/* Return the GpgOL Attach Type for attachment OBJ.  Tag needs to be
   the tag of that property. */
attachtype_t
get_gpgolattachtype (LPATTACH obj, ULONG tag)
{
  TSTART;
  HRESULT hr;
  LPSPropValue propval = NULL;
  attachtype_t retval;

  hr = HrGetOneProp ((LPMAPIPROP)obj, tag, &propval);
  if (FAILED (hr))
    {
      if (hr != MAPI_E_NOT_FOUND)
        log_error ("%s:%s: error getting GpgOL Attach Type: hr=%#lx",
                   SRCNAME, __func__, hr);
      TRETURN ATTACHTYPE_UNKNOWN;
    }
  retval = (attachtype_t)propval->Value.l;
  MAPIFreeBuffer (propval);
  TRETURN retval;
}


/* Gather information about attachments and TRETURN a new table of
   attachments.  Caller must release the returned table.s The routine
   will TRETURN NULL in case of an error or if no attachments are
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
    {
      TRETURN NULL;
    }

  /* Open the attachment table.  */
  hr = message->GetAttachmentTable (0, &mapitable);
  memdbg_addRef (mapitable);
  if (FAILED (hr))
    {
      log_debug ("%s:%s: GetAttachmentTable failed: hr=%#lx",
                 SRCNAME, __func__, hr);
      TRETURN NULL;
    }
      
  hr = HrQueryAllRows (mapitable, (LPSPropTagArray)&propAttNum,
                       NULL, NULL, 0, &mapirows);
  if (FAILED (hr))
    {
      log_debug ("%s:%s: HrQueryAllRows failed: hr=%#lx",
                 SRCNAME, __func__, hr);
      gpgol_release (mapitable);
      TRETURN NULL;
    }
  n_attach = mapirows->cRows > 0? mapirows->cRows : 0;

  log_debug ("%s:%s: message has %u attachments\n",
             SRCNAME, __func__, n_attach);
  if (!n_attach)
    {
      FreeProws (mapirows);
      gpgol_release (mapitable);
      TRETURN NULL;
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
      memdbg_addRef (att);
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
      table[pos].content_id = fast? NULL : get_attach_content_id (att);
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
      gpgol_release (att);
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
                     anonstr (table[pos].filename), table[pos].content_type,
                     table[pos].content_type_parms);
        }
    }

  TRETURN table;
}


/* Release a table as created by mapi_create_attach_table. */
void
mapi_release_attach_table (mapi_attach_item_t *table)
{
  TSTART;
  unsigned int pos;
  LPMAPITABLE mapitable;

  if (!table)
    {
      TRETURN;
    }

  mapitable = (LPMAPITABLE)table[0].private_mapitable;
  if (mapitable)
    gpgol_release (mapitable);
  for (pos=0; !table[pos].end_of_table; pos++)
    {
      xfree (table[pos].filename);
      xfree (table[pos].content_type);
      xfree (table[pos].content_id);
    }
  xfree (table);
  TRETURN;
}


/* Return an attachment as a new IStream object.  Returns NULL on
   failure.  If R_ATTACH is not NULL the actual attachment will not be
   released but stored at that address; the caller needs to release it
   in this case.  */
LPSTREAM
mapi_get_attach_as_stream (LPMESSAGE message, mapi_attach_item_t *item,
                           LPATTACH *r_attach)
{
  TSTART;
  HRESULT hr;
  LPATTACH att;
  LPSTREAM stream;

  if (r_attach)
    *r_attach = NULL;

  if (!item || item->end_of_table || item->mapipos == -1)
    {
      TRETURN NULL;
    }

  hr = message->OpenAttach (item->mapipos, NULL, MAPI_BEST_ACCESS, &att);
  memdbg_addRef (att);
  if (FAILED (hr))
    {
      log_error ("%s:%s: can't open attachment at %d: hr=%#lx",
                 SRCNAME, __func__, item->mapipos, hr);
      TRETURN NULL;
    }
  if (item->method != ATTACH_BY_VALUE)
    {
      log_error ("%s:%s: attachment: method not supported", SRCNAME, __func__);
      gpgol_release (att);
      TRETURN NULL;
    }

  hr = gpgol_openProperty (att, PR_ATTACH_DATA_BIN, &IID_IStream,
                          0, 0, (LPUNKNOWN*) &stream);
  if (FAILED (hr))
    {
      log_error ("%s:%s: can't open data stream of attachment: hr=%#lx",
                 SRCNAME, __func__, hr);
      gpgol_release (att);
      TRETURN NULL;
    }

  if (r_attach)
    *r_attach = att;
  else
    gpgol_release (att);

  TRETURN stream;
}


/* Return a malloced buffer with the content of the attachment. If
   R_NBYTES is not NULL the number of bytes will get stored there.
   ATT must have an attachment method of ATTACH_BY_VALUE.  Returns
   NULL on error.  If UNPROTECT is set and the appropriate crypto
   attribute is available, the function Returns the unprotected
   version of the atatchment. */
static char *
attach_to_buffer (LPATTACH att, size_t *r_nbytes)
{
  TSTART;
  HRESULT hr;
  LPSTREAM stream;
  STATSTG statInfo;
  ULONG nread;
  char *buffer;

  hr = gpgol_openProperty (att, PR_ATTACH_DATA_BIN, &IID_IStream,
                          0, 0, (LPUNKNOWN*) &stream);
  if (FAILED (hr))
    {
      log_error ("%s:%s: can't open data stream of attachment: hr=%#lx",
                 SRCNAME, __func__, hr);
      TRETURN NULL;
    }

  hr = stream->Stat (&statInfo, STATFLAG_NONAME);
  if ( hr != S_OK )
    {
      log_error ("%s:%s: Stat failed: hr=%#lx", SRCNAME, __func__, hr);
      gpgol_release (stream);
      TRETURN NULL;
    }

  /* Allocate one byte more so that we can terminate the string.  */
  buffer = (char*)xmalloc ((size_t)statInfo.cbSize.QuadPart + 1);

  hr = stream->Read (buffer, (size_t)statInfo.cbSize.QuadPart, &nread);
  if ( hr != S_OK )
    {
      log_error ("%s:%s: Read failed: hr=%#lx", SRCNAME, __func__, hr);
      xfree (buffer);
      gpgol_release (stream);
      TRETURN NULL;
    }
  if (nread != statInfo.cbSize.QuadPart)
    {
      log_error ("%s:%s: not enough bytes returned\n", SRCNAME, __func__);
      xfree (buffer);
      buffer = NULL;
    }
  gpgol_release (stream);

  /* Make sure that the buffer is a C string.  */
  if (buffer)
    buffer[nread] = 0;

  if (r_nbytes)
    *r_nbytes = nread;
  TRETURN buffer;
}



/* Return an attachment as a malloced buffer.  The size of the buffer
   will be stored at R_NBYTES.  If unprotect is true, the atatchment
   will be unprotected.  Returns NULL on failure. */
char *
mapi_get_attach (LPMESSAGE message,
                 mapi_attach_item_t *item, size_t *r_nbytes)
{
  TSTART;
  HRESULT hr;
  LPATTACH att;
  char *buffer;

  if (!item || item->end_of_table || item->mapipos == -1)
    {
      TRETURN NULL;
    }

  hr = message->OpenAttach (item->mapipos, NULL, MAPI_BEST_ACCESS, &att);
  memdbg_addRef (att);
  if (FAILED (hr))
    {
      log_error ("%s:%s: can't open attachment at %d: hr=%#lx",
                 SRCNAME, __func__, item->mapipos, hr);
      TRETURN NULL;
    }
  if (item->method != ATTACH_BY_VALUE)
    {
      log_error ("%s:%s: attachment: method not supported", SRCNAME, __func__);
      gpgol_release (att);
      TRETURN NULL;
    }

  buffer = attach_to_buffer (att, r_nbytes);
  gpgol_release (att);

  TRETURN buffer;
}


/* Mark this attachment as the original MOSS message.  We set a custom
   property as well as the hidden flag.  */
int 
mapi_mark_moss_attach (LPMESSAGE message, mapi_attach_item_t *item)
{
  TSTART;
  int retval = -1;
  HRESULT hr;
  LPATTACH att;
  SPropValue prop;

  if (!item || item->end_of_table || item->mapipos == -1)
    {
      TRETURN -1;
    }

  log_debug ("%s:%s: Marking %i as MOSS attachment",
             SRCNAME, __func__, item->mapipos);

  hr = message->OpenAttach (item->mapipos, NULL, MAPI_BEST_ACCESS, &att);
  memdbg_addRef (att);
  if (FAILED (hr))
    {
      log_error ("%s:%s: can't open attachment at %d: hr=%#lx",
                 SRCNAME, __func__, item->mapipos, hr);
      TRETURN -1;
    }

  if (get_gpgolattachtype_tag (message, &prop.ulPropTag) )
    goto leave;
  prop.Value.l = ATTACHTYPE_MOSS;
  hr = HrSetOneProp (att, &prop);	
  if (hr)
    {
      log_error ("%s:%s: can't set %s property: hr=%#lx\n",
                 SRCNAME, __func__, "GpgOL Attach Type", hr); 
      TRETURN false;
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
  gpgol_release (att);
  TRETURN retval;
}


/* If the hidden property has not been set on ATTACH, set it and save
   the changes. */
int 
mapi_set_attach_hidden (LPATTACH attach)
{
  TSTART;
  int retval = -1;
  HRESULT hr;
  LPSPropValue propval;
  SPropValue prop;

  hr = HrGetOneProp ((LPMAPIPROP)attach, PR_ATTACHMENT_HIDDEN, &propval);
  if (SUCCEEDED (hr) 
      && PROP_TYPE (propval->ulPropTag) == PT_BOOLEAN
      && propval->Value.b)
    {
      TRETURN 0;/* Already set to hidden. */
    }

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
  TRETURN retval;
}


/* Returns true if ATTACH has the hidden flag set to true.  */
int
mapi_test_attach_hidden (LPATTACH attach)
{
  TSTART;
  HRESULT hr;
  LPSPropValue propval = NULL;
  int result = 0;
  
  hr = HrGetOneProp ((LPMAPIPROP)attach, PR_ATTACHMENT_HIDDEN, &propval);
  if (FAILED (hr))
    {
      TRETURN result; /* No.  */
    }
  
  if (PROP_TYPE (propval->ulPropTag) == PT_BOOLEAN && propval->Value.b)
    result = 1; /* Yes.  */

  MAPIFreeBuffer (propval);
  TRETURN result;
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
  TSTART;
  HRESULT hr;
  SPropValue prop;

  if (get_gpgolmsgclass_tag (message, &prop.ulPropTag) )
    {
      TRETURN -1;
    }
  prop.Value.lpszA = xstrdup (name);
  hr = HrSetOneProp (message, &prop);	
  xfree (prop.Value.lpszA);
  if (hr)
    {
      log_error ("%s:%s: can't set %s property: hr=%#lx\n",
                 SRCNAME, __func__, "GpgOL Msg Class", hr); 
      TRETURN -1;
    }

  TRETURN 0;
}


/* Return the charset as assigned by GpgOL to an attachment.  This may
   return NULL it is has not been assigned or is the standard
   (UTF-8). */
char *
mapi_get_gpgol_charset (LPMESSAGE obj)
{
  TSTART;
  HRESULT hr;
  LPSPropValue propval = NULL;
  ULONG tag;
  char *retstr;

  if (get_gpgolcharset_tag (obj, &tag) )
    {
      TRETURN NULL; /* Error.  */
    }
  hr = HrGetOneProp ((LPMAPIPROP)obj, tag, &propval);
  if (FAILED (hr))
    {
      TRETURN NULL;
    }
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
  TRETURN retstr;
}


/* Set the GpgOl charset property to an attachment. 
   Note that this function does not call SaveChanges.  */
int 
mapi_set_gpgol_charset (LPMESSAGE obj, const char *charset)
{
  TSTART;
  HRESULT hr;
  SPropValue prop;
  char *p;

  /* Note that we lowercase the value and cut it to a max of 32
     characters.  The latter is required to make sure that
     HrSetOneProp will always work.  */
  if (get_gpgolcharset_tag (obj, &prop.ulPropTag) )
    {
      TRETURN -1;
    }
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
      TRETURN -1;
    }

  TRETURN 0;
}



/* Return GpgOL's draft info string as an allocated string.  If no
   draft info is available, NULL is returned.  */
char *
mapi_get_gpgol_draft_info (LPMESSAGE msg)
{
  TSTART;
  HRESULT hr;
  LPSPropValue propval = NULL;
  ULONG tag;
  char *retstr;

  if (get_gpgoldraftinfo_tag (msg, &tag) )
    {
      TRETURN NULL;
    }
  hr = HrGetOneProp ((LPMAPIPROP)msg, tag, &propval);
  if (FAILED (hr))
    {
      TRETURN NULL;
    }
  if (PROP_TYPE (propval->ulPropTag) == PT_STRING8)
    retstr = xstrdup (propval->Value.lpszA);
  else
    retstr = NULL;

  MAPIFreeBuffer (propval);
  TRETURN retstr;
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
  TSTART;
  HRESULT hr;
  SPropValue prop;
  SPropTagArray proparray;

  if (get_gpgoldraftinfo_tag (message, &prop.ulPropTag) )
    {
      TRETURN -1;
    }
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
      TRETURN -1;
    }

  TRETURN 0;
}


/* Return the MIME info as an allocated string.  Will never return
   NULL.  */
char *
mapi_get_mime_info (LPMESSAGE msg)
{
  TSTART;
  HRESULT hr;
  LPSPropValue propval = NULL;
  ULONG tag;
  char *retstr;

  if (get_gpgolmimeinfo_tag (msg, &tag) )
    {
      TRETURN xstrdup ("[Error getting tag for MIME info]");
    }
  hr = HrGetOneProp ((LPMAPIPROP)msg, tag, &propval);
  if (FAILED (hr))
    {
      TRETURN xstrdup ("");
    }
  if (PROP_TYPE (propval->ulPropTag) == PT_STRING8)
    retstr = xstrdup (propval->Value.lpszA);
  else
    retstr = xstrdup ("[MIME info has an invalid type]");

  MAPIFreeBuffer (propval);
  TRETURN retstr;
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
  TSTART;
  char *buf = mapi_get_gpgol_draft_info (message);
  int ret = 0;
  if (!buf)
    {
      TRETURN 0;
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
  TRETURN ret;
}

/* Sets the draft info flags. Protocol is always Auto.
   flags should be the same as defined by
   get_gpgol_draft_info_flags
*/
int
set_gpgol_draft_info_flags (LPMESSAGE message, int flags)
{
  TSTART;
  char buf[4];
  buf[3] = '\0';
  buf[2] = 'A'; /* Protocol */
  buf[1] = flags & 2 ? 'S' : 's';
  buf[0] = flags & 1 ? 'E' : 'e';

  TRETURN mapi_set_gpgol_draft_info (message, buf);
}


/* Helper for mapi_get_msg_content_type() */
static int
parse_headers_cb (void *dummy_arg,
                             rfc822parse_event_t event, rfc822parse_t msg)
{
  TSTART;
  (void)dummy_arg;
  (void)msg;

  if (event == RFC822PARSE_T2BODY)
    {
      TRETURN 42; /* Hack to stop the parsing after having read the
                     outer headers. */
    }
  TRETURN 0;
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
  TSTART;
  rfc822parse_t msg;
  const char *header_lines, *s;
  rfc822parse_field_t ctx;
  size_t length;
  char *retstr = NULL;

  if (r_protocol)
    *r_protocol = NULL;
  if (r_smtype)
    *r_smtype = NULL;

  /* Read the headers into an rfc822 object. */
  msg = rfc822parse_open (parse_headers_cb, NULL);
  if (!msg)
    {
      log_error ("%s:%s: rfc822parse_open failed",
                 SRCNAME, __func__);
      TRETURN NULL;
    }

  const std::string hdrStr = mapi_get_header (message);
  if (hdrStr.empty())
    {

      log_error ("%s:%s: failed to get headers",
                 SRCNAME, __func__);
      rfc822parse_close (msg);
      TRETURN NULL;
    }

  header_lines = hdrStr.c_str();
  while ((s = strchr (header_lines, '\n')))
    {
      length = (s - header_lines);
      if (length && s[-1] == '\r')
        length--;

      if (!strncmp ("Wks-Phase: confirm", header_lines,
                    std::max (18, (int) length)))
        {
          log_debug ("%s:%s: detected wks confirmation mail",
                     SRCNAME, __func__);
          retstr = xstrdup ("wks.confirmation.mail");
          rfc822parse_close (msg);
          TRETURN retstr;
        }

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
  TRETURN retstr;
}


/* Returns True if MESSAGE has a GpgOL Last Decrypted property with any value.
   This indicates that there should be no PR_BODY tag.  */
int
mapi_has_last_decrypted (LPMESSAGE message)
{
  TSTART;
  HRESULT hr;
  LPSPropValue propval = NULL;
  ULONG tag;
  int yes = 0;
  
  if (get_gpgollastdecrypted_tag (message, &tag) )
    {
      TRETURN 0; /* No.  */
    }
  hr = HrGetOneProp ((LPMAPIPROP)message, tag, &propval);
  if (FAILED (hr))
    {
      TRETURN 0; /* No.  */
    }
  
  if (PROP_TYPE (propval->ulPropTag) == PT_BINARY)
    yes = 1;

  MAPIFreeBuffer (propval);
  TRETURN yes;
}


/* Helper to check whether the file name of OBJ is "smime.p7m".
   returns on true if so.  */
static int
has_smime_filename (LPATTACH obj)
{
  TSTART;
  HRESULT hr;
  LPSPropValue propval;
  int yes = 0;

  hr = HrGetOneProp ((LPMAPIPROP)obj, PR_ATTACH_FILENAME, &propval);
  if (FAILED(hr))
    {
      hr = HrGetOneProp ((LPMAPIPROP)obj, PR_ATTACH_LONG_FILENAME, &propval);
      if (FAILED(hr))
        {
          TRETURN 0;
        }
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
  TRETURN yes;
}


/* Copy the MAPI body to a PGPBODY type attachment. */
int
mapi_body_to_attachment (LPMESSAGE message)
{
  TSTART;
  HRESULT hr;
  LPSTREAM instream;
  ULONG newpos;
  LPATTACH newatt = NULL;
  SPropValue prop;
  LPSTREAM outstream = NULL;
  LPUNKNOWN punk;
  char body_filename[] = PGPBODYFILENAME;

  instream = mapi_get_body_as_stream (message);
  if (!instream)
    {
      TRETURN -1;
    }

  log_debug ("%s:%s: Creating MOSS body attachment",
             SRCNAME, __func__);

  hr = message->CreateAttach (NULL, 0, &newpos, &newatt);
  if (hr)
    {
      log_error ("%s:%s: can't create attachment: hr=%#lx\n",
                 SRCNAME, __func__, hr);
      goto leave;
    }

  prop.ulPropTag = PR_ATTACH_METHOD;
  prop.Value.ul = ATTACH_BY_VALUE;
  hr = HrSetOneProp ((LPMAPIPROP)newatt, &prop);
  if (hr)
    {
      log_error ("%s:%s: can't set attach method: hr=%#lx\n",
                 SRCNAME, __func__, hr);
      goto leave;
    }

  /* Mark that attachment so that we know why it has been created.  */
  if (get_gpgolattachtype_tag (message, &prop.ulPropTag) )
    goto leave;
  prop.Value.l = ATTACHTYPE_PGPBODY;
  hr = HrSetOneProp ((LPMAPIPROP)newatt, &prop);
  if (hr)
    {
      log_error ("%s:%s: can't set %s property: hr=%#lx\n",
                 SRCNAME, __func__, "GpgOL Attach Type", hr);
      goto leave;
    }

  prop.ulPropTag = PR_ATTACHMENT_HIDDEN;
  prop.Value.b = TRUE;
  hr = HrSetOneProp ((LPMAPIPROP)newatt, &prop);
  if (hr)
    {
      log_error ("%s:%s: can't set hidden attach flag: hr=%#lx\n",
                 SRCNAME, __func__, hr);
      goto leave;
    }

  prop.ulPropTag = PR_ATTACH_FILENAME_A;
  prop.Value.lpszA = body_filename;
  hr = HrSetOneProp ((LPMAPIPROP)newatt, &prop);
  if (hr)
    {
      log_error ("%s:%s: can't set attach filename: hr=%#lx\n",
                 SRCNAME, __func__, hr);
      goto leave;
    }

  punk = (LPUNKNOWN)outstream;
  hr = gpgol_openProperty (newatt, PR_ATTACH_DATA_BIN, &IID_IStream, 0,
                             MAPI_CREATE|MAPI_MODIFY, &punk);
  if (FAILED (hr))
    {
      log_error ("%s:%s: can't create output stream: hr=%#lx\n",
                 SRCNAME, __func__, hr);
      goto leave;
    }
  outstream = (LPSTREAM)punk;

  /* Insert a blank line so that our mime parser skips over the mail
     headers.  */
  hr = outstream->Write ("\r\n", 2, NULL);
  if (hr)
    {
      log_error ("%s:%s: Write failed: hr=%#lx", SRCNAME, __func__, hr);
      goto leave;
    }

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
      log_error ("%s:%s: Commiting output stream failed: hr=%#lx",
                 SRCNAME, __func__, hr);
      goto leave;
    }
  gpgol_release (outstream);
  outstream = NULL;
  hr = newatt->SaveChanges (0);
  if (hr)
    {
      log_error ("%s:%s: SaveChanges of the attachment failed: hr=%#lx\n",
                 SRCNAME, __func__, hr);
      goto leave;
    }
  gpgol_release (newatt);
  newatt = NULL;
  hr = mapi_save_changes (message, KEEP_OPEN_READWRITE);

 leave:
  if (outstream)
    {
      outstream->Revert ();
      gpgol_release (outstream);
    }
  if (newatt)
    gpgol_release (newatt);
  gpgol_release (instream);
  TRETURN hr? -1:0;
}

int
mapi_mark_or_create_moss_attach (LPMESSAGE message, msgtype_t msgtype)
{
  TSTART;
  int i;
  if (msgtype == MSGTYPE_UNKNOWN ||
      msgtype == MSGTYPE_GPGOL)
    {
      TRETURN 0;
    }

  /* First check if we already have one marked. */
  mapi_attach_item_t *table = mapi_create_attach_table (message, 0);
  int part1 = 0,
      part2 = 0;
  for (i = 0; table && !table[i].end_of_table; i++)
    {
      if (table[i].attach_type == ATTACHTYPE_PGPBODY ||
          table[i].attach_type == ATTACHTYPE_MOSS ||
          table[i].attach_type == ATTACHTYPE_MOSSTEMPL)
        {
          if (!part1)
            {
              part1 = i + 1;
            }
          else if (!part2)
            {
              /* If we have two MOSS attachments we use
                 the second one. */
              part2 = i + 1;
              break;
            }
        }
    }
  if (part1 || part2)
    {
      /* Found existing moss attachment */
      mapi_release_attach_table (table);
      /* Remark to ensure that it is hidden. As our revert
         code must unhide it so that it is not stored in winmail.dat
         but used as the mosstmpl. */
      mapi_attach_item_t *item = table - 1 + (part2 ? part2 : part1);
      LPATTACH att;
      if (message->OpenAttach (item->mapipos, NULL, MAPI_BEST_ACCESS, &att) != S_OK)
        {
          log_error ("%s:%s: can't open attachment at %d",
                     SRCNAME, __func__, item->mapipos);
          TRETURN -1;
        }
      memdbg_addRef (att);
      if (!mapi_test_attach_hidden (att))
        {
          mapi_set_attach_hidden (att);
        }
      gpgol_release (att);
      if (part2)
        {
          TRETURN part2;
        }
      TRETURN part1;
    }

  if (msgtype == MSGTYPE_GPGOL_CLEAR_SIGNED ||
      msgtype == MSGTYPE_GPGOL_PGP_MESSAGE)
    {
      /* Inline message we need to create body attachment so that we
         are able to restore the content. */
      if (mapi_body_to_attachment (message))
        {
          log_error ("%s:%s: Failed to create body attachment.",
                     SRCNAME, __func__);
          TRETURN 0;
        }
      log_debug ("%s:%s: Created body attachment. Repeating lookup.",
                 SRCNAME, __func__);
      /* The position of the MOSS attach might change depending on
         the attachment count of the mail. So repeat the check to get
         the right position. */
      TRETURN mapi_mark_or_create_moss_attach (message, msgtype);
    }
  if (!table)
    {
      log_debug ("%s:%s: Neither pgp inline nor an attachment table.",
                 SRCNAME, __func__);
      TRETURN 0;
    }

  /* MIME Mails check for S/MIME first. */
  for (i = 0; !table[i].end_of_table; i++)
    {
      if (table[i].content_type
          && (!strcmp (table[i].content_type, "application/pkcs7-mime")
              || !strcmp (table[i].content_type,
                          "application/x-pkcs7-mime"))
          && table[i].filename
          && !strcmp (table[i].filename, "smime.p7m"))
        break;
    }
  if (!table[i].end_of_table)
    {
      mapi_mark_moss_attach (message, table + i);
      mapi_release_attach_table (table);
      TRETURN i + 1;
    }

  /* PGP/MIME or S/MIME stuff.  */
  /* Multipart/encrypted message: We expect 2 attachments.
     The first one with the version number and the second one
     with the ciphertext.  As we don't know wether we are
     called the first time, we first try to find these
     attachments by looking at all attachments.  Only if this
     fails we identify them by their order (i.e. the first 2
     attachments) and mark them as part1 and part2.  */
  for (i = 0; !table[i].end_of_table; i++); /* Count entries */
  if (i >= 2)
    {
      int part1_idx = -1,
          part2_idx = -1;
      /* At least 2 attachments but none are marked.  Thus we
         assume that this is the first time we see this
         message and we will set the mark now if we see
         appropriate content types. */
      if (table[0].content_type
          && !strcmp (table[0].content_type,
                      "application/pgp-encrypted"))
        part1_idx = 0;
      if (table[1].content_type
          && !strcmp (table[1].content_type,
                      "application/octet-stream"))
        part2_idx = 1;
      if (part1_idx != -1 && part2_idx != -1)
        {
          mapi_mark_moss_attach (message, table+part1_idx);
          mapi_mark_moss_attach (message, table+part2_idx);
          mapi_release_attach_table (table);
          TRETURN 2;
        }
    }

  if (!table[0].end_of_table && table[1].end_of_table)
    {
      /* No MOSS flag found in the table but there is only one
         attachment.  Due to the message type we know that this is
         the original MOSS message.  We mark this attachment as
         hidden, so that it won't get displayed.  We further mark
         it as our original MOSS attachment so that after parsing
         we have a mean to find it again (see above).  */
      mapi_mark_moss_attach (message, table + 0);
      mapi_release_attach_table (table);
      TRETURN 1;
    }

   mapi_release_attach_table (table);
   TRETURN 0; /* No original attachment - this should not happen.  */
}

static std::string
ac_get_value (const char *header, const char *what)
{
  TSTART;
  if (!header || !what)
    {
      STRANGEPOINT;
      TRETURN std::string();
    }

  const char *s = strstr (header, what);
  if (!s)
    {
      log_debug ("%s:%s: could not find %s in autocrypt header",
                 SRCNAME, __func__, what);
      TRETURN std::string();
    }
  /* As we found it we can be sure that this is not out
     of bounds. */
  s += strlen (what);

  if (*s != '=')
    {
      log_debug ("%s:%s: No equal sign after %s in autocrypt header %s",
                 SRCNAME, __func__, what, s);
      TRETURN std::string();
    }

  /* Move over the = sign. */
  s++;

  /* Find the sep */
  const char *s2 = strchr (s, ';');
  if (!s2)
    {
      /* No seperator found. Assume the rest is the value */
      TRETURN s;
    }

  /* From the equal to the ; is our value. */
  TRETURN std::string (s, s2 - s);
}

static GpgME::Data
prepare_key_data (const std::string &d)
{
  TSTART;
  if (d.empty())
    {
      STRANGEPOINT;
      TRETURN GpgME::Data();
    }

  /* Prepare the keydata */
  b64_state_t base64;     /* The state of the Base-64 decoder.  */
  b64_init (&base64);

  /* strdup and not xstrdup as we want GpgME to take over */
  char *b64decoded = strdup (d.c_str());
  size_t len = b64_decode (&base64, b64decoded, strlen(b64decoded));

  if (!len)
    {
      log_error ("%s:%s: Invalid base64 in %s", SRCNAME, __func__,
                 b64decoded);
      xfree (b64decoded);
      TRETURN GpgME::Data();
    }

  auto data = GpgME::Data (b64decoded, len, false /* take ownership */);

  TRETURN data;
}

bool
mapi_get_header_info (LPMESSAGE message,
                      autocrypt_s &r_autocrypt)
{
  TSTART;
  rfc822parse_t msg;

  /* Read the headers into an rfc822 object. */
  msg = rfc822parse_open (parse_headers_cb, NULL);
  if (!msg)
    {
      log_error ("%s:%s: rfc822parse_open failed",
                 SRCNAME, __func__);
      TRETURN false;
    }

  const std::string hdrStr = mapi_get_header (message);
  if (hdrStr.empty())
    {

      log_error ("%s:%s: failed to get headers",
                 SRCNAME, __func__);
      rfc822parse_close (msg);
      TRETURN false;
    }

  size_t length;
  const char *header_lines = hdrStr.c_str();
  const char *s;
  while ((s = strchr (header_lines, '\n')))
    {
      length = (s - header_lines);
      if (length && s[-1] == '\r')
        length--;

      rfc822parse_insert (msg, (const unsigned char*)header_lines, length);
      header_lines = s+1;
    }

  const char *ac_field = rfc822parse_get_field (msg, "Autocrypt", -1, 0);
  if (ac_field)
    {
      r_autocrypt.exists = true;
      r_autocrypt.addr = ac_get_value (ac_field, "addr");
      r_autocrypt.data = prepare_key_data (ac_get_value (ac_field, "keydata"));
      r_autocrypt.pref = ac_get_value (ac_field, "prefer-encrypt");
    }

  rfc822parse_close (msg);
  TRETURN true;
}
