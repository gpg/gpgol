/* mapihelp.cpp - Helper functions for MAPI
 *	Copyright (C) 2005, 2007 g10 Code GmbH
 * 
 * This file is part of GpgOL.
 * 
 * GpgOL is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2 of the License, or (at your option) any later version.
 * 
 * GpgOL is distributed in the hope that it will be useful,
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

#include "mymapi.h"
#include "mymapitags.h"
#include "common.h"
#include "rfc822parse.h"
#include "serpent.h"
#include "mapihelp.h"


#define TRACEPOINT() do { log_debug ("%s:%s:%d: tracepoint\n", \
                                     SRCNAME, __func__, __LINE__); \
                        } while (0)


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

    default:
      log_debug ("%s:%s: HrGetOneProp(%s) property type %lu not supported\n",
                 SRCNAME, __func__, propname,
                 PROP_TYPE (propval->ulPropTag) );
      return;
    }
  MAPIFreeBuffer (propval);
}


/* Return the property tag for GpgOL Attach Type. */
int 
get_gpgolattachtype_tag (LPMESSAGE message, ULONG *r_tag)
{
  HRESULT hr;
  LPSPropTagArray proparr = NULL;
  MAPINAMEID mnid, *pmnid;	
  /* {31805ab8-3e92-11dc-879c-00061b031004}: GpgOL custom properties.  */
  GUID guid = {0x31805ab8, 0x3e92, 0x11dc, {0x87, 0x9c, 0x00, 0x06,
                                            0x1b, 0x03, 0x10, 0x04}};

  memset (&mnid, 0, sizeof mnid);
  mnid.lpguid = &guid;
  mnid.ulKind = MNID_STRING;
  mnid.Kind.lpwstrName = L"GpgOL Attach Type";
  pmnid = &mnid;
  hr = message->GetIDsFromNames (1, &pmnid, MAPI_CREATE, &proparr);
  if (FAILED (hr)) 
    {
      log_error ("%s:%s: can't map %s property: hr=%#lx\n",
                 SRCNAME, __func__, "GpgOL Attach Type", hr); 
      return -1;
    }
    
  *r_tag = ((proparr->aulPropTag[0] & 0xFFFF0000) | PT_LONG);
  return 0;
}


/* Return the property tag for GpgOL Sig Status. */
int 
get_gpgolsigstatus_tag (LPMESSAGE message, ULONG *r_tag)
{
  HRESULT hr;
  LPSPropTagArray proparr = NULL;
  MAPINAMEID mnid, *pmnid;	
  /* {31805ab8-3e92-11dc-879c-00061b031004}: GpgOL custom properties.  */
  GUID guid = {0x31805ab8, 0x3e92, 0x11dc, {0x87, 0x9c, 0x00, 0x06,
                                            0x1b, 0x03, 0x10, 0x04}};

  memset (&mnid, 0, sizeof mnid);
  mnid.lpguid = &guid;
  mnid.ulKind = MNID_STRING;
  mnid.Kind.lpwstrName = L"GpgOL Sig Status";
  pmnid = &mnid;
  hr = message->GetIDsFromNames (1, &pmnid, MAPI_CREATE, &proparr);
  if (FAILED (hr)) 
    {
      log_error ("%s:%s: can't map %s property: hr=%#lx\n",
                 SRCNAME, __func__, "GpgOL Sig Status", hr); 
      return -1;
    }
    
  *r_tag = ((proparr->aulPropTag[0] & 0xFFFF0000) | PT_STRING8);
  return 0;
}


/* Return the property tag for GpgOL Protect IV. */
int 
get_gpgolprotectiv_tag (LPMESSAGE message, ULONG *r_tag)
{
  HRESULT hr;
  LPSPropTagArray proparr = NULL;
  MAPINAMEID mnid, *pmnid;	
  /* {31805ab8-3e92-11dc-879c-00061b031004}: GpgOL custom properties.  */
  GUID guid = {0x31805ab8, 0x3e92, 0x11dc, {0x87, 0x9c, 0x00, 0x06,
                                            0x1b, 0x03, 0x10, 0x04}};

  memset (&mnid, 0, sizeof mnid);
  mnid.lpguid = &guid;
  mnid.ulKind = MNID_STRING;
  mnid.Kind.lpwstrName = L"GpgOL Protect IV";
  pmnid = &mnid;
  hr = message->GetIDsFromNames (1, &pmnid, MAPI_CREATE, &proparr);
  if (FAILED (hr)) 
    {
      log_error ("%s:%s: can't map %s property: hr=%#lx\n",
                 SRCNAME, __func__, "GpgOL Protect IV", hr); 
      return -1;
    }
    
  *r_tag = ((proparr->aulPropTag[0] & 0xFFFF0000) | PT_BINARY);
  return 0;
}



/* This function checks whether MESSAGE requires processing by us and
   adjusts the message class to our own.  Return true if the message
   was changed. */
int
mapi_change_message_class (LPMESSAGE message)
{
  HRESULT hr;
  SPropValue prop;
  LPSPropValue propval = NULL;
  
  char *newvalue = NULL;

  if (!message)
    return 0; /* No message: Nop. */

  hr = HrGetOneProp ((LPMAPIPROP)message, PR_MESSAGE_CLASS_A, &propval);
  if (FAILED (hr))
    {
      log_error ("%s:%s: HrGetOneProp() failed: hr=%#lx\n",
                 SRCNAME, __func__, hr);
      return 0;
    }
    
  if ( PROP_TYPE (propval->ulPropTag) == PT_STRING8 )
    {
      const char *s = propval->Value.lpszA;

      if (!strcmp (s, "IPM.Note"))
        {
          /* Most message today are of this type.  However a PGP/MIME
             encrypted message although has this class here.  We need
             to see whether we can detect such a mail right here and
             change the message class accordingly. */
          char *ct, *proto;

          ct = mapi_get_message_content_type (message, &proto, NULL);
          if (!ct)
            log_debug ("%s:%s: message has no content type", 
                       SRCNAME, __func__);
          else
            {
              log_debug ("%s:%s: content type is '%s'", 
                         SRCNAME, __func__, ct);
              if (proto)
                {
                  log_debug ("%s:%s:     protocol is '%s'", 
                             SRCNAME, __func__, proto);
              
                  if (!strcmp (ct, "multipart/encrypted")
                      && !strcmp (proto, "application/pgp-encrypted"))
                    newvalue = xstrdup ("IPM.Note.GpgOL.MultipartEncrypted");
                  xfree (proto);
                }
              xfree (ct);
            }
        }
      else if (!strcmp (s, "IPM.Note.SMIME"))
        {
          /* This is an S/MIME opaque encrypted or signed message.
             Check what it really is.  */
          char *ct, *smtype;

          ct = mapi_get_message_content_type (message, NULL, &smtype);
          if (!ct)
            log_debug ("%s:%s: message has no content type", 
                       SRCNAME, __func__);
          else
            {
              log_debug ("%s:%s: content type is '%s'", 
                         SRCNAME, __func__, ct);
              if (smtype)
                {
                  log_debug ("%s:%s:   smime-type is '%s'", 
                             SRCNAME, __func__, smtype);
              
                  if (!strcmp (ct, "application/pkcs7-mime")
                      || !strcmp (ct, "application/x-pkcs7-mime"))
                    {
                      if (!strcmp (smtype, "signed-data"))
                        newvalue = xstrdup ("IPM.Note.GpgOL.OpaqueSigned");
                      else if (!strcmp (smtype, "enveloped-data"))
                        newvalue = xstrdup ("IPM.Note.GpgOL.OpaqueEncrypted");
                    }
                  xfree (smtype);
                }
              xfree (ct);
            }
          if (!newvalue)
            newvalue = xstrdup ("IPM.Note.GpgOL");
        }
      else if (!strncmp (s, "IPM.Note.SMIME", 14) && (!s[14] || s[14] =='.'))
        {
          /* This is "IPM.Note.SMIME.foo" (where ".foo" is optional
             but the previous condition has already taken care of
             this.  Note that we can't just insert a new part and keep
             the SMIME; we need to change the SMIME part of the class
             name so that Outlook does not proxcess it as an SMIME
             message. */
          newvalue = (char*)xmalloc (strlen (s) + 1);
          strcpy (stpcpy (newvalue, "IPM.Note.GpgOL"), s+14);
        }
    }
  MAPIFreeBuffer (propval);
  if (!newvalue)
    return 0;

  prop.ulPropTag = PR_MESSAGE_CLASS_A;
  prop.Value.lpszA = newvalue; 
  hr = message->SetProps (1, &prop, NULL);
  xfree (newvalue);
  if (hr != S_OK)
    {
      log_error ("%s:%s: can't set message class: hr=%#lx\n",
                 SRCNAME, __func__, hr);
      return 0;
    }

  hr = message->SaveChanges (KEEP_OPEN_READONLY);
  if (hr != S_OK)
    {
      log_error ("%s:%s: SaveChanges() failed: hr=%#lx\n",
                 SRCNAME, __func__, hr); 
      return 0;
    }

  return 1;
}


/* Return the message type.  This function knows only about our own
   message types.  Returns MSGTYPE_UNKNOWN for any MESSAGE we have
   no special support for.  */
msgtype_t
mapi_get_message_type (LPMESSAGE message)
{
  HRESULT hr;
  LPSPropValue propval = NULL;
  msgtype_t msgtype = MSGTYPE_UNKNOWN;

  if (!message)
    return msgtype; 

  hr = HrGetOneProp ((LPMAPIPROP)message, PR_MESSAGE_CLASS_A, &propval);
  if (FAILED (hr))
    {
      log_error ("%s:%s: HrGetOneProp() failed: hr=%#lx\n",
                 SRCNAME, __func__, hr);
      return msgtype;
    }
    
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
          else
            log_debug ("%s:%s: message class `%s' not supported",
                       SRCNAME, __func__, s-14);
        }
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


/* Return a binary propery in a malloced buffer with its length stored
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
static attachtype_t
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

      hr = message->OpenAttach (pos, NULL, MAPI_BEST_ACCESS, &att);	
      if (FAILED (hr))
        {
          log_error ("%s:%s: can't open attachment %d: hr=%#lx",
                     SRCNAME, __func__, pos, hr);
          table[pos].mapipos = -1;
          continue;
        }

      table[pos].mapipos = pos;
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
  table[pos].end_of_table = 1;
  FreeProws (mapirows);
  mapitable->Release ();

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

  if (!table)
    return;

  for (pos=0; !table[pos].end_of_table; pos++)
    {
      xfree (table[pos].filename);
      xfree (table[pos].content_type);
    }
  xfree (table);
}


/* Return an attachment as a new IStream object.  Returns NULL on
   failure. */
LPSTREAM
mapi_get_attach_as_stream (LPMESSAGE message, mapi_attach_item_t *item)
{
  HRESULT hr;
  LPATTACH att;
  LPSTREAM stream;

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

  att->Release ();

  return stream;
}


/* Return a malloced buffer with the content of the attachment. If
   R_NBYTES is not NULL the number of bytes will get stored there.
   ATT must have an attachment method of ATATCH_BY_VALUE.  Returns
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



/* Return an attachment as a malloced buffer; The size of the buffer
   will be stored at R_NBYTES.  Returns NULL on failure. */
char *
mapi_get_attach (LPMESSAGE message, mapi_attach_item_t *item, size_t *r_nbytes)
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

  buffer = attach_to_buffer (att, r_nbytes, 0, NULL);
  att->Release ();

  return buffer;
}


/* Mark this attachment as the orginal MOSS message.  We set a custom
   property as weel ast the hidden hidden flag on ot..  */
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

  if (FAILED (hr)) 
    {
      log_error ("%s:%s: can't map %s property: hr=%#lx\n",
                 SRCNAME, __func__, "GpgOL Attach Type", hr); 
      goto leave;
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



/* Returns True if MESSAGE has a GpgOL Sig Status property.  */
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
    yes = !(propval->Value.lpszA && !strcmp (propval->Value.lpszA, "?"));
  else
    yes = 0;

  MAPIFreeBuffer (propval);
  return yes;
}


/* Set the signature status property to STATUS_STRING.  To set the
   status to not checked, set it to "?".  Note that this function does
   not call SaveChanges.  */
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


/* Helper for mapi_get_msg_content_type() */
static int
get_message_content_type_cb (void *dummy_arg,
                             rfc822parse_event_t event, rfc822parse_t msg)
{
  if (event == RFC822PARSE_T2BODY)
    return 42; /* Hack to stop the parsing after having read the
                  outher headers. */
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
         to cope withy unicode etc. */
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


/* Return the content of the body attachment of MESSAGE.  The body
   attachment is a hidden attachment created by us for later display.
   If R_NBYTES is not NULL the number of bytes in the returned buffer
   is stored there.  If R_ISHTML is not NULL a flag indicating whether
   the HTML is html formatted is stored there.  If R_PROTECTED is not
   NULL a flag indicating whethernthe message was protected is store
   there.  If no body attachment can be found or on any other error
   NULL is returned.  Caller must free the returned string. */
char *
mapi_get_gpgol_body_attachment (LPMESSAGE message, size_t *r_nbytes, 
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

  if (r_ishtml)
    *r_ishtml = 0;
  if (r_protected)
    *r_protected = 0;

  if (get_gpgolattachtype_tag (message, &moss_tag) )
    return NULL;

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
  if (!n_attach)
    {
      FreeProws (mapirows);
      mapitable->Release ();
      log_debug ("%s:%s: No attachments at all", SRCNAME, __func__);
      return NULL;
    }

  for (pos=0; pos < n_attach; pos++) 
    {
      LPATTACH att;

      hr = message->OpenAttach (pos, NULL, MAPI_BEST_ACCESS, &att);	
      if (FAILED (hr))
        {
          log_error ("%s:%s: can't open attachment %d: hr=%#lx",
                     SRCNAME, __func__, pos, hr);
          continue;
        }
      if ( (bodytype=has_gpgol_body_name (att))
           && get_gpgolattachtype (att, moss_tag) == ATTACHTYPE_FROMMOSS)
        {
          if (get_attach_method (att) == ATTACH_BY_VALUE)
            body = attach_to_buffer (att, r_nbytes, 1, r_protected);
          att->Release ();
          if (r_ishtml)
            *r_ishtml = (bodytype == 2);
          break;
        }
      att->Release ();
    }
  FreeProws (mapirows);
  mapitable->Release ();
  if (!body)
    log_error ("%s:%s: no suitable body attachment found", SRCNAME, __func__);
  
  return body;
}


