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
#include "intern.h"
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
      if (!strncmp (s, "IPM.Note.SMIME", 14) && (!s[14] || s[14] =='.'))
        {
          /* This is either "IPM.Note.SMIME" or "IPM.Note.SMIME.foo".
             Note that we ncan't just insert a new aprt and keep the
             SMIME; we need to change the SMIME part of the class name
             so that Outlook does not proxcess it as an SMIME
             message. */
          newvalue = (char*)xmalloc (strlen (s) + 1);
          strcpy (stpcpy (newvalue, "IPM.Note.GpgSM"), s+14);
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


/* Return the message class as a scalar.  This function knows only
   about out own message classes.  Returns MSGCLS_UNKNOWN for any
   MSGCLASS we have no special support for.  */
msgclass_t
mapi_get_message_class (LPMESSAGE message)
{
  HRESULT hr;
  LPSPropValue propval = NULL;
  msgclass_t msgcls = MSGCLS_UNKNOWN;

  if (!message)
    return msgcls; 

  hr = HrGetOneProp ((LPMAPIPROP)message, PR_MESSAGE_CLASS_A, &propval);
  if (FAILED (hr))
    {
      log_error ("%s:%s: HrGetOneProp() failed: hr=%#lx\n",
                 SRCNAME, __func__, hr);
      return msgcls;
    }
    
  if ( PROP_TYPE (propval->ulPropTag) == PT_STRING8 )
    {
      const char *s = propval->Value.lpszA;
      if (!strncmp (s, "IPM.Note.GpgSM", 14) && (!s[14] || s[14] =='.'))
        {
          s += 14;
          if (!*s)
            msgcls = MSGCLS_GPGSM;
          else if (!strcmp (s, ".MultipartSigned"))
            msgcls = MSGCLS_GPGSM_MULTIPART_SIGNED;
          else
            log_debug ("%s:%s: message class `%s' not supported",
                       SRCNAME, __func__, s-14);
        }
    }
  MAPIFreeBuffer (propval);
  return msgcls;
}


/* This function is pretty useless because IConverterSession won't
   take attachments in to account.  Need to write our own version.  */
// int
// mapi_to_mime (LPMESSAGE message, const char *filename)
// {
//   HRESULT hr;
//   LPCONVERTERSESSION session;
//   LPSTREAM stream;

//   hr = CoCreateInstance (CLSID_IConverterSession, NULL, CLSCTX_INPROC_SERVER,
//                          IID_IConverterSession, (void **) &session);
//   if (FAILED (hr))
//     {
//       log_error ("%s:%s: can't create new IConverterSession object: hr=%#lx",
//                  SRCNAME, __func__, hr);
//       return -1;
//     }


//   hr = OpenStreamOnFile (MAPIAllocateBuffer, MAPIFreeBuffer,
//                          (STGM_CREATE | STGM_READWRITE),
//                          (char*)filename, NULL, &stream); 
//   if (FAILED (hr)) 
//     {
//       log_error ("%s:%s: can't create file `%s': hr=%#lx\n",
//                  SRCNAME, __func__, filename, hr); 
//       hr = -1;
//     }
//   else
//     {
//       hr = session->MAPIToMIMEStm (message, stream, 0);
//       if (FAILED (hr))
//         {
//           log_error ("%s:%s: MAPIToMIMEStm failed: hr=%#lx",
//                      SRCNAME, __func__, hr);
//           stream->Revert ();
//           hr = -1;
//         }
//       else
//         {
//           stream->Commit (0);
//           hr = 0;
//         }

//       stream->Release ();
//     }

//   session->Release ();
//   return hr;
// }


