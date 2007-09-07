/* message-events.cpp - Subclass impl of IExchExtMessageEvents
 *	Copyright (C) 2004, 2005, 2007 g10 Code GmbH
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
#include "myexchext.h"
#include "display.h"
#include "common.h"
#include "msgcache.h"
#include "mapihelp.h"

#include "olflange-ids.h"
#include "olflange-def.h"
#include "olflange.h"
#include "ol-ext-callback.h"
#include "mimeparser.h"
#include "mimemaker.h"
#include "message.h"
#include "message-events.h"


#define TRACEPOINT() do { log_debug ("%s:%s:%d: tracepoint\n", \
                                     SRCNAME, __func__, __LINE__); \
                        } while (0)



/* Wrapper around UlRelease with error checking. */
/* FIXME: Duplicated code.  */
static void 
ul_release (LPVOID punk)
{
  ULONG res;
  
  if (!punk)
    return;
  res = UlRelease (punk);
//   log_debug ("%s UlRelease(%p) had %lu references\n", __func__, punk, res);
}






/* Our constructor.  */
GpgolMessageEvents::GpgolMessageEvents (GpgolExt *pParentInterface)
{ 
  m_pExchExt = pParentInterface;
  m_lRef = 0; 
  m_bOnSubmitActive = FALSE;
  m_want_html = FALSE;
  m_processed = FALSE;
}


/* The QueryInterfac.  */
STDMETHODIMP 
GpgolMessageEvents::QueryInterface (REFIID riid, LPVOID FAR *ppvObj)
{   
  *ppvObj = NULL;
  if (riid == IID_IExchExtMessageEvents)
    {
      *ppvObj = (LPVOID)this;
      AddRef();
      return S_OK;
    }
  if (riid == IID_IUnknown)
    {
      *ppvObj = (LPVOID)m_pExchExt;  
      m_pExchExt->AddRef();
      return S_OK;
    }
  return E_NOINTERFACE;
}



/* Called from Exchange on reading a message.  Returns: S_FALSE to
   signal Exchange to continue calling extensions.  PEECB is a pointer
   to the IExchExtCallback interface. */
STDMETHODIMP 
GpgolMessageEvents::OnRead (LPEXCHEXTCALLBACK pEECB) 
{
  LPMDB mdb = NULL;
  LPMESSAGE message = NULL;
  
  log_debug ("%s:%s: received\n", SRCNAME, __func__);
  pEECB->GetObject (&mdb, (LPMAPIPROP *)&message);
  log_mapi_property (message, PR_CONVERSATION_INDEX,"PR_CONVERSATION_INDEX");
  switch (m_pExchExt->getMsgtype (pEECB))
    {
    case MSGTYPE_UNKNOWN: 
      break;
    case MSGTYPE_GPGOL:
      log_debug ("%s:%s: ignoring unknown message of original SMIME class\n",
                 SRCNAME, __func__);
      break;
    case MSGTYPE_GPGOL_MULTIPART_SIGNED:
      log_debug ("%s:%s: processing multipart signed message\n", 
                 SRCNAME, __func__);
      m_processed = TRUE;
      message_verify (message, m_pExchExt->getMsgtype (pEECB), 0);
      break;
    case MSGTYPE_GPGOL_MULTIPART_ENCRYPTED:
      log_debug ("%s:%s: processing multipart encrypted message\n",
                 SRCNAME, __func__);
      m_processed = TRUE;
      message_decrypt (message, m_pExchExt->getMsgtype (pEECB), 0);
      /* Hmmm, we might want to abort it and run our own inspector
         instead.  */
      break;
    case MSGTYPE_GPGOL_OPAQUE_SIGNED:
      log_debug ("%s:%s: processing opaque signed message\n", 
                 SRCNAME, __func__);
      m_processed = TRUE;
      message_verify (message, m_pExchExt->getMsgtype (pEECB), 0);
      break;
    case MSGTYPE_GPGOL_CLEAR_SIGNED:
      log_debug ("%s:%s: processing clear signed pgp message\n", 
                 SRCNAME, __func__);
      m_processed = TRUE;
      message_verify (message, m_pExchExt->getMsgtype (pEECB), 0);
      break;
    case MSGTYPE_GPGOL_OPAQUE_ENCRYPTED:
      log_debug ("%s:%s: processing opaque encrypted message\n",
                 SRCNAME, __func__);
      m_processed = TRUE;
      message_decrypt (message, m_pExchExt->getMsgtype (pEECB), 0);
      /* Hmmm, we might want to abort it and run our own inspector
         instead.  */
      break;
    case MSGTYPE_GPGOL_PGP_MESSAGE:
      log_debug ("%s:%s: processing pgp message\n", SRCNAME, __func__);
      m_processed = TRUE;
      message_decrypt (message, m_pExchExt->getMsgtype (pEECB), 0);
      /* Hmmm, we might want to abort it and run our own inspector
         instead.  */
      break;
    }
  
  ul_release (message);
  ul_release (mdb);

  return S_FALSE;
}


/* Called by Exchange after a message has been read.  Returns: S_FALSE
   to signal Exchange to continue calling extensions.  PEECB is a
   pointer to the IExchExtCallback interface. LFLAGS are some flags. */
STDMETHODIMP 
GpgolMessageEvents::OnReadComplete (LPEXCHEXTCALLBACK pEECB, ULONG lFlags)
{
  log_debug ("%s:%s: received\n", SRCNAME, __func__);

  /* If the message has been processed by is (i.e. in OnRead), we now
     use our own display code.  */
  if (m_processed)
    {
      HRESULT hr;
      HWND hwnd = NULL;
      LPMESSAGE message = NULL;
      LPMDB mdb = NULL;

      if (FAILED (pEECB->GetWindow (&hwnd)))
        hwnd = NULL;
      hr = pEECB->GetObject (&mdb, (LPMAPIPROP *)&message);
      if (SUCCEEDED (hr))
        {
          int ishtml, wasprotected;
          char *body;

          /* If the message was protected we don't allow a fallback to
             the OOM display methods.  FIXME:  This is currently disabled. */
          body = mapi_get_gpgol_body_attachment (message, NULL,
                                                 &ishtml, &wasprotected);
          if (body)
            update_display (hwnd, /*wasprotected?NULL:*/pEECB, ishtml, body);
          else
            update_display (hwnd, NULL, 0, 
                            _("[Crypto operation failed - "
                              "can't show the body of the message]"));
        }
      ul_release (message);
      ul_release (mdb);
    }
  

#if 1
    {
      HWND hWnd = NULL;

      if (FAILED (pEECB->GetWindow (&hWnd)))
        hWnd = NULL;
      else
        log_window_hierarchy (hWnd, "%s:%s:%d: Windows hierarchy:",
                              SRCNAME, __func__, __LINE__);
    }
#endif

  return S_FALSE;
}


/* Called by Exchange when a message will be written. Returns: S_FALSE
   to signal Exchange to continue calling extensions.  PEECB is a
   pointer to the IExchExtCallback interface. */
STDMETHODIMP 
GpgolMessageEvents::OnWrite (LPEXCHEXTCALLBACK pEECB)
{
  log_debug ("%s:%s: received\n", SRCNAME, __func__);

  HRESULT hr;
  LPDISPATCH pDisp;
  DISPID dispid;
  VARIANT aVariant;
  DISPPARAMS dispparamsNoArgs = {NULL, NULL, 0, 0};
  HWND hWnd = NULL;


  /* If we are going to encrypt, check that the BodyFormat is
     something we support.  This helps avoiding surprise by sending
     out unencrypted messages. */
  if (m_pExchExt->m_gpgEncrypt || m_pExchExt->m_gpgSign)
    {
      pDisp = find_outlook_property (pEECB, "BodyFormat", &dispid);
      if (!pDisp)
        {
          log_debug ("%s:%s: BodyFormat not found\n", SRCNAME, __func__);
          m_bWriteFailed = TRUE;	
          return E_FAIL;
        }
      
      aVariant.bstrVal = NULL;
      hr = pDisp->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                          DISPATCH_PROPERTYGET, &dispparamsNoArgs,
                          &aVariant, NULL, NULL);
      if (hr != S_OK)
        {
          log_debug ("%s:%s: retrieving BodyFormat failed: %#lx",
                     SRCNAME, __func__, hr);
          m_bWriteFailed = TRUE;	
          pDisp->Release();
          return E_FAIL;
        }
  
      if (aVariant.vt != VT_INT && aVariant.vt != VT_I4)
        {
          log_debug ("%s:%s: BodyFormat is not an integer (%d)",
                     SRCNAME, __func__, aVariant.vt);
          m_bWriteFailed = TRUE;	
          pDisp->Release();
          return E_FAIL;
        }
  
      if (aVariant.intVal == 1)
        m_want_html = 0;
      else if (aVariant.intVal == 2)
        m_want_html = 1;
      else
        {

          log_debug ("%s:%s: BodyFormat is %d",
                     SRCNAME, __func__, aVariant.intVal);
          
          if (FAILED(pEECB->GetWindow (&hWnd)))
            hWnd = NULL;
          MessageBox (hWnd,
                      _("Sorry, we can only encrypt plain text messages and\n"
                      "no RTF messages. Please make sure that only the text\n"
                      "format has been selected."),
                      "GpgOL", MB_ICONERROR|MB_OK);

          m_bWriteFailed = TRUE;	
          pDisp->Release();
          return E_FAIL;
        }
      pDisp->Release();
      
    }
  
  
  return S_FALSE;
}




/* Called by Exchange when the data has been written to the message.
   Encrypts and signs the message if the options are set.  PEECB is a
   pointer to the IExchExtCallback interface.  Returns: S_FALSE to
   signal Exchange to continue calling extensions.  We return E_FAIL
   to signals Exchange an error; the message will not be sent.  Note
   that we don't return S_OK because this would mean that we rolled
   back the write operation and that no further extensions should be
   called. */
STDMETHODIMP 
GpgolMessageEvents::OnWriteComplete (LPEXCHEXTCALLBACK pEECB, ULONG lFlags)
{
  HRESULT hrReturn = S_FALSE;
  LPMESSAGE msg = NULL;
  LPMDB pMDB = NULL;
  HWND hWnd = NULL;
  int rc;

  log_debug ("%s:%s: received\n", SRCNAME, __func__);


  if (lFlags & (EEME_FAILED|EEME_COMPLETE_FAILED))
    return S_FALSE; /* We don't need to rollback anything in case
                       other extensions flagged a failure. */
          
  if (!m_bOnSubmitActive) /* The user is just saving the message. */
    return S_FALSE;
  
  if (m_bWriteFailed)     /* Operation failed already. */
    return S_FALSE;

  /* Try to get the current window. */
  if (FAILED(pEECB->GetWindow (&hWnd)))
    hWnd = NULL;

  /* Get the object and call the encryption or signing fucntion.  */
  HRESULT hr = pEECB->GetObject (&pMDB, (LPMAPIPROP *)&msg);
  if (SUCCEEDED (hr))
    {
      if (m_pExchExt->m_gpgEncrypt && m_pExchExt->m_gpgSign)
        rc = mime_sign_encrypt (msg, PROTOCOL_OPENPGP);
      else if (m_pExchExt->m_gpgEncrypt && !m_pExchExt->m_gpgSign)
        rc = mime_encrypt (msg, PROTOCOL_OPENPGP);
      else if (!m_pExchExt->m_gpgEncrypt && m_pExchExt->m_gpgSign)
        rc = mime_sign (msg, PROTOCOL_OPENPGP);
      else
        rc = 0;
      
      if (rc)
        {
          hrReturn = E_FAIL;
          m_bWriteFailed = TRUE;	
        }
    }
  
  ul_release (msg);
  ul_release (pMDB);
  
  return hrReturn;
}


/* Called by Exchange when the user selects the "check names" command.
   PEECB is a pointer to the IExchExtCallback interface.  Returns
   S_FALSE to signal Exchange to continue calling extensions. */
STDMETHODIMP 
GpgolMessageEvents::OnCheckNames(LPEXCHEXTCALLBACK pEECB)
{
  log_debug ("%s:%s: received\n", SRCNAME, __func__);
  return S_FALSE;
}


/* Called by Exchange when "check names" command is complete.
   PEECB is a pointer to the IExchExtCallback interface.  Returns
   S_FALSE to signal Exchange to continue calling extensions. */
STDMETHODIMP 
GpgolMessageEvents::OnCheckNamesComplete (LPEXCHEXTCALLBACK pEECB,ULONG lFlags)
{
  log_debug ("%s:%s: received\n", SRCNAME, __func__);
  return S_FALSE;
}


/* Called by Exchange before the message data will be written and
   submitted to MAPI.  PEECB is a pointer to the IExchExtCallback
   interface.  Returns S_FALSE to signal Exchange to continue calling
   extensions. */
STDMETHODIMP 
GpgolMessageEvents::OnSubmit (LPEXCHEXTCALLBACK pEECB)
{
  log_debug ("%s:%s: received\n", SRCNAME, __func__);
  m_bOnSubmitActive = TRUE;
  m_bWriteFailed = FALSE;
  return S_FALSE;
}


/* Called by Exchange after the message has been submitted to MAPI.
   PEECB is a pointer to the IExchExtCallback interface. */
STDMETHODIMP_ (VOID) 
GpgolMessageEvents::OnSubmitComplete (LPEXCHEXTCALLBACK pEECB,
                                            ULONG lFlags)
{
  log_debug ("%s:%s: received\n", SRCNAME, __func__);
  m_bOnSubmitActive = FALSE; 
}


