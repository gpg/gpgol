/* message-events.cpp - Subclass impl of IExchExtMessageEvents
 *	Copyright (C) 2004, 2005, 2007, 2008 g10 Code GmbH
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
#include <windows.h>

#include "mymapi.h"
#include "mymapitags.h"
#include "myexchext.h"
#include "display.h"
#include "common.h"
#include "msgcache.h"
#include "mapihelp.h"

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
static void 
ul_release (LPVOID punk, const char *func, int lnr)
{
  ULONG res;
  
  if (!punk)
    return;
  res = UlRelease (punk);
  if (opt.enable_debug & DBG_MEMORY)
    log_debug ("%s:%s:%d: UlRelease(%p) had %lu references\n", 
               SRCNAME, func, lnr, punk, res);
}






/* Our constructor.  */
GpgolMessageEvents::GpgolMessageEvents (GpgolExt *pParentInterface)
{ 
  m_pExchExt = pParentInterface;
  m_lRef = 0; 
  m_bOnSubmitActive = false;
  m_want_html = false;
  m_processed = false;
  m_wasencrypted = false;
  m_gotinspector = false;
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
   signal Exchange to continue calling extensions.  EECB is a pointer
   to the IExchExtCallback interface. */
STDMETHODIMP 
GpgolMessageEvents::OnRead (LPEXCHEXTCALLBACK eecb) 
{
  HWND hwnd = NULL;
  LPMDB mdb = NULL;
  LPMESSAGE message = NULL;
  
  m_wasencrypted = false;
  if (FAILED (eecb->GetWindow (&hwnd)))
    hwnd = NULL;

  m_gotinspector = !!is_inspector_display (hwnd);

  log_debug ("%s:%s: received (hwnd=%p) %s\n", 
             SRCNAME, __func__, hwnd, m_gotinspector? "got_inspector":"");

  /* Fixme: If preview decryption is not enabled and we have an
     encrypted message, we might want to show a greyed out preview
     window.  There are two ways to clear the preview window: 
     - Change the message class to something unknown to Outlook, like 
       IPM.GpgOL.  This shows a grey and empty preview window.
     - Set flag 0x2000 in the 0x8510 property (SideEffects).  This
       shows a grey window with a notice that the message can't be 
       shown due to active content.  */  

  /* The is_inspector_display function is not reliable enough.
     Missing another solution we set it to true for now with the
     result that the preview decryption can't be disabled.  */
  m_gotinspector = 1;
  
  if (m_gotinspector || opt.preview_decrypt)
    {
      eecb->GetObject (&mdb, (LPMAPIPROP *)&message);
      switch (message_incoming_handler (message, hwnd, false))
        {
        case 1: 
          m_processed = true;
          break;
        case 2: 
          m_processed = true;
          m_wasencrypted = true;
          break;
        default:
          ;
        }
      
      ul_release (message, __func__, __LINE__);
      ul_release (mdb, __func__, __LINE__);
    }
  
  return S_FALSE;
}


/* Called by Exchange after a message has been read.  Returns: S_FALSE
   to signal Exchange to continue calling extensions.  EECB is a
   pointer to the IExchExtCallback interface. FLAGS are some flags. */
STDMETHODIMP 
GpgolMessageEvents::OnReadComplete (LPEXCHEXTCALLBACK eecb, ULONG flags)
{
  log_debug ("%s:%s: received; flags=%#lx m_processed=%d \n",
             SRCNAME, __func__, flags, m_processed);

  /* If the message has been processed by us (i.e. in OnRead), we now
     use our own display code.  */
  if (!flags && m_processed)
    {
      HWND hwnd = NULL;

      if (FAILED (eecb->GetWindow (&hwnd)))
        hwnd = NULL;
      log_debug ("%s:%s: (hwnd=%p)\n", SRCNAME, __func__, hwnd);
      message_display_handler (eecb, hwnd);
    }
  

  return S_FALSE;
}


/* Called by Exchange when a message will be written. Returns: S_FALSE
   to signal Exchange to continue calling extensions.  EECB is a
   pointer to the IExchExtCallback interface. */
STDMETHODIMP 
GpgolMessageEvents::OnWrite (LPEXCHEXTCALLBACK eecb)
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
      pDisp = find_outlook_property (eecb, "BodyFormat", &dispid);
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
          
          if (FAILED(eecb->GetWindow (&hWnd)))
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
   Encrypts and signs the message if the options are set.  EECB is a
   pointer to the IExchExtCallback interface.  Returns: S_FALSE to
   signal Exchange to continue calling extensions.  We return E_FAIL
   to signals Exchange an error; the message will not be sent.  Note
   that we don't return S_OK because this would mean that we rolled
   back the write operation and that no further extensions should be
   called. */
STDMETHODIMP 
GpgolMessageEvents::OnWriteComplete (LPEXCHEXTCALLBACK eecb, ULONG flags)
{
  HRESULT hrReturn = S_FALSE;
  LPMESSAGE msg = NULL;
  LPMDB pMDB = NULL;
  HWND hWnd = NULL;
  int rc;

  log_debug ("%s:%s: received\n", SRCNAME, __func__);


  if (flags & (EEME_FAILED|EEME_COMPLETE_FAILED))
    return S_FALSE; /* We don't need to rollback anything in case
                       other extensions flagged a failure. */
          
  if (!m_bOnSubmitActive) /* The user is just saving the message. */
    return S_FALSE;
  
  if (m_bWriteFailed)     /* Operation failed already. */
    return S_FALSE;

  /* Try to get the current window. */
  if (FAILED(eecb->GetWindow (&hWnd)))
    hWnd = NULL;

  /* Get the object and call the encryption or signing function.  */
  HRESULT hr = eecb->GetObject (&pMDB, (LPMAPIPROP *)&msg);
  if (SUCCEEDED (hr))
    {
      protocol_t proto = m_pExchExt->m_protoSelection;

      if (m_pExchExt->m_gpgEncrypt && m_pExchExt->m_gpgSign)
        rc = message_sign_encrypt (msg, proto, hWnd);
      else if (m_pExchExt->m_gpgEncrypt && !m_pExchExt->m_gpgSign)
        rc = message_encrypt (msg, proto, hWnd);
      else if (!m_pExchExt->m_gpgEncrypt && m_pExchExt->m_gpgSign)
        rc = message_sign (msg, proto, hWnd);
      else
        rc = 0;
      
      if (rc)
        {
          hrReturn = E_FAIL;
          m_bWriteFailed = TRUE;	
        }
    }
  
  ul_release (msg, __func__, __LINE__);
  ul_release (pMDB, __func__, __LINE__);
  
  return hrReturn;
}


/* Called by Exchange when the user selects the "check names" command.
   EECB is a pointer to the IExchExtCallback interface.  Returns
   S_FALSE to signal Exchange to continue calling extensions. */
STDMETHODIMP 
GpgolMessageEvents::OnCheckNames(LPEXCHEXTCALLBACK eecb)
{
  log_debug ("%s:%s: received\n", SRCNAME, __func__);
  return S_FALSE;
}


/* Called by Exchange when "check names" command is complete.
   EECB is a pointer to the IExchExtCallback interface.  Returns
   S_FALSE to signal Exchange to continue calling extensions. */
STDMETHODIMP 
GpgolMessageEvents::OnCheckNamesComplete (LPEXCHEXTCALLBACK eecb,ULONG flags)
{
  log_debug ("%s:%s: received\n", SRCNAME, __func__);
  return S_FALSE;
}


/* Called by Exchange before the message data will be written and
   submitted to MAPI.  EECB is a pointer to the IExchExtCallback
   interface.  Returns S_FALSE to signal Exchange to continue calling
   extensions. */
STDMETHODIMP 
GpgolMessageEvents::OnSubmit (LPEXCHEXTCALLBACK eecb)
{
  log_debug ("%s:%s: received\n", SRCNAME, __func__);
  m_bOnSubmitActive = TRUE;
  m_bWriteFailed = FALSE;
  return S_FALSE;
}


/* Called by Exchange after the message has been submitted to MAPI.
   EECB is a pointer to the IExchExtCallback interface. */
STDMETHODIMP_ (VOID) 
GpgolMessageEvents::OnSubmitComplete (LPEXCHEXTCALLBACK eecb, ULONG flags)
{
  log_debug ("%s:%s: received\n", SRCNAME, __func__);
  m_bOnSubmitActive = FALSE; 
}


