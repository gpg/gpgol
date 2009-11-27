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
#include "mimeparser.h"
#include "mimemaker.h"
#include "message.h"
#include "message-events.h"

#include "explorers.h"
#include "inspectors.h"

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


static int
get_crypto_flags (HWND hwnd, bool *r_sign, bool *r_encrypt)
{
  LPDISPATCH inspector;
  int rc;

  inspector = get_inspector_from_hwnd (hwnd);
  if (!inspector)
    {
      log_error ("%s:%s: inspector not found", SRCNAME, __func__);
      rc = -1;
    }
  else
    {
      rc = get_inspector_composer_flags (inspector, r_sign, r_encrypt);
      inspector->Release ();
    }
  return rc;
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
  
  if ( (m_gotinspector || opt.preview_decrypt) && !opt.disable_gpgol )
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
  if (!flags && m_processed && !opt.disable_gpgol)
    {
      HRESULT hr;
      LPMDB mdb = NULL;
      LPMESSAGE message = NULL;
      HWND hwnd = NULL;

      if (FAILED (eecb->GetWindow (&hwnd)))
        hwnd = NULL;
      hr = eecb->GetObject (&mdb, (LPMAPIPROP *)&message);
      if (hr != S_OK || !message) 
        log_error ("%s:%s: error getting message: hr=%#lx",
                   SRCNAME, __func__, hr);
      else
        {
          LPDISPATCH inspector = get_inspector_from_hwnd (hwnd);
          message_display_handler (message, inspector, hwnd);
          if (inspector)
            inspector->Release ();
        }
      if (message)
        message->Release ();
      if (mdb)
        mdb->Release ();
    }
  

  return S_FALSE;
}


/* Called by Exchange when a message will be written. Returns: S_FALSE
   to signal Exchange to continue calling extensions.  EECB is a
   pointer to the IExchExtCallback interface. */
STDMETHODIMP 
GpgolMessageEvents::OnWrite (LPEXCHEXTCALLBACK eecb)
{
  LPDISPATCH obj;
  HWND hwnd = NULL;
  bool sign, encrypt, need_crypto;
  int bodyfmt;

  if (FAILED (eecb->GetWindow (&hwnd)))
    hwnd = NULL;
  log_debug ("%s:%s: received (hwnd=%p)", SRCNAME, __func__, hwnd);

  need_crypto = (!get_crypto_flags (hwnd, &sign, &encrypt)
                 && (sign || encrypt));
    
  /* If we are going to encrypt, check that the BodyFormat is
     something we support.  This helps avoiding surprise by sending
     out unencrypted messages. */
  if (need_crypto && !opt.disable_gpgol)
    {
      obj = get_eecb_object (eecb);
      if (!obj)
        bodyfmt = -1;
      else
        {
          bodyfmt = get_oom_int (obj, "BodyFormat");
          obj->Release ();
        }

      if (bodyfmt == 1)
        m_want_html = 0;
      else if (bodyfmt == 2)
        m_want_html = 1;
      else
        {
          log_debug ("%s:%s: BodyFormat is %d", SRCNAME, __func__, bodyfmt);
          MessageBox (hwnd,
                      _("Sorry, we can only encrypt plain text messages and\n"
                      "no RTF messages. Please make sure that only the text\n"
                      "format has been selected."),
                      "GpgOL", MB_ICONERROR|MB_OK);

          m_bWriteFailed = TRUE;	
          return E_FAIL;
        }
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
  HWND hwnd = NULL;
  int rc;

  if (FAILED(eecb->GetWindow (&hwnd)))
    hwnd = NULL;
  log_debug ("%s:%s: received (hwnd=%p)", SRCNAME, __func__, hwnd);


  if (flags & (EEME_FAILED|EEME_COMPLETE_FAILED))
    return S_FALSE; /* We don't need to rollback anything in case
                       other extensions flagged a failure. */

  if (opt.disable_gpgol)
    return S_FALSE;
          
  if (!m_bOnSubmitActive) /* The user is just saving the message. */
    return S_FALSE;
  
  if (m_bWriteFailed)     /* Operation failed already. */
    return S_FALSE;


  /* Get the object and call the encryption or signing function.  */
  HRESULT hr = eecb->GetObject (&pMDB, (LPMAPIPROP *)&msg);
  if (SUCCEEDED (hr))
    {
      protocol_t proto = PROTOCOL_UNKNOWN; /* Let the UI server select
                                              the protocol.  */
      bool sign, encrypt;

      if (get_crypto_flags (hwnd, &sign, &encrypt))
        rc = -1;
      else if (encrypt && sign)
        rc = message_sign_encrypt (msg, proto, hwnd);
      else if (encrypt && !sign)
        rc = message_encrypt (msg, proto, hwnd);
      else if (!encrypt && sign)
        rc = message_sign (msg, proto, hwnd);
      else
        {
          /* In case this is a forward message which is not to be
             signed or encrypted we need to remove a possible body
             attachment.  */
          if (mapi_delete_gpgol_body_attachment (msg))
            mapi_save_changes (msg, KEEP_OPEN_READWRITE);
          rc = 0;
        }
      
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
  (void)eecb;

  log_debug ("%s:%s: received\n", SRCNAME, __func__);
  return S_FALSE;
}


/* Called by Exchange when "check names" command is complete.
   EECB is a pointer to the IExchExtCallback interface.  Returns
   S_FALSE to signal Exchange to continue calling extensions. */
STDMETHODIMP 
GpgolMessageEvents::OnCheckNamesComplete (LPEXCHEXTCALLBACK eecb,ULONG flags)
{
  (void)eecb;
  (void)flags;

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
  (void)eecb;

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
  (void)eecb;
  (void)flags;

  log_debug ("%s:%s: received\n", SRCNAME, __func__);
  m_bOnSubmitActive = FALSE; 
}


