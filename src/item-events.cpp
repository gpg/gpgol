/* item-events.cpp - GpgolItemEvents implementation
 *	Copyright (C) 2007 g10 Code GmbH
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

#error not used because it requires an ECF

#ifdef HAVE_CONFIG_H
#include <config.h>
#endif
#include <windows.h>

#include "mymapi.h"
#include "mymapitags.h"
#include "myexchext.h"
#include "common.h"
#include "olflange-def.h"
#include "olflange.h"
#include "message.h"
#include "mapihelp.h"
#include "display.h"
#include "item-events.h"

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
GpgolItemEvents::GpgolItemEvents (GpgolExt *pParentInterface)
{ 
  m_pExchExt = pParentInterface;
  m_ref = 0;
  m_processed = false;
  m_wasencrypted = false;
}


/* The QueryInterfac.  */
STDMETHODIMP 
GpgolItemEvents::QueryInterface (REFIID riid, LPVOID FAR *ppvObj)
{
  *ppvObj = NULL;
  if (riid == IID_IOutlookExtItemEvents)
    {
      *ppvObj = (LPVOID)this;
      AddRef ();
      return S_OK;
    }
  if (riid == IID_IUnknown)
    {
      *ppvObj = (LPVOID)m_pExchExt;
      m_pExchExt->AddRef ();
      return S_OK;
    }
  return E_NOINTERFACE;
}
 

/* This method is called if an item is about to being displayed.
   Possible return values are: 

      S_FALSE - Let Outlook continue the operation.

      E_ABORT - Abort the open operation and the item is not being
                displayed.
*/
STDMETHODIMP 
GpgolItemEvents::OnOpen (LPEXCHEXTCALLBACK eecb)
{
  LPMDB mdb = NULL;
  LPMESSAGE message = NULL;
  HWND hwnd = NULL;
  
  log_debug ("%s:%s: received\n", SRCNAME, __func__);

  m_wasencrypted = false;
  eecb->GetObject (&mdb, (LPMAPIPROP *)&message);
  if (message_incoming_handler (message, hwnd, false))
    m_processed = TRUE;
  ul_release (message, __func__, __LINE__);
  ul_release (mdb, __func__, __LINE__);

  return S_FALSE;
}


/* Like all the other Complete methods this one is called after OnOpen
   has been called for all registered extensions.  FLAGS may have these
   values:

    0 - The open action has not been canceled.

    EEME_FAILED or EEME_COMPLETE_FAILED - both indicate that the
        open action has been canceled.  

    Note that this method may be called more than once for each OnOpen
    and may even occur without an OnOpen (i.e. while setting up a new
    extension).

    The same return values as for OnOpen may be used..
*/ 
STDMETHODIMP 
GpgolItemEvents::OnOpenComplete (LPEXCHEXTCALLBACK eecb, ULONG flags)
{
  log_debug ("%s:%s: received, flags=%#lx", SRCNAME, __func__, flags);

  /* If the message has been processed by us (i.e. in OnOpen), we now
     use our own display code.  */
  if (!flags && m_processed)
    {
      HWND hwnd = NULL;

      if (FAILED (eecb->GetWindow (&hwnd)))
        hwnd = NULL;
      if (message_display_handler (eecb, hwnd))
        m_wasencrypted = true;
    }
  
  return S_FALSE;
}


/* This method is called if an item's window received a close request.
   Possible return values are: 

      S_FALSE - Let Outlook continue the operation.

      E_ABORT - Abort the close operation and don't dismiss the window.
*/
STDMETHODIMP 
GpgolItemEvents::OnClose (LPEXCHEXTCALLBACK eecb, ULONG save_options)
{
  log_debug ("%s:%s: received, options=%#lx", SRCNAME, __func__, save_options);

  return S_FALSE;
}


/* This is the corresponding Complete method for OnClose.  See
   OnOpenComplete for a description.  */
STDMETHODIMP 
GpgolItemEvents::OnCloseComplete (LPEXCHEXTCALLBACK eecb, ULONG flags)
{
  log_debug ("%s:%s: received, flags=%#lx", SRCNAME, __func__, flags);

  if (m_wasencrypted)
    {
      /* Delete any body parts so that encrypted stuff won't show up
         in the clear. */
      HRESULT hr;
      LPMESSAGE message = NULL;
      LPMDB mdb = NULL;
      
      log_debug ("%s:%s: deleting body properties", SRCNAME, __func__);
      hr = eecb->GetObject (&mdb, (LPMAPIPROP *)&message);
      if (SUCCEEDED (hr))
        {
          mapi_delete_body_parts (message, KEEP_OPEN_READWRITE);
          m_wasencrypted = false;
        }  
      else
        log_debug_w32 (hr, "%s:%s: error getting message", 
                       SRCNAME, __func__);
     
      ul_release (message, __func__, __LINE__);
      ul_release (mdb, __func__, __LINE__);
    }

  return S_FALSE;
}
