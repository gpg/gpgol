/* item-events.cpp - GpgolItemEvents implementation
 *	Copyright (C) 2007 g10 Code GmbH
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
#include "common.h"
#include "olflange-def.h"
#include "olflange.h"
#include "item-events.h"

#define TRACEPOINT() do { log_debug ("%s:%s:%d: tracepoint\n", \
                                     SRCNAME, __func__, __LINE__); \
                        } while (0)



/* Our constructor.  */
GpgolItemEvents::GpgolItemEvents (GpgolExt *pParentInterface)
{ 
  m_pExchExt = pParentInterface;
  m_ref = 0;
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
GpgolItemEvents::OnOpen (LPEXCHEXTCALLBACK peecb)
{
  log_debug ("%s:%s: received", SRCNAME, __func__);
  return S_FALSE;
}


/* Like all the other Complete methods this one is called after OnOpen
   has been called for all registred extensions.  FLAGS may have tehse
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
GpgolItemEvents::OnOpenComplete (LPEXCHEXTCALLBACK peecb, ULONG flags)
{
  log_debug ("%s:%s: received, flags=%#lx", SRCNAME, __func__, flags);
  return S_FALSE;
}


/* This method is called if an item's window received a close request.
   Possible return values are: 

      S_FALSE - Let Outlook continue the operation.

      E_ABORT - Abort the close operation and don't dismiss the window.
*/
STDMETHODIMP 
GpgolItemEvents::OnClose (LPEXCHEXTCALLBACK peecb, ULONG save_options)
{
  log_debug ("%s:%s: received, options=%#lx", SRCNAME, __func__, save_options);
  return S_FALSE;
}


/* This is the corresponding Complete method for OnClose.  See
   OnOpenComplete for a description.  */
STDMETHODIMP 
GpgolItemEvents::OnCloseComplete (LPEXCHEXTCALLBACK peecb, ULONG flags)
{
  log_debug ("%s:%s: received, flags=%#lx", SRCNAME, __func__, flags);
  return S_FALSE;
}
