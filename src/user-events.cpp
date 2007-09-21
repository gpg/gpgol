/* user-events.cpp - Subclass impl of IExchExtUserEvents
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
#include "display.h"
#include "common.h"
#include "msgcache.h"
#include "engine.h"
#include "mapihelp.h"

#include "olflange-def.h"
#include "olflange.h"
#include "user-events.h"


#define TRACEPOINT() do { log_debug ("%s:%s:%d: tracepoint\n", \
                                     SRCNAME, __func__, __LINE__); \
                        } while (0)


/* Wrapper around UlRelease with error checking. */
/* FIXME: Duplicated code.  */
#if 0
static void 
ul_release (LPVOID punk)
{
  ULONG res;
  
  if (!punk)
    return;
  res = UlRelease (punk);
//   log_debug ("%s UlRelease(%p) had %lu references\n", __func__, punk, res);
}
#endif




/* Our constructor.  */
GpgolUserEvents::GpgolUserEvents (GpgolExt *pParentInterface)
{ 
  m_pExchExt = pParentInterface;
  m_lRef = 0; 
}


/* The QueryInterface which does the actual subclassing.  */
STDMETHODIMP 
GpgolUserEvents::QueryInterface (REFIID riid, LPVOID FAR *ppvObj)
{   
  *ppvObj = NULL;
  if (riid == IID_IExchExtUserEvents)
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



/* Called from Outlook for all selection changes.

   PEECB is a pointer to the IExchExtCallback interface.  */
STDMETHODIMP_ (VOID)
GpgolUserEvents::OnSelectionChange (LPEXCHEXTCALLBACK eecb) 
{
  HRESULT hr;
  ULONG count, objtype;
  char msgclass[256];

  log_debug ("%s:%s: received\n", SRCNAME, __func__);

  hr = eecb->GetSelectionCount (&count);
  if (SUCCEEDED (hr) && count > 0)
    {
      hr = eecb->GetSelectionItem (0L, NULL, NULL, &objtype,
                                   msgclass, sizeof msgclass -1, NULL, 0L);
      if (SUCCEEDED(hr) && objtype == MAPI_MESSAGE)
        {
          log_debug ("%s:%s: message class: %s\n",
            SRCNAME, __func__, msgclass);
        }
    }

}

/* I assume this is called from Outlook for all object changes.

   PEECB is a pointer to the IExchExtCallback interface.  */
STDMETHODIMP_ (VOID)
GpgolUserEvents::OnObjectChange (LPEXCHEXTCALLBACK eecb) 
{ 
  log_debug ("%s:%s: received\n", SRCNAME, __func__);

}

