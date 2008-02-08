/* session-events.cpp - Subclass impl of IExchExtSessionEvents
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
#include "session-events.h"


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
  if (opt.enable_debug & DBG_MEMORY)
    log_debug ("%s UlRelease(%p) had %lu references\n", __func__, punk, res);
}







/* Our constructor.  */
GpgolSessionEvents::GpgolSessionEvents (GpgolExt *pParentInterface)
{ 
  m_pExchExt = pParentInterface;
  m_lRef = 0; 
}


/* The QueryInterface which does the actual subclassing.  */
STDMETHODIMP 
GpgolSessionEvents::QueryInterface (REFIID riid, LPVOID FAR *ppvObj)
{   
  *ppvObj = NULL;
  if (riid == IID_IExchExtSessionEvents)
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



/* Called from Exchange when a new message arrives.  Returns: S_FALSE
   to signal Exchange to continue calling extensions.  PEECB is a
   pointer to the IExchExtCallback interface. */
STDMETHODIMP 
GpgolSessionEvents::OnDelivery (LPEXCHEXTCALLBACK pEECB) 
{
  LPMDB pMDB = NULL;
  LPMESSAGE pMessage = NULL;

  log_debug ("%s:%s: received\n", SRCNAME, __func__);
  pEECB->GetObject (&pMDB, (LPMAPIPROP *)&pMessage);
  log_mapi_property (pMessage, PR_MESSAGE_CLASS,"PR_MESSAGE_CLASS");
  /* Note, that at this point even an OpenPGP signed message has the
     message class IPM.Note.SMIME.MultipartSigned.  If we would not
     change the message class here, OL will change it later (before an
     OnRead) to IPM.Note. */
  mapi_change_message_class (pMessage, 0);
  log_mapi_property (pMessage, PR_MESSAGE_CLASS,"PR_MESSAGE_CLASS");
  ul_release (pMessage);
  ul_release (pMDB);

  return S_FALSE;
}

