/* user-events.cpp - Subclass impl of IExchExtUserEvents
 *	Copyright (C) 2007, 2008 g10 Code GmbH
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
#include "user-events.h"


#define TRACEPOINT() do { log_debug ("%s:%s:%d: tracepoint\n", \
                                     SRCNAME, __func__, __LINE__); \
                        } while (0)


/* Wrapper around UlRelease with error checking. */
// static void 
// ul_release (LPVOID punk, const char *func, int lnr)
// {
//   ULONG res;
  
//   if (!punk)
//     return;
//   res = UlRelease (punk);
//   log_debug ("%s:%s:%d: UlRelease(%p) had %lu references\n", 
//              SRCNAME, func, lnr, punk, res);
// }




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
      /* Get the first selected item.  */
      hr = eecb->GetSelectionItem (0L, NULL, NULL, &objtype,
                                   msgclass, sizeof msgclass -1, NULL, 0L);
      if (SUCCEEDED(hr) && objtype == MAPI_MESSAGE)
        {
          log_debug ("%s:%s: message class: %s\n",
                     SRCNAME, __func__, msgclass);

          /* If SMIME has been enabled and the current message is of
             class SMIME or in the past processed by CryptoEx, we
             change the message class. */ 
          // Unfortunaltely we can't use this because:
          // 1. GetSelectionItem is as usual heavily undocumented and
          // we need to guess a bit to see how to get message from the
          // EntryID (2nd and 3rd arg).  2.  There are reports that
          // OL2007 crashes when changing the message here.
//           if (opt.enable_smime 
//               && (!strncmp (msgclass, "IPM.Note.SMIME", 14)
//                   || !strncmp (msgclass, "IPM.Note.Secure.Cex", 19)))
//             {
//               LPMESSAGE message = NULL;
//               LPMDB mdb = NULL;

//               hr = eecb->GetObject (&mdb, (LPMAPIPROP *)&message);
//               if (SUCCEEDED (hr) && !mapi_has_sig_status (message))
//                 {
//                   log_debug ("%s:%s: message class not yet checked"
//                              " - doing now\n", SRCNAME, __func__);
//                   mapi_change_message_class (message);
//                 }
//               ul_release (message, __func__, __LINE__);
//               ul_release (mdb, __func__, __LINE__);
//             }
        }
      else if (SUCCEEDED(hr) && objtype == MAPI_FOLDER)
        {
          log_debug ("%s:%s: objtype: %lu\n",
                     SRCNAME, __func__, objtype);
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

