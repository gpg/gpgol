/* attached-file-events.cpp - GpgolAttachedFileEvents implementation
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
#include "myexchext.h"
#include "intern.h"
#include "olflange-def.h"
#include "olflange.h"
#include "attached-file-events.h"

#define TRACEPOINT() do { log_debug ("%s:%s:%d: tracepoint\n", \
                                     SRCNAME, __func__, __LINE__); \
                        } while (0)



/* Our constructor.  */
GpgolAttachedFileEvents::GpgolAttachedFileEvents (GpgolExt *pParentInterface)
{ 
  m_pExchExt = pParentInterface;
  m_ref = 0;
}


/* The QueryInterfac.  */
STDMETHODIMP 
GpgolAttachedFileEvents::QueryInterface (REFIID riid, LPVOID FAR *ppvObj)
{
  *ppvObj = NULL;
  if (riid == IID_IExchExtAttachedFileEvents)
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
 

/* Fixme: We need to figure out what this exactly does.  There is no
   public information available exepct for the MAPI book which is out
   of print.  */
STDMETHODIMP 
GpgolAttachedFileEvents::OnReadPattFromSzFile 
  (LPATTACH att, LPTSTR file, ULONG flags)
{
  log_debug ("%s:%s: att=%p file=`%s' flags=%lx\n", 
	     SRCNAME, __func__, att, file, flags);
  return S_FALSE;
}
  
 
STDMETHODIMP 
GpgolAttachedFileEvents::OnWritePattToSzFile 
  (LPATTACH att, LPTSTR file, ULONG flags)
{
  log_debug ("%s:%s: att=%p file=`%s' flags=%lx\n", 
	     SRCNAME, __func__, att, file, flags);
  return S_FALSE;
}


STDMETHODIMP
GpgolAttachedFileEvents::QueryDisallowOpenPatt (LPATTACH att)
{
  log_debug ("%s:%s: att=%p\n", SRCNAME, __func__, att);
  return S_FALSE;
}


STDMETHODIMP 
GpgolAttachedFileEvents::OnOpenPatt (LPATTACH att)
{
  log_debug ("%s:%s: att=%p\n", SRCNAME, __func__, att);
  return S_FALSE;
}


STDMETHODIMP 
GpgolAttachedFileEvents::OnOpenSzFile (LPTSTR file, ULONG flags)
{
  log_debug ("%s:%s: file=`%s' flags=%lx\n", SRCNAME, __func__, file, flags);
  return S_FALSE;
}
