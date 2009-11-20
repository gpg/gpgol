/* mailitem.cpp - Code to handle the Outlook MailItem
 *	Copyright (C) 2009 g10 Code GmbH
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
#include <olectl.h>

#include "common.h"
#include "oomhelp.h"
#include "eventsink.h"
#include "mailitem.h"


/* Subclass of ItemEvents so that we can hook into the events.  */
BEGIN_EVENT_SINK(GpgolItemEvents, IOOMItemEvents)
  STDMETHOD(Read)(THIS_ );
  STDMETHOD(Write)(THIS_ PBOOL cancel);
  STDMETHOD(Open)(THIS_ PBOOL cancel);
  STDMETHOD(Close)(THIS_ PBOOL cancel);
EVENT_SINK_DEFAULT_CTOR(GpgolItemEvents)
EVENT_SINK_DEFAULT_DTOR(GpgolItemEvents)
EVENT_SINK_INVOKE(GpgolItemEvents)
{
  HRESULT hr;
  (void)lcid; (void)riid; (void)result; (void)exepinfo; (void)argerr;

  switch (dispid)
    {
    case 0xf001:
      if (!(flags & DISPATCH_METHOD))
        goto badflags;
      if (parms->cArgs != 0)
        hr = DISP_E_BADPARAMCOUNT;
      else
        {
          Read ();
          hr = S_OK;
        }
      break;

    case 0xf004:
      if (!(flags & DISPATCH_METHOD))
        goto badflags;
      if (!parms) 
        hr = DISP_E_PARAMNOTOPTIONAL;
      else if (parms->cArgs != 1)
        hr = DISP_E_BADPARAMCOUNT;
      else if (parms->rgvarg[0].vt != (VT_BOOL|VT_BYREF))
        hr = DISP_E_BADVARTYPE;
      else
        {
          BOOL cancel_default = !!*parms->rgvarg[0].pboolVal;
          switch (dispid)
            {
            case 0xf002: Write (&cancel_default); break;
            case 0xf003: Open (&cancel_default); break;
            case 0xf004: Close (&cancel_default); break;
            }
          *parms->rgvarg[0].pboolVal = (cancel_default 
                                        ? VARIANT_TRUE:VARIANT_FALSE);
          hr = S_OK;
        }
      break;

    badflags:
    default:
      hr = DISP_E_MEMBERNOTFOUND;
    }
  return hr;
}
END_EVENT_SINK(GpgolItemEvents, IID_IOOMItemEvents)


/* This is the event sink for a read event.  */
STDMETHODIMP
GpgolItemEvents::Read (void)
{
  log_debug ("%s:%s: Called", SRCNAME, __func__);

  return S_OK;
}


/* This is the event sink for a write event.  */
STDMETHODIMP
GpgolItemEvents::Write (PBOOL cancel_default)
{
  (void)cancel_default;
  log_debug ("%s:%s: Called", SRCNAME, __func__);

  return S_OK;
}


/* This is the event sink for an open event.  */
STDMETHODIMP
GpgolItemEvents::Open (PBOOL cancel_default)
{
  (void)cancel_default;
  log_debug ("%s:%s: Called", SRCNAME, __func__);

  return S_OK;
}

/* This is the event sink for a close event.  */
STDMETHODIMP
GpgolItemEvents::Close (PBOOL cancel_default)
{
  (void)cancel_default;
  log_debug ("%s:%s: Called", SRCNAME, __func__);

  return S_OK;
}

