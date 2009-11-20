/* cmdbarcontrols.h - Defs to handle the CommandBarControls
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

#ifndef CMDBARCONTROLS_H
#define CMDBARCONTROLS_H

#include <commctrl.h>
#include "oomhelp.h"

DEFINE_OLEGUID(IID_IOOMCommandBarButtonEvents,  0x000c0351, 0, 0);

typedef struct IOOMCommandBarButtonEvents IOOMCommandBarButtonEvents;
typedef IOOMCommandBarButtonEvents *LPOOMCOMMANDBARBUTTONEVENTS;

struct IOOMCommandBarButton;
typedef IOOMCommandBarButton *LPOOMCOMMANDBARBUTTON;

struct GpgolCommandBarButtonEvents;
typedef GpgolCommandBarButtonEvents *LPGPGOLCOMMANDBARBUTTONEVENTS;


EXTERN_C const IID IID_IOOMCommandBarButtonEvents;
#undef INTERFACE
#define INTERFACE  IOOMCommandBarButtonEvents
DECLARE_INTERFACE_(IOOMCommandBarButtonEvents, IDispatch)
{
  /*** IUnknown methods ***/
  STDMETHOD(QueryInterface) (THIS_ REFIID riid, LPVOID FAR * lppvObj) PURE;
  STDMETHOD_(ULONG,AddRef) (THIS)  PURE;
  STDMETHOD_(ULONG,Release) (THIS) PURE;

  /*** IDispatch methods ***/
  STDMETHOD(GetTypeInfoCount)(THIS_ UINT*) PURE;
  STDMETHOD(GetTypeInfo)(THIS_ UINT, LCID, LPTYPEINFO*) PURE;
  STDMETHOD(GetIDsOfNames)(THIS_ REFIID, LPOLESTR*, UINT, LCID, DISPID*) PURE;
  STDMETHOD(Invoke)(THIS_ DISPID, REFIID, LCID, WORD,
                    DISPPARAMS*, VARIANT*, EXCEPINFO*, UINT*) PURE;

  /*** IOOMCommandBarButtonEvents methods ***/
  /* dispid=1 */
  STDMETHOD(Click)(THIS_ LPDISPATCH, PBOOL) PURE;
};



LPDISPATCH install_GpgolCommandBarButtonEvents_sink (LPDISPATCH button);
void       detach_GpgolCommandBarButtonEvents_sink (LPDISPATCH sink);



#endif /*CMDBARCONTROLS_H*/
