/* explorers.h - Defs to handle the OOM Explorers
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

#ifndef EXPLORERS_H
#define EXPLORERS_H

#include "myexchext.h"

DEFINE_OLEGUID(IID_IOOMExplorer,              0x00063003, 0, 0);
DEFINE_OLEGUID(IID_IOOMExplorers,             0x0006300A, 0, 0);
DEFINE_OLEGUID(IID_IOOMExplorersEvents,       0x00063078, 0, 0);


typedef struct IOOMExplorer IOOMExplorer;
typedef IOOMExplorer *LPOOMEXPLORER;

typedef struct IOOMExplorersEvents IOOMExplorersEvents;
typedef IOOMExplorersEvents *LPOOMEXPLORERSEVENTS;


EXTERN_C const IID IID_IOOMExplorer;
#undef INTERFACE
#define INTERFACE  IOOMExplorer
DECLARE_INTERFACE_(IOOMExplorer, IDispatch)
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

  /*** IOOM_Explorer methods ***/
  /* Activate, Close, .... */
};


EXTERN_C const IID IID_IOOMExplorersEvents;
#undef INTERFACE
#define INTERFACE  IOOMExplorersEvents
DECLARE_INTERFACE_(IOOMExplorersEvents, IDispatch)
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

  /*** IOOMExplorersEvents methods ***/
  /* dispid=0xf001 */
  STDMETHOD(NewExplorer)(THIS_ LPOOMEXPLORER) PURE;
};


/* Install an explorers collection event sink in OBJ.  Returns the
   event sink object which must be released by the caller if it is not
   anymore required.  */
LPDISPATCH install_GpgolExplorersEvents_sink (LPDISPATCH obj);

void add_explorer_controls (LPOOMEXPLORER explorer);



#endif /*EXPLORERS_H*/
