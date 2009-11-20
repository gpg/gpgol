/* inspectors.h - Defs to handle the OOM Inspectors
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

#ifndef INSPECTORS_H
#define INSPECTORS_H

#include "myexchext.h"


DEFINE_OLEGUID(IID_IOOMInspector,             0x00063005, 0, 0);
DEFINE_OLEGUID(IID_IOOMInspectorEvents,       0x0006302A, 0, 0);
DEFINE_OLEGUID(IID_IOOMInspectors,            0x00063008, 0, 0);
DEFINE_OLEGUID(IID_IOOMInspectorsEvents,      0x00063079, 0, 0);

typedef struct IOOMInspector IOOMInspector;
typedef IOOMInspector *LPOOMINSPECTOR;

typedef struct IOOMInspectorEvents IOOMInspectorEvents;
typedef IOOMInspectorEvents *LPOOMINSPECTOREVENTS;

typedef struct IOOMInspectorsEvents IOOMInspectorsEvents;
typedef IOOMInspectorsEvents *LPOOMINSPECTORSEVENTS;


EXTERN_C const IID IID_IOOMInspector;
#undef INTERFACE
#define INTERFACE  IOOMInspector
DECLARE_INTERFACE_(IOOMInspector, IDispatch)
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

  /*** IOOM_Inspector methods ***/
  /* Fixme: Don't know.  */
};


EXTERN_C const IID IID_IOOMInspectorEvents;
#undef INTERFACE
#define INTERFACE  IOOMInspectorEvents
DECLARE_INTERFACE_(IOOMInspectorEvents, IDispatch)
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

  /*** IOOM_Inspector methods ***/
  /* dispid=0xf001 */
  STDMETHOD_(void, Activate)(THIS_);
  /* dispid=0xfa11 */
  STDMETHODIMP(BeforeMaximize)(THIS_ PBOOL);
  /* dispid=0xfa12 */
  STDMETHODIMP(BeforeMinimize)(THIS_ PBOOL);
  /* dispid=0xfa13 */
  STDMETHODIMP(BeforeMove)(THIS_ PBOOL);
  /* dispid=0xfa14 */
  STDMETHODIMP(BeforeSize)(THIS_ PBOOL);
  /* dispid=0xf008 */
  STDMETHOD_(void, Close)(THIS_);
  /* dispid=0xf006 */
  STDMETHOD_(void, Deactivate)(THIS_);
};


EXTERN_C const IID IID_IOOMInspectorsEvents;
#undef INTERFACE
#define INTERFACE  IOOMInspectorsEvents
DECLARE_INTERFACE_(IOOMInspectorsEvents, IDispatch)
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

  /*** IOOMInspectorsEvents methods ***/
  /* dispid=0xf001 */
  STDMETHOD(NewInspector)(THIS_ LPOOMINSPECTOR) PURE;
};


/* Create a new sink and attach it to OBJECT.  */
LPDISPATCH install_GpgolInspectorsEvents_sink (LPDISPATCH object);

LPDISPATCH install_GpgolInspectorEvents_sink (LPDISPATCH object);
void detach_GpgolInspectorEvents_sink (LPDISPATCH sink);


void proc_inspector_button_click (LPDISPATCH button,
                                  const char *tag, int instid);

int get_inspector_composer_flags (LPDISPATCH inspector,
                                  bool *r_sign, bool *r_encrypt);
int set_inspector_composer_flags (LPDISPATCH inspector,
                                  bool sign, bool encrypt);




#endif /*INSPECTORS_H*/
