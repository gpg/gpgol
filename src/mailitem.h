/* mailitem.h - Defs to handle the MailItem
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

#ifndef MAILITEM_H
#define MAILITEM_H

#include <commctrl.h>
#include "oomhelp.h"

DEFINE_OLEGUID(IID_IOOMItemEvents,            0x0006302B, 0, 0);
DEFINE_OLEGUID(IID_IOOMMailItem,              0x00063034, 0, 0);


typedef struct IOOMItemEvents IOOMItemEvents;
typedef IOOMItemEvents *LPOOMITEMEVENTS;

struct IOOMMailItem;
typedef IOOMMailItem *LPOOMMAILITEM;


EXTERN_C const IID IID_IOOMItemEvents;
#undef INTERFACE
#define INTERFACE  IOOMItemEvents
DECLARE_INTERFACE_(IOOMItemEvents, IDispatch)
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

  /*** IOOMItemEvents methods ***/
  /* WARNING: This is for documentation only; I have no idea about the
     vtable layout.  However it doesn't matter because we only use it
     via the IDispatch interface.  */
  /* dispid=0xf001 */
  STDMETHOD(Read)(THIS_ ) PURE;

  /* dispid=0xf002 */
  STDMETHOD(Write)(THIS_ PBOOL cancel) PURE;

  /* dispid=0xf003 */
  STDMETHOD(Open)(THIS_ PBOOL cancel) PURE;

  /* dispid=0xf004 */
  STDMETHOD(Close)(THIS_ PBOOL cancel) PURE;

  /* dispid=0xf005 */
  STDMETHOD(Send)(THIS_ PBOOL cancel) PURE;

  /* dispid=0xf006 */
  //STDMETHOD(CustomAction)(THIS_ LPDISPATCH action, LPDISPATCH response,
  //                        PBOOL cancel);
  /* dispid=0xf008 */
  //STDMETHOD(CustomPropertyChange)(THIS_ VARIANT name);
  /* dispid=0xf009 */
  //STDMETHOD(PropertyChange)(THIS_ VARIANT name);

  /* dispid=0xf00a 
     OL2003: Called between the ECE OnCheckNames and OnCheckNamesComplete.  */
  //STDMETHOD(BeforeCheckName)(THIS_ PBOOL cancel);

  /* dispid=0xf00b */
  //STDMETHOD(AttachmentAdd)(THIS_ LPDISPATCH att);
  /* dispid=0xf00c */
  //STDMETHOD(AttachmentRead)(THIS_ LPDISPATCH att);
  /* dispid=0xf00d */
  //STDMETHOD(BeforeAttachmentSave)(THIS_ LPDISPATCH att, PBOOL cancel);
  /* dispid=0xf468 */
  //STDMETHOD(Forward)(THIS_ LPDISPATCH forward, PBOOL cancel);
  /* dispid=0xf466 */
  //STDMETHOD(Reply)(THIS_ LPDISPATCH response, PBOOL cancel);
  /* dispid=0xf467 */
  //STDMETHOD(ReplyAll)(THIS_ LPDISPATCH response, PBOOL cancel);
  /* dispid=0xfa75 */
  //STDMETHOD(BeforeDelete)(THIS_ LPDISPATCH item, PBOOL cancel);
};


LPDISPATCH install_GpgolItemEvents_sink (LPDISPATCH item);



#endif /*MAILITEM_H*/
