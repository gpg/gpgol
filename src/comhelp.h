/* comhelp.h - Helper macros to define / declare COM interfaces.
 *    Copyright (C) 2015 Intevation GmbH
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
#ifndef COMHELP_H
#define COMHELP_H

/*** IUnknown methods ***/
#define DECLARE_IUNKNOWN_METHODS                        \
  STDMETHOD(QueryInterface)(THIS_ REFIID, PVOID*) PURE; \
  STDMETHOD_(ULONG,AddRef)(THIS) PURE;                  \
  STDMETHOD_(ULONG,Release)(THIS) PURE

/*** IDispatch methods ***/
#define DECLARE_IDISPATCH_METHODS                                             \
  STDMETHOD(GetTypeInfoCount)(THIS_ UINT*) PURE;                              \
  STDMETHOD(GetTypeInfo)(THIS_ UINT, LCID, LPTYPEINFO*) PURE;                 \
  STDMETHOD(GetIDsOfNames)(THIS_ REFIID, LPOLESTR*, UINT, LCID, DISPID*) PURE;\
  STDMETHOD(Invoke)(THIS_ DISPID, REFIID, LCID, WORD,                         \
                    DISPPARAMS*, VARIANT*, EXCEPINFO*, UINT*) PURE

#endif // COMHELP_H
