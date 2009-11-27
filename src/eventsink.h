/* eventsink.h - Macros to implement an OLE event sink
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

/* This is our lame implementation of event sinks.  It is sufficient
   for our purpose but far from optimal; e.g. we should not simply
   provide stubs but do real sub-classing by modifying the vtables.  */

#ifndef EVENTSINK_H
#define EVENTSINK_H

#define debug_oom        (opt.enable_debug & DBG_OOM)
#define debug_oom_extra  (opt.enable_debug & DBG_OOM_EXTRA)


#define BEGIN_EVENT_SINK(subcls,parentcls)                               \
class subcls : public parentcls                                          \
{                                                                        \
 public:                                                                 \
  subcls (void);                                                         \
  ~subcls (void);                                                        \
  LPDISPATCH m_object;                                                   \
  LPCONNECTIONPOINT m_pCP;                                               \
  DWORD m_cookie;                                                        \
 private:                                                                \
  ULONG     m_ref;                                                       \
 public:                                                                 \
  STDMETHODIMP QueryInterface (REFIID riid, LPVOID FAR *ppvObj);         \
  inline STDMETHODIMP_(ULONG) AddRef (void)                              \
    {                                                                    \
      ++m_ref;                                                           \
      if (debug_oom_extra)                                               \
        log_debug ("%s:" #subcls ":%s: m_ref now %lu",                   \
                   SRCNAME,__func__, m_ref);                             \
      return m_ref;                                                      \
    }                                                                    \
  inline STDMETHODIMP_(ULONG) Release (void)                             \
    {                                                                    \
      ULONG count = --m_ref;                                             \
      if (debug_oom_extra)                                               \
        log_debug ("%s:" #subcls ":%s: mref now %lu",                    \
                   SRCNAME,__func__,count);                              \
      if (!count)                                                        \
        delete this;                                                     \
      return count;                                                      \
    }                                                                    \
  STDMETHODIMP GetTypeInfoCount (UINT*);                                 \
  STDMETHODIMP GetTypeInfo (UINT, LCID, LPTYPEINFO*);                    \
  STDMETHODIMP GetIDsOfNames (REFIID, LPOLESTR*, UINT, LCID, DISPID*);   \
  STDMETHODIMP Invoke (DISPID, REFIID, LCID, WORD,                       \
                       DISPPARAMS*, VARIANT*, EXCEPINFO*, UINT*);        \
/* End of macro BEGIN_EVENT_SINK.  */

#define EVENT_SINK_CTOR(subcls)                                          \
};                                                                       \
subcls::subcls ()                                                        \
{                                                                        \
  m_ref = 1;                                                             \
/* End of macro EVENT_SINK_CTOR.  */

#define EVENT_SINK_DTOR(subcls)                                          \
}                                                                        \
subcls::~subcls ()                                                       \
/* End of macro EVENT_SINK_DTOR.  */

#define EVENT_SINK_INVOKE(subcls)                                        \
STDMETHODIMP subcls::Invoke (DISPID dispid, REFIID riid, LCID lcid,      \
                WORD flags, DISPPARAMS *parms, VARIANT *result,          \
                EXCEPINFO *exepinfo, UINT *argerr)                       \
/* End of macro EVENT_SINK_INVOKE.  */

#define EVENT_SINK_DEFAULT_DTOR_CODE(subcls)                             \
{                                                                        \
  if (debug_oom)                                                         \
    log_debug ("%s:" #subcls ":%s: tdor", SRCNAME, __func__);            \
  if (m_pCP)                                                             \
    log_error ("%s:%s: Unadvise missing", SRCNAME, __func__);            \
  if (m_object)                                                          \
    log_error ("%s:%s: Object not released", SRCNAME,__func__);          \
}                                                                        \
/* End of macro EVENT_SINK_DTOR_DEFAULT_CODE.  */


#define EVENT_SINK_DEFAULT_CTOR(subcls)                                  \
  EVENT_SINK_CTOR(subcls)                                                \
/* End of macro EVENT_SINK_STD_DTOR.  */

#define EVENT_SINK_DEFAULT_DTOR(subcls)                                  \
  EVENT_SINK_DTOR(subcls)                                                \
  EVENT_SINK_DEFAULT_DTOR_CODE(subcls)                                   \
/* End of macro EVENT_SINK_STD_DTOR.  */


#define END_EVENT_SINK(subcls,iidcls)                                    \
STDMETHODIMP subcls::QueryInterface (REFIID riid, LPVOID FAR *ppvObj)    \
{                                                                        \
  /*log_debug ("%s:%s:%s: called riid=%08lx-%04hx-%04hx"                 \
             "-%02x-%02x-%02x-%02x-%02x-%02x-%02x-%02x",                 \
             SRCNAME, #subcls, __func__,                                 \
             riid.Data1, riid.Data2, riid.Data3,                         \
             riid.Data4[0], riid.Data4[1], riid.Data4[2],                \
             riid.Data4[3], riid.Data4[4], riid.Data4[5],                \
             riid.Data4[6], riid.Data4[7] );*/                           \
  if (riid == IID_IUnknown || riid == iidcls || riid == IID_IDispatch)   \
    {                                                                    \
      *ppvObj = (LPUNKNOWN) this;                                        \
      ((LPUNKNOWN)*ppvObj)->AddRef();                                    \
      return S_OK;                                                       \
    }                                                                    \
  *ppvObj = NULL;                                                        \
  return E_NOINTERFACE;                                                  \
}                                                                        \
STDMETHODIMP subcls::GetTypeInfoCount (UINT *r_count)                    \
{                                                                        \
  *r_count = 0;                                                          \
  return S_OK;                                                           \
}                                                                        \
STDMETHODIMP subcls::GetTypeInfo (UINT iTypeInfo, LCID lcid,             \
                                  LPTYPEINFO *r_typeinfo)                \
{                                                                        \
  (void)iTypeInfo;                                                       \
  (void)lcid;                                                            \
  (void)r_typeinfo;                                                      \
  return E_NOINTERFACE;                                                  \
}                                                                        \
STDMETHODIMP subcls::GetIDsOfNames (REFIID riid, LPOLESTR *rgszNames,    \
                                    UINT cNames, LCID lcid,              \
                                    DISPID *rgDispId)                    \
{                                                                        \
  (void)riid;                                                            \
  (void)rgszNames;                                                       \
  (void)cNames;                                                          \
  (void)lcid;                                                            \
  (void)rgDispId;                                                        \
  return E_NOINTERFACE;                                                  \
}                                                                        \
LPDISPATCH install_ ## subcls ## _sink (LPDISPATCH object)               \
{                                                                        \
  HRESULT hr;                                                            \
  LPDISPATCH disp;                                                       \
  LPCONNECTIONPOINTCONTAINER pCPC;                                       \
  LPCONNECTIONPOINT pCP;                                                 \
  subcls *sink;                                                          \
  DWORD cookie;                                                          \
                                                                         \
  disp = NULL;                                                           \
  hr = object->QueryInterface (IID_IConnectionPointContainer,            \
                               (LPVOID*)&disp);                          \
  if (hr != S_OK || !disp)                                               \
    {                                                                    \
      log_error ("%s:%s:%s: IConnectionPoint not supported: hr=%#lx",    \
                 SRCNAME, #subcls, __func__, hr);                        \
      return NULL;                                                       \
    }                                                                    \
  pCPC = (LPCONNECTIONPOINTCONTAINER)disp;                               \
  pCP = NULL;                                                            \
  hr = pCPC->FindConnectionPoint (iidcls, &pCP);                         \
  if (hr != S_OK || !pCP)                                                \
    {                                                                    \
      log_error ("%s:%s:%s: ConnectionPoint not found: hr=%#lx",         \
                 SRCNAME,#subcls,  __func__, hr);                        \
      pCPC->Release ();                                                  \
      return NULL;                                                       \
    }                                                                    \
  sink = new subcls; /* Note: Advise does another AddRef.  */            \
  hr = pCP->Advise ((LPUNKNOWN)sink, &cookie);                           \
  pCPC->Release ();                                                      \
  if (hr != S_OK)                                                        \
    {                                                                    \
      log_error ("%s:%s:%s: Advice failed: hr=%#lx",                     \
                 SRCNAME, #subcls, __func__, hr);                        \
      pCP->Release ();                                                   \
      sink->Release ();                                                  \
      return NULL;                                                       \
    }                                                                    \
  if (debug_oom)                                                         \
    log_debug ("%s:%s:%s: Advice succeeded", SRCNAME, #subcls, __func__);\
  sink->m_cookie = cookie;                                               \
  sink->m_pCP = pCP;                                                     \
  object->AddRef ();                                                     \
  sink->m_object = object;                                               \
  return (LPDISPATCH)sink;                                               \
}                                                                        \
void detach_ ## subcls ## _sink (LPDISPATCH obj)                         \
{                                                                        \
  HRESULT hr;                                                            \
  subcls *sink;                                                          \
                                                                         \
  if (debug_oom_extra)                                                   \
    log_debug ("%s:%s:%s: Called", SRCNAME, #subcls, __func__);          \
  hr = obj->QueryInterface (iidcls, (void**)&sink);                      \
  if (hr != S_OK || !sink)                                               \
    {                                                                    \
      log_error ("%s:%s:%s: invalid object passed: hr=%#lx",             \
                 SRCNAME, #subcls, __func__, hr);                        \
      return;                                                            \
    }                                                                    \
  if (sink->m_pCP)                                                       \
    {                                                                    \
      if (debug_oom_extra)                                               \
        log_debug ("%s:%s:%s: Unadvising", SRCNAME, #subcls, __func__);  \
      hr = sink->m_pCP->Unadvise (sink->m_cookie);                       \
      if (hr != S_OK)                                                    \
        log_error ("%s:%s:%s: Unadvice failed: hr=%#lx",                 \
                   SRCNAME, #subcls, __func__, hr);                      \
      if (debug_oom_extra)                                               \
        log_debug ("%s:%s:%s: Releasing connt point",                    \
                   SRCNAME, #subcls, __func__);                          \
      sink->m_pCP->Release ();                                           \
      sink->m_pCP = NULL;                                                \
    }                                                                    \
  if (sink->m_object)                                                    \
    {                                                                    \
      if (debug_oom_extra)                                               \
        log_debug ("%s:%s:%s: Releasing actual object",                  \
                   SRCNAME, #subcls, __func__);                          \
      sink->m_object->Release ();                                        \
      sink->m_object = NULL;                                             \
    }                                                                    \
  sink->Release ();                                                      \
}                                                                        \
/* End of macro END_EVENT_SINK.  */


#endif /*EVENTSINK_H*/
