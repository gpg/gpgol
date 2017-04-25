/* user-events.h - Definitions for our subclass of IExchExtUserEvents
 * Copyright (C) 2007 g10 Code GmbH
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

#ifndef USER_EVENTS_H
#define USER_EVENTS_H

/*
   GpgolUserEvents 
 
   The GpgolUserEvents class implements the reaction on the certain
   user events.
 */
class GpgolUserEvents : public IExchExtUserEvents
{
  /* Constructor. */
public:
  GpgolUserEvents (GpgolExt *pParentInterface);
  virtual ~GpgolUserEvents () {}

  /* Attributes. */
private:
  ULONG   m_lRef;
  ULONG   m_lContext;
  GpgolExt *m_pExchExt;
  
public:
  STDMETHODIMP QueryInterface (REFIID riid, LPVOID *ppvObj);
  inline STDMETHODIMP_(ULONG) AddRef (void)
  {
    ++m_lRef; 
    return m_lRef; 
  };
  inline STDMETHODIMP_(ULONG) Release (void) 
  {
    ULONG lCount = --m_lRef;
    if (!lCount) 
      delete this;
    return lCount;	
  };

  STDMETHODIMP_ (VOID) OnSelectionChange (LPEXCHEXTCALLBACK eecb);
  STDMETHODIMP_ (VOID) OnObjectChange (LPEXCHEXTCALLBACK eecb);

  inline void SetContext (ULONG lContext)
  { 
    m_lContext = lContext;
  };
};


#endif /*USER_EVENTS_H*/
