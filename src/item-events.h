/* item-events.h - GpgolItemEvents definitions.
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


#ifndef ITEM_EVENTS_H
#define ITEM_EVENTS_H

/*
   GpgolItemEvents

   The GpgolItemEvents class implements the processing of events
   related to certain items.  This is currently open and close of an
   inspector.
 */
class GpgolItemEvents : public IOutlookExtItemEvents
{
 public:
  GpgolItemEvents (GpgolExt *pParentInterface);
  virtual ~GpgolItemEvents () {}

 private:
  GpgolExt *m_pExchExt;
  ULONG     m_ref;
  bool    m_processed;       /* The message has been porcessed by us.  */
  bool    m_wasencrypted;    /* The original message was encrypted.  */
  
 public:
  STDMETHODIMP QueryInterface (REFIID riid, LPVOID FAR *ppvObj);
  inline STDMETHODIMP_(ULONG) AddRef (void)
    {
      ++m_ref;
      return m_ref;
    }
  inline STDMETHODIMP_(ULONG) Release (void)
    {
      ULONG count = --m_ref;
      if (!count)
	delete this;
      return count;
    }
  
  
  STDMETHODIMP OnOpen (LPEXCHEXTCALLBACK peecb);
  STDMETHODIMP OnOpenComplete (LPEXCHEXTCALLBACK peecb, ULONG flags);
  STDMETHODIMP OnClose (LPEXCHEXTCALLBACK peecb, ULONG save_options);
  STDMETHODIMP OnCloseComplete (LPEXCHEXTCALLBACK peecb, ULONG flags);
};

#endif /*ITEM_EVENTS_H*/
