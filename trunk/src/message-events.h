/* message-events.h - Definitions for out subclass of IExchExtMessageEvents
 *	Copyright (C) 2005, 2007 g10 Code GmbH
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

#ifndef MESSAGE_EVENTS_H
#define MESSAGE_EVENTS_H

/*
   GpgolMessageEvents 
 
   The GpgolMessageEvents class implements the reaction of the exchange 
   message events.
 */
class GpgolMessageEvents : public IExchExtMessageEvents
{
  /* Constructor. */
 public:
  GpgolMessageEvents (GpgolExt* pParentInterface);
  
  /* Attributes.  */
 private:
  ULONG   m_lRef;
  ULONG   m_lContext;
  BOOL    m_bOnSubmitActive;
  GpgolExt* m_pExchExt;
  bool    m_bWriteFailed;
  bool    m_want_html;       /* Encryption of HTML is desired. */
  bool    m_processed;       /* The message has been porcessed by us.  */
  bool    m_wasencrypted;    /* The original message was encrypted.  */
  bool    m_gotinspector;    /* We are working on a real inspector.  */
  
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
  
  STDMETHODIMP OnRead (LPEXCHEXTCALLBACK pEECB);
  STDMETHODIMP OnReadComplete (LPEXCHEXTCALLBACK pEECB, ULONG lFlags);
  STDMETHODIMP OnWrite (LPEXCHEXTCALLBACK pEECB);
  STDMETHODIMP OnWriteComplete (LPEXCHEXTCALLBACK pEECB, ULONG lFlags);
  STDMETHODIMP OnCheckNames (LPEXCHEXTCALLBACK pEECB);
  STDMETHODIMP OnCheckNamesComplete (LPEXCHEXTCALLBACK pEECB, ULONG lFlags);
  STDMETHODIMP OnSubmit (LPEXCHEXTCALLBACK pEECB);
  STDMETHODIMP_ (VOID)OnSubmitComplete (LPEXCHEXTCALLBACK pEECB, ULONG lFlags);

  inline void SetContext (ULONG lContext)
    { 
      m_lContext = lContext;
    };
};


#endif /*MESSAGE_EVENTS_H*/
