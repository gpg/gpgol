/* property-sheets.h - Definitions for our subclass of IExchExtPropertySheets
 * Copyright (C) 2005, 2007 g10 Code GmbH
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

#ifndef PROPERTY_SHEETS_H
#define PROPERTY_SHEETS_H


/*
   GpgolPropertySheets 
 
   The GpgolPropertySheets implements the exchange property
   sheet extension to put the GPG options page in the exchanges
   options property sheet.
 */
class GpgolPropertySheets : public IExchExtPropertySheets
{
  // constructor
public:
  GpgolPropertySheets(GpgolExt* pParentInterface);
  virtual ~GpgolPropertySheets () {}

  // attibutes
private:
  ULONG m_lRef;
  GpgolExt* m_pExchExt;
  
public:    	
  STDMETHODIMP QueryInterface(REFIID riid, LPVOID *ppvObj);
  inline STDMETHODIMP_(ULONG) AddRef()
  { 
    ++m_lRef;
    return m_lRef; 
  };
  inline STDMETHODIMP_(ULONG) Release() 
  {
    ULONG lCount = --m_lRef;
    if (!lCount)
      delete this;
    return lCount;	
  };

  STDMETHODIMP_ (ULONG) GetMaxPageCount(ULONG lFlags);          
  STDMETHODIMP  GetPages(LPEXCHEXTCALLBACK pEECB, ULONG lFlags,
                         LPPROPSHEETPAGE pPSP, ULONG FAR * pcpsp);
  STDMETHODIMP_ (void) FreePages(LPPROPSHEETPAGE pPSP, 
                                 ULONG lFlags, ULONG cpsp);          

};



#endif /*PROPERTY_SHEETS_H*/
