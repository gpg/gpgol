/* olflange.h - Flange between Outlook and the MapiGPGME class
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

#ifndef OLFLANGE_H
#define OLFLANGE_H

#include "mapihelp.h"


/*
 GpgolExt 

 The GpgolExt class is the main exchange extension class. The other 
 extensions will be created in the constructor of this class.
*/
class GpgolExt : public IExchExt
{
public:
  GpgolExt();
  virtual ~GpgolExt();

public:	
  HWND m_hWndExchange;  /* Handle of the exchange window. */

  /* Parameters for sending mails.  */
  protocol_t  m_protoSelection;
  BOOL  m_gpgEncrypt;
  BOOL  m_gpgSign;
  
private:
  ULONG m_lRef;
  ULONG m_lContext;
  
  /* Pointer to the other extension objects.  */
  GpgolExtCommands        *m_pExchExtCommands;
  GpgolUserEvents         *m_pExchExtUserEvents;
  GpgolSessionEvents      *m_pExchExtSessionEvents;
  GpgolMessageEvents      *m_pExchExtMessageEvents;
  GpgolPropertySheets     *m_pExchExtPropertySheets;
  GpgolAttachedFileEvents *m_pExchExtAttachedFileEvents;
  GpgolItemEvents         *m_pOutlookExtItemEvents;

public:
  STDMETHODIMP QueryInterface(REFIID riid, LPVOID* ppvObj);
  inline STDMETHODIMP_(ULONG) AddRef() { ++m_lRef;  return m_lRef; };
  inline STDMETHODIMP_(ULONG) Release()
    {
      ULONG lCount = --m_lRef;
      if (!lCount) 
	delete this;
      return lCount;	
    };
  STDMETHODIMP Install(LPEXCHEXTCALLBACK pEECB, ULONG lContext, ULONG lFlags);
};


const char *ext_context_name (unsigned long no);

#endif /*OLFLANGE_H*/
