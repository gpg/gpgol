/* olflange.h - Flange between Outlook and the MapiGPGME class
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

#ifndef OLFLANGE_H
#define OLFLANGE_H

#include "mymapi.h"
#include "mymapitags.h"
#include "myexchext.h"
#include "mapihelp.h"

#include "olflange-def.h"

/* The GUID for this plugin.  */
#define CLSIDSTR_GPGOL   "{42d30988-1a3a-11da-c687-000d6080e735}"
DEFINE_GUID(CLSID_GPGOL, 0x42d30988, 0x1a3a, 0x11da,
            0xc6, 0x87, 0x00, 0x0d, 0x60, 0x80, 0xe7, 0x35);

/* For documentation: The GUID used for our custom properties:
   {31805ab8-3e92-11dc-879c-00061b031004}
 */

/* The ProgID used by us */
#define GPGOL_PROGID "Z.GNU.GpgOL"
/* User friendly add in name */
#define GPGOL_PRETTY "GpgOL - The GnuPG Outlook Plugin"
/* Short description of the addin */
#define GPGOL_DESCRIPTION "Cryptography for Outlook"



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

EXTERN_C const char * __stdcall gpgol_check_version (const char *req_version);

EXTERN_C int get_ol_main_version (void);

void install_sinks (LPEXCHEXTCALLBACK eecb);

void install_forms (void);

LPDISPATCH get_eecb_object (LPEXCHEXTCALLBACK eecb);

#endif /*OLFLANGE_H*/
