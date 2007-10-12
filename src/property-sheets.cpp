/* property-sheets.cpp - Subclass impl of IExchExtPropertySheets
 *	Copyright (C) 2004, 2005, 2007 g10 Code GmbH
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

#ifdef HAVE_CONFIG_H
#include <config.h>
#endif

#include <windows.h>

#include "mymapi.h"
#include "mymapitags.h"
#include "myexchext.h"
#include "common.h"
#include "display.h"
#include "msgcache.h"
#include "engine.h"
#include "mapihelp.h"

#include "olflange-def.h"
#include "olflange.h"
#include "dialogs.h"
#include "property-sheets.h"


#define TRACEPOINT() do { log_debug ("%s:%s:%d: tracepoint\n", \
                                     SRCNAME, __func__, __LINE__); \
                        } while (0)




GpgolPropertySheets::GpgolPropertySheets (GpgolExt* pParentInterface)
{ 
    m_pExchExt = pParentInterface;
    m_lRef = 0; 
}


STDMETHODIMP 
GpgolPropertySheets::QueryInterface(REFIID riid, LPVOID FAR * ppvObj)
{   
    *ppvObj = NULL;
    if (riid == IID_IExchExtPropertySheets) {
        *ppvObj = (LPVOID)this;
        AddRef();
        return S_OK;
    }
    if (riid == IID_IUnknown) {
        *ppvObj = (LPVOID)m_pExchExt;
        m_pExchExt->AddRef();
        return S_OK;
    }

    return E_NOINTERFACE;
}


/* Called by Echange to get the maximum number of property pages which
   are to be added. LFLAGS is a bitmask indicating what type of
   property sheet is being displayed. Return value: The maximum number
   of custom pages for the property sheet.  */
STDMETHODIMP_ (ULONG) 
GpgolPropertySheets::GetMaxPageCount(ULONG lFlags)
{
    if (lFlags == EEPS_TOOLSOPTIONS)
	return 1;	
    return 0;
}


/* Called by Exchange to request information about the property page.
   Return value: S_OK to signal Echange to use the pPSP
   information. */
STDMETHODIMP 
GpgolPropertySheets::GetPages(
	LPEXCHEXTCALLBACK pEECB, // A pointer to Exchange callback interface.
	ULONG lFlags,            // A  bitmask indicating what type of
                                 //  property sheet is being displayed.
	LPPROPSHEETPAGE pPSP,    // The output parm pointing to pointer
                                 //  to list of property sheets.
	ULONG FAR * plPSP)       // The output parm pointing to buffer 
                                 //  containing the number of property 
                                 //  sheets actually used.
{
  pPSP[0].dwSize = sizeof (PROPSHEETPAGE);
  pPSP[0].dwFlags = PSP_DEFAULT | PSP_HASHELP;
  pPSP[0].hInstance = glob_hinst;
  pPSP[0].pszTemplate = MAKEINTRESOURCE (IDD_GPG_OPTIONS);
  pPSP[0].hIcon = NULL;     
  pPSP[0].pszTitle = NULL;  
  pPSP[0].pfnDlgProc = (DLGPROC) GPGOptionsDlgProc;
  pPSP[0].lParam = 0;     
  pPSP[0].pfnCallback = NULL;
  pPSP[0].pcRefParent = NULL; 

  *plPSP = 1;

  return S_OK;
}


STDMETHODIMP_ (VOID) 
GpgolPropertySheets::FreePages (LPPROPSHEETPAGE pPSP,
                                      ULONG lFlags, ULONG lPSP)
{
}
