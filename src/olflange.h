/* olflange.h - Flange between Outlook and the MapiGPGME class
 *	Copyright (C) 2001 G Data Software AG, http://www.gdata.de
 *	Copyright (C) 2005 g10 Code GmbH
 * 
 * This file is part of GPGol.
 * 
 * GPGol is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2 of the License, or (at your option) any later version.
 * 
 * GPGol is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Lesser General Public License for more details.
 * 
 * You should have received a copy of the GNU Lesser General Public
 * License along with this program; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA
 * 02110-1301, USA.
 */

#ifndef OLFLANGE_H
#define OLFLANGE_H


/*
 CGPGExchExt 

 The CGPGExchExt class is the main exchange extension class. The other 
 extensions will be created in the constructor of this class.
*/
class CGPGExchExt : public IExchExt
{
public:
  CGPGExchExt();
  virtual ~CGPGExchExt();

public:	
  HWND m_hWndExchange;  /* Handle of the exchange window. */
  /* parameter for sending mails */
  BOOL  m_gpgEncrypt;
  BOOL  m_gpgSign;
  
private:
  ULONG m_lRef;
  ULONG m_lContext;
  
  /* pointer to the other extension objects */
  CGPGExchExtMessageEvents* m_pExchExtMessageEvents;
  CGPGExchExtCommands* m_pExchExtCommands;
  CGPGExchExtPropertySheets* m_pExchExtPropertySheets;
  CGPGExchExtAttachedFileEvents *m_pExchExtAttachedFileEvents;

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

/*
   CGPGExchExtMessageEvents 
 
   The CGPGExchExtMessageEvents class implements the reaction of the exchange 
   message events.
 */
class CGPGExchExtMessageEvents : public IExchExtMessageEvents
{
  /* constructor */
public:
  CGPGExchExtMessageEvents (CGPGExchExt* pParentInterface);

  /* attributes */
private:
  ULONG   m_lRef;
  ULONG   m_lContext;
  BOOL    m_bOnSubmitActive;
  CGPGExchExt* m_pExchExt;
  BOOL    m_bWriteFailed;
  
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

/*
   CGPGExchExtCommands 

   Makes the menu and toolbar extensions. Implements the own commands.
 */
class CGPGExchExtCommands : public IExchExtCommands
{
public:
  CGPGExchExtCommands (CGPGExchExt* pParentInterface);
  
private:
  ULONG m_lRef;
  ULONG m_lContext;
  
  UINT  m_nCmdEncrypt;
  UINT  m_nCmdSign;

  UINT  m_nToolbarButtonID1;
  UINT  m_nToolbarButtonID2;     
  UINT  m_nToolbarBitmap1;
  UINT  m_nToolbarBitmap2;
  
  HWND  m_hWnd;
  
  CGPGExchExt* m_pExchExt;
  
public:
  STDMETHODIMP QueryInterface(REFIID riid, LPVOID *ppvObj);
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

  STDMETHODIMP InstallCommands (LPEXCHEXTCALLBACK pEECB, HWND hWnd,
                                HMENU hMenu, UINT FAR * pnCommandIDBase,
                                LPTBENTRY pTBEArray,
                                UINT nTBECnt, ULONG lFlags);
  STDMETHODIMP DoCommand (LPEXCHEXTCALLBACK pEECB, UINT nCommandID);
  STDMETHODIMP_(void) InitMenu (LPEXCHEXTCALLBACK pEECB);
  STDMETHODIMP Help (LPEXCHEXTCALLBACK pEECB, UINT nCommandID);
  STDMETHODIMP QueryHelpText (UINT nCommandID, ULONG lFlags,
                              LPTSTR szText, UINT nCharCnt);
  STDMETHODIMP QueryButtonInfo (ULONG lToolbarID, UINT nToolbarButtonID, 
                                LPTBBUTTON pTBB, LPTSTR lpszDescription,
                                UINT nCharCnt, ULONG lFlags);
  STDMETHODIMP ResetToolbar (ULONG nToolbarID, ULONG lFlags);

  inline void SetContext (ULONG lContext)
  { 
    m_lContext = lContext;
  };
};


/*
   CGPGExchExtPropertySheets 
 
   The CGPGExchExtPropertySheets implements the exchange property
   sheet extension to put the GPG options page in the exchanges
   options property sheet.
 */
class CGPGExchExtPropertySheets : public IExchExtPropertySheets
{
  // constructor
public:
  CGPGExchExtPropertySheets(CGPGExchExt* pParentInterface);

  // attibutes
private:
  ULONG m_lRef;
  CGPGExchExt* m_pExchExt;
  
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

#endif /*OLFLANGE_H*/
