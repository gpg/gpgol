/* ext-commands.h - Definitions for our subclass of IExchExtCommands
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

#ifndef EXT_COMMANDS_H
#define EXT_COMMANDS_H


struct toolbar_info_s;
typedef struct toolbar_info_s *toolbar_info_t;


/*
   GpgolExtCommands 

   Makes the menu and toolbar extensions. Implements the own commands.
 */
class GpgolExtCommands : public IExchExtCommands
{
public:
  GpgolExtCommands (GpgolExt* pParentInterface);
  virtual ~GpgolExtCommands (void);

private:
  ULONG m_lRef;
  ULONG m_lContext;
  
  UINT  m_nCmdProtoAuto;
  UINT  m_nCmdProtoPgpmime;
  UINT  m_nCmdProtoSmime;
  UINT  m_nCmdEncrypt;
  UINT  m_nCmdSign;
  UINT  m_nCmdKeyManager;
  UINT  m_nCmdCryptoState;
  UINT  m_nCmdDebug0;
  UINT  m_nCmdDebug1;
  UINT  m_nCmdDebug2;

  /* A list of all active toolbar items.  */
  toolbar_info_t m_toolbar_info;
  
  HWND  m_hWnd;
  
  GpgolExt* m_pExchExt;

  void add_toolbar (LPTBENTRY tbearr, UINT n_tbearr, ...);
  void update_protocol_menu (LPEXCHEXTCALLBACK eecb);
  
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



#endif /*EXT_COMMANDS_H*/
