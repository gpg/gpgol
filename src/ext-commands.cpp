/* ext-commands.cpp - Subclass impl of IExchExtCommands
 *	Copyright (C) 2004, 2005, 2007 g10 Code GmbH
 * 
 * This file is part of GpgOL.
 * 
 * GpgOL is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2 of the License, or (at your option) any later version.
 * 
 * GpgOL is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Lesser General Public License for more details.
 * 
 * You should have received a copy of the GNU Lesser General Public
 * License along with this program; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA
 * 02110-1301, USA.
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
#include "mapihelp.h"

#include "olflange-ids.h"
#include "olflange-def.h"
#include "olflange.h"
#include "ol-ext-callback.h"
#include "message.h"
#include "ext-commands.h"


#define TRACEPOINT() do { log_debug ("%s:%s:%d: tracepoint\n", \
                                     SRCNAME, __func__, __LINE__); \
                        } while (0)

static void toolbar_add_menu (LPEXCHEXTCALLBACK pEECB, 
                              UINT FAR *pnCommandIDBase, ...)
#if __GNUC__ >= 4 
                               __attribute__ ((sentinel))
#endif
  ;




/* Wrapper around UlRelease with error checking. */
/* FIXME: Duplicated code.  */
static void 
ul_release (LPVOID punk)
{
  ULONG res;
  
  if (!punk)
    return;
  res = UlRelease (punk);
//   log_debug ("%s UlRelease(%p) had %lu references\n", __func__, punk, res);
}





/* Constructor */
GpgolExtCommands::GpgolExtCommands (GpgolExt* pParentInterface)
{ 
  m_pExchExt = pParentInterface; 
  m_lRef = 0; 
  m_lContext = 0; 
  m_nCmdEncrypt = 0;  
  m_nCmdDecrypt = 0;  
  m_nCmdSign = 0; 
  m_nCmdShowInfo = 0;  
  m_nCmdCheckSig = 0;
  m_nCmdKeyManager = 0;
  m_nCmdDebug1 = 0;
  m_nCmdDebug2 = 0;
  m_nToolbarButtonID1 = 0; 
  m_nToolbarButtonID2 = 0; 
  m_nToolbarBitmap1 = 0;
  m_nToolbarBitmap2 = 0; 
  m_hWnd = NULL; 
}



STDMETHODIMP 
GpgolExtCommands::QueryInterface (REFIID riid, LPVOID FAR * ppvObj)
{
    *ppvObj = NULL;
    if ((riid == IID_IExchExtCommands) || (riid == IID_IUnknown)) {
        *ppvObj = (LPVOID)this;
        AddRef ();
        return S_OK;
    }
    return E_NOINTERFACE;
}


/* Return the toolbar's window from the button entry array.  On
   success the handle of the window is return as well as the
   corresponding index at R_IDX.  On error NULL is returned and
   the value at R_IDX is not changed. */
static HWND
toolbar_from_tbe (LPTBENTRY pTBEArray, UINT nTBECnt, int *r_idx)
{
  int nTBIndex;

  for (nTBIndex = nTBECnt-1; nTBIndex > -1; --nTBIndex)
    {	
      if (EETBID_STANDARD == pTBEArray[nTBIndex].tbid)
        {
          *r_idx = nTBIndex;
          return pTBEArray[nTBIndex].hwnd;
        }	
    }
  return NULL;
}

/* Add a new menu.  The variable entries are made up of pairs of
   strings and UINT *.  A NULL is used to terminate this list. An empty
   string is translated to a separator menu item. */
static void
toolbar_add_menu (LPEXCHEXTCALLBACK pEECB, 
                  UINT FAR *pnCommandIDBase, ...)
{
  va_list arg_ptr;
  HMENU menu;
  const char *string;
  UINT *cmdptr;
  
  va_start (arg_ptr, pnCommandIDBase);
  pEECB->GetMenuPos (EECMDID_ToolsCustomizeToolbar, &menu, NULL, NULL, 0);
  while ( (string = va_arg (arg_ptr, const char *)) )
    {
      cmdptr = va_arg (arg_ptr, UINT*);

      if (!*string)
        AppendMenu (menu, MF_SEPARATOR, 0, NULL);
      else
	{
          AppendMenu (menu, MF_STRING, *pnCommandIDBase, string);
          if (cmdptr)
            *cmdptr = *pnCommandIDBase;
          (*pnCommandIDBase)++;
        }
    }
  va_end (arg_ptr);
}



/* Called by Exchange to install commands and toolbar buttons.  Returns
   S_FALSE to signal Exchange to continue calling extensions. */
STDMETHODIMP 
GpgolExtCommands::InstallCommands (
	LPEXCHEXTCALLBACK pEECB, // The Exchange Callback Interface.
	HWND hWnd,               // The window handle to the main window
                                 // of context.
	HMENU hMenu,             // The menu handle to main menu of context.
	UINT FAR *pnCommandIDBase,  // The base command id.
	LPTBENTRY pTBEArray,     // The array of toolbar button entries.
	UINT nTBECnt,            // The count of button entries in array.
	ULONG lFlags)            // reserved
{
  HRESULT hr;
  m_hWnd = hWnd;
  LPDISPATCH pDisp;
  DISPID dispid;
  DISPID dispid_put = DISPID_PROPERTYPUT;
  DISPPARAMS dispparams;
  VARIANT aVariant;
  int force_encrypt = 0;
  int tb_idx;
  HWND hwnd_toolbar;
  TBADDBITMAP tbab;

  
  log_debug ("%s:%s: context=%s flags=0x%lx\n", SRCNAME, __func__, 
             ext_context_name (m_lContext), lFlags);


  /* Outlook 2003 sometimes displays the plaintext and sometimes the
     original undecrypted text when doing a reply.  This seems to
     depend on the size of the message; my guess it that only short
     messages are locally saved in the process and larger ones are
     fetched again from the backend - or the other way around.
     Anyway, we can't rely on that and thus me make sure to update the
     Body object right here with our own copy of the plaintext.  To
     match the text we use the ConversationIndex property.

     Unfortunately there seems to be no way of resetting the saved
     property after updating the body, thus even without entering a
     single byte the user will be asked when cancelling a reply
     whether he really wants to do that.

     Note, that we can't optimize the code here by first reading the
     body because this would pop up the securiy window, telling the
     user that someone is trying to read this data.
  */
  if (m_lContext == EECONTEXT_SENDNOTEMESSAGE)
    {
      LPMDB mdb = NULL;
      LPMESSAGE message = NULL;
      
      /*  Note that for read and send the object returned by the
          outlook extension callback is of class 43 (MailItem) so we
          only need to ask for Body then. */
      hr = pEECB->GetObject (&mdb, (LPMAPIPROP *)&message);
      if (FAILED(hr))
        log_debug ("%s:%s: getObject failed: hr=%#lx\n", SRCNAME,__func__,hr);
      else if (!opt.compat.no_msgcache)
        {
          const char *body;
          char *key = NULL;
          size_t keylen = 0;
          void *refhandle = NULL;
     
          pDisp = find_outlook_property (pEECB, "ConversationIndex", &dispid);
          if (pDisp)
            {
              DISPPARAMS dispparamsNoArgs = {NULL, NULL, 0, 0};

              aVariant.bstrVal = NULL;
              hr = pDisp->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                                  DISPATCH_PROPERTYGET, &dispparamsNoArgs,
                                  &aVariant, NULL, NULL);
              if (hr != S_OK)
                log_debug ("%s:%s: retrieving ConversationIndex failed: %#lx",
                           SRCNAME, __func__, hr);
              else if (aVariant.vt != VT_BSTR)
                log_debug ("%s:%s: ConversationIndex is not a string (%d)",
                           SRCNAME, __func__, aVariant.vt);
              else if (aVariant.bstrVal)
                {
                  char *p;

                  key = wchar_to_utf8 (aVariant.bstrVal);
                  log_debug ("%s:%s: ConversationIndex is `%s'",
                           SRCNAME, __func__, key);
                  /* The key is a hex string.  Convert it to binary. */
                  for (keylen=0,p=key; hexdigitp(p) && hexdigitp(p+1); p += 2)
                    ((unsigned char*)key)[keylen++] = xtoi_2 (p);
                  
		  SysFreeString (aVariant.bstrVal);
                }

              pDisp->Release();
              pDisp = NULL;
            }
          
          if (key && keylen
              && (body = msgcache_get (key, keylen, &refhandle)) 
              && (pDisp = find_outlook_property (pEECB, "Body", &dispid)))
            {
#if 1
              dispparams.cNamedArgs = 1;
              dispparams.rgdispidNamedArgs = &dispid_put;
              dispparams.cArgs = 1;
              dispparams.rgvarg = &aVariant;
              dispparams.rgvarg[0].vt = VT_LPWSTR;
              dispparams.rgvarg[0].bstrVal = utf8_to_wchar (body);
              hr = pDisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                                 DISPATCH_PROPERTYPUT, &dispparams,
                                 NULL, NULL, NULL);
              xfree (dispparams.rgvarg[0].bstrVal);
              log_debug ("%s:%s: PROPERTYPUT(body) result -> %#lx\n",
                         SRCNAME, __func__, hr);
#else
              log_window_hierarchy (hWnd, "%s:%s:%d: Windows hierarchy:",
                                    SRCNAME, __func__, __LINE__);
#endif
              pDisp->Release();
              pDisp = NULL;
              
              /* Because we found the plaintext in the cache we can assume
                 that the orginal message has been encrypted and thus we
                 now set a flag to make sure that by default the reply
                 gets encrypted too. */
              force_encrypt = 1;
            }
          msgcache_unref (refhandle);
          xfree (key);
        }
      
      ul_release (message);
      ul_release (mdb);
    }


  if (m_lContext == EECONTEXT_READNOTEMESSAGE)
    {

      switch (m_pExchExt->getMsgtype (pEECB))
        {
        case MSGTYPE_GPGOL_MULTIPART_ENCRYPTED:
        case MSGTYPE_GPGOL_OPAQUE_ENCRYPTED:
        case MSGTYPE_GPGOL_PGP_MESSAGE:
          toolbar_add_menu (pEECB, pnCommandIDBase, "", NULL,
                            _("&Decrypt and verify message"), &m_nCmdDecrypt,
                            NULL);
          break;
        default:
          break;
        }

      /* We always enable the verify button as it might be useful on
         an already decryopted message. */
      toolbar_add_menu (pEECB, pnCommandIDBase,
                        _("&Verify signature"), &m_nCmdCheckSig,
                        _("&Display crypto information"), &m_nCmdShowInfo,
                        NULL);

      toolbar_add_menu (pEECB, pnCommandIDBase, "", NULL,
                        _("Debug-1 (open_inspector)"), &m_nCmdDebug1,
                        _("Debug-2 (n/a)"), &m_nCmdDebug2,
                        NULL);
      
      hwnd_toolbar = toolbar_from_tbe (pTBEArray, nTBECnt, &tb_idx);
      if (hwnd_toolbar)
        {
          m_nToolbarButtonID1 = pTBEArray[tb_idx].itbbBase;
          pTBEArray[tb_idx].itbbBase++;

          tbab.hInst = glob_hinst;
          tbab.nID = IDB_DECRYPT;
          m_nToolbarBitmap1 = SendMessage(hwnd_toolbar, TB_ADDBITMAP,
                                          1, (LPARAM)&tbab);
          m_nToolbarButtonID2 = pTBEArray[tb_idx].itbbBase;
          pTBEArray[tb_idx].itbbBase++;
        }
    }

  if (m_lContext == EECONTEXT_SENDNOTEMESSAGE) 
    {
      toolbar_add_menu (pEECB, pnCommandIDBase, "", NULL,
                        _("&encrypt message with GnuPG"), &m_nCmdEncrypt,
                        _("&sign message with GnuPG"), &m_nCmdSign,
                        NULL );
      

      hwnd_toolbar = toolbar_from_tbe (pTBEArray, nTBECnt, &tb_idx);
      if (hwnd_toolbar) 
        {
          m_nToolbarButtonID1 = pTBEArray[tb_idx].itbbBase;
          pTBEArray[tb_idx].itbbBase++;

          tbab.hInst = glob_hinst;
          tbab.nID = IDB_ENCRYPT;
          m_nToolbarBitmap1 = SendMessage (hwnd_toolbar, TB_ADDBITMAP,
                                           1, (LPARAM)&tbab);

          m_nToolbarButtonID2 = pTBEArray[tb_idx].itbbBase;
          pTBEArray[tb_idx].itbbBase++;

          tbab.nID = IDB_SIGN;
          m_nToolbarBitmap2 = SendMessage (hwnd_toolbar, TB_ADDBITMAP,
                                           1, (LPARAM)&tbab);
        }

      m_pExchExt->m_gpgEncrypt = opt.encrypt_default;
      m_pExchExt->m_gpgSign    = opt.sign_default;
      if (force_encrypt)
        m_pExchExt->m_gpgEncrypt = true;
    }

  if (m_lContext == EECONTEXT_VIEWER) 
    {
      toolbar_add_menu (pEECB, pnCommandIDBase, "", NULL,
                        _("GnuPG Key &Manager"), &m_nCmdKeyManager,
                        NULL);

      hwnd_toolbar = toolbar_from_tbe (pTBEArray, nTBECnt, &tb_idx);
      if (hwnd_toolbar)
        {
          m_nToolbarButtonID1 = pTBEArray[tb_idx].itbbBase;
          pTBEArray[tb_idx].itbbBase++;

          tbab.hInst = glob_hinst;
          tbab.nID = IDB_KEY_MANAGER;
          m_nToolbarBitmap1 = SendMessage(hwnd_toolbar, TB_ADDBITMAP,
                                          1, (LPARAM)&tbab);
        }	
    }
  return S_FALSE;
}


/* Called by Exchange when a user selects a command.  Return value:
   S_OK if command is handled, otherwise S_FALSE. */
STDMETHODIMP 
GpgolExtCommands::DoCommand (
                  LPEXCHEXTCALLBACK pEECB, // The Exchange Callback Interface.
                  UINT nCommandID)         // The command id.
{
  HRESULT hr;
  HWND hWnd = NULL;
  LPMESSAGE message = NULL;
  LPMDB mdb = NULL;
      
  if (FAILED (pEECB->GetWindow (&hWnd)))
    hWnd = NULL;

  log_debug ("%s:%s: commandID=%u (%#x) hwnd=%p\n",
             SRCNAME, __func__, nCommandID, nCommandID, hWnd);

  if (nCommandID == SC_CLOSE && m_lContext == EECONTEXT_READNOTEMESSAGE)
    {
      /* This is the system close command. Replace it with our own to
         avoid the "save changes" query, apparently induced by OL
         internal syncronisation of our SetWindowText message with its
         own OOM (in this case Body). */
      LPDISPATCH pDisp;
      DISPID dispid;
      DISPPARAMS dispparams;
      VARIANT aVariant;
      
      pDisp = find_outlook_property (pEECB, "Close", &dispid);
      if (pDisp)
        {
          dispparams.rgvarg = &aVariant;
          dispparams.rgvarg[0].vt = VT_INT;
          dispparams.rgvarg[0].intVal = 1; /* olDiscard */
          dispparams.cArgs = 1;
          dispparams.cNamedArgs = 0;
          hr = pDisp->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                              DISPATCH_METHOD, &dispparams,
                              NULL, NULL, NULL);
          pDisp->Release();
          pDisp = NULL;
          if (hr == S_OK)
            {
              log_debug ("%s:%s: invoking Close succeeded", SRCNAME,__func__);
              return S_OK; /* We handled the close command. */
            }

          log_debug ("%s:%s: invoking Close failed: %#lx",
                     SRCNAME, __func__, hr);
        }

      /* Closing on our own failed - pass it on. */
      return S_FALSE; 
    }
  else if (nCommandID == 154)
    {
      log_debug ("%s:%s: command Reply called\n", SRCNAME, __func__);
      /* What we might want to do is to call Reply, then GetInspector
         and then Activate - this allows us to get full control over
         the quoted message and avoids the ugly msgcache. */
      return S_FALSE; /* Pass it on.  */
    }
  else if (nCommandID == 155)
    {
      log_debug ("%s:%s: command ReplyAll called\n", SRCNAME, __func__);
      return S_FALSE; /* Pass it on.  */
    }
  else if (nCommandID == 156)
    {
      log_debug ("%s:%s: command Forward called\n", SRCNAME, __func__);
      return S_FALSE; /* Pass it on.  */
    }
  else if (nCommandID == m_nCmdEncrypt 
           && m_lContext == EECONTEXT_READNOTEMESSAGE)
    {
      hr = pEECB->GetObject (&mdb, (LPMAPIPROP *)&message);
      if (SUCCEEDED (hr))
        {
//           GpgMsg *m = CreateGpgMsg (message);
//           m->setExchangeCallback ((void*)pEECB);
//           m->decrypt (hWnd, 0);
//           delete m;
	}
      ul_release (message);
      ul_release (mdb);
    }
  else if (nCommandID == m_nCmdShowInfo
           && m_lContext == EECONTEXT_READNOTEMESSAGE)
    {
      MessageBox (NULL, 
                  _("Here you should see crypto related info"),
                  "GpgOL", MB_ICONINFORMATION|MB_OK);
    }
  else if (nCommandID == m_nCmdCheckSig
           && m_lContext == EECONTEXT_READNOTEMESSAGE)
    {
      hr = pEECB->GetObject (&mdb, (LPMAPIPROP *)&message);
      if (SUCCEEDED (hr))
        {
          message_verify (message, m_pExchExt->getMsgtype (pEECB), 1);
	}
      ul_release (message);
      ul_release (mdb);
    }
  else if (nCommandID == m_nCmdEncrypt
           && m_lContext == EECONTEXT_SENDNOTEMESSAGE) 
    {
      m_pExchExt->m_gpgEncrypt = !m_pExchExt->m_gpgEncrypt;
    }
  else if (nCommandID == m_nCmdSign
           && m_lContext == EECONTEXT_SENDNOTEMESSAGE) 
    {
      m_pExchExt->m_gpgSign = !m_pExchExt->m_gpgSign;
    }
  else if (nCommandID == m_nCmdKeyManager
           && m_lContext == EECONTEXT_VIEWER)
    {
      if (start_key_manager ())
        MessageBox (NULL, _("Could not start Key-Manager"),
                    "GpgOL", MB_ICONERROR|MB_OK);
    }
  else if (nCommandID == m_nCmdDebug1
           && m_lContext == EECONTEXT_READNOTEMESSAGE)
    {
      hr = pEECB->GetObject (&mdb, (LPMAPIPROP *)&message);
      if (SUCCEEDED (hr))
        {
          open_inspector (pEECB, message);
	}
      ul_release (message);
      ul_release (mdb);

    }
  else
    return S_FALSE; /* Pass on unknown command. */

  return S_OK; 
}


/* Called by Exchange when it receives a WM_INITMENU message, allowing
   the extension object to enable, disable, or update its menu
   commands before the user sees them. This method is called
   frequently and should be written in a very efficient manner. */
STDMETHODIMP_(VOID) 
GpgolExtCommands::InitMenu(LPEXCHEXTCALLBACK pEECB) 
{
#if 0
  log_debug ("%s:%s: context=%s\n",
             SRCNAME, __func__, ext_context_name (m_lContext));
#endif
}


/* Called by Exchange when the user requests help for a menu item.
   PEECP is the pointer to Exchange Callback Interface.  NCOMMANDID is
   the command id.  Return value: S_OK when it is a menu item of this
   plugin and the help was shown; otherwise S_FALSE.  */
STDMETHODIMP 
GpgolExtCommands::Help (LPEXCHEXTCALLBACK pEECB, UINT nCommandID)
{
  if (nCommandID == m_nCmdDecrypt && 
      m_lContext == EECONTEXT_READNOTEMESSAGE) 
    {
      MessageBox (m_hWnd,
                  _("Select this option to decrypt and verify the message."),
                  "GpgOL", MB_OK);
    }
  else if (nCommandID == m_nCmdShowInfo
           && m_lContext == EECONTEXT_READNOTEMESSAGE) 
    {
      MessageBox (m_hWnd,
                  _("Select this option to show information"
                    " on the crypto status"),
                  "GpgOL", MB_OK);
    }
  else if (nCommandID == m_nCmdCheckSig
           && m_lContext == EECONTEXT_READNOTEMESSAGE) 
    {
      MessageBox (m_hWnd,
                  _("Check the signature now and display the result"),
                  "GpgOL", MB_OK);
    }
  else if (nCommandID == m_nCmdEncrypt 
           && m_lContext == EECONTEXT_SENDNOTEMESSAGE) 
    {
      MessageBox (m_hWnd,
                  _("Select this option to encrypt the message."),
                  "GpgOL", MB_OK);	
    } 
  else if (nCommandID == m_nCmdSign
           && m_lContext == EECONTEXT_SENDNOTEMESSAGE) 
    {
      MessageBox (m_hWnd,
                  _("Select this option to sign the message."),
                  "GpgOL", MB_OK);	
    }
  else if (nCommandID == m_nCmdKeyManager
           && m_lContext == EECONTEXT_VIEWER) 
    {
      MessageBox (m_hWnd, _("Select this option to open GpgOL Key Manager"),
                  "GpgOL", MB_OK);
    }
  else
    return S_FALSE;

  return S_OK;
}


/* Called by Exchange to get the status bar text or the tooltip of a
   menu item.  NCOMMANDID is the command id corresponding to the menu
   item activated.  LFLAGS identifies either EECQHT_STATUS or
   EECQHT_TOOLTIP.  PSZTEXT is a pointer to buffer to received the
   text to be displayed.  NCHARCNT is the available space in PSZTEXT.

   Returns S_OK when it is a menu item of this plugin and the text was
   set; otherwise S_FALSE.  */
STDMETHODIMP 
GpgolExtCommands::QueryHelpText(UINT nCommandID, ULONG lFlags,
                                LPTSTR pszText,  UINT nCharCnt)    
{
	
  if (nCommandID == m_nCmdDecrypt
      && m_lContext == EECONTEXT_READNOTEMESSAGE) 
    {
      if (lFlags == EECQHT_STATUS)
        lstrcpyn (pszText, ".", nCharCnt);
      if (lFlags == EECQHT_TOOLTIP)
        lstrcpyn (pszText, _("Decrypt message and verify signature"), 
                  nCharCnt);
    }
  else if (nCommandID == m_nCmdShowInfo
           && m_lContext == EECONTEXT_READNOTEMESSAGE) 
    {
      if (lFlags == EECQHT_STATUS)
        lstrcpyn (pszText, ".", nCharCnt);
      if (lFlags == EECQHT_TOOLTIP)
        lstrcpyn (pszText, _("Show S/MIME status info"),
                  nCharCnt);
    }
  else if (nCommandID == m_nCmdCheckSig
           && m_lContext == EECONTEXT_READNOTEMESSAGE) 
    {
      if (lFlags == EECQHT_STATUS)
        lstrcpyn (pszText, ".", nCharCnt);
      if (lFlags == EECQHT_TOOLTIP)
        lstrcpyn (pszText,
                  _("Check the signature now and display the result"),
                  nCharCnt);
    }
  else if (nCommandID == m_nCmdEncrypt
           && m_lContext == EECONTEXT_SENDNOTEMESSAGE) 
    {
      if (lFlags == EECQHT_STATUS)
        lstrcpyn (pszText, ".", nCharCnt);
      if (lFlags == EECQHT_TOOLTIP)
        lstrcpyn (pszText,
                  _("Encrypt message with GPG"),
                  nCharCnt);
    }
  else if (nCommandID == m_nCmdSign
           && m_lContext == EECONTEXT_SENDNOTEMESSAGE) 
    {
      if (lFlags == EECQHT_STATUS)
        lstrcpyn (pszText, ".", nCharCnt);
      if (lFlags == EECQHT_TOOLTIP)
        lstrcpyn (pszText,
                  _("Sign message with GPG"),
                  nCharCnt);
    }
  else if (nCommandID == m_nCmdKeyManager
           && m_lContext == EECONTEXT_VIEWER) 
    {
      if (lFlags == EECQHT_STATUS)
        lstrcpyn (pszText, ".", nCharCnt);
      if (lFlags == EECQHT_TOOLTIP)
        lstrcpyn (pszText,
                  _("Open the GpgOL Key Manager"),
                  nCharCnt);
    }
  else 
    return S_FALSE;

  return S_OK;
}


/* Called by Exchange to get toolbar button infos.  TOOLBARID is the
   toolbar identifier.  BUTTONID is the toolbar button index.  PTBB is
   a pointer to toolbar button structure DESCRIPTION is a pointer to
   buffer receiving the text for the button.  DESCRIPTION_SIZE is the
   maximum size of DESCRIPTION.  FLAGS are flags which might have the
   EXCHEXT_UNICODE bit set.

   Returns S_OK when it is a button of this plugin and the requested
   info was delivered; otherwise S_FALSE.  */
STDMETHODIMP 
GpgolExtCommands::QueryButtonInfo (ULONG toolbarid, UINT buttonid, 
                                   LPTBBUTTON pTBB, 
                                   LPTSTR description, UINT description_size,
                                   ULONG flags)          
{
  if (buttonid == m_nToolbarButtonID1
      && m_lContext == EECONTEXT_READNOTEMESSAGE)
    {
      pTBB->iBitmap = m_nToolbarBitmap1;             
      pTBB->idCommand = m_nCmdEncrypt;
      pTBB->fsState = TBSTATE_ENABLED;
      pTBB->fsStyle = TBSTYLE_BUTTON;
      pTBB->dwData = 0;
      pTBB->iString = -1;
      lstrcpyn (description,
                _("Decrypt message and verify signature"),
                description_size);
    }
  else if (buttonid == m_nToolbarButtonID1
           && m_lContext == EECONTEXT_SENDNOTEMESSAGE)
    {
      pTBB->iBitmap = m_nToolbarBitmap1;             
      pTBB->idCommand = m_nCmdEncrypt;
      pTBB->fsState = TBSTATE_ENABLED;
      if (m_pExchExt->m_gpgEncrypt)
        pTBB->fsState |= TBSTATE_CHECKED;
      pTBB->fsStyle = TBSTYLE_BUTTON | TBSTYLE_CHECK;
      pTBB->dwData = 0;
      pTBB->iString = -1;
      lstrcpyn (description, _("Encrypt message with GPG"),
                description_size);
    }
  else if (buttonid == m_nToolbarButtonID2
           && m_lContext == EECONTEXT_SENDNOTEMESSAGE)
    {
      pTBB->iBitmap = m_nToolbarBitmap2;             
      pTBB->idCommand = m_nCmdSign;
      pTBB->fsState = TBSTATE_ENABLED;
      if (m_pExchExt->m_gpgSign)
        pTBB->fsState |= TBSTATE_CHECKED;
      pTBB->fsStyle = TBSTYLE_BUTTON | TBSTYLE_CHECK;
      pTBB->dwData = 0;
      pTBB->iString = -1;
      lstrcpyn (description, _("Sign message with GPG"),
                description_size);
    }
  else if (buttonid == m_nToolbarButtonID1
           && m_lContext == EECONTEXT_VIEWER)
    {
      pTBB->iBitmap = m_nToolbarBitmap1;             
      pTBB->idCommand = m_nCmdEncrypt;
      pTBB->fsState = TBSTATE_ENABLED;
      pTBB->fsStyle = TBSTYLE_BUTTON;
      pTBB->dwData = 0;
      pTBB->iString = -1;
      lstrcpyn (description, _("Open GPG Key Manager"),
                description_size);
    }
  else
    return S_FALSE;

  return S_OK;
}



STDMETHODIMP 
GpgolExtCommands::ResetToolbar (ULONG lToolbarID, ULONG lFlags)
{	
  return S_OK;
}


