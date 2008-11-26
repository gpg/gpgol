/* ext-commands.cpp - Subclass impl of IExchExtCommands
 *	Copyright (C) 2004, 2005, 2007, 2008 g10 Code GmbH
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

#define _WIN32_IE 0x400 /* Need TBIF_COMMAND et al.  */
#include <windows.h>

#include "mymapi.h"
#include "mymapitags.h"
#include "myexchext.h"
#include "common.h"
#include "display.h"
#include "msgcache.h"
#include "mapihelp.h"

#include "dialogs.h"       /* For IDB_foo. */
#include "olflange-def.h"
#include "olflange.h"
#include "ol-ext-callback.h"
#include "message.h"
#include "engine.h"
#include "ext-commands.h"
#include "revert.h"

#define TRACEPOINT() do { log_debug ("%s:%s:%d: tracepoint\n", \
                                     SRCNAME, __func__, __LINE__); \
                        } while (0)

/* An object to store information about active (installed) toolbar
   buttons.  */
struct toolbar_info_s
{
  toolbar_info_t next;

  UINT button_id;/* The ID of the button as assigned by Outlook.  */
  UINT bitmap;   /* The bitmap of the button.  */
  UINT cmd_id;   /* The ID of the command to send on a click.  */
  const char *desc;/* The description text.  */
  ULONG context; /* Context under which this entry will be used.  */ 
  int did_qbi;   /* Has been processed by QueryButtonInfo.  */
};


/* Keep copies of some bitmaps.  */
static int bitmaps_initialized;
static HBITMAP my_check_bitmap, my_uncheck_bitmap;



static void add_menu (LPEXCHEXTCALLBACK eecb, 
                      UINT FAR *pnCommandIDBase, ...)
#if __GNUC__ >= 4 
                               __attribute__ ((sentinel))
#endif
  ;




/* Wrapper around UlRelease with error checking. */
static void 
ul_release (LPVOID punk, const char *func, int lnr)
{
  ULONG res;
  
  if (!punk)
    return;
  res = UlRelease (punk);
  if (opt.enable_debug & DBG_MEMORY)
    log_debug ("%s:%s:%d: UlRelease(%p) had %lu references\n", 
               SRCNAME, func, lnr, punk, res);
}



/* Constructor */
GpgolExtCommands::GpgolExtCommands (GpgolExt* pParentInterface)
{ 
  m_pExchExt = pParentInterface; 
  m_lRef = 0; 
  m_lContext = 0; 
  m_nCmdProtoAuto = 0;
  m_nCmdProtoPgpmime = 0;
  m_nCmdProtoSmime = 0;
  m_nCmdEncrypt = 0;  
  m_nCmdSign = 0; 
  m_nCmdKeyManager = 0;
  m_nCmdRevertFolder = 0;
  m_nCmdCryptoState = 0;
  m_nCmdDebug0 = 0;
  m_nCmdDebug1 = 0;
  m_nCmdDebug2 = 0;
  m_nCmdDebug3 = 0;
  m_toolbar_info = NULL; 
  m_hWnd = NULL; 

  if (!bitmaps_initialized)
    {
      my_uncheck_bitmap = get_system_check_bitmap (0);
      my_check_bitmap = get_system_check_bitmap (1);
      bitmaps_initialized = 1;
    }
}

/* Destructor */
GpgolExtCommands::~GpgolExtCommands (void)
{
  while (m_toolbar_info)
    {
      toolbar_info_t tmp = m_toolbar_info->next;
      xfree (m_toolbar_info);
      m_toolbar_info = tmp;
    }
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


/* Add a new menu.  The variable entries are made up of pairs of
   strings and UINT *.  A NULL is used to terminate this list.  An
   empty string is translated to a separator menu item.  One level of
   submenus are supported. */
static void
add_menu (LPEXCHEXTCALLBACK eecb, UINT FAR *pnCommandIDBase, ...)
{
  va_list arg_ptr;
  HMENU mainmenu, submenu, menu;
  const char *string;
  UINT *cmdptr;
  
  va_start (arg_ptr, pnCommandIDBase);
  /* We put all new entries into the tools menu.  To make this work we
     need to pass the id of an existing item from that menu.  */
  eecb->GetMenuPos (EECMDID_ToolsCustomizeToolbar, &mainmenu, NULL, NULL, 0);
  menu = mainmenu;
  submenu = NULL;
  while ( (string = va_arg (arg_ptr, const char *)) )
    {
      cmdptr = va_arg (arg_ptr, UINT*);

      if (!*string)
        ; /* Ignore this entry.  */
      else if (*string == '@' && !string[1])
        AppendMenu (menu, MF_SEPARATOR, 0, NULL);
      else if (*string == '>')
        {
          submenu = CreatePopupMenu ();
          AppendMenu (menu, MF_STRING|MF_POPUP, (UINT_PTR)submenu, string+1);
          menu = submenu;
        }
      else if (*string == '<')
        {
          menu = mainmenu;
          submenu = NULL;
        }
      else
	{
          AppendMenu (menu, MF_STRING, *pnCommandIDBase, string);
          if (menu == submenu)
            SetMenuItemBitmaps (menu, *pnCommandIDBase, MF_BYCOMMAND,
                                my_uncheck_bitmap, my_check_bitmap);
          if (cmdptr)
            *cmdptr = *pnCommandIDBase;
          (*pnCommandIDBase)++;
        }
    }
  va_end (arg_ptr);
}


static void
check_menu (LPEXCHEXTCALLBACK eecb, UINT menu_id, int checked)
{
  HMENU menu;

  eecb->GetMenuPos (EECMDID_ToolsCustomizeToolbar, &menu, NULL, NULL, 0);
  if (debug_commands)
    log_debug ("check_menu: eecb=%p menu_id=%u checked=%d -> menu=%p\n", 
               eecb, menu_id, checked, menu);
  CheckMenuItem (menu, menu_id, 
                 MF_BYCOMMAND | (checked?MF_CHECKED:MF_UNCHECKED));
}


static void
check_toolbar (LPEXCHEXTCALLBACK eecb, struct toolbar_info_s *toolbar_info,
               UINT cmd_id, int checked)
{
  HWND hwnd;
  toolbar_info_t tb_info;
  TBBUTTONINFOA tbb;

  eecb->GetToolbar (EETBID_STANDARD, &hwnd);
  if (debug_commands)
    log_debug ("check_toolbar: eecb=%p cmd_id=%u checked=%d -> hwnd=%p\n", 
               eecb, cmd_id, checked, hwnd);

  for (tb_info = toolbar_info; tb_info; tb_info = tb_info->next )
    if (tb_info->cmd_id == cmd_id)
      break;
  if (!tb_info)
    {
      log_error ("check_toolbar: no such toolbar button");
      return;
    }
  if (!tb_info->did_qbi)
    {
      if(debug_commands)
        log_debug ("check_toolbar: button(cmd_id=%d) not yet initialized",
                   cmd_id);
      return;
    }

  tbb.cbSize = sizeof (tbb);
  tbb.dwMask = TBIF_COMMAND | TBIF_STATE | TBIF_STYLE;
  if (!SendMessage (hwnd, TB_GETBUTTONINFO, cmd_id, (LPARAM)&tbb))
    log_error_w32 (-1, "TB_GETBUTTONINFO failed");
  else
    {
      tbb.cbSize = sizeof (tbb);
      tbb.dwMask = TBIF_STATE;
      if (checked)
        tbb.fsState |= TBSTATE_CHECKED;
      else
        tbb.fsState &= ~TBSTATE_CHECKED;
      if (!SendMessage (hwnd, TB_SETBUTTONINFO, cmd_id, (LPARAM)&tbb))
        log_error_w32 (-1, "TB_SETBUTTONINFO failed");
    }
}


static void
check_menu_toolbar (LPEXCHEXTCALLBACK eecb,
                    struct toolbar_info_s *toolbar_info,
                    UINT cmd_id, int checked)
{
  check_menu (eecb, cmd_id, checked);
  check_toolbar (eecb, toolbar_info, cmd_id, checked);
}


void
GpgolExtCommands::update_protocol_menu (LPEXCHEXTCALLBACK eecb)
{
  if (debug_commands)
    log_debug ("update_protocol_menu called\n");
  switch (m_pExchExt->m_protoSelection)
    {
    case PROTOCOL_OPENPGP:
      check_menu_toolbar (eecb, m_toolbar_info, m_nCmdProtoAuto, FALSE);
      check_menu_toolbar (eecb, m_toolbar_info, m_nCmdProtoPgpmime, TRUE);
      check_menu_toolbar (eecb, m_toolbar_info, m_nCmdProtoSmime, FALSE);
      break;
    case PROTOCOL_SMIME:
      check_menu_toolbar (eecb, m_toolbar_info, m_nCmdProtoAuto, FALSE);
      check_menu_toolbar (eecb, m_toolbar_info, m_nCmdProtoPgpmime, FALSE);
      check_menu_toolbar (eecb, m_toolbar_info, m_nCmdProtoSmime, TRUE);
      break;
    default:
      check_menu_toolbar (eecb, m_toolbar_info, m_nCmdProtoAuto, TRUE);
      check_menu_toolbar (eecb, m_toolbar_info, m_nCmdProtoPgpmime, FALSE);
      check_menu_toolbar (eecb, m_toolbar_info, m_nCmdProtoSmime, FALSE);
      break;
    }
}


void
GpgolExtCommands::add_toolbar (LPTBENTRY tbearr, UINT n_tbearr, ...)
{
  va_list arg_ptr;
  const char *desc;
  UINT bmapid;
  UINT cmdid;
  int tbeidx;
  toolbar_info_t tb_info;
  int rc;

  for (tbeidx = n_tbearr-1; tbeidx > -1; tbeidx--)
    if (tbearr[tbeidx].tbid == EETBID_STANDARD)
      break;
  if (!(tbeidx > -1))
    {
      log_error ("standard toolbar not found");
      return;
    }
  
  SendMessage (tbearr[tbeidx].hwnd, TB_BUTTONSTRUCTSIZE,
               (WPARAM)(int)sizeof (TBBUTTON), 0);

  
  va_start (arg_ptr, n_tbearr);

  while ( (desc = va_arg (arg_ptr, const char *)) )
    {
      bmapid = va_arg (arg_ptr, UINT);
      cmdid = va_arg (arg_ptr, UINT);

      if (!*desc)
        ; /* Empty description - ignore this item.  */
      else if (*desc == '|' && !desc[1])
        {
          /* Separator. Ignore BMAPID and CMDID.  */
          /* Not yet implemented.  */
        }
      else
	{
          TBADDBITMAP tbab;
  
          tb_info = (toolbar_info_t)xcalloc (1, sizeof *tb_info);
          tb_info->button_id = tbearr[tbeidx].itbbBase++;

          tbab.hInst = glob_hinst;
          tbab.nID = bmapid;
          rc = SendMessage (tbearr[tbeidx].hwnd, TB_ADDBITMAP,1,(LPARAM)&tbab);
          if (rc == -1)
            log_error_w32 (-1, "TB_ADDBITMAP failed for `%s'", desc);
          tb_info->bitmap = rc;
          tb_info->cmd_id = cmdid;
          tb_info->desc = desc;
          tb_info->context = m_lContext;

          tb_info->next = m_toolbar_info;
          m_toolbar_info = tb_info;
          if (debug_commands)
            log_debug ("%s:%s: ctx=%lx button_id=%d cmd_id=%d '%s'\n", 
                       SRCNAME, __func__, m_lContext,
                       tb_info->button_id, tb_info->cmd_id, tb_info->desc);
        }
    }
  va_end (arg_ptr);

}




/* Called by Exchange to install commands and toolbar buttons.  Returns
   S_FALSE to signal Exchange to continue calling extensions. */
STDMETHODIMP 
GpgolExtCommands::InstallCommands (
	LPEXCHEXTCALLBACK eecb, // The Exchange Callback Interface.
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

  (void)hMenu;
  
  if (debug_commands)
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
     body because this would pop up the security window, telling the
     user that someone is trying to read this data.
  */
  if (m_lContext == EECONTEXT_SENDNOTEMESSAGE)
    {
      LPMDB mdb = NULL;
      LPMESSAGE message = NULL;
      
      /*  Note that for read and send the object returned by the
          outlook extension callback is of class 43 (MailItem) so we
          only need to ask for Body then. */
      hr = eecb->GetObject (&mdb, (LPMAPIPROP *)&message);
      if (FAILED(hr))
        log_debug ("%s:%s: getObject failed: hr=%#lx\n", SRCNAME,__func__,hr);
      else if (!opt.compat.no_msgcache)
        {
          const char *body;
          char *key = NULL;
          size_t keylen = 0;
          void *refhandle = NULL;
     
          pDisp = find_outlook_property (eecb, "ConversationIndex", &dispid);
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
              && (pDisp = find_outlook_property (eecb, "Body", &dispid)))
            {
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
      
      ul_release (message, __func__, __LINE__);
      ul_release (mdb, __func__, __LINE__);
    }

  /* Now install menu and toolbar items.  */
  if (m_lContext == EECONTEXT_READNOTEMESSAGE)
    {
      int is_encrypted = 0;
      LPMDB mdb = NULL;
      LPMESSAGE message = NULL;

      hr = eecb->GetObject (&mdb, (LPMAPIPROP *)&message);
      if (FAILED(hr))
        log_debug ("%s:%s: getObject failed: hr=%#lx\n", SRCNAME, __func__, hr);
      else
        {
          switch (mapi_get_message_type (message))
            {
            case MSGTYPE_GPGOL_MULTIPART_ENCRYPTED:
            case MSGTYPE_GPGOL_OPAQUE_ENCRYPTED:
            case MSGTYPE_GPGOL_PGP_MESSAGE:
              is_encrypted = 1;
              if ( mapi_test_sig_status (message) )
                is_encrypted++;
              break;
            default:
              break;
            }
        }
      ul_release (message, __func__, __LINE__);
      ul_release (mdb, __func__, __LINE__);
      
      /* We always enable the verify button as it might be useful on
         an already decrypted message. */
      add_menu (eecb, pnCommandIDBase,
        "@", NULL,
        opt.disable_gpgol? "":_("GpgOL Decrypt/Verify"), &m_nCmdCryptoState,
        opt.enable_debug? "GpgOL Debug-0 (display crypto info)":"", 
                &m_nCmdDebug0,
        (opt.enable_debug && !opt.disable_gpgol)?
                "GpgOL Debug-1 (open_inspector)":"", &m_nCmdDebug1,
        (opt.enable_debug && !opt.disable_gpgol)? 
                "GpgOL Debug-2 (change msg class)":"", &m_nCmdDebug2,
        opt.enable_debug? "GpgOL Debug-3 (revert message class)":"",
                &m_nCmdDebug3,
        NULL);

      if (!opt.disable_gpgol)
        add_toolbar (pTBEArray, nTBECnt, 
                     is_encrypted == 2 
                     ? _("This is a signed and encrypted message.\n"
                         "Click for more information. ")
                     : is_encrypted
                     ? _("This is an encrypted message.\n"
                         "Click for more information. ")
                     : _("This is a signed message.\n"
                       "Click for more information. "),
                     IDB_CRYPTO_STATE, m_nCmdCryptoState,
                     NULL, 0, 0);

    }
  else if (m_lContext == EECONTEXT_SENDNOTEMESSAGE && !opt.disable_gpgol) 
    {
      add_menu (eecb, pnCommandIDBase,
        "@", NULL,
        _(">GnuPG protocol"), NULL,
        _("auto"),   &m_nCmdProtoAuto,        
        _("PGP/MIME"),&m_nCmdProtoPgpmime,        
        _("S/MIME"), &m_nCmdProtoSmime,        
          "<", NULL,
        _("&encrypt message with GnuPG"), &m_nCmdEncrypt,
        _("&sign message with GnuPG"), &m_nCmdSign,
        NULL );

      /* We display the protocol icons only for Outlook 2007 becuase
         there seems to be no way to add a plain menu item. */
      if (get_ol_main_version () > 11)
        add_toolbar (pTBEArray, nTBECnt,
                     "Encrypt", IDB_ENCRYPT, m_nCmdEncrypt,
                     "Sign",    IDB_SIGN,    m_nCmdSign,
                     "Autoselect", IDB_PROTO_AUTO, m_nCmdProtoAuto,
                     "Use PGP/MIME", IDB_PROTO_PGPMIME, m_nCmdProtoPgpmime,
                     "Use/MIME", IDB_PROTO_SMIME, m_nCmdProtoSmime,
                     NULL, 0, 0);
      else
        add_toolbar (pTBEArray, nTBECnt,
                     "Encrypt", IDB_ENCRYPT, m_nCmdEncrypt,
                     "Sign",    IDB_SIGN,    m_nCmdSign,
                     NULL, 0, 0);
      

      m_pExchExt->m_protoSelection = opt.default_protocol;
      update_protocol_menu (eecb);

      m_pExchExt->m_gpgEncrypt = opt.encrypt_default;

      m_pExchExt->m_gpgSign    = opt.sign_default;
      if (force_encrypt)
        m_pExchExt->m_gpgEncrypt = true;
      check_menu (eecb, m_nCmdEncrypt, m_pExchExt->m_gpgEncrypt);
      check_menu (eecb, m_nCmdSign, m_pExchExt->m_gpgSign);
    }
  else if (m_lContext == EECONTEXT_VIEWER) 
    {
      add_menu (eecb, pnCommandIDBase, 
        "@", NULL,
        _("GnuPG Certificate &Manager"), &m_nCmdKeyManager,
        _("Remove GpgOL flags from this folder"), &m_nCmdRevertFolder,
        NULL);

      add_toolbar (pTBEArray, nTBECnt, 
        _("Open the certificate manager"), IDB_KEY_MANAGER, m_nCmdKeyManager,
        NULL, 0, 0);
    }
  return S_FALSE;
}


/* Called by Exchange when a user selects a command.  Return value:
   S_OK if command is handled, otherwise S_FALSE. */
STDMETHODIMP 
GpgolExtCommands::DoCommand (LPEXCHEXTCALLBACK eecb, UINT nCommandID)
{
  HRESULT hr;
  HWND hwnd = NULL;
  LPMESSAGE message = NULL;
  LPMDB mdb = NULL;
      
  if (FAILED (eecb->GetWindow (&hwnd)))
    hwnd = NULL;

  if (debug_commands)
    log_debug ("%s:%s: commandID=%u (%#x) context=%s hwnd=%p\n",
               SRCNAME, __func__, nCommandID, nCommandID, 
               ext_context_name (m_lContext), hwnd);

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
      
      if (debug_commands)
        log_debug ("%s:%s: command Close called\n", SRCNAME, __func__);
      pDisp = find_outlook_property (eecb, "Close", &dispid);
      if (pDisp)
        {
          /* Note that there is a report on the Net from 2005 by Amit
             Joshi where he claims that in Outlook XP olDiscard does
             not work but is treated like olSave.  */ 
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
              message_wipe_body_cruft (eecb);
              return S_OK; /* We handled the close command. */
            }

          log_debug ("%s:%s: invoking Close failed: %#lx",
                     SRCNAME, __func__, hr);
        }
      else
        log_debug ("%s:%s: invoking Close failed: no Close method)",
                   SRCNAME, __func__);

      message_wipe_body_cruft (eecb);

      /* Closing on our own failed - pass it on. */
      return S_FALSE; 
    }
  else if (nCommandID == EECMDID_ComposeReplyToSender)
    {
      if (debug_commands)
        log_debug ("%s:%s: command Reply called\n", SRCNAME, __func__);
      /* What we might want to do is to call Reply, then GetInspector
         and then Activate - this allows us to get full control over
         the quoted message and avoids the ugly msgcache. */
      return S_FALSE; /* Pass it on.  */
    }
  else if (nCommandID == EECMDID_ComposeReplyToAll)
    {
      if (debug_commands)
        log_debug ("%s:%s: command ReplyAll called\n", SRCNAME, __func__);
      return S_FALSE; /* Pass it on.  */
    }
  else if (nCommandID == EECMDID_ComposeForward)
    {
      if (debug_commands)
        log_debug ("%s:%s: command Forward called\n", SRCNAME, __func__);
      return S_FALSE; /* Pass it on.  */
    }
  else if (nCommandID == m_nCmdCryptoState
           && m_lContext == EECONTEXT_READNOTEMESSAGE)
    {
      log_debug ("%s:%s: command CryptoState called\n", SRCNAME, __func__);
      hr = eecb->GetObject (&mdb, (LPMAPIPROP *)&message);
      if (SUCCEEDED (hr))
        {
          if (message_incoming_handler (message, hwnd, true))
            message_display_handler (eecb, hwnd);
	}
      else
        log_debug_w32 (hr, "%s:%s: command CryptoState failed", 
                       SRCNAME, __func__);
      ul_release (message, __func__, __LINE__);
      ul_release (mdb, __func__, __LINE__);
    }
  else if (nCommandID == m_nCmdProtoAuto
           && m_lContext == EECONTEXT_SENDNOTEMESSAGE) 
    {
      log_debug ("%s:%s: command ProtoAuto called\n", SRCNAME, __func__);
      m_pExchExt->m_protoSelection = PROTOCOL_UNKNOWN;
      update_protocol_menu (eecb);
    }
  else if (nCommandID == m_nCmdProtoPgpmime
           && m_lContext == EECONTEXT_SENDNOTEMESSAGE) 
    {
      log_debug ("%s:%s: command ProtoPgpmime called\n", SRCNAME, __func__);
      m_pExchExt->m_protoSelection = PROTOCOL_OPENPGP;
      update_protocol_menu (eecb);
    }
  else if (nCommandID == m_nCmdProtoSmime
           && m_lContext == EECONTEXT_SENDNOTEMESSAGE) 
    {
      log_debug ("%s:%s: command ProtoSmime called\n", SRCNAME, __func__);
      if (opt.enable_smime)
        {
          m_pExchExt->m_protoSelection = PROTOCOL_SMIME;
          update_protocol_menu (eecb);
        }
      else
        {
          MessageBox (hwnd,
                      _("Support for S/MIME has not been enabled.\n"
                        "\n"
                        "To enable S/MIME support, open the option dialog"
                        " and check \"Enable the S/MIME support\".  The"
                        " option dialog can be found in the main menu at:"
                       " Extras->Options->GpgOL.\n"),
                      "GpgOL", MB_ICONHAND|MB_OK);
        }
    }
  else if (nCommandID == m_nCmdEncrypt
           && m_lContext == EECONTEXT_SENDNOTEMESSAGE) 
    {
      log_debug ("%s:%s: command Encrypt called\n", SRCNAME, __func__);
      m_pExchExt->m_gpgEncrypt = !m_pExchExt->m_gpgEncrypt;
      check_menu (eecb, m_nCmdEncrypt, m_pExchExt->m_gpgEncrypt);
    }
  else if (nCommandID == m_nCmdSign
           && m_lContext == EECONTEXT_SENDNOTEMESSAGE) 
    {
      log_debug ("%s:%s: command Sign called\n", SRCNAME, __func__);
      m_pExchExt->m_gpgSign = !m_pExchExt->m_gpgSign;
      check_menu (eecb, m_nCmdSign, m_pExchExt->m_gpgSign);
    }
  else if (nCommandID == m_nCmdKeyManager
           && m_lContext == EECONTEXT_VIEWER)
    {
      log_debug ("%s:%s: command KeyManager called\n", SRCNAME, __func__);
      if (engine_start_keymanager (hwnd))
        MessageBox (NULL, _("Could not start certificate manager"),
                    _("GpgOL"), MB_ICONERROR|MB_OK);
    }
  else if (nCommandID == m_nCmdRevertFolder
           && m_lContext == EECONTEXT_VIEWER)
    {
      log_debug ("%s:%s: command ReverFoldert called\n", SRCNAME, __func__);
      /* Notify the user that the general GpgOl fucntionaly will be
         disabled when calling this function the first time.  */
      if ( opt.disable_gpgol
           || (MessageBox 
               (hwnd,
                _("You are about to start the process of reversing messages "
                  "created by GpgOL to prepare deinstalling of GpgOL. "
                  "Running this command will put GpgOL into a disabled state "
                  "so that messages are not anymore processed by GpgOL.\n"
                  "\n"
                  "You should convert all folders one after the other with "
                  "this command, close Outlook and then deinstall GpgOL.\n"
                  "\n"
                  "Note that if you start Outlook again with GpgOL still "
                  "being installed, GpgOL will again process messages."),
                _("GpgOL"), MB_ICONWARNING|MB_OKCANCEL) == IDOK))
        {
          if ( MessageBox 
               (hwnd,
                _("Do you want to revert this folder?"),
                _("GpgOL"), MB_ICONQUESTION|MB_YESNO) == IDYES )
            {
              if (!opt.disable_gpgol)
                opt.disable_gpgol = 1;
          
              gpgol_folder_revert (eecb);
            }
        }
    }
  else if (opt.enable_debug && nCommandID == m_nCmdDebug0
           && m_lContext == EECONTEXT_READNOTEMESSAGE)
    {
      log_debug ("%s:%s: command Debug0 (showInfo) called\n",
                 SRCNAME, __func__);
      hr = eecb->GetObject (&mdb, (LPMAPIPROP *)&message);
      if (SUCCEEDED (hr))
        {
          message_show_info (message, hwnd);
	}
      ul_release (message, __func__, __LINE__);
      ul_release (mdb, __func__, __LINE__);
    }
  else if (opt.enable_debug && nCommandID == m_nCmdDebug1
           && m_lContext == EECONTEXT_READNOTEMESSAGE)
    {
      log_debug ("%s:%s: command Debug1 (open inspector) called\n",
                 SRCNAME, __func__);
      hr = eecb->GetObject (&mdb, (LPMAPIPROP *)&message);
      if (SUCCEEDED (hr))
        {
          open_inspector (eecb, message);
	}
      ul_release (message, __func__, __LINE__);
      ul_release (mdb, __func__, __LINE__);
    }
  else if (opt.enable_debug && nCommandID == m_nCmdDebug2
           && m_lContext == EECONTEXT_READNOTEMESSAGE)
    {
      log_debug ("%s:%s: command Debug2 (change message class) called\n", 
                 SRCNAME, __func__);
      hr = eecb->GetObject (&mdb, (LPMAPIPROP *)&message);
      if (SUCCEEDED (hr))
        {
          /* We sync here. */
          mapi_change_message_class (message, 1);
	}
      ul_release (message, __func__, __LINE__);
      ul_release (mdb, __func__, __LINE__);
    }
  else if (opt.enable_debug && nCommandID == m_nCmdDebug3
           && m_lContext == EECONTEXT_READNOTEMESSAGE)
    {
      log_debug ("%s:%s: command Debug3 (revert_message_class) called\n", 
                 SRCNAME, __func__);
      hr = eecb->GetObject (&mdb, (LPMAPIPROP *)&message);
      if (SUCCEEDED (hr))
        {
          int rc = gpgol_message_revert (message, 1, 
                                         KEEP_OPEN_READWRITE|FORCE_SAVE);
          log_debug ("%s:%s: gpgol_message_revert returns %d\n", 
                     SRCNAME, __func__, rc);
	}
      ul_release (message, __func__, __LINE__);
      ul_release (mdb, __func__, __LINE__);
    }
  else
    {
      if (debug_commands)
        log_debug ("%s:%s: command passed on\n", SRCNAME, __func__);
      return S_FALSE; /* Pass on unknown command. */
    }
  

  return S_OK; 
}


/* Called by Exchange when it receives a WM_INITMENU message, allowing
   the extension object to enable, disable, or update its menu
   commands before the user sees them. This method is called
   frequently and should be written in a very efficient manner. */
STDMETHODIMP_(VOID) 
GpgolExtCommands::InitMenu(LPEXCHEXTCALLBACK eecb) 
{
  (void)eecb;
}


/* Called by Exchange when the user requests help for a menu item.
   EECB is the pointer to Exchange Callback Interface.  NCOMMANDID is
   the command id.  Return value: S_OK when it is a menu item of this
   plugin and the help was shown; otherwise S_FALSE.  */
STDMETHODIMP 
GpgolExtCommands::Help (LPEXCHEXTCALLBACK eecb, UINT nCommandID)
{
  (void)eecb;

  if (nCommandID == m_nCmdProtoAuto
           && m_lContext == EECONTEXT_SENDNOTEMESSAGE) 
    {
      MessageBox (m_hWnd,
                 _("Select this option to automatically select the protocol."),
                  "GpgOL", MB_OK);	
    } 
  else if (nCommandID == m_nCmdProtoPgpmime
           && m_lContext == EECONTEXT_SENDNOTEMESSAGE) 
    {
      MessageBox (m_hWnd,
                  _("Select this option to select the PGP/MIME protocol."),
                  "GpgOL", MB_OK);	
    } 
  else if (nCommandID == m_nCmdProtoSmime
           && m_lContext == EECONTEXT_SENDNOTEMESSAGE) 
    {
      MessageBox (m_hWnd,
                  _("Select this option to select the S/MIME protocol."),
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
      MessageBox (m_hWnd,
                  _("Select this option to open the certificate manager"),
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
	
  if (nCommandID == m_nCmdProtoAuto
           && m_lContext == EECONTEXT_SENDNOTEMESSAGE) 
    {
      if (lFlags == EECQHT_STATUS)
        lstrcpyn (pszText, ".", nCharCnt);
      if (lFlags == EECQHT_TOOLTIP)
        lstrcpyn (pszText,
                  _("Automatically select the protocol for sign/encrypt"),
                  nCharCnt);
    }
  else if (nCommandID == m_nCmdProtoPgpmime
           && m_lContext == EECONTEXT_SENDNOTEMESSAGE) 
    {
      if (lFlags == EECQHT_STATUS)
        lstrcpyn (pszText, ".", nCharCnt);
      if (lFlags == EECQHT_TOOLTIP)
        lstrcpyn (pszText,
                  _("Use PGP/MIME for sign/encrypt"),
                  nCharCnt);
    }
  else if (nCommandID == m_nCmdProtoSmime
           && m_lContext == EECONTEXT_SENDNOTEMESSAGE) 
    {
      if (lFlags == EECQHT_STATUS)
        lstrcpyn (pszText, ".", nCharCnt);
      if (lFlags == EECQHT_TOOLTIP)
        lstrcpyn (pszText,
                  _("Use S/MIME for sign/encrypt"),
                  nCharCnt);
    }
  else if (nCommandID == m_nCmdEncrypt
           && m_lContext == EECONTEXT_SENDNOTEMESSAGE) 
    {
      if (lFlags == EECQHT_STATUS)
        lstrcpyn (pszText, ".", nCharCnt);
      if (lFlags == EECQHT_TOOLTIP)
        lstrcpyn (pszText,
                  _("Encrypt message with GnuPG"),
                  nCharCnt);
    }
  else if (nCommandID == m_nCmdSign
           && m_lContext == EECONTEXT_SENDNOTEMESSAGE) 
    {
      if (lFlags == EECQHT_STATUS)
        lstrcpyn (pszText, ".", nCharCnt);
      if (lFlags == EECQHT_TOOLTIP)
        lstrcpyn (pszText,
                  _("Sign message with GnuPG"),
                  nCharCnt);
    }
  else if (nCommandID == m_nCmdKeyManager
           && m_lContext == EECONTEXT_VIEWER) 
    {
      if (lFlags == EECQHT_STATUS)
        lstrcpyn (pszText, ".", nCharCnt);
      if (lFlags == EECQHT_TOOLTIP)
        lstrcpyn (pszText,
                  _("Open the GpgOL certificate manager"),
                  nCharCnt);
    }
  else 
    return S_FALSE;

  return S_OK;
}


/* Called by Exchange to get toolbar button infos.  TOOLBARID is the
   toolbar identifier.  BUTTONID is the toolbar button index.  PTBB is
   a pointer to the toolbar button structure.  DESCRIPTION is a pointer to
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
  toolbar_info_t tb_info;
  size_t n;
  

  (void)description_size;
  (void)flags;

  for (tb_info = m_toolbar_info; tb_info; tb_info = tb_info->next )
    if (tb_info->button_id == buttonid
        && tb_info->context == m_lContext)
      break;
  if (!tb_info)
    return S_FALSE; /* Not one of our toolbar buttons.  */

  if (debug_commands)
    log_debug ("%s:%s: ctx=%lx tbid=%ld button_id(req)=%d got=%d"
               " cmd_id=%d '%s'\n", 
               SRCNAME, __func__, m_lContext, toolbarid, buttonid,
               tb_info->button_id, tb_info->cmd_id, tb_info->desc);

  /* Mark that this button has passed this function.  */
  tb_info->did_qbi = 1;
  
  pTBB->iBitmap = tb_info->bitmap;
  pTBB->idCommand = tb_info->cmd_id;
  pTBB->fsState = TBSTATE_ENABLED;
  pTBB->fsStyle = TBSTYLE_BUTTON;
  pTBB->dwData = 0;
  pTBB->iString = -1;
  
  n = strlen (tb_info->desc);
  if (n > description_size)
    n = description_size;
  lstrcpyn (description, tb_info->desc, n);

  if (tb_info->cmd_id == m_nCmdEncrypt)
    {
      pTBB->fsStyle |= TBSTYLE_CHECK;
      if (m_pExchExt->m_gpgEncrypt)
        pTBB->fsState |= TBSTATE_CHECKED;
    }
  else if (tb_info->cmd_id == m_nCmdSign)
    {
      pTBB->fsStyle |= TBSTYLE_CHECK;
      if (m_pExchExt->m_gpgSign)
        pTBB->fsState |= TBSTATE_CHECKED;
    }
  else if (tb_info->cmd_id == m_nCmdProtoAuto)
    {
      pTBB->fsStyle |= TBSTYLE_CHECK;
      if (m_pExchExt->m_protoSelection != PROTOCOL_OPENPGP
          && m_pExchExt->m_protoSelection != PROTOCOL_SMIME)
        pTBB->fsState |= TBSTATE_CHECKED;
    }
  else if (tb_info->cmd_id == m_nCmdProtoPgpmime)
    {
      pTBB->fsStyle |= TBSTYLE_CHECK;
      if (m_pExchExt->m_protoSelection == PROTOCOL_OPENPGP)
        pTBB->fsState |= TBSTATE_CHECKED;
    }
  else if (tb_info->cmd_id == m_nCmdProtoSmime)
    {
      pTBB->fsStyle |= TBSTYLE_CHECK;
      if (m_pExchExt->m_protoSelection == PROTOCOL_SMIME)
        pTBB->fsState |= TBSTATE_CHECKED;
    }

  return S_OK;
}



STDMETHODIMP 
GpgolExtCommands::ResetToolbar (ULONG lToolbarID, ULONG lFlags)
{	
  (void)lToolbarID;
  (void)lFlags;
  
  return S_OK;
}

