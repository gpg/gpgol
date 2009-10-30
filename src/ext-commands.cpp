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
#include "message.h"
#include "engine.h"
#include "revert.h"
#include "ext-commands.h"
#include "explorers.h"
#include "inspectors.h"

#define TRACEPOINT() do { log_debug ("%s:%s:%d: tracepoint\n", \
                                     SRCNAME, __func__, __LINE__); \
                        } while (0)


/* Keep copies of some bitmaps.  */
static int bitmaps_initialized;
static HBITMAP my_check_bitmap, my_uncheck_bitmap;


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


/* Note: Duplicated from message-events.cpp.  Eventually we should get
   rid of this module.  */
static LPDISPATCH
get_inspector (LPEXCHEXTCALLBACK eecb)
{
  LPDISPATCH obj;
  LPDISPATCH inspector = NULL;
  
  obj = get_eecb_object (eecb);
  if (obj)
    {
      /* This should be MailItem; use the getInspector method.  */
      inspector = get_oom_object (obj, "GetInspector");
      obj->Release ();
    }
  return inspector;
}


/* Note: Duplicated from message-events.cpp.  Eventually we should get
   rid of this module.  */
static int
get_crypto_flags (LPEXCHEXTCALLBACK eecb, bool *r_sign, bool *r_encrypt)
{
  LPDISPATCH inspector;
  int rc;

  inspector = get_inspector (eecb);
  if (!inspector)
    {
      log_error ("%s:%s: inspector not found", SRCNAME, __func__);
      rc = -1;
    }
  else
    {
      rc = get_inspector_composer_flags (inspector, r_sign, r_encrypt);
      inspector->Release ();
    }
  return rc;
}


static void
set_crypto_flags (LPEXCHEXTCALLBACK eecb, bool sign, bool encrypt)
{
  LPDISPATCH inspector;

  inspector = get_inspector (eecb);
  if (!inspector)
    log_error ("%s:%s: inspector not found", SRCNAME, __func__);
  else
    {
      set_inspector_composer_flags (inspector, sign, encrypt);
      inspector->Release ();
    }
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
  LPDISPATCH obj;

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
      int force_encrypt = 0;
      char *draft_info = NULL;
      
      /*  Note that for read and send the object returned by the
          outlook extension callback is of class 43 (MailItem) so we
          only need to ask for Body then. */
      hr = eecb->GetObject (&mdb, (LPMAPIPROP *)&message);
      if (FAILED(hr))
        log_debug ("%s:%s: getObject failed: hr=%#lx\n", SRCNAME,__func__,hr);
      else if (!opt.compat.no_msgcache)
        {
          obj = get_eecb_object (eecb);
          if (obj)
            {
              const char *body;
              char *key, *p;
              size_t keylen;
              void *refhandle = NULL;

              key = get_oom_string (obj, "ConversationIndex");
              if (key)
                {
                  log_debug ("%s:%s: ConversationIndex is `%s'",
                             SRCNAME, __func__, key);
                  /* The key is a hex string.  Convert it to binary. */
                  for (keylen=0,p=key; hexdigitp(p) && hexdigitp(p+1); p += 2)
                    ((unsigned char*)key)[keylen++] = xtoi_2 (p);
                  
                  if (keylen && (body = msgcache_get (key, keylen, &refhandle)))
                    {
                      put_oom_string (obj, "Body", body);
                      /* Because we found the plaintext in the cache
                         we can assume that the orginal message has
                         been encrypted and thus we now set a flag to
                         make sure that by default the reply gets
                         encrypted too. */
                      force_encrypt = 1;
                    }
                  msgcache_unref (refhandle);
                  xfree (key);
                }
              obj->Release ();
            }
        }
      
      /* Because we have the message open, we use it to get the draft
         info property.  */
      if (message)
        draft_info = mapi_get_gpgol_draft_info (message);

      ul_release (message, __func__, __LINE__);
      ul_release (mdb, __func__, __LINE__);
      
      if (!opt.disable_gpgol) 
        {
          bool sign, encrypt;
          
          if (draft_info && draft_info[0] == 'E')
            encrypt = true;
          else if (draft_info && draft_info[0] == 'e')
            encrypt = false;
          else
            encrypt = !!opt.encrypt_default;
          
          if (draft_info && draft_info[0] && draft_info[1] == 'S')
            sign = true;
          else if (draft_info && draft_info[0] && draft_info[1] == 's')
            sign = false;
          else
            sign = !!opt.sign_default;
          
          if (force_encrypt)
            encrypt = true;
          
          /* FIXME:  ove that to the inspector activation.  */
          //set_crypto_flags (eecb, sign, encrypt);
        }
      xfree (draft_info);
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
      
      if (debug_commands)
        log_debug ("%s:%s: command Close called\n", SRCNAME, __func__);
      
      pDisp = get_eecb_object (eecb);
      dispid = lookup_oom_dispid (pDisp, "Close");
      if (pDisp && dispid != DISPID_UNKNOWN)
        {
          DISPPARAMS dispparams;
          VARIANT aVariant;
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
        {
          if (pDisp)
            pDisp->Release ();
          log_debug ("%s:%s: invoking Close failed: no Close method)",
                     SRCNAME, __func__);
        }

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
  else if (nCommandID == EECMDID_SaveMessage
           && m_lContext == EECONTEXT_SENDNOTEMESSAGE) 
    {
      char buf[4];
      bool sign, encrypt;
      
      log_debug ("%s:%s: command SaveMessage called\n", SRCNAME, __func__);

      if (get_crypto_flags (eecb, &sign, &encrypt))
        buf[0] = buf[1] = '?';
      else
        {
          buf[0] = encrypt? 'E':'e';
          buf[1] = sign?    'S':'s';
        }
      buf[2] = 'A'; /* Automatic.  */
      buf[3] = 0;

      hr = eecb->GetObject (&mdb, (LPMAPIPROP *)&message);
      if (SUCCEEDED (hr))
        mapi_set_gpgol_draft_info (message, buf);
      else
        log_debug ("%s:%s: getObject failed: hr=%#lx\n",SRCNAME, __func__, hr);
      ul_release (message, __func__, __LINE__);
      ul_release (mdb, __func__, __LINE__);
      return S_FALSE; /* Pass on to next handler.  */
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
  (void)nCommandID;

  return S_FALSE;
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

  return S_FALSE;
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
  (void)description_size;
  (void)flags;

  return S_FALSE; /* Not one of our toolbar buttons.  */
}



STDMETHODIMP 
GpgolExtCommands::ResetToolbar (ULONG lToolbarID, ULONG lFlags)
{	
  (void)lToolbarID;
  (void)lFlags;
  
  return S_OK;
}

