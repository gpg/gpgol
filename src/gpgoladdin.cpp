/* gpgoladdin.cpp - Connect GpgOL to Outlook as an addin
 * Copyright (C) 2013 Intevation GmbH
 *    2015 by Bundesamt für Sicherheit in der Informationstechnik
 * Software engineering by Intevation GmbH
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
#include <stdio.h>
#include <string.h>

#include "common.h"
#include "gpgoladdin.h"

#include "mymapi.h"
#include "mymapitags.h"

#include "mapihelp.h"

#include "oomhelp.h"

#include "olflange.h"

#include "gpgol-ids.h"
#include "ribbon-callbacks.h"
#include "eventsinks.h"
#include "eventsink.h"
#include "windowmessages.h"
#include "mail.h"
#include "addin-options.h"

#include <gpg-error.h>
#include <list>

#define ICON_SIZE_LARGE  32
#define ICON_SIZE_NORMAL 16

/* We use UTF-8 internally. */
#undef _
#define _(a) utf8_gettext (a)

ULONG addinLocks = 0;

bool can_unload = false;

static std::list<LPDISPATCH> g_ribbon_uis;

static GpgolAddin * addin_instance = NULL;

/* This is the main entry point for the addin
   Outlook uses this function to query for an Object implementing
   the IClassFactory interface.
*/
STDAPI DllGetClassObject (REFCLSID rclsid, REFIID riid, LPVOID* ppvObj)
{
  if (!ppvObj)
    return E_POINTER;

  *ppvObj = NULL;
  if (rclsid != CLSID_GPGOL)
    return CLASS_E_CLASSNOTAVAILABLE;

  /* Let the factory give the requested interface. */
  GpgolAddinFactory* factory = new GpgolAddinFactory();
  if (!factory)
    return E_OUTOFMEMORY;

  HRESULT hr = factory->QueryInterface (riid, ppvObj);
  if(FAILED(hr))
    {
      *ppvObj = NULL;
      delete factory;
    }

  return hr;
}


STDAPI DllCanUnloadNow()
{
  /* This is called regularly to check if memory can be freed
     by unloading the dll. The following unload will not call
     any addin methods like disconnect etc. It will just
     unload the Library. Any callbacks will become invalid.
     So we _only_ say it's ok to unload if we were disconnected.
     For the epic story behind the next line see GnuPG-Bug-Id 1837 */
  return can_unload ? S_OK : S_FALSE;
}

/* Class factory */
STDMETHODIMP GpgolAddinFactory::QueryInterface (REFIID riid, LPVOID* ppvObj)
{
  HRESULT hr = S_OK;

  *ppvObj = NULL;

  if ((IID_IUnknown == riid) || (IID_IClassFactory == riid))
    *ppvObj = static_cast<IClassFactory*>(this);
  else
    {
      hr = E_NOINTERFACE;
      LPOLESTR sRiid = NULL;
      StringFromIID (riid, &sRiid);
      /* Should not happen */
      log_debug ("GpgolAddinFactory queried for unknown interface: %S \n", sRiid);
    }

  if (*ppvObj)
    ((LPUNKNOWN)*ppvObj)->AddRef();

  return hr;
}


/* This actually creates the instance of our COM object */
STDMETHODIMP GpgolAddinFactory::CreateInstance (LPUNKNOWN punk, REFIID riid,
                                                LPVOID* ppvObj)
{
  (void)punk;
  *ppvObj = NULL;

  GpgolAddin* obj = GpgolAddin::get_instance();
  if (NULL == obj)
    return E_OUTOFMEMORY;

  HRESULT hr = obj->QueryInterface (riid, ppvObj);

  if (FAILED(hr))
    {
      LPOLESTR sRiid = NULL;
      StringFromIID (riid, &sRiid);
      fprintf(stderr, "failed to create instance for: %S", sRiid);
    }

  return hr;
}

/* GpgolAddin definition */


/* Constructor of GpgolAddin

   Initializes members and creates the interface objects for the new
   context.  Does the DLL initialization if it has not been done
   before.

   The ref count is set by the factory after creation.
*/
GpgolAddin::GpgolAddin (void) : m_lRef(0),
  m_application(nullptr),
  m_addin(nullptr),
  m_applicationEventSink(nullptr),
  m_explorersEventSink(nullptr),
  m_disabled(false),
  m_hook(nullptr)
{
  read_options ();
  /* RibbonExtender is it's own object to avoid the pitfalls of
     multiple inheritance
  */
  m_ribbonExtender = new GpgolRibbonExtender();
}

GpgolAddin::~GpgolAddin (void)
{
  if (m_disabled)
    {
      return;
    }
  log_debug ("%s:%s: Releasing Application Event Sink;",
             SRCNAME, __func__);
  gpgol_release (m_explorersEventSink);
  gpgol_release (m_applicationEventSink);

  write_options ();

  UnhookWindowsHookEx (m_hook);

  addin_instance = NULL;

  log_debug ("%s:%s: Object deleted\n", SRCNAME, __func__);
}

STDMETHODIMP
GpgolAddin::QueryInterface (REFIID riid, LPVOID* ppvObj)
{
  HRESULT hr = S_OK;

  *ppvObj = NULL;

  if (m_disabled)
    return E_NOINTERFACE;

  if ((riid == IID_IUnknown) || (riid == IID_IDTExtensibility2) ||
      (riid == IID_IDispatch))
    {
      *ppvObj = (LPUNKNOWN) this;
    }
  else if (riid == IID_IRibbonExtensibility)
    {
      return m_ribbonExtender->QueryInterface (riid, ppvObj);
    }
  else
    {
      hr = E_NOINTERFACE;
#if 0
      LPOLESTR sRiid = NULL;
      StringFromIID(riid, &sRiid);
      log_debug ("%s:%s: queried for unimplmented interface: %S",
                 SRCNAME, __func__, sRiid);
#endif
    }

  if (*ppvObj)
    ((LPUNKNOWN)*ppvObj)->AddRef();

  return hr;
}

static void
addGpgOLToReg (const std::string &path)
{
  HKEY h;
  int err = RegOpenKeyEx (HKEY_CURRENT_USER, path.c_str(), 0,
                          KEY_ALL_ACCESS, &h);
  if (err != ERROR_SUCCESS)
    {
      log_debug ("%s:%s: no DoNotDisableAddinList entry '%s' creating it",
                 SRCNAME, __func__, path.c_str ());
      err = RegCreateKeyEx (HKEY_CURRENT_USER, path.c_str (), 0, NULL,
                            REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL,
                            &h, NULL);
    }
  if (err != ERROR_SUCCESS)
    {
      log_error ("%s:%s: failed to create key.",
                 SRCNAME, __func__);
      return;
    }

  DWORD type;
  err = RegQueryValueEx (h, GPGOL_PROGID, NULL, &type, NULL, NULL);
  if (err == ERROR_SUCCESS)
    {
      log_debug ("%s:%s: Found gpgol reg key. Leaving it unchanged.",
                 SRCNAME, __func__);
      RegCloseKey (h);
      return;
    }

  // No key exists. Create one.
  DWORD dwTemp = 1;
  err = RegSetValueEx (h, GPGOL_PROGID, 0, REG_DWORD, (BYTE*)&dwTemp, 4);
  RegCloseKey (h);

  if (err != ERROR_SUCCESS)
    {
      log_error ("%s:%s: failed to set registry value.",
                 SRCNAME, __func__);
    }
  else
    {
      log_debug ("%s:%s: added gpgol to %s",
                 SRCNAME, __func__, path.c_str ());
    }
}

/* This is a bit evil as we basically disable outlooks resiliency
   for us. But users are still able to manually disable the addon
   or change the donotdisable setting to zero and we won't change
   it.

   It has been much requested by users that we do this automatically.
*/
static void
setupDoNotDisable ()
{
  std::string path = "Software\\Microsoft\\Office\\";
  path += std::to_string (g_ol_version_major);
  path += ".0\\Outlook\\Resiliency\\DoNotDisableAddinList";

  addGpgOLToReg (path);

  path = "Software\\Microsoft\\Office\\";
  path += std::to_string (g_ol_version_major);
  path += ".0\\Outlook\\Resiliency\\AddinList";

  addGpgOLToReg (path);
}

STDMETHODIMP
GpgolAddin::OnConnection (LPDISPATCH Application, ext_ConnectMode ConnectMode,
                          LPDISPATCH AddInInst, SAFEARRAY ** custom)
{
  (void)custom;
  char* version;

  log_debug ("%s:%s: this is GpgOL %s\n",
             SRCNAME, __func__, PACKAGE_VERSION);

  can_unload = false;
  m_application = Application;
  m_application->AddRef();
  m_addin = AddInInst;

  version = get_oom_string (Application, "Version");

  log_debug ("%s:%s:   using GPGME %s\n",
             SRCNAME, __func__, gpgme_check_version (NULL));
  log_debug ("%s:%s:   in Outlook %s\n",
             SRCNAME, __func__, version);

  g_ol_version_major = atoi (version);

  if (!version || !strlen (version) ||
      (strncmp (version, "14", 2) &&
       strncmp (version, "15", 2) &&
       strncmp (version, "16", 2)))
    {
      m_disabled = true;
      log_debug ("%s:%s: Disabled addin for unsupported version.",
                 SRCNAME, __func__);

      xfree (version);
      return S_OK;
    }

  setupDoNotDisable ();

  if (ConnectMode != ext_cm_Startup)
    {
      OnStartupComplete (custom);
    }
  return S_OK;
}

STDMETHODIMP
GpgolAddin::OnDisconnection (ext_DisconnectMode RemoveMode,
                             SAFEARRAY** custom)
{
  (void)custom;
  (void)RemoveMode;
  log_debug ("%s:%s: cleaning up GpgolAddin object;",
             SRCNAME, __func__);

  /* Doing the wipe in the dtor is too late. Outlook
     does not allow us any OOM calls then and only returns
     "Unexpected error" in that case. Weird. */

  if (Mail::closeAllMails_o ())
    {
      MessageBox (NULL,
                  "Failed to remove plaintext from at least one message.\n\n"
                  "Until GpgOL is activated again it is possible that the "
                  "plaintext of messages decrypted in this Session is saved "
                  "or transfered back to your mailserver.",
                  _("GpgOL"),
                  MB_ICONINFORMATION|MB_OK);
    }

  write_options();
  can_unload = true;
  return S_OK;
}

STDMETHODIMP
GpgolAddin::OnAddInsUpdate (SAFEARRAY** custom)
{
  (void)custom;
  return S_OK;
}

static void
check_html_preferred()
{
  /* Check if HTML Mail should be enabled. */
  HKEY h;
  std::string path = "Software\\Microsoft\\Office\\";
  path += std::to_string (g_ol_version_major);
  path += ".0\\Outlook\\Options\\Mail";
  opt.prefer_html = 1;
  int err = RegOpenKeyEx (HKEY_CURRENT_USER, path.c_str() , 0, KEY_READ, &h);
  if (err != ERROR_SUCCESS)
    {
      log_debug ("%s:%s: no mail options under %s",
                 SRCNAME, __func__, path.c_str());
      return;
    }
  else
    {
      DWORD type;
      err = RegQueryValueEx (h, "ReadAsPlain", NULL, &type, NULL, NULL);
      if (err != ERROR_SUCCESS || type != REG_DWORD)
        {
          log_debug ("%s:%s: No type or key for ReadAsPlain",
                     SRCNAME, __func__);
          return;
        }
      else
        {
          DWORD data;
          DWORD size = sizeof (DWORD);
          err = RegQueryValueEx (h, "ReadAsPlain", NULL, NULL, (LPBYTE)&data,
                                 &size);
          if (err != ERROR_SUCCESS)
            {
              log_debug ("%s:%s: Failed to find out ReadAsPlain",
                         SRCNAME, __func__);
              return;
            }
          opt.prefer_html = data ? 0 : 1;
          return;
        }
    }
}

static LPDISPATCH
install_explorer_sinks (LPDISPATCH application)
{

  LPDISPATCH explorers = get_oom_object (application, "Explorers");

  if (!explorers)
    {
      log_error ("%s:%s: No explorers object",
                 SRCNAME, __func__);
      return nullptr;
    }
  int count = get_oom_int (explorers, "Count");

  for (int i = 1; i <= count; i++)
    {
      std::string item = "Item(";
      item += std::to_string (i) + ")";
      LPDISPATCH explorer = get_oom_object (explorers, item.c_str());
      if (!explorer)
        {
          log_error ("%s:%s: failed to get explorer %i",
                     SRCNAME, __func__, i);
          continue;
        }
      /* Explorers delete themself in the close event of the explorer. */
      LPDISPATCH sink = install_ExplorerEvents_sink (explorer);
      if (!sink)
        {
          log_error ("%s:%s: failed to create eventsink for explorer %i",
                     SRCNAME, __func__, i);

        }
      else
        {
          log_oom_extra ("%s:%s: created sink %p for explorer %i",
                         SRCNAME, __func__, sink, i);
        }
      add_explorer (explorer);
      gpgol_release (explorer);
    }
  /* Now install the event sink to handle new explorers */
  return install_ExplorersEvents_sink (explorers);
}

STDMETHODIMP
GpgolAddin::OnStartupComplete (SAFEARRAY** custom)
{
  (void)custom;
  TRACEPOINT;

  i18n_init ();

  if (!create_responder_window())
    {
      log_error ("%s:%s: Failed to create the responder window;",
                 SRCNAME, __func__);
    }

  if (!m_application)
    {
      /* Should not happen as OnConnection should be called before */
      log_error ("%s:%s: no application set;",
                 SRCNAME, __func__);
      return E_NOINTERFACE;
    }

  if (!(m_hook = create_message_hook ()))
    {
      log_error ("%s:%s: Failed to create messagehook. ",
                 SRCNAME, __func__);
    }

  /* Set up categories */
  const char *decCategory = _("GpgOL: Encrypted Message");
  const char *verifyCategory = _("GpgOL: Trusted Sender Address");
  ensure_category_exists (m_application, decCategory, 8);
  ensure_category_exists (m_application, verifyCategory, 5);
  install_forms ();
  m_applicationEventSink = install_ApplicationEvents_sink (m_application);
  m_explorersEventSink = install_explorer_sinks (m_application);
  check_html_preferred ();
  return S_OK;
}

STDMETHODIMP
GpgolAddin::OnBeginShutdown (SAFEARRAY * * custom)
{
  (void)custom;
  TRACEPOINT;
  return S_OK;
}

STDMETHODIMP
GpgolAddin::GetTypeInfoCount (UINT *r_count)
{
  *r_count = 0;
  TRACEPOINT; /* Should not happen */
  return S_OK;
}

STDMETHODIMP
GpgolAddin::GetTypeInfo (UINT iTypeInfo, LCID lcid,
                                  LPTYPEINFO *r_typeinfo)
{
  (void)iTypeInfo;
  (void)lcid;
  (void)r_typeinfo;
  TRACEPOINT; /* Should not happen */
  return S_OK;
}

STDMETHODIMP
GpgolAddin::GetIDsOfNames (REFIID riid, LPOLESTR *rgszNames,
                                    UINT cNames, LCID lcid,
                                    DISPID *rgDispId)
{
  (void)riid;
  (void)rgszNames;
  (void)cNames;
  (void)lcid;
  (void)rgDispId;
  TRACEPOINT; /* Should not happen */
  return E_NOINTERFACE;
}

STDMETHODIMP
GpgolAddin::Invoke (DISPID dispid, REFIID riid, LCID lcid,
                    WORD flags, DISPPARAMS *parms, VARIANT *result,
                    EXCEPINFO *exepinfo, UINT *argerr)
{
  USE_INVOKE_ARGS
  TRACEPOINT; /* Should not happen */
  return DISP_E_MEMBERNOTFOUND;
}



/* Definition of GpgolRibbonExtender */

GpgolRibbonExtender::GpgolRibbonExtender (void) : m_lRef(0)
{
}

GpgolRibbonExtender::~GpgolRibbonExtender (void)
{
  log_debug ("%s:%s: cleaning up GpgolRibbonExtender object;",
             SRCNAME, __func__);
  log_debug ("%s:%s: Object deleted\n", SRCNAME, __func__);
}

STDMETHODIMP
GpgolRibbonExtender::QueryInterface(REFIID riid, LPVOID* ppvObj)
{
  HRESULT hr = S_OK;

  *ppvObj = NULL;

  if ((riid == IID_IUnknown) || (riid == IID_IRibbonExtensibility) ||
      (riid == IID_IDispatch))
    {
      *ppvObj = (LPUNKNOWN) this;
    }
  else
    {
      LPOLESTR sRiid = NULL;
      StringFromIID (riid, &sRiid);
      log_debug ("%s:%s: queried for unknown interface: %S",
                 SRCNAME, __func__, sRiid);
    }

  if (*ppvObj)
    ((LPUNKNOWN)*ppvObj)->AddRef();

  return hr;
}

STDMETHODIMP
GpgolRibbonExtender::GetTypeInfoCount (UINT *r_count)
{
  *r_count = 0;
  TRACEPOINT; /* Should not happen */
  return S_OK;
}

STDMETHODIMP
GpgolRibbonExtender::GetTypeInfo (UINT iTypeInfo, LCID lcid,
                                  LPTYPEINFO *r_typeinfo)
{
  (void)iTypeInfo;
  (void)lcid;
  (void)r_typeinfo;
  TRACEPOINT; /* Should not happen */
  return S_OK;
}

/* Good documentation of what this function is supposed to do can
   be found at: http://msdn.microsoft.com/en-us/library/cc237568.aspx

   There is also a very good blog explaining how Ribbon Extensibility
   is supposed to work.
   http://blogs.msdn.com/b/andreww/archive/2007/03/09/
why-is-it-so-hard-to-shim-iribbonextensibility.aspx
   */

#define ID_MAPPER(name,id)                      \
  if (!wcscmp (rgszNames[i], name))             \
    {                                           \
      found = true;                             \
      rgDispId[i] = id;                         \
      break;                                    \
    }                                           \


STDMETHODIMP
GpgolRibbonExtender::GetIDsOfNames (REFIID riid, LPOLESTR *rgszNames,
                                    UINT cNames, LCID lcid,
                                    DISPID *rgDispId)
{
  (void)riid;
  (void)lcid;
  bool found = false;

  if (!rgszNames || !cNames || !rgDispId)
    {
      return E_POINTER;
    }

  for (unsigned int i = 0; i < cNames; i++)
    {
      log_debug ("%s:%s: GetIDsOfNames for: %S",
                 SRCNAME, __func__, rgszNames[i]);
      /* How this is supposed to work with cNames > 1 is unknown,
         but we can just say that we won't support callbacks with
         different parameters and just match the name (the first element)
         and we give it one of our own dispIds's that are later handled in
         the invoke part */
      ID_MAPPER (L"btnDecrypt", ID_BTN_DECRYPT)
      ID_MAPPER (L"btnDecryptLarge", ID_BTN_DECRYPT_LARGE)
      ID_MAPPER (L"btnEncrypt", ID_BTN_ENCRYPT)
      ID_MAPPER (L"btnEncryptLarge", ID_BTN_ENCRYPT_LARGE)
      ID_MAPPER (L"btnEncryptSmall", IDI_ENCRYPT_20_PNG)
      ID_MAPPER (L"btnSignSmall", IDI_SIGN_20_PNG)
      ID_MAPPER (L"btnSignEncryptLarge", IDI_SIGN_ENCRYPT_40_PNG)
      ID_MAPPER (L"btnEncryptFileLarge", ID_BTN_ENCSIGN_LARGE)
      ID_MAPPER (L"btnSignLarge", ID_BTN_SIGN_LARGE)
      ID_MAPPER (L"btnVerifyLarge", ID_BTN_VERIFY_LARGE)
      ID_MAPPER (L"btnSigstateLarge", ID_BTN_SIGSTATE_LARGE)

      /* MIME support: */
      ID_MAPPER (L"encryptMime", ID_CMD_MIME_ENCRYPT)
      ID_MAPPER (L"encryptMimeEx", ID_CMD_MIME_ENCRYPT_EX)
      ID_MAPPER (L"signMime", ID_CMD_MIME_SIGN)
      ID_MAPPER (L"signMimeEx", ID_CMD_MIME_SIGN_EX)
      ID_MAPPER (L"encryptSignMime", ID_CMD_SIGN_ENCRYPT_MIME)
      ID_MAPPER (L"encryptSignMimeEx", ID_CMD_SIGN_ENCRYPT_MIME_EX)
      ID_MAPPER (L"getEncryptPressed", ID_GET_ENCRYPT_PRESSED)
      ID_MAPPER (L"getEncryptPressedEx", ID_GET_ENCRYPT_PRESSED_EX)
      ID_MAPPER (L"getSignPressed", ID_GET_SIGN_PRESSED)
      ID_MAPPER (L"getSignPressedEx", ID_GET_SIGN_PRESSED_EX)
      ID_MAPPER (L"getSignEncryptPressed", ID_GET_SIGN_ENCRYPT_PRESSED)
      ID_MAPPER (L"getSignEncryptPressedEx", ID_GET_SIGN_ENCRYPT_PRESSED_EX)
      ID_MAPPER (L"ribbonLoaded", ID_ON_LOAD)
      ID_MAPPER (L"openOptions", ID_CMD_OPEN_OPTIONS)
      ID_MAPPER (L"getSigLabel", ID_GET_SIG_LABEL)
      ID_MAPPER (L"getSigSTip", ID_GET_SIG_STIP)
      ID_MAPPER (L"getSigTip", ID_GET_SIG_TTIP)
      ID_MAPPER (L"launchDetails", ID_LAUNCH_CERT_DETAILS)
      ID_MAPPER (L"getIsDetailsEnabled", ID_GET_IS_DETAILS_ENABLED)
      ID_MAPPER (L"getIsCrypto", ID_GET_IS_CRYPTO_MAIL)
      ID_MAPPER (L"printDecrypted", ID_CMD_PRINT_DECRYPTED)
    }

  if (cNames > 1)
    {
      log_debug ("More then one name provided. Should not happen");
    }

  return found ? S_OK : E_NOINTERFACE;
}

STDMETHODIMP
GpgolRibbonExtender::Invoke (DISPID dispid, REFIID riid, LCID lcid,
                             WORD flags, DISPPARAMS *parms, VARIANT *result,
                             EXCEPINFO *exepinfo, UINT *argerr)
{
  USE_INVOKE_ARGS
  log_debug ("%s:%s: enter with dispid: %x",
             SRCNAME, __func__, (int)dispid);

  if (!(flags & DISPATCH_METHOD))
    {
      log_debug ("%s:%s: not called in method mode. Bailing out.",
                 SRCNAME, __func__);
      return DISP_E_MEMBERNOTFOUND;
    }

  switch (dispid)
    {
      case ID_CMD_SIGN_ENCRYPT_MIME:
        return mark_mime_action (parms->rgvarg[1].pdispVal,
                                 OP_SIGN|OP_ENCRYPT, false);
      case ID_CMD_SIGN_ENCRYPT_MIME_EX:
        return mark_mime_action (parms->rgvarg[1].pdispVal,
                                 OP_SIGN|OP_ENCRYPT, true);
      case ID_CMD_MIME_ENCRYPT:
        return mark_mime_action (parms->rgvarg[1].pdispVal, OP_ENCRYPT,

                                 false);
      case ID_CMD_MIME_SIGN:
        return mark_mime_action (parms->rgvarg[1].pdispVal, OP_SIGN,
                                 false);
      case ID_GET_ENCRYPT_PRESSED:
        return get_crypt_pressed (parms->rgvarg[0].pdispVal, OP_ENCRYPT,
                                  result, false);
      case ID_GET_SIGN_PRESSED:
        return get_crypt_pressed (parms->rgvarg[0].pdispVal, OP_SIGN,
                                  result, false);
      case ID_GET_SIGN_ENCRYPT_PRESSED:
        return get_crypt_pressed (parms->rgvarg[0].pdispVal,
                                  OP_SIGN | OP_ENCRYPT,
                                  result, false);
      case ID_CMD_MIME_SIGN_EX:
        return mark_mime_action (parms->rgvarg[1].pdispVal, OP_SIGN, true);
      case ID_CMD_MIME_ENCRYPT_EX:
        return mark_mime_action (parms->rgvarg[1].pdispVal, OP_ENCRYPT, true);
      case ID_GET_ENCRYPT_PRESSED_EX:
        return get_crypt_pressed (parms->rgvarg[0].pdispVal, OP_ENCRYPT,
                                  result, true);
      case ID_GET_SIGN_PRESSED_EX:
        return get_crypt_pressed (parms->rgvarg[0].pdispVal, OP_SIGN,
                                  result, true);
      case ID_GET_SIGN_ENCRYPT_PRESSED_EX:
        return get_crypt_pressed (parms->rgvarg[0].pdispVal, OP_SIGN | OP_ENCRYPT,
                                  result, true);
      case ID_GET_SIG_STIP:
        return get_sig_stip (parms->rgvarg[0].pdispVal, result);
      case ID_GET_SIG_TTIP:
        return get_sig_ttip (parms->rgvarg[0].pdispVal, result);
      case ID_GET_SIG_LABEL:
        return get_sig_label (parms->rgvarg[0].pdispVal, result);
      case ID_LAUNCH_CERT_DETAILS:
        return launch_cert_details (parms->rgvarg[0].pdispVal);
      case ID_GET_IS_DETAILS_ENABLED:
        return get_is_details_enabled (parms->rgvarg[0].pdispVal, result);

      case ID_ON_LOAD:
          {
            g_ribbon_uis.push_back (parms->rgvarg[0].pdispVal);
            return S_OK;
          }
      case ID_CMD_OPEN_OPTIONS:
          {
            options_dialog_box (NULL);
            return S_OK;
          }
      case ID_CMD_PRINT_DECRYPTED:
        return print_decrypted (parms->rgvarg[0].pdispVal);
      case ID_GET_IS_CRYPTO_MAIL:
        return get_is_crypto_mail (parms->rgvarg[0].pdispVal, result);
      case ID_BTN_ENCRYPT:
      case ID_BTN_DECRYPT:
      case ID_BTN_DECRYPT_LARGE:
      case ID_BTN_ENCRYPT_LARGE:
      case ID_BTN_ENCSIGN_LARGE:
      case ID_BTN_SIGN_LARGE:
      case ID_BTN_VERIFY_LARGE:
      case IDI_SIGN_ENCRYPT_40_PNG:
      case IDI_ENCRYPT_20_PNG:
      case IDI_SIGN_20_PNG:
        return getIcon (dispid, result);
      case ID_BTN_SIGSTATE_LARGE:
        return get_crypto_icon (parms->rgvarg[0].pdispVal, result);
    }

  log_debug ("%s:%s: leave", SRCNAME, __func__);

  return DISP_E_MEMBERNOTFOUND;
}

/* Returns the XML markup for the various RibbonID's

   The custom ui syntax is documented at:
   http://msdn.microsoft.com/en-us/library/dd926139%28v=office.12%29.aspx

   The outlook specific elements are documented at:
   http://msdn.microsoft.com/en-us/library/office/ee692172%28v=office.14%29.aspx
*/
static STDMETHODIMP
GetCustomUI_MIME (BSTR RibbonID, BSTR * RibbonXml)
{
  char * buffer = NULL;

/*  const char *certManagerTTip =
    _("Start the Certificate Management Software");
  const char *certManagerSTip =
    _("Open GPA or Kleopatra to manage your certificates. "
      "You can use this you to generate your "
      "own certificates. ");*/
  const char *encryptTTip =
    _("Encrypt the message");
  const char *encryptSTip =
    _("Encrypts the message and all attachments before sending");
  const char *signTTip =
    _("Sign the message");
  const char *signSTip =
    _("Sign the message and all attachments before sending");

  const char *secureTTip =
    _("Sign and encrypt the message");
  const char *secureSTip =
    _("Encrypting and cryptographically signing a message means that the "
      "recipients can be sure that no one modified the message and only the "
      "recipients can read it");
  const char *optsSTip =
    _("Open the settings dialog for GpgOL");
  log_debug ("%s:%s: GetCustomUI_MIME for id: %ls", SRCNAME, __func__, RibbonID);

  if (!RibbonXml || !RibbonID)
    return E_POINTER;

  if (!wcscmp (RibbonID, L"Microsoft.Outlook.Mail.Compose"))
    {
      gpgrt_asprintf (&buffer,
        "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\""
        " onLoad=\"ribbonLoaded\">"
        " <ribbon>"
        "   <tabs>"
        "    <tab idMso=\"TabNewMailMessage\">"
        "     <group id=\"general\""
        "            label=\"%s\">"
        "       <splitButton id=\"splitty\" size=\"large\">"
        "         <toggleButton id=\"mimeSignEncrypt\""
        "                 label=\"%s\""
        "                 screentip=\"%s\""
        "                 supertip=\"%s\""
        "                 getPressed=\"getSignEncryptPressed\""
        "                 getImage=\"btnSignEncryptLarge\""
        "                 onAction=\"encryptSignMime\"/>"
        "         <menu id=\"encMenu\" showLabel=\"true\">"
        "         <toggleButton id=\"mimeSign\""
        "                 getImage=\"btnSignSmall\""
        "                 label=\"%s\""
        "                 screentip=\"%s\""
        "                 supertip=\"%s\""
        "                 onAction=\"signMime\""
        "                 getPressed=\"getSignPressed\"/>"
        "         <toggleButton id=\"mimeEncrypt\""
        "                 getImage=\"btnEncryptSmall\""
        "                 label=\"%s\""
        "                 screentip=\"%s\""
        "                 supertip=\"%s\""
        "                 onAction=\"encryptMime\""
        "                 getPressed=\"getEncryptPressed\"/>"
        "       </menu></splitButton>"
        "       <dialogBoxLauncher>"
        "         <button id=\"optsBtn\""
        "                 onAction=\"openOptions\""
        "                 screentip=\"%s\"/>"
        "       </dialogBoxLauncher>"
        "     </group>"
        "    </tab>"
        "   </tabs>"
        " </ribbon>"
        "</customUI>", _("GpgOL"),
        _("Secure"), secureTTip, secureSTip,
        _("Sign"), signTTip, signSTip,
        _("Encrypt"), encryptTTip, encryptSTip,
        optsSTip
        );
    }
  else if (!wcscmp (RibbonID, L"Microsoft.Outlook.Mail.Read"))
    {
      gpgrt_asprintf (&buffer,
        "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\""
        " onLoad=\"ribbonLoaded\">"
        " <ribbon>"
        "   <tabs>"
        "    <tab idMso=\"TabReadMessage\">"
        "     <group id=\"general\""
        "            label=\"%s\">"
        "       <button id=\"idSigned\""
        "               getImage=\"btnSigstateLarge\""
        "               size=\"large\""
        "               getLabel=\"getSigLabel\""
        "               getScreentip=\"getSigTip\""
        "               getSupertip=\"getSigSTip\""
        "               onAction=\"launchDetails\"/>"
       /* "               getEnabled=\"getIsCrypto\"/>" */
        "       <dialogBoxLauncher>"
        "         <button id=\"optsBtn\""
        "                 onAction=\"openOptions\""
        "                 screentip=\"%s\"/>"
        "       </dialogBoxLauncher>"
        "     </group>"
        "    </tab>"
        "   </tabs>"
        " </ribbon>"
        "</customUI>",
        _("GpgOL"),
        optsSTip
        );
    }
  /* We don't use this code currently because calling the send
     event for Inline Response mailitems fails. */
  else if (!wcscmp (RibbonID, L"Microsoft.Outlook.Explorer") && g_ol_version_major > 14)
    {
      gpgrt_asprintf (&buffer,
        "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\""
        " onLoad=\"ribbonLoaded\">"
        " <ribbon>"
        "   <tabs>"
        "    <tab idMso=\"TabMail\">"
        "     <group id=\"general_read\""
        "            label=\"%s\">"
        "       <button id=\"idSigned\""
        "               getImage=\"btnSigstateLarge\""
        "               size=\"large\""
        "               getLabel=\"getSigLabel\""
        "               getScreentip=\"getSigTip\""
        "               getSupertip=\"getSigSTip\""
        "               onAction=\"launchDetails\""
        "               getEnabled=\"getIsDetailsEnabled\"/>"
        "       <dialogBoxLauncher>"
        "         <button id=\"optsBtn_read\""
        "                 onAction=\"openOptions\""
        "                 screentip=\"%s\"/>"
        "       </dialogBoxLauncher>"
        "     </group>"
        "    </tab>"
        "   </tabs>"
        "   <contextualTabs>"
        "   <tabSet idMso=\"TabComposeTools\">"
        "    <tab idMso=\"TabMessage\">"
        "     <group id=\"general\""
        "            label=\"%s\">"
        "       <splitButton id=\"splitty\" size=\"large\">"
        "         <toggleButton id=\"mimeSignEncrypt\""
        "                 label=\"%s\""
        "                 screentip=\"%s\""
        "                 supertip=\"%s\""
        "                 getPressed=\"getSignEncryptPressedEx\""
        "                 getImage=\"btnSignEncryptLarge\""
        "                 onAction=\"encryptSignMimeEx\"""/>"
        "         <menu id=\"encMenu\" showLabel=\"true\">"
        "         <toggleButton id=\"mimeSign\""
        "                 getImage=\"btnSignSmall\""
        "                 label=\"%s\""
        "                 screentip=\"%s\""
        "                 supertip=\"%s\""
        "                 onAction=\"signMimeEx\""
        "                 getPressed=\"getSignPressedEx\"/>"
        "         <toggleButton id=\"mimeEncrypt\""
        "                 getImage=\"btnEncryptSmall\""
        "                 label=\"%s\""
        "                 screentip=\"%s\""
        "                 supertip=\"%s\""
        "                 onAction=\"encryptMimeEx\""
        "                 getPressed=\"getEncryptPressedEx\"/>"
        "       </menu></splitButton>"
        "       <dialogBoxLauncher>"
        "         <button id=\"optsBtn\""
        "                 onAction=\"openOptions\""
        "                 screentip=\"%s\"/>"
        "       </dialogBoxLauncher>"
        "      </group>"
        "    </tab>"
        "   </tabSet>"
        "  </contextualTabs>"
        " </ribbon>"
        " <contextMenus>"
        "  <contextMenu idMso=\"ContextMenuMailItem\">"
        "   <button id=\"decryptedPrintBtn\""
//        "idMso=\"FilePrintQuick\""
        "           label=\"%s\""
        "           onAction=\"printDecrypted\""
        "           getImage=\"btnEncryptSmall\""
        "           getVisible=\"getIsCrypto\""
        "           insertAfterMso=\"FilePrintQuick\""
        "   />"
        "  </contextMenu>"
        " </contextMenus>"
        "</customUI>",
        _("GpgOL"),
        optsSTip,
        _("GpgOL"),
        _("Secure"), secureTTip, secureSTip,
        _("Sign"), signTTip, signSTip,
        _("Encrypt"), encryptTTip, encryptSTip,
        optsSTip,
        _("&amp;Print decrypted")
        );
    }
  else if (!wcscmp (RibbonID, L"Microsoft.Outlook.Explorer"))
    {
      // No TabComposeTools in Outlook 2010
      gpgrt_asprintf (&buffer,
        "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\""
        " onLoad=\"ribbonLoaded\">"
        " <ribbon>"
        "   <tabs>"
        "    <tab idMso=\"TabMail\">"
        "     <group id=\"general_read\""
        "            label=\"%s\">"
        "       <button id=\"idSigned\""
        "               getImage=\"btnSigstateLarge\""
        "               size=\"large\""
        "               getLabel=\"getSigLabel\""
        "               getScreentip=\"getSigTip\""
        "               getSupertip=\"getSigSTip\""
        "               onAction=\"launchDetails\""
        "               getEnabled=\"getIsDetailsEnabled\"/>"
        "       <dialogBoxLauncher>"
        "         <button id=\"optsBtn_read\""
        "                 onAction=\"openOptions\""
        "                 screentip=\"%s\"/>"
        "       </dialogBoxLauncher>"
        "     </group>"
        "    </tab>"
        "   </tabs>"
        " </ribbon>"
        "</customUI>",
        _("GpgOL"),
        optsSTip
        );
    }

  if (buffer)
    {
      wchar_t *wbuf = utf8_to_wchar2 (buffer, strlen(buffer));
      xfree (buffer);
      *RibbonXml = SysAllocString (wbuf);
      xfree (wbuf);
    }
  else
    *RibbonXml = NULL;

  return S_OK;
}

STDMETHODIMP
GpgolRibbonExtender::GetCustomUI (BSTR RibbonID, BSTR * RibbonXml)
{
  return GetCustomUI_MIME (RibbonID, RibbonXml);
}

/* RibbonUi elements are created on demand but they are reused
   in different inspectors. So far and from all documentation
   I could find RibbonUi elments are never
   deleted. When they are created the onLoad callback is called
   to register them.
   The callbacks registered in the XML description are only
   executed on Load. So to have different information depending
   on the available mails we have to invalidate the UI ourself.
   This means that the callbacks will be reevaluated and the UI
   Updated. Sadly we don't know which ribbon_ui needs updates
   so we have to invalidate everything.
*/
void gpgoladdin_invalidate_ui ()
{
  std::list<LPDISPATCH>::iterator it;

  for (it = g_ribbon_uis.begin(); it != g_ribbon_uis.end(); ++it)
    {
      log_debug ("%s:%s: Invalidating ribbon: %p",
                 SRCNAME, __func__, *it);
      invoke_oom_method (*it, "Invalidate", NULL);
    }
}

GpgolAddin *
GpgolAddin::get_instance ()
{
  if (!addin_instance)
    {
      addin_instance = new GpgolAddin ();
    }
  return addin_instance;
}
