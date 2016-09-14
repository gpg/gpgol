/* gpgoladdin.cpp - Connect GpgOL to Outlook as an addin
 *    Copyright (C) 2013, 2015 Intevation GmbH
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
#include "myexchext.h"

#include "display.h"
#include "msgcache.h"
#include "engine.h"
#include "engine-assuan.h"
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

/* Invalidating the interface does not take a nice effect so we store
   this option in a global variable. */
bool use_mime_ui = false;

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
GpgolAddin::GpgolAddin (void) : m_lRef(0), m_application(0),
  m_addin(0), m_applicationEventSink(0), m_disabled(false)
{
  read_options ();
  use_mime_ui = opt.mime_ui;
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
  gpgol_release (m_applicationEventSink);

  engine_deinit ();
  write_options ();

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

STDMETHODIMP
GpgolAddin::OnConnection (LPDISPATCH Application, ext_ConnectMode ConnectMode,
                          LPDISPATCH AddInInst, SAFEARRAY ** custom)
{
  (void)custom;
  char* version;

  log_debug ("%s:%s: this is GpgOL %s\n",
             SRCNAME, __func__, PACKAGE_VERSION);
  log_debug ("%s:%s:   in Outlook %s\n",
             SRCNAME, __func__, gpgme_check_version (NULL));

  can_unload = false;
  m_application = Application;
  m_application->AddRef();
  m_addin = AddInInst;

  version = get_oom_string (Application, "Version");

  log_debug ("%s:%s:   using GPGME %s\n",
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
  engine_init ();

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

  if (Mail::revert_all_mails ())
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

STDMETHODIMP
GpgolAddin::OnStartupComplete (SAFEARRAY** custom)
{
  (void)custom;
  TRACEPOINT;

  install_forms ();

  if (!create_responder_window())
    {
      log_error ("%s:%s: Failed to create the responder window;",
                 SRCNAME, __func__);
    }

  if (m_application)
    {
      m_applicationEventSink = install_ApplicationEvents_sink(m_application);
      return S_OK;
    }
  /* Should not happen as OnConnection should be called before */
  log_error ("%s:%s: no application set;",
             SRCNAME, __func__);
  return E_NOINTERFACE;
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
      ID_MAPPER (L"attachmentDecryptCallback", ID_CMD_DECRYPT)
      ID_MAPPER (L"encryptSelection", ID_CMD_ENCRYPT_SELECTION)
      ID_MAPPER (L"decryptSelection", ID_CMD_DECRYPT_SELECTION)
      ID_MAPPER (L"startCertManager", ID_CMD_CERT_MANAGER)
      ID_MAPPER (L"btnCertManager", ID_BTN_CERTMANAGER)
      ID_MAPPER (L"btnDecrypt", ID_BTN_DECRYPT)
      ID_MAPPER (L"btnDecryptLarge", ID_BTN_DECRYPT_LARGE)
      ID_MAPPER (L"btnEncrypt", ID_BTN_ENCRYPT)
      ID_MAPPER (L"btnEncryptLarge", ID_BTN_ENCRYPT_LARGE)
      ID_MAPPER (L"btnEncryptFileLarge", ID_BTN_ENCSIGN_LARGE)
      ID_MAPPER (L"btnSignLarge", ID_BTN_SIGN_LARGE)
      ID_MAPPER (L"btnVerifyLarge", ID_BTN_VERIFY_LARGE)
      ID_MAPPER (L"encryptBody", ID_CMD_ENCRYPT_BODY)
      ID_MAPPER (L"decryptBody", ID_CMD_DECRYPT_BODY)
      ID_MAPPER (L"addEncSignedAttachment", ID_CMD_ATT_ENCSIGN_FILE)
      ID_MAPPER (L"addEncAttachment", ID_CMD_ATT_ENC_FILE)
      ID_MAPPER (L"signBody", ID_CMD_SIGN_BODY)
      ID_MAPPER (L"verifyBody", ID_CMD_VERIFY_BODY)

      /* MIME support: */
      ID_MAPPER (L"encryptMime", ID_CMD_MIME_ENCRYPT)
      ID_MAPPER (L"signMime", ID_CMD_MIME_SIGN)
      ID_MAPPER (L"getEncryptPressed", ID_GET_ENCRYPT_PRESSED)
      ID_MAPPER (L"getSignPressed", ID_GET_SIGN_PRESSED)
      ID_MAPPER (L"encryptMimeEx", ID_CMD_MIME_ENCRYPT_EX)
      ID_MAPPER (L"signMimeEx", ID_CMD_MIME_SIGN_EX)
      ID_MAPPER (L"getEncryptPressedEx", ID_GET_ENCRYPT_PRESSED_EX)
      ID_MAPPER (L"getSignPressedEx", ID_GET_SIGN_PRESSED_EX)
      ID_MAPPER (L"ribbonLoaded", ID_ON_LOAD);
      ID_MAPPER (L"openOptions", ID_CMD_OPEN_OPTIONS)
      ID_MAPPER (L"getSigStatus", ID_GET_SIG_STATUS)
      ID_MAPPER (L"getEncStatus", ID_GET_ENC_STATUS)
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
      case ID_CMD_DECRYPT:
        /* We can assume that this points to an implementation of
           IRibbonControl as we know the callback dispid. */
        return decryptAttachments (parms->rgvarg[0].pdispVal);
      case ID_CMD_ENCRYPT_SELECTION:
        return encryptSelection (parms->rgvarg[0].pdispVal);
      case ID_CMD_DECRYPT_SELECTION:
        return decryptSelection (parms->rgvarg[0].pdispVal);
      case ID_CMD_CERT_MANAGER:
        return startCertManager (parms->rgvarg[0].pdispVal);
      case ID_CMD_ENCRYPT_BODY:
        return encryptBody (parms->rgvarg[0].pdispVal);
      case ID_CMD_DECRYPT_BODY:
        return decryptBody (parms->rgvarg[0].pdispVal);
      case ID_CMD_ATT_ENCSIGN_FILE:
        return addEncSignedAttachment (parms->rgvarg[0].pdispVal);
      case ID_CMD_ATT_ENC_FILE:
        return addEncAttachment (parms->rgvarg[0].pdispVal);
      case ID_CMD_SIGN_BODY:
        return signBody (parms->rgvarg[0].pdispVal);
      case ID_CMD_VERIFY_BODY:
        return verifyBody (parms->rgvarg[0].pdispVal);
      case ID_CMD_MIME_SIGN:
        return mark_mime_action (parms->rgvarg[1].pdispVal, OP_SIGN, false);
      case ID_CMD_MIME_ENCRYPT:
        return mark_mime_action (parms->rgvarg[1].pdispVal, OP_ENCRYPT,
                                 false);
      case ID_GET_ENCRYPT_PRESSED:
        return get_crypt_pressed (parms->rgvarg[0].pdispVal, OP_ENCRYPT,
                                  result, false);
      case ID_GET_SIGN_PRESSED:
        return get_crypt_pressed (parms->rgvarg[0].pdispVal, OP_SIGN,
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
      case ID_GET_ENC_STATUS:
        return get_crypt_status (parms->rgvarg[0].pdispVal, OP_ENCRYPT,
                                 result);
      case ID_GET_SIG_STATUS:
        return get_crypt_status (parms->rgvarg[0].pdispVal, OP_SIGN, result);
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
      case ID_BTN_CERTMANAGER:
      case ID_BTN_ENCRYPT:
      case ID_BTN_DECRYPT:
      case ID_BTN_DECRYPT_LARGE:
      case ID_BTN_ENCRYPT_LARGE:
      case ID_BTN_ENCSIGN_LARGE:
      case ID_BTN_SIGN_LARGE:
      case ID_BTN_VERIFY_LARGE:
        return getIcon (dispid, result);
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
    _("Encrypt the message.");
  const char *encryptSTip =
    _("Encrypts the message and all attachments before sending.");
  const char *signTTip =
    _("Sign the message.");
  const char *signSTip =
    _("Sign the message and all attachments before sending.");
  const char *optsSTip =
    _("Open the settings dialog for GpgOL.");
#if 0
  const char *encryptedTTip =
    "If this is toggled the message was encrypted. (this is development UI)";
  const char *encryptedSTip =
    "TODO insert more details here";
  const char *signedTTip =
    "If this is toggled the message was signed. (this is development UI)";
  const char *signedSTip =
    "TODO insert more details here";
#endif
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
        "       <toggleButton id=\"mimeEncrypt\""
        "               getImage=\"btnEncryptLarge\""
        "               size=\"large\""
        "               label=\"%s\""
        "               screentip=\"%s\""
        "               supertip=\"%s\""
        "               onAction=\"encryptMime\""
        "               getPressed=\"getEncryptPressed\"/>"
        "       <toggleButton id=\"mimeSign\""
        "               getImage=\"btnSignLarge\""
        "               size=\"large\""
        "               label=\"%s\""
        "               screentip=\"%s\""
        "               supertip=\"%s\""
        "               onAction=\"signMime\""
        "               getPressed=\"getSignPressed\"/>"
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
        _("Encrypt"), encryptTTip, encryptSTip,
        _("Sign"), signTTip, signSTip,
        optsSTip
        );
    }
  else if (!wcscmp (RibbonID, L"Microsoft.Outlook.Mail.Read"))
    {
      gpgrt_asprintf (&buffer,
        "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\""
        " onLoad=\"ribbonLoaded\">"
#if 0
        " <ribbon>"
        "   <tabs>"
        "    <tab idMso=\"TabReadMessage\">"
        "     <group id=\"general\""
        "            label=\"%s\">"
        "       <toggleButton id=\"idEncrypted\""
        "               getImage=\"btnCryptStatus\""
        "               size=\"large\""
        "               label=\"%s\""
        "               screentip=\"%s\""
        "               supertip=\"%s\""
        "               getEnabled=\"getEncStatus\"/>"
        "       <button id=\"idSigned\""
        "               getImage=\"btnSignLarge\""
        "               size=\"large\""
        "               label=\"%s\""
        "               screentip=\"%s\""
        "               supertip=\"%s\""
        "               onAction=\"verifyBody\""
        "               getEnabled=\"getSigStatus\"/>"
        "       <dialogBoxLauncher>"
        "         <button id=\"optsBtn\""
        "                 onAction=\"openOptions\""
        "                 screentip=\"%s\"/>"
        "       </dialogBoxLauncher>"
        "     </group>"
        "    </tab>"
        "   </tabs>"
        " </ribbon>"
#endif
        "<contextMenus>"
        " <contextMenu idMso=\"ContextMenuAttachments\">"
        "   <button id=\"gpgol_decrypt\""
        "           label=\"%s\""
        "           getImage=\"btnDecrypt\""
        "           onAction=\"attachmentDecryptCallback\"/>"
        " </contextMenu>"
        "</contextMenus>"
        "</customUI>",
#if 0
        _("GpgOL"),
        _("Encrypted"), encryptedTTip, encryptedSTip,
        _("Signed"), signedTTip, signedSTip,
        optsSTip,
#endif
        _("Decrypt")
        );
    }
#if 0
  /* We don't use this code currently because calling the send
     event for Inline Response mailitems fails. */
  else if (!wcscmp (RibbonID, L"Microsoft.Outlook.Explorer"))
    {
      gpgrt_asprintf (&buffer,
        "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\""
        " onLoad=\"ribbonLoaded\">"
        " <ribbon>"
        "   <contextualTabs>"
        "   <tabSet idMso=\"TabComposeTools\">"
        "    <tab idMso=\"TabMessage\">"
        "     <group id=\"general\""
        "            label=\"%s\">"
        "       <toggleButton id=\"mimeEncryptEx\""
        "               getImage=\"btnEncryptLarge\""
        "               size=\"large\""
        "               label=\"%s\""
        "               screentip=\"%s\""
        "               supertip=\"%s\""
        "               onAction=\"encryptMimeEx\""
        "               getPressed=\"getEncryptPressedEx\"/>"
        "       <toggleButton id=\"mimeSignEx\""
        "               getImage=\"btnSignLarge\""
        "               size=\"large\""
        "               label=\"%s\""
        "               screentip=\"%s\""
        "               supertip=\"%s\""
        "               onAction=\"signMimeEx\""
        "               getPressed=\"getSignPressedEx\"/>"
        "       <dialogBoxLauncher>"
        "         <button id=\"optsBtn\""
        "                 onAction=\"openOptions\""
        "                 screentip=\"%s\"/>"
        "       </dialogBoxLauncher>"
        "     </group>"
        "    </tab>"
        "   </tabSet>"
        "   </contextualTabs>"
        " </ribbon>"
        "</customUI>", _("GpgOL"),
        _("Encrypt"), encryptTTip, encryptSTip,
        _("Sign"), signTTip, signSTip,
        optsSTip
        );
    }
#endif

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

/* This is the old pre-mime adding UI code. It will be removed once we have a
   stable version that can also send mime messages.
*/
static STDMETHODIMP
GetCustomUI_old (BSTR RibbonID, BSTR * RibbonXml)
{
  char *buffer = NULL;
  const char *certManagerTTip =
    _("Start the Certificate Management Software");
  const char *certManagerSTip =
    _("Open GPA or Kleopatra to manage your certificates. "
      "You can use this you to generate your "
      "own certificates. ");
  const char *encryptTextTTip =
    _("Encrypt the text of the message");
  const char *encryptTextSTip =
    _("Choose the certificates for which the message "
      "should be encrypted and replace the text "
      "with the encrypted message.");
  const char *encryptFileTTip =
    _("Add a file as an encrypted attachment");
  const char *encryptFileSTip =
    _("Encrypts a file and adds it as an attachment to the "
      "message. ");
  const char *encryptSignFileTTip =
    _("Add a file as an encrypted attachment with a signature");
  const char *encryptSignFileSTip =
    _("Encrypts a file, signs it and adds both the encrypted file "
      "and the signature as attachments to the message. ");
  const char *decryptTextTTip=
    _("Decrypt the message");
  const char *decryptTextSTip =
    _("Look for PGP or S/MIME encrypted data in the message text "
      "and decrypt it.");
  const char *signTextTTip =
    _("Add a signature of the message");
  const char *signTextSTip =
    _("Appends a signed copy of the message text in an opaque signature. "
      "An opaque signature ensures that the signed text is not modified by "
      "embedding it in the signature itself. "
      "The combination of the signed message text and your signature is "
      "added below the plain text. "
      "The message will not be encrypted!");
  const char *optsSTip =
    _("Open the settings dialog for GpgOL.");

  log_debug ("%s:%s: GetCustomUI for id: %ls", SRCNAME, __func__, RibbonID);

  if (!RibbonXml)
    return E_POINTER;

  if (!wcscmp (RibbonID, L"Microsoft.Outlook.Mail.Compose"))
    {
      gpgrt_asprintf (&buffer,
        "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\""
        " onLoad=\"ribbonLoaded\">"
        " <ribbon>"
        "   <tabs>"
        "    <tab id=\"gpgolTab\""
        "         label=\"%s\">"
        "     <group id=\"general\""
        "            label=\"%s\">"
        "       <button id=\"CustomButton\""
        "               getImage=\"btnCertManager\""
        "               size=\"large\""
        "               label=\"%s\""
        "               screentip=\"%s\""
        "               supertip=\"%s\""
        "               onAction=\"startCertManager\"/>"
        "       <dialogBoxLauncher>"
        "         <button id=\"optsBtn\""
        "                 onAction=\"openOptions\""
        "                 screentip=\"%s\"/>"
        "       </dialogBoxLauncher>"
        "     </group>"
        "     <group id=\"textGroup\""
        "            label=\"%s\">"
        "       <button id=\"fullTextEncrypt\""
        "               getImage=\"btnEncryptLarge\""
        "               size=\"large\""
        "               label=\"%s\""
        "               screentip=\"%s\""
        "               supertip=\"%s\""
        "               onAction=\"encryptBody\"/>"
        "       <button id=\"fullTextDecrypt\""
        "               getImage=\"btnDecryptLarge\""
        "               size=\"large\""
        "               label=\"%s\""
        "               screentip=\"%s\""
        "               supertip=\"%s\""
        "               onAction=\"decryptBody\"/>"
        "       <button id=\"fullTextSign\""
        "               getImage=\"btnSignLarge\""
        "               size=\"large\""
        "               label=\"%s\""
        "               screentip=\"%s\""
        "               supertip=\"%s\""
        "               onAction=\"signBody\"/>"
        "       <button id=\"fullTextVerify\""
        "               getImage=\"btnVerifyLarge\""
        "               size=\"large\""
        "               label=\"%s\""
        "               onAction=\"verifyBody\"/>"
        "     </group>"
        "     <group id=\"attachmentGroup\""
        "            label=\"%s\">"
        "       <button id=\"encryptedFile\""
        "               getImage=\"btnEncryptLarge\""
        "               size=\"large\""
        "               label=\"%s\""
        "               screentip=\"%s\""
        "               supertip=\"%s\""
        "               onAction=\"addEncAttachment\"/>"
        "       <button id=\"encryptSignFile\""
        "               getImage=\"btnEncryptFileLarge\""
        "               size=\"large\""
        "               label=\"%s\""
        "               screentip=\"%s\""
        "               supertip=\"%s\""
        "               onAction=\"addEncSignedAttachment\"/>"
        "     </group>"
        "    </tab>"
        "   </tabs>"
        " </ribbon>"
        " <contextMenus>"
        "  <contextMenu idMso=\"ContextMenuText\">"
        "    <button id=\"encryptButton\""
        "            label=\"%s\""
        "            getImage=\"btnEncrypt\""
        "            onAction=\"encryptSelection\"/>"
        "    <button id=\"decryptButton\""
        "            label=\"%s\""
        "            getImage=\"btnDecrypt\""
        "            onAction=\"decryptSelection\"/>"
        " </contextMenu>"
        "</contextMenus>"
        "</customUI>", _("GpgOL"), _("General"),
        _("Start Certificate Manager"), certManagerTTip, certManagerSTip,
        optsSTip,
        _("Textbody"),
        _("Encrypt"), encryptTextTTip, encryptTextSTip,
        _("Decrypt"), decryptTextTTip, decryptTextSTip,
        _("Sign"), signTextTTip, signTextSTip,
        _("Verify"),
        _("Attachments"),
        _("Encrypted file"), encryptFileTTip, encryptFileSTip,
        _("Encrypted file and Signature"), encryptSignFileTTip, encryptSignFileSTip,
        _("Encrypt"), _("Decrypt")
        );
    }
  else if (!wcscmp (RibbonID, L"Microsoft.Outlook.Mail.Read"))
    {
      gpgrt_asprintf (&buffer,
        "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">"
        " <ribbon>"
        "   <tabs>"
        "    <tab id=\"gpgolTab\""
        "         label=\"%s\">"
        "     <group id=\"general\""
        "            label=\"%s\">"
        "       <button id=\"CustomButton\""
        "               getImage=\"btnCertManager\""
        "               size=\"large\""
        "               label=\"%s\""
        "               screentip=\"%s\""
        "               supertip=\"%s\""
        "               onAction=\"startCertManager\"/>"
        "       <dialogBoxLauncher>"
        "         <button id=\"optsBtn\""
        "                 onAction=\"openOptions\""
        "                 screentip=\"%s\"/>"
        "       </dialogBoxLauncher>"
        "     </group>"
        "     <group id=\"textGroup\""
        "            label=\"%s\">"
        "       <button id=\"fullTextDecrypt\""
        "               getImage=\"btnDecryptLarge\""
        "               size=\"large\""
        "               label=\"%s\""
        "               screentip=\"%s\""
        "               supertip=\"%s\""
        "               onAction=\"decryptBody\"/>"
        "       <button id=\"fullTextVerify\""
        "               getImage=\"btnVerifyLarge\""
        "               size=\"large\""
        "               label=\"%s\""
        "               onAction=\"verifyBody\"/>"
        "     </group>"
        "    </tab>"
        "   </tabs>"
        "  <contextualTabs>"
        "    <tabSet idMso=\"TabSetAttachments\">"
        "        <tab idMso=\"TabAttachments\">"
        "            <group label=\"%s\" id=\"gnupgLabel\">"
        "                <button id=\"gpgol_contextual_decrypt\""
        "                    size=\"large\""
        "                    label=\"%s\""
        "                    getImage=\"btnDecryptLarge\""
        "                    onAction=\"attachmentDecryptCallback\" />"
        "            </group>"
        "        </tab>"
        "    </tabSet>"
        "  </contextualTabs>"
        " </ribbon>"
        "<contextMenus>"
        "<contextMenu idMso=\"ContextMenuReadOnlyMailText\">"
        "   <button id=\"decryptReadButton\""
        "           label=\"%s\""
        "           getImage=\"btnDecrypt\""
        "           onAction=\"decryptSelection\"/>"
        " </contextMenu>"
        " <contextMenu idMso=\"ContextMenuAttachments\">"
        "   <button id=\"gpgol_decrypt\""
        "           label=\"%s\""
        "           getImage=\"btnDecrypt\""
        "           onAction=\"attachmentDecryptCallback\"/>"
        " </contextMenu>"
        "</contextMenus>"
        "</customUI>",
        _("GpgOL"), _("General"),
        _("Start Certificate Manager"), certManagerTTip, certManagerSTip,
        optsSTip,
        _("Textbody"),
        _("Decrypt"), decryptTextTTip, decryptTextSTip,
        _("Verify"),
        _("GpgOL"), _("Save and decrypt"),
        _("Decrypt"),
        _("Decrypt"));
    }
  else if (!wcscmp (RibbonID, L"Microsoft.Outlook.Explorer"))
    {
      gpgrt_asprintf (&buffer,
        "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">"
        " <ribbon>"
        "   <tabs>"
        "    <tab id=\"gpgolTab\""
        "         label=\"%s\">"
        "     <group id=\"general\""
        "            label=\"%s\">"
        "       <button id=\"CustomButton\""
        "               getImage=\"btnCertManager\""
        "               size=\"large\""
        "               label=\"%s\""
        "               screentip=\"%s\""
        "               supertip=\"%s\""
        "               onAction=\"startCertManager\"/>"
        "       <dialogBoxLauncher>"
        "         <button id=\"optsBtn\""
        "                 onAction=\"openOptions\""
        "                 screentip=\"%s\"/>"
        "       </dialogBoxLauncher>"
        "     </group>"
        /* This would be totally nice but Outlook
           saves the decrypted text aftewards automatically.
           Yay,..
        "     <group id=\"textGroup\""
        "            label=\"%s\">"
        "       <button id=\"fullTextDecrypt\""
        "               getImage=\"btnDecryptLarge\""
        "               size=\"large\""
        "               label=\"%s\""
        "               onAction=\"decryptBody\"/>"
        "     </group>"
        */
        "    </tab>"
        "   </tabs>"
        "  <contextualTabs>"
        "    <tabSet idMso=\"TabSetAttachments\">"
        "        <tab idMso=\"TabAttachments\">"
        "            <group label=\"%s\" id=\"gnupgLabel\">"
        "                <button id=\"gpgol_contextual_decrypt\""
        "                    size=\"large\""
        "                    label=\"%s\""
        "                    getImage=\"btnDecryptLarge\""
        "                    onAction=\"attachmentDecryptCallback\" />"
        "            </group>"
        "        </tab>"
        "    </tabSet>"
        "  </contextualTabs>"
        " </ribbon>"
        " <contextMenus>"
        /*
           There appears to be no way to access the word editor
           / get the selected text from that Context.
        " <contextMenu idMso=\"ContextMenuReadOnlyMailText\">"
        " <button id=\"decryptReadButton1\""
        "         label=\"%s\""
        "         onAction=\"decryptSelection\"/>"
        " </contextMenu>"
        */
        " <contextMenu idMso=\"ContextMenuAttachments\">"
        "   <button id=\"gpgol_decrypt\""
        "           label=\"%s\""
        "           getImage=\"btnDecrypt\""
        "           onAction=\"attachmentDecryptCallback\"/>"
        " </contextMenu>"
        " </contextMenus>"
        "</customUI>",
        _("GpgOL"), _("General"),
        _("Start Certificate Manager"), certManagerTTip, certManagerSTip,
        optsSTip,
        /*_("Mail Body"), _("Decrypt"),*/
        _("GpgOL"), _("Save and decrypt"),/*_("Decrypt"), */
        _("Save and decrypt"));
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
  if (use_mime_ui)
    {
      return GetCustomUI_MIME (RibbonID, RibbonXml);
    }
  else
    {
      return GetCustomUI_old (RibbonID, RibbonXml);
    }
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
      log_debug ("Invalidating ribbon: %p", *it);
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
