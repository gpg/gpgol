/* gpgoladdin.cpp - Connect GpgOL to Outlook as an addin
 *    Copyright (C) 2013 Intevation GmbH
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

#include "util.h"
#include "gpgoladdin.h"

#include "mymapi.h"
#include "mymapitags.h"
#include "myexchext.h"

#include "common.h"
#include "display.h"
#include "msgcache.h"
#include "engine.h"
#include "engine-assuan.h"
#include "mapihelp.h"

#include "oomhelp.h"

#include "olflange.h"

#include "gpgol-ids.h"
#include "ribbon-callbacks.h"

#define TRACEPOINT() do { log_debug ("%s:%s:%d: tracepoint\n", \
                                     SRCNAME, __func__, __LINE__); \
                        } while (0)

#define ICON_SIZE_LARGE  32
#define ICON_SIZE_NORMAL 16

ULONG addinLocks = 0;

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
    return addinLocks == 0 ? S_OK : S_FALSE;
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
  *ppvObj = NULL;

  GpgolAddin* obj = new GpgolAddin();
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
  m_addin(0), m_disabled(false)
{
  read_options ();
  /* RibbonExtender is it's own object to avoid the pitfalls of
     multiple inheritance
  */
  m_ribbonExtender = new GpgolRibbonExtender();
}

GpgolAddin::~GpgolAddin (void)
{
  log_debug ("%s:%s: cleaning up GpgolAddin object;",
             SRCNAME, __func__);

  delete m_ribbonExtender;

  if (!m_disabled)
    {
      engine_deinit ();
      write_options ();
    }

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

  m_application = Application;
  m_application->AddRef();
  m_addin = AddInInst;

  version = get_oom_string (Application, "Version");

  log_debug ("%s:%s:   using GPGME %s\n",
             SRCNAME, __func__, version);

  if (!version || !strlen (version) || strncmp (version, "14", 2))
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

  write_options();
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
  TRACEPOINT();

  if (m_application)
    {
      /*
         An install_sinks here works this but we
         don't implement all the old extension feature
         in the addin yet.
         install_sinks ((LPEXCHEXTCALLBACK)m_application);
      */
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
  TRACEPOINT();
  return S_OK;
}

STDMETHODIMP
GpgolAddin::GetTypeInfoCount (UINT *r_count)
{
  *r_count = 0;
  TRACEPOINT(); /* Should not happen */
  return S_OK;
}

STDMETHODIMP
GpgolAddin::GetTypeInfo (UINT iTypeInfo, LCID lcid,
                                  LPTYPEINFO *r_typeinfo)
{
  (void)iTypeInfo;
  (void)lcid;
  (void)r_typeinfo;
  TRACEPOINT(); /* Should not happen */
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
  TRACEPOINT(); /* Should not happen */
  return E_NOINTERFACE;
}

STDMETHODIMP
GpgolAddin::Invoke (DISPID dispid, REFIID riid, LCID lcid,
                    WORD flags, DISPPARAMS *parms, VARIANT *result,
                    EXCEPINFO *exepinfo, UINT *argerr)
{
  TRACEPOINT(); /* Should not happen */
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
  TRACEPOINT(); /* Should not happen */
  return S_OK;
}

STDMETHODIMP
GpgolRibbonExtender::GetTypeInfo (UINT iTypeInfo, LCID lcid,
                                  LPTYPEINFO *r_typeinfo)
{
  (void)iTypeInfo;
  (void)lcid;
  (void)r_typeinfo;
  TRACEPOINT(); /* Should not happen */
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
      ID_MAPPER (L"encryptBody", ID_CMD_ENCRYPT_BODY)
      ID_MAPPER (L"decryptBody", ID_CMD_DECRYPT_BODY)
      ID_MAPPER (L"addEncSignedAttachment", ID_CMD_ATT_ENCSIGN_FILE)
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
      case ID_BTN_CERTMANAGER:
      case ID_BTN_ENCRYPT:
      case ID_BTN_DECRYPT:
      case ID_BTN_DECRYPT_LARGE:
      case ID_BTN_ENCRYPT_LARGE:
      case ID_BTN_ENCSIGN_LARGE:
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
STDMETHODIMP
GpgolRibbonExtender::GetCustomUI (BSTR RibbonID, BSTR * RibbonXml)
{
  wchar_t buffer[4096];

  memset(buffer, 0, sizeof buffer);

  log_debug ("%s:%s: GetCustomUI for id: %S", SRCNAME, __func__, RibbonID);

  if (!RibbonXml)
    return E_POINTER;

  if (!wcscmp (RibbonID, L"Microsoft.Outlook.Mail.Compose"))
    {
      swprintf (buffer,
        L"<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">"
        L" <ribbon>"
        L"   <tabs>"
        L"    <tab id=\"gpgolTab\""
        L"         label=\"%S\">"
        L"     <group id=\"general\""
        L"            label=\"%S\">"
        L"       <button id=\"CustomButton\""
        L"               getImage=\"btnCertManager\""
        L"               size=\"large\""
        L"               label=\"%S\""
        L"               onAction=\"startCertManager\"/>"
        L"     </group>"
        L"     <group id=\"textGroup\""
        L"            label=\"%S\">"
        L"       <button id=\"fullTextEncrypt\""
        L"               getImage=\"btnEncryptLarge\""
        L"               size=\"large\""
        L"               label=\"%S\""
        L"               onAction=\"encryptBody\"/>"
        L"       <button id=\"fullTextDecrypt\""
        L"               getImage=\"btnDecryptLarge\""
        L"               size=\"large\""
        L"               label=\"%S\""
        L"               onAction=\"decryptBody\"/>"
        L"     </group>"
        L"     <group id=\"attachmentGroup\""
        L"            label=\"%S\">"
        L"       <button id=\"encryptSignFile\""
        L"               getImage=\"btnEncryptFileLarge\""
        L"               size=\"large\""
        L"               label=\"%S\""
        L"               onAction=\"attachEncryptFile\"/>"
        L"     </group>"
        L"    </tab>"
        L"   </tabs>"
        L" </ribbon>"
        L" <contextMenus>"
        L"  <contextMenu idMso=\"ContextMenuText\">"
        L"    <button id=\"encryptButton\""
        L"            label=\"%S\""
        L"            getImage=\"btnEncrypt\""
        L"            onAction=\"encryptSelection\"/>"
        L"    <button id=\"decryptButton\""
        L"            label=\"%S\""
        L"            getImage=\"btnDecrypt\""
        L"            onAction=\"decryptSelection\"/>"
        L" </contextMenu>"
        L"</contextMenus>"
        L"</customUI>", _("GpgOL"), _("General"),
        _("Start Certificate Manager"),
        _("Textbody"),
        _("Encrypt"),
        _("Decrypt"),
        _("Attachments"),
        _("Encrypted file"),
        _("Encrypt"), _("Decrypt")
        );
    }
  else if (!wcscmp (RibbonID, L"Microsoft.Outlook.Mail.Read"))
    {
      swprintf (buffer,
        L"<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">"
        L"<contextMenus>"
        L"<contextMenu idMso=\"ContextMenuReadOnlyMailText\">"
        L" <button id=\"decryptReadButton\""
        L"         label=\"%S\""
        L"         onAction=\"decryptSelection\"/>"
        L" </contextMenu>"
        L"</contextMenus>"
        L"</customUI>", _("Decrypt"));
    }
  else if (!wcscmp (RibbonID, L"Microsoft.Outlook.Explorer"))
    {
      swprintf (buffer,
        L"<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">"
        L" <ribbon>"
        L"   <tabs>"
        L"    <tab id=\"gpgolTab\""
        L"         label=\"%S\">"
        L"     <group id=\"general\""
        L"            label=\"%S\">"
        L"       <button id=\"CustomButton\""
        L"               getImage=\"btnCertManager\""
        L"               size=\"large\""
        L"               label=\"%S\""
        L"               onAction=\"startCertManager\"/>"
        L"     </group>"
        L"    </tab>"
        L"   </tabs>"
        L"  <contextualTabs>"
        L"    <tabSet idMso=\"TabSetAttachments\">"
        L"        <tab idMso=\"TabAttachments\">"
        L"            <group label=\"%S\" id=\"gnupgLabel\">"
        L"                <button id=\"gpgol_contextual_decrypt\""
        L"                    size=\"large\""
        L"                    label=\"%S\""
        L"                    getImage=\"btnDecryptLarge\""
        L"                    onAction=\"attachmentDecryptCallback\" />"
        L"            </group>"
        L"        </tab>"
        L"    </tabSet>"
        L"  </contextualTabs>"
        L" </ribbon>"
        L" <contextMenus>"
        /*
           There appears to be no way to access the word editor
           / get the selected text from that Context.
        L" <contextMenu idMso=\"ContextMenuReadOnlyMailText\">"
        L" <button id=\"decryptReadButton1\""
        L"         label=\"%S\""
        L"         onAction=\"decryptSelection\"/>"
        L" </contextMenu>"
        */
        L" <contextMenu idMso=\"ContextMenuAttachments\">"
        L"   <button id=\"gpgol_decrypt\""
        L"           label=\"%S\""
        L"           getImage=\"btnDecrypt\""
        L"           onAction=\"attachmentDecryptCallback\"/>"
        L" </contextMenu>"
        L" </contextMenus>"
        L"</customUI>",
        _("GpgOL"), _("General"), _("Start Certificate Manager"),
        _("GpgOL"), _("Save and decrypt"),/*_("Decrypt"), */
        _("Save and decrypt"));
    }

  if (wcslen (buffer))
    *RibbonXml = SysAllocString (buffer);
  else
    *RibbonXml = NULL;

  return S_OK;
}
