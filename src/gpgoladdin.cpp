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

#define TRACEPOINT() do { log_debug ("%s:%s:%d: tracepoint\n", \
                                     SRCNAME, __func__, __LINE__); \
                        } while (0)

/* Id's of our callbacks */
#define ID_CMD_DECRYPT_VERIFY   1
#define ID_CMD_DECRYPT          2
#define ID_CMD_VERIFY           3

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
GpgolAddin::GpgolAddin (void) : m_lRef(0), m_application(0), m_addin(0)
{
  /* Create the COM Extension Object that handles the startup and
     endinge initialization
  */
  m_gpgolext = new GpgolExt();

  /* RibbonExtender is it's own object to avoid the pitfalls of
     multiple inheritance
  */
  m_ribbonExtender = new GpgolRibbonExtender();
}

GpgolAddin::~GpgolAddin (void)
{
  log_debug ("%s:%s: cleaning up GpgolAddin object;",
             SRCNAME, __func__);

  engine_deinit ();
  write_options ();
  delete m_gpgolext;
  delete m_ribbonExtender;

  log_debug ("%s:%s: Object deleted\n", SRCNAME, __func__);
}

STDMETHODIMP
GpgolAddin::QueryInterface (REFIID riid, LPVOID* ppvObj)
{
  HRESULT hr = S_OK;

  *ppvObj = NULL;

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
      hr = m_gpgolext->QueryInterface (riid, ppvObj);
#if 0
      if (FAILED(hr))
        {
          LPOLESTR sRiid = NULL;
          StringFromIID(riid, &sRiid);
          log_debug ("%s:%s: queried for unimplmented interface: %S",
                     SRCNAME, __func__, sRiid);
        }
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
  TRACEPOINT();

  if (!m_application)
    {
      m_application = Application;
      m_application->AddRef();
      m_addin = AddInInst;
    }
  else
    {
      /* This should not happen but happened during development when
         the vtable was incorrect and the wrong function was called */
      log_debug ("%s:%s: Application already set. Ignoring new value.",
                 SRCNAME, __func__);
      return S_OK;
    }

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
                 SRCNAME, __func__, rgszNames[0]);
      /* How this is supposed to work with cNames > 1 is unknown,
         but we can just say that we won't support callbacks with
         different parameters and just match the name (the first element)
         and we give it one of our own dispIds's that are later handled in
         the invoke part */
      if (!wcscmp (rgszNames[i], L"AttachmentDecryptCallback"))
        {
          found = true;
          rgDispId[i] = ID_CMD_DECRYPT;
        }
    }

  if (cNames > 1)
    {
      log_debug ("More then one name provided. Should not happen");
    }

  return found ? S_OK : E_NOINTERFACE;
}

HRESULT
GpgolRibbonExtender::decryptAttachments(LPRIBBONCONTROL ctrl)
{
  BSTR idStr = NULL;
  LPDISPATCH context = NULL;
  int attachmentCount;
  HRESULT hr = 0;
  int i = 0;
  HWND curWindow;
  LPOLEWINDOW actExplorer;
  int err;

  /* We got the vtable right so we can save us the invoke and
     property lookup hassle and call it directly */
  hr = ctrl->get_Id (&idStr);
  hr |= ctrl->get_Context (&context);

  if (FAILED(hr))
    {
      log_debug ("%s:%s:Context / ID lookup failed. hr: %x",
                 SRCNAME, __func__, (unsigned int) hr);
      SysFreeString (idStr);
      return E_FAIL;
    }
  else
    {
      log_debug ("%s:%s: contextId: %S, contextObj: %s",
                 SRCNAME, __func__, idStr, get_object_name (context));
      SysFreeString (idStr);
    }

  attachmentCount = get_oom_int (context, "Count");
  log_debug ("Count: %i ", attachmentCount);

  actExplorer = (LPOLEWINDOW) get_oom_object(context,
                                             "Application.ActiveExplorer");
  if (actExplorer)
    actExplorer->GetWindow (&curWindow);
  else
    {
      log_debug ("%s:%s: Could not find active window",
                 SRCNAME, __func__);
      curWindow = NULL;
    }

  char *filenames[attachmentCount + 1];
  filenames[attachmentCount] = NULL;
  /* Yes the items start at 1! */
  for (i = 1; i <= attachmentCount; i++)
    {
      char buf[16];
      char *filename;
      wchar_t *wcsOutFilename;
      DISPPARAMS saveParams;
      VARIANT aVariant[1];
      LPDISPATCH attachmentObj;
      DISPID saveID;

      snprintf (buf, sizeof (buf), "Item(%i)", i);
      attachmentObj = get_oom_object (context, buf);
      filename = get_oom_string (attachmentObj, "FileName");

      saveID = lookup_oom_dispid (attachmentObj, "SaveAsFile");

      saveParams.rgvarg = aVariant;
      saveParams.rgvarg[0].vt = VT_BSTR;
      filenames[i-1] = get_save_filename (NULL, filename);
      xfree (filename);

      wcsOutFilename = utf8_to_wchar2 (filenames[i-1],
                                       strlen(filenames[i-1]));
      saveParams.rgvarg[0].bstrVal = SysAllocString (wcsOutFilename);
      saveParams.cArgs = 1;
      saveParams.cNamedArgs = 0;

      hr = attachmentObj->Invoke (saveID, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                                  DISPATCH_METHOD, &saveParams,
                                  NULL, NULL, NULL);
      SysFreeString (saveParams.rgvarg[0].bstrVal);
      if (FAILED(hr))
        {
          int j;
          log_debug ("%s:%s: Saving to file failed. hr: %x",
                     SRCNAME, __func__, (unsigned int) hr);
          for (j = 0; j < i; j++)
            xfree (filenames[j]);
          return hr;
        }
    }
  err = op_assuan_start_decrypt_files (curWindow, filenames);

  for (i = 0; i < attachmentCount; i++)
    xfree (filenames[i]);

  return err ? E_FAIL : S_OK;
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
        return decryptAttachments ((LPRIBBONCONTROL)
                                   parms->rgvarg[0].pdispVal);
    }

  log_debug ("%s:%s: leave", SRCNAME, __func__);

  return DISP_E_MEMBERNOTFOUND;
}

BSTR
loadXMLResource (int id)
{
  /* XXX I do not know how to get the handle of the currently
     executed code as we never had a chance in DllMain to save
     that handle. */

  /* FIXME this does not work as intended */
  HMODULE hModule = GetModuleHandle("gpgol.dll");

  HRSRC hRsrc = FindResourceEx (hModule, MAKEINTRESOURCE(id), TEXT("XML"),
                                MAKELANGID(LANG_NEUTRAL, SUBLANG_NEUTRAL));

  if (!hRsrc)
    {
      log_error_w32 (-1, "%s:%s: FindResource(%d) failed\n",
                     SRCNAME, __func__, id);
      return NULL;
    }

  HGLOBAL hGlobal = LoadResource(hModule, hRsrc);

  if (!hGlobal)
    {
      log_error_w32 (-1, "%s:%s: LoadResource(%d) failed\n",
                     SRCNAME, __func__, id);
      return NULL;
    }

  LPVOID xmlData = LockResource (hGlobal);

  return SysAllocString (reinterpret_cast<OLECHAR*>(xmlData));
}

STDMETHODIMP
GpgolRibbonExtender::GetCustomUI (BSTR RibbonID, BSTR * RibbonXml)
{
  log_debug ("%s:%s: GetCustomUI for id: %S", SRCNAME, __func__, RibbonID);

  if (!RibbonXml)
    return E_POINTER;

  /*if (!wcscmp (RibbonID, L"Microsoft.Outlook.Explorer"))
    {*/
     // *RibbonXml = loadXMLResource (IDR_XML_EXPLORER);
  /* TODO use callback for label's and Icons, load xml from resource */
      *RibbonXml = SysAllocString (
        L"<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">"
        L" <ribbon>"
        L"   <tabs>"
        L"    <tab id=\"gpgolTab\""
        L"         label=\"GnuPG\">"
        L"     <group id=\"general\""
        L"            label=\"Allgemein\">"
        L"       <button id=\"CustomButton\""
        L"               imageMso=\"HappyFace\""
        L"               size=\"large\""
        L"               label=\"Zertifikatsverwaltung\""
        L"               onAction=\"startCertManager\"/>"
        L"     </group>"
        L"    </tab>"
        L"   </tabs>"
        L" </ribbon>"
        L" <contextMenus>"
        L" <contextMenu idMso=\"ContextMenuAttachments\">"
            L"<button id=\"gpgol_decrypt\""
                L" label=\"Save and decrypt\""
                L" onAction=\"AttachmentDecryptCallback\"/>"
        L" </contextMenu>"
        L" </contextMenus>"
        L"</customUI>"
        );
  /*  } */

  return S_OK;
}
