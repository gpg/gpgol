/* gpgoladdin.h - Connect GpgOL to Outlook as an addin
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

#ifndef GPGOLADDIN_H
#define GPGOLADDIN_H

#include <windows.h>

#include "mymapi.h"

#include <vector>
#include <memory>

class GpgolAddinRibbonExt;
class ApplicationEventListener;
class DispCache;
class CategoryManager;

/* Enums for the IDTExtensibility2 interface*/
typedef enum
  {
    ext_cm_AfterStartup = 0,
    ext_cm_Startup,
    ext_cm_External,
    ext_cm_CommandLine,
    ext_cm_Solution,
    ext_cm_UISetup
  }
ext_ConnectMode;

typedef enum
  {
    ext_dm_HostShutdown = 0,
    ext_dm_UserClosed,
    ext_dm_UISetupComplete,
    ext_dm_SolutionClosed
  }
ext_DisconnectMode;


/* Global variables */
extern bool g_ignore_next_load;

/* Global class locks */
extern ULONG addinLocks;

struct IDTExtensibility2;
typedef struct IDTExtensibility2 *LEXTENSIBILTY2;

/* Interface definitions */
DEFINE_GUID(IID_IDTExtensibility2, 0xB65AD801, 0xABAF, 0x11D0, 0xBB, 0x8B,
            0x00, 0xA0, 0xC9, 0x0F, 0x27, 0x44);

#undef INTERFACE
#define INTERFACE IDTExtensibility2
DECLARE_INTERFACE_(IDTExtensibility2, IDispatch)
{
  DECLARE_IUNKNOWN_METHODS;
  DECLARE_IDISPATCH_METHODS;
  /*** IDTExtensibility2 methods ***/

  STDMETHOD(OnConnection)(LPDISPATCH, ext_ConnectMode, LPDISPATCH,
                          SAFEARRAY**) PURE;
  STDMETHOD(OnDisconnection)(ext_DisconnectMode, SAFEARRAY**) PURE;
  STDMETHOD(OnAddInsUpdate)(SAFEARRAY **) PURE;
  STDMETHOD(OnStartupComplete)(SAFEARRAY**) PURE;
  STDMETHOD(OnBeginShutdown)(SAFEARRAY**) PURE;
};

DEFINE_GUID(IID_IRibbonExtensibility, 0x000C0396, 0x0000, 0x0000, 0xC0, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x46);

struct IRibbonExtensibility;
typedef struct IRibbonExtensibility *LRIBBONEXTENSIBILITY;

#undef INTERFACE
#define INTERFACE IRibbonExtensibility
DECLARE_INTERFACE_(IRibbonExtensibility, IDispatch)
{
  DECLARE_IUNKNOWN_METHODS;
  DECLARE_IDISPATCH_METHODS;

  /*** IRibbonExtensibility methods ***/
  STDMETHOD(GetCustomUI)(BSTR RibbonID, BSTR * RibbonXml) PURE;
};

DEFINE_GUID(IID_IRibbonCallback, 0xCE895442, 0x9981, 0x4315, 0xAA, 0x85,
            0x4B, 0x9A, 0x5C, 0x77, 0x39, 0xD8);

struct IRibbonCallback;
typedef struct IRibbonCallback *LRIBBONCALLBACK;

#undef INTERFACE
#define INTERFACE IRibbonCallback
DECLARE_INTERFACE_(IRibbonCallback, IUnknown)
{
  DECLARE_IUNKNOWN_METHODS;

  /*** IRibbonCallback methods ***/
  STDMETHOD(OnRibbonLoad)(IUnknown* pRibbonUIUnk) PURE;
  STDMETHOD(ButtonClicked)(IDispatch* ribbon) PURE;
};

DEFINE_GUID(IID_IRibbonControl, 0x000C0395, 0x0000, 0x0000, 0xC0, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x46);

struct IRibbonControl;
typedef struct IRibbonControl *LPRIBBONCONTROL;

#undef INTERFACE
#define INTERFACE IRibbonControl
DECLARE_INTERFACE_(IRibbonControl, IDispatch)
{
  DECLARE_IUNKNOWN_METHODS;
  DECLARE_IDISPATCH_METHODS;

  STDMETHOD(get_Id)(BSTR* id) PURE;
  STDMETHOD(get_Context)(IDispatch** context) PURE;
  STDMETHOD(get_Tag)(BSTR* Tag) PURE;
};


DEFINE_GUID(IID_ICustomTaskPaneConsumer, 0x000C033E, 0x0000, 0x0000, 0xC0,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);

class GpgolRibbonExtender : public IRibbonExtensibility
{
public:
  GpgolRibbonExtender(void);
  virtual ~GpgolRibbonExtender();

  /* IUnknown */
  STDMETHODIMP QueryInterface (REFIID riid, LPVOID* ppvObj);
  inline STDMETHODIMP_(ULONG) AddRef() { ++m_lRef;  return m_lRef; };
  inline STDMETHODIMP_(ULONG) Release()
    {
      ULONG lCount = --m_lRef;
      if (!lCount)
        delete this;
      return lCount;
    };

  /* IDispatch */
  STDMETHODIMP GetTypeInfoCount (UINT*);
  STDMETHODIMP GetTypeInfo (UINT, LCID, LPTYPEINFO*);
  STDMETHODIMP GetIDsOfNames (REFIID, LPOLESTR*, UINT, LCID, DISPID*);
  STDMETHODIMP Invoke (DISPID, REFIID, LCID, WORD,
                       DISPPARAMS*, VARIANT*, EXCEPINFO*, UINT*);

  /* IRibbonExtensibility */
  STDMETHODIMP GetCustomUI (BSTR RibbonID, BSTR* RibbonXml);

private:
  ULONG m_lRef;

};

class GpgolAddin : public IDTExtensibility2
{
public:
  GpgolAddin(void);
  virtual ~GpgolAddin();

public:

  /* IUnknown */
  STDMETHODIMP QueryInterface (REFIID riid, LPVOID* ppvObj);
  inline STDMETHODIMP_(ULONG) AddRef() { ++m_lRef;  return m_lRef; };
  inline STDMETHODIMP_(ULONG) Release()
    {
      ULONG lCount = --m_lRef;
      if (!lCount)
        delete this;
      return lCount;
    };

  /* IDispatch */
  STDMETHODIMP GetTypeInfoCount (UINT*);
  STDMETHODIMP GetTypeInfo (UINT, LCID, LPTYPEINFO*);
  STDMETHODIMP GetIDsOfNames (REFIID, LPOLESTR*, UINT, LCID, DISPID*);
  STDMETHODIMP Invoke (DISPID, REFIID, LCID, WORD,
                       DISPPARAMS*, VARIANT*, EXCEPINFO*, UINT*);

  /* IDTExtensibility */
  STDMETHODIMP OnConnection (LPDISPATCH Application,
                             ext_ConnectMode ConnectMode,
                             LPDISPATCH AddInInst,
                             SAFEARRAY** custom);
  STDMETHODIMP OnDisconnection (ext_DisconnectMode RemoveMode,
                                SAFEARRAY**  custom);
  STDMETHODIMP OnAddInsUpdate (SAFEARRAY** custom);
  STDMETHODIMP OnStartupComplete (SAFEARRAY** custom);
  STDMETHODIMP OnBeginShutdown (SAFEARRAY** custom);

public:
  static GpgolAddin * get_instance ();

  void registerExplorerSink (LPDISPATCH sink);
  void unregisterExplorerSink (LPDISPATCH sink);
  /* Start the shutdown. Unregisters everything and closes all
     crypto mails. */
  void shutdown ();
  LPDISPATCH get_application () { return m_application; }
  std::shared_ptr<DispCache> get_dispcache () { return m_dispcache; }
  bool isShutdown() { return m_shutdown; };

  /* Register a ribbon ui component */
  void addRibbon (LPDISPATCH ribbon);

  /* Invalidate the ribbons. */
  void invalidateRibbons ();

  std::shared_ptr<CategoryManager> get_category_mngr ();

private:
  ULONG m_lRef;
  GpgolRibbonExtender* m_ribbonExtender;

  LPDISPATCH m_application;
  LPDISPATCH m_addin;
  LPDISPATCH m_applicationEventSink;
  LPDISPATCH m_explorersEventSink;
  LPDISPATCH m_ribbon_control;
  bool m_disabled;
  bool m_shutdown;
  HHOOK m_hook;
  std::vector<LPDISPATCH> m_explorerEventSinks;
  std::shared_ptr<DispCache> m_dispcache;
  std::vector<LPDISPATCH> m_ribbon_uis;
  std::shared_ptr<CategoryManager> m_category_mngr;
};

class GpgolAddinFactory: public IClassFactory
{
public:
  GpgolAddinFactory(): m_lRef(0){}
  virtual ~GpgolAddinFactory();

  STDMETHODIMP QueryInterface (REFIID riid, LPVOID* ppvObj);
  inline STDMETHODIMP_(ULONG) AddRef() { ++m_lRef;  return m_lRef; };
  inline STDMETHODIMP_(ULONG) Release()
    {
      ULONG lCount = --m_lRef;
      if (!lCount)
        delete this;
      return lCount;
    };

  /* IClassFactory */
  STDMETHODIMP CreateInstance (LPUNKNOWN unknown, REFIID riid,
                               LPVOID* ppvObj);
  STDMETHODIMP LockServer (BOOL lock)
    {
      if (lock)
        ++addinLocks;
      else
        --addinLocks;
      return S_OK;
    }

private:
  ULONG m_lRef;
};

STDAPI DllGetClassObject (REFCLSID rclsid, REFIID riid, LPVOID* ppvObj);

/* Invalidates the UI XML to trigger a reload of the UI Elements. */
void gpgoladdin_invalidate_ui ();

/* Check in the Outlook settings HTML is enabled. Modifies opt.prefer_html */
void check_html_preferred ();

/* Check in the GpgOL settings whether auto vd (verify-decrypt) is
   enabled.  Modifies opt.dont_autodecrypt_preview.  */
void check_auto_vd_mail ();
#endif /*GPGOLADDIN_H*/
