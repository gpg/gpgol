/* myexchext.h - Simple replacement for exchext.h

 * This file defines the interface used by Exchange extensions.  It
 * has been compiled by g10 Code GmbH from several sources describing
 * the interface.
 *
 * Revisions:
 * 2005-08-12  Initial version.
 *
 */

#ifndef EXCHEXT_H
#define EXCHEXT_H

#ifdef __cplusplus
extern "C" {
#if 0
}
#endif
#endif

#include <commctrl.h>
#include <unknwn.h>
#include "mymapi.h"


/* Constants used by Install. */
#define EECONTEXT_SESSION               0x00000001
#define EECONTEXT_VIEWER                0x00000002
#define EECONTEXT_REMOTEVIEWER          0x00000003
#define EECONTEXT_SEARCHVIEWER          0x00000004
#define EECONTEXT_ADDRBOOK              0x00000005
#define EECONTEXT_SENDNOTEMESSAGE       0x00000006
#define EECONTEXT_READNOTEMESSAGE       0x00000007
#define EECONTEXT_SENDPOSTMESSAGE       0x00000008
#define EECONTEXT_READPOSTMESSAGE       0x00000009
#define EECONTEXT_READREPORTMESSAGE     0x0000000A
#define EECONTEXT_SENDRESENDMESSAGE     0x0000000B
#define EECONTEXT_PROPERTYSHEETS        0x0000000C
#define EECONTEXT_ADVANCEDCRITERIA      0x0000000D
#define EECONTEXT_TASK                  0x0000000E

/* Constants for GetVersion. */
#define EECBGV_GETBUILDVERSION          0x00000001
#define EECBGV_GETACTUALVERSION         0x00000002
#define EECBGV_GETVIRTUALVERSION        0x00000004
#define EECBGV_BUILDVERSION_MAJOR       0x000d0000
#define EECBGV_BUILDVERSION_MAJOR_MASK  0xffff0000
#define EECBGV_BUILDVERSION_MINOR_MASK  0x0000ffff

/* Some toolbar IDs. */
#define EETBID_STANDARD                 0x00000001

/* Constants use for QueryHelpText. */
#define EECQHT_STATUS                   0x00000001
#define EECQHT_TOOLTIP                  0x00000002

/* Flags use by the methods of IExchExtPropertySheets.  */
#define EEPS_MESSAGE                    0x00000001
#define EEPS_FOLDER                     0x00000002
#define EEPS_STORE                      0x00000003
#define EEPS_TOOLSOPTIONS               0x00000004

/* Flags used by OnFooComplete. */
#define EEME_FAILED                     0x00000001
#define EEME_COMPLETE_FAILED            0x00000002


/* Command IDs. */
#define EECMDID_ToolsCustomizeToolbar          134
#define EECMDID_ToolsOptions                   136


/* GUIDs */
DEFINE_GUID(GUID_NULL, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);

DEFINE_OLEGUID(IID_IUnknown,                  0x00000000, 0, 0);
DEFINE_OLEGUID(IID_IDispatch,                 0x00020400, 0, 0);

DEFINE_OLEGUID(IID_IExchExtCallback,          0x00020d10, 0, 0);
DEFINE_OLEGUID(IID_IExchExt,                  0x00020d11, 0, 0);
DEFINE_OLEGUID(IID_IExchExtCommands,          0x00020d12, 0, 0);
DEFINE_OLEGUID(IID_IExchExtUserEvents,        0x00020d13, 0, 0);
DEFINE_OLEGUID(IID_IExchExtSessionEvents,     0x00020d14, 0, 0);
DEFINE_OLEGUID(IID_IExchExtMessageEvents,     0x00020d15, 0, 0);
DEFINE_OLEGUID(IID_IExchExtAttachedFileEvents,0x00020d16, 0, 0);
DEFINE_OLEGUID(IID_IExchExtPropertySheets,    0x00020d17, 0, 0);
DEFINE_OLEGUID(IID_IExchExtAdvancedCriteria,  0x00020d18, 0, 0);
DEFINE_OLEGUID(IID_IExchExtModeless,          0x00020d19, 0, 0);
DEFINE_OLEGUID(IID_IExchExtModelessCallback,  0x00020d1a, 0, 0);
DEFINE_OLEGUID(IID_IOutlookExtCallback,       0x0006720d, 0, 0);


/* Type definitions. */

/* Parameters for the toolbar entries for
   IExchExtCommands::InstallCommands. */
struct TBENTRY
{
  HWND hwnd;
  ULONG tbid;
  ULONG ulFlags;
  UINT itbbBase;
};
typedef struct TBENTRY TBENTRY;
typedef struct TBENTRY *LPTBENTRY;






/**** Class declarations.  ***/
typedef struct IExchExt IExchExt;
typedef IExchExt *LPEXCHEXT;

typedef struct IExchExtMessageEvents IExchExtMessageEvents;
typedef IExchExtMessageEvents *LPEXCHEXTMESSAGEEVENTS;

typedef struct IExchExtCommands IExchExtCommands;
typedef IExchExtCommands *LPEXCHEXTCOMMANDS;

typedef struct IExchExtPropertySheets IExchExtPropertySheets;
typedef IExchExtPropertySheets *LPEXCHEXTPROPERTYSHEETS;

typedef struct IExchExtCallback IExchExtCallback;
typedef IExchExtCallback *LPEXCHEXTCALLBACK;

typedef struct IOutlookExtCallback IOutlookExtCallback;
typedef IOutlookExtCallback *LPOUTLOOKEXTCALLBACK;

/* The next classes are not yet defined. but if so they should go here. */
typedef struct IExchExtModeless IExchExtModeless; 
typedef IExchExtModeless *LPEXCHEXTMODELESS;
typedef struct IExchExtModelessCallback IExchExtModelessCallback;
typedef IExchExtModelessCallback *LPEXCHEXTMODELESSCALLBACK;




/*** Class declarations of classes defined elsewhere. ***/
struct IMAPISession;
typedef struct IMAPISession *LPMAPISESSION;

struct IAddrBook;
typedef struct IAddrBook *LPADRBOOK;

struct IMAPIFolder;
typedef struct IMAPIFolder *LPMAPIFOLDER;

struct IMAPIProp;
typedef struct IMAPIProp *LPMAPIPROP;

struct IPersistMessage;
typedef struct IPersistMessage *LPPERSISTMESSAGE;

struct IMAPIMessageSite;
typedef struct IMAPIMessageSite *LPMAPIMESSAGESITE;

struct IMAPIViewContext;
typedef struct IMAPIViewContext *LPMAPIVIEWCONTEXT;



/*** Types derived from the above class definitions. ***/

/* Callback used to load an extension. */
typedef LPEXCHEXT (CALLBACK *LPFNEXCHEXTENTRY)(void);

/* Parameters for the IExchExtCallback::ChooseFolder. */
typedef UINT (STDAPICALLTYPE *LPEECFHOOKPROC)(HWND, UINT, WPARAM, LPARAM);

struct EXCHEXTCHOOSEFOLDER
{
  UINT cbLength;
  HWND hwnd;
  LPTSTR szCaption;
  LPTSTR szLabel;
  LPTSTR szHelpFile;
  ULONG ulHelpID;
  HINSTANCE hinst;
  UINT uiDlgID;
  LPEECFHOOKPROC lpeecfhp;
  DWORD dwHookData;
  ULONG ulFlags;
  LPMDB pmdb;
  LPMAPIFOLDER pfld;
  LPTSTR szName;
  DWORD dwReserved1;
  DWORD dwReserved2;
  DWORD dwReserved3;
};
typedef struct EXCHEXTCHOOSEFOLDER EXCHEXTCHOOSEFOLDER;
typedef struct EXCHEXTCHOOSEFOLDER *LPEXCHEXTCHOOSEFOLDER;




/**** Class definitions.  ***/

EXTERN_C const IID IID_IExchExt;
#undef INTERFACE
#define INTERFACE IExchExt
DECLARE_INTERFACE_(IExchExt, IUnknown)
{
  /*** IUnknown methods. ***/
  STDMETHOD(QueryInterface)(THIS_ REFIID, PVOID*) PURE;
  STDMETHOD_(ULONG, AddRef)(THIS) PURE;
  STDMETHOD_(ULONG, Release)(THIS) PURE;

  /*** IExchExt methods. ***/
  STDMETHOD(Install)(THIS_ LPEXCHEXTCALLBACK, ULONG, ULONG) PURE;
};



EXTERN_C const IID IID_IExchExtMessageEvents;
#undef INTERFACE
#define INTERFACE IExchExtMessageEvents
DECLARE_INTERFACE_(IExchExtMessageEvents, IUnknown)
{
  /*** IUnknown methods. ***/
  STDMETHOD(QueryInterface)(THIS_ REFIID, PVOID*) PURE;
  STDMETHOD_(ULONG, AddRef)(THIS) PURE;
  STDMETHOD_(ULONG, Release)(THIS) PURE;

  /*** IExchExtMessageEvents methods. ***/
  STDMETHOD(OnRead)(THIS_ LPEXCHEXTCALLBACK) PURE;
  STDMETHOD(OnReadComplete)(THIS_ LPEXCHEXTCALLBACK, ULONG) PURE;
  STDMETHOD(OnWrite)(THIS_ LPEXCHEXTCALLBACK) PURE;
  STDMETHOD(OnWriteComplete)(THIS_ LPEXCHEXTCALLBACK, ULONG) PURE;
  STDMETHOD(OnCheckNames)(THIS_ LPEXCHEXTCALLBACK) PURE;
  STDMETHOD(OnCheckNamesComplete)(THIS_ LPEXCHEXTCALLBACK, ULONG) PURE;
  STDMETHOD(OnSubmit)(THIS_ LPEXCHEXTCALLBACK) PURE;
  STDMETHOD_(void, OnSubmitComplete)(THIS_ LPEXCHEXTCALLBACK, ULONG) PURE;
};



EXTERN_C const IID IID_IExchExtCommands;
#undef INTERFACE
#define INTERFACE IExchExtCommands
DECLARE_INTERFACE_(IExchExtCommands, IUnknown)
{
  /*** IUnknown methods. ***/
  STDMETHOD(QueryInterface)(THIS_ REFIID, PVOID*) PURE;
  STDMETHOD_(ULONG, AddRef)(THIS) PURE;
  STDMETHOD_(ULONG, Release)(THIS) PURE;

  /*** IExchExtCommands methods. ***/
  STDMETHOD(InstallCommands)(THIS_ LPEXCHEXTCALLBACK, HWND, HMENU,
                             UINT*, LPTBENTRY, UINT, ULONG) PURE;
  STDMETHOD_(void, InitMenu)(THIS_ LPEXCHEXTCALLBACK) PURE;
  STDMETHOD(DoCommand)(THIS_ LPEXCHEXTCALLBACK, UINT) PURE;
  STDMETHOD(Help)(THIS_ LPEXCHEXTCALLBACK, UINT) PURE;
  STDMETHOD(QueryHelpText)(THIS_ UINT, ULONG, LPTSTR, UINT) PURE;
  STDMETHOD(QueryButtonInfo)(THIS_ ULONG, UINT, LPTBBUTTON,
                             LPTSTR, UINT, ULONG) PURE;
  STDMETHOD(ResetToolbar)(THIS_ ULONG, ULONG) PURE;
};



EXTERN_C const IID IID_IExchExtPropertySheets;
#undef INTERFACE
#define INTERFACE IExchExtPropertySheets
DECLARE_INTERFACE_(IExchExtPropertySheets, IUnknown)
{
  /*** IUnknown methods. ***/
  STDMETHOD(QueryInterface)(THIS_ REFIID, PVOID*) PURE;
  STDMETHOD_(ULONG, AddRef)(THIS) PURE;
  STDMETHOD_(ULONG, Release)(THIS) PURE;

  /*** IExchExtPropertySheet methods. ***/
  STDMETHOD_(ULONG, GetMaxPageCount)(THIS_ ULONG) PURE;
  STDMETHOD(GetPages)(THIS_ LPEXCHEXTCALLBACK, ULONG,
                      LPPROPSHEETPAGE, ULONG*) PURE;
  STDMETHOD_(void, FreePages)(THIS_ LPPROPSHEETPAGE, ULONG, ULONG) PURE;
};



EXTERN_C const IID IID_IExchExtCallback;
#undef INTERFACE
#define INTERFACE IExchExtCallback
DECLARE_INTERFACE_(IExchExtCallback, IUnknown)
{
  /*** IUnknown methods. ***/
  STDMETHOD(QueryInterface)(THIS_ REFIID, PVOID*) PURE;
  STDMETHOD_(ULONG, AddRef)(THIS) PURE;
  STDMETHOD_(ULONG, Release)(THIS) PURE;

  /*** IExchExtCallback methods. ***/
  STDMETHOD(GetVersion)(THIS_ ULONG*, ULONG) PURE;
  STDMETHOD(GetWindow)(THIS_ HWND*) PURE;
  STDMETHOD(GetMenu)(THIS_ HMENU*) PURE;
  STDMETHOD(GetToolbar)(THIS_ ULONG, HWND*) PURE;
  STDMETHOD(GetSession)(THIS_ LPMAPISESSION*, LPADRBOOK*) PURE;
  STDMETHOD(GetObject)(THIS_ LPMDB*, LPMAPIPROP*) PURE;
  STDMETHOD(GetSelectionCount)(THIS_ ULONG*) PURE;
  STDMETHOD(GetSelectionItem)(THIS_ ULONG, ULONG*, LPENTRYID*, ULONG*,
                              LPTSTR, ULONG, ULONG*, ULONG) PURE;
  STDMETHOD(GetMenuPos)(THIS_ ULONG, HMENU*, ULONG*, ULONG*, ULONG) PURE;
  STDMETHOD(GetSharedExtsDir)(THIS_ LPTSTR, ULONG, ULONG) PURE;
  STDMETHOD(GetRecipients)(THIS_ LPADRLIST*) PURE;
  STDMETHOD(SetRecipients)(THIS_ LPADRLIST) PURE;
  STDMETHOD(GetNewMessageSite)(THIS_ ULONG, LPMAPIFOLDER, LPPERSISTMESSAGE,
                               LPMESSAGE*, LPMAPIMESSAGESITE*,
                               LPMAPIVIEWCONTEXT*, ULONG) PURE;
  STDMETHOD(RegisterModeless)(THIS_ LPEXCHEXTMODELESS,
                              LPEXCHEXTMODELESSCALLBACK*) PURE;
  STDMETHOD(ChooseFolder)(THIS_ LPEXCHEXTCHOOSEFOLDER) PURE;
};



EXTERN_C const IID IID_IOutlookExtCallback;
#undef INTERFACE
#define INTERFACE IOutlookExtCallback
DECLARE_INTERFACE_(IOutlookExtCallback, IUnknown)
{
  /*** IUnknown methods. ***/
  STDMETHOD(QueryInterface)(THIS_ REFIID, PVOID*) PURE;
  STDMETHOD_(ULONG, AddRef)(THIS) PURE;
  STDMETHOD_(ULONG, Release)(THIS) PURE;

  /*** IOutlookExtCallback.  **/
  STDMETHOD(GetObject)(LPUNKNOWN *ppunk);
  STDMETHOD(GetOfficeCharacter)(void **ppmsotfc);
};



#ifdef __cplusplus
}
#endif
#endif /*EXCHEXT_H*/
