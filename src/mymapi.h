/* mymapi.h - MAPI definitions required for OutlGPG and Mingw32
 * Copyright (C) 1998 Justin Bradford
 * Copyright (C) 2000 François Gouget
 * Copyright (C) 2005 g10 Code GmbH
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this program; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA
 * 02110-1301, USA.
 */

/* This header has been put together from the mingw32 and Wine
   headers.  The first were missing the MAPI definitions and the
   latter one is not compatible to the first one.  We only declare
   stuff we really need for this project.

   Revisions:
   2005-07-26  Initial version (wk at g10code).
   2005-08-14  Tweaked for use with myexchext.h.
*/

#ifndef MAPI_H
#define MAPI_H

#ifdef __cplusplus
extern "C" {
#endif

/* Some types */

#ifndef __LHANDLE
#define __LHANDLE
typedef unsigned long           LHANDLE, *LPLHANDLE;
#endif
#define lhSessionNull           ((LHANDLE)0)

typedef unsigned long           FLAGS;

struct MapiFileDesc_s
{
  ULONG ulReserved;
  ULONG flFlags;
  ULONG nPosition;
  LPSTR lpszPathName;
  LPSTR lpszFileName;
  LPVOID lpFileType;
};
typedef struct MapiFileDesc_s MapiFileDesc;
typedef struct MapiFileDesc_s *lpMapiFileDesc;

struct MapiRecipDesc_s
{
  ULONG ulReserved;
  ULONG ulRecipClass;
  LPSTR lpszName;
  LPSTR lpszAddress;
  ULONG ulEIDSize;
  LPVOID lpEntryID;
};
typedef struct MapiRecipDesc_s MapiRecipDesc;
typedef struct MapiRecipDesc_s *lpMapiRecipDesc;

struct MapiMessage_s
{
  ULONG ulReserved;
  LPSTR lpszSubject;
  LPSTR lpszNoteText;
  LPSTR lpszMessageType;
  LPSTR lpszDateReceived;
  LPSTR lpszConversationID;
  FLAGS flFlags;
  lpMapiRecipDesc lpOriginator;
  ULONG nRecipCount;
  lpMapiRecipDesc lpRecips;
  ULONG nFileCount;
  lpMapiFileDesc lpFiles;
};
typedef struct MapiMessage_s MapiMessage;
typedef struct MapiMessage_s *lpMapiMessage;


/* Error codes */
#define MAPI_E_NO_ACCESS                   E_ACCESSDENIED
#define MAPI_E_NOT_FOUND                   ((SCODE)0x8004010F)


/* MAPILogon */

#define MAPI_LOGON_UI           0x00000001
#define MAPI_PASSWORD_UI        0x00020000
#define MAPI_NEW_SESSION        0x00000002
#define MAPI_FORCE_DOWNLOAD     0x00001000
#define MAPI_EXTENDED           0x00000020


/* MAPISendMail */

#define MAPI_DIALOG             0x00000008


/* Flags for various calls */
#define MAPI_MODIFY                   0x00000001U /* Object can be modified. */
#define MAPI_CREATE                   0x00000002U 
#define MAPI_ACCESS_MODIFY            MAPI_MODIFY /* Want write access. */
#define MAPI_ACCESS_READ              0x00000002U /* Want read access. */
#define MAPI_ACCESS_DELETE            0x00000004U /* Want delete access. */
#define MAPI_ACCESS_CREATE_HIERARCHY  0x00000008U
#define MAPI_ACCESS_CREATE_CONTENTS   0x00000010U
#define MAPI_ACCESS_CREATE_ASSOCIATED 0x00000020U
#define MAPI_UNICODE                  0x80000000U /* Strings in this
                                                     call are Unicode. */

#define MAPI_BEST_ACCESS              0x00000010U


#define ATTACH_BY_VALUE       1 
#define ATTACH_BY_REFERENCE   2 
#define ATTACH_BY_REF_RESOLVE 3 
#define ATTACH_BY_REF_ONLY    4 
#define ATTACH_EMBEDDED_MSG   5
#define ATTACH_OLE            6  


#define FORCE_SAVE                    0x00000004U

#define RTF_SYNC_BODY_CHANGED         1 /* FIXME FIXME */


#ifndef MAPI_DIM
# define MAPI_DIM 1 /* Default to one dimension for variable length arrays */
#endif

DEFINE_OLEGUID(IID_IABContainer,0x2030D,0,0);
DEFINE_OLEGUID(IID_IABLogon,0x20314,0,0);
DEFINE_OLEGUID(IID_IABProvider,0x20311,0,0);
DEFINE_OLEGUID(IID_IAddrBook,0x20309,0,0);
DEFINE_OLEGUID(IID_IAttachment,0x20308,0,0);
DEFINE_OLEGUID(IID_IDistList,0x2030E,0,0);
DEFINE_OLEGUID(IID_IEnumMAPIFormProp,0x20323,0,0);
DEFINE_OLEGUID(IID_IMailUser,0x2030A,0,0);
DEFINE_OLEGUID(IID_IMAPIAdviseSink,0x20302,0,0);
DEFINE_OLEGUID(IID_IMAPIContainer,0x2030B,0,0);
DEFINE_OLEGUID(IID_IMAPIControl,0x2031B,0,0);
DEFINE_OLEGUID(IID_IMAPIFolder,0x2030C,0,0);
DEFINE_OLEGUID(IID_IMAPIForm,0x20327,0,0);
DEFINE_OLEGUID(IID_IMAPIFormAdviseSink,0x2032F,0,0);
DEFINE_OLEGUID(IID_IMAPIFormContainer,0x2032E,0,0);
DEFINE_OLEGUID(IID_IMAPIFormFactory,0x20350,0,0);
DEFINE_OLEGUID(IID_IMAPIFormInfo,0x20324,0,0);
DEFINE_OLEGUID(IID_IMAPIFormMgr,0x20322,0,0);
DEFINE_OLEGUID(IID_IMAPIFormProp,0x2032D,0,0);
DEFINE_OLEGUID(IID_IMAPIMessageSite,0x20370,0,0);
DEFINE_OLEGUID(IID_IMAPIProgress,0x2031F,0,0);
DEFINE_OLEGUID(IID_IMAPIProp,0x20303,0,0);
DEFINE_OLEGUID(IID_IMAPIPropData,0x2031A,0,0);
DEFINE_OLEGUID(IID_IMAPISession,0x20300,0,0);
DEFINE_OLEGUID(IID_IMAPISpoolerInit,0x20317,0,0);
DEFINE_OLEGUID(IID_IMAPISpoolerService,0x2031E,0,0);
DEFINE_OLEGUID(IID_IMAPISpoolerSession,0x20318,0,0);
DEFINE_OLEGUID(IID_IMAPIStatus,0x20305,0,0);
DEFINE_OLEGUID(IID_IMAPISup,0x2030F,0,0);
DEFINE_OLEGUID(IID_IMAPITable,0x20301,0,0);
DEFINE_OLEGUID(IID_IMAPITableData,0x20316,0,0);
DEFINE_OLEGUID(IID_IMAPIViewAdviseSink,0x2032B,0,0);
DEFINE_OLEGUID(IID_IMAPIViewContext,0x20321,0,0);
DEFINE_OLEGUID(IID_IMessage,0x20307,0,0);
DEFINE_OLEGUID(IID_IMsgServiceAdmin,0x2031D,0,0);
DEFINE_OLEGUID(IID_IMsgStore,0x20306,0,0);
DEFINE_OLEGUID(IID_IMSLogon,0x20313,0,0);
DEFINE_OLEGUID(IID_IMSProvider,0x20310,0,0);
DEFINE_OLEGUID(IID_IPersistMessage,0x2032A,0,0);
DEFINE_OLEGUID(IID_IProfAdmin,0x2031C,0,0);
DEFINE_OLEGUID(IID_IProfSect,0x20304,0,0);
DEFINE_OLEGUID(IID_IProviderAdmin,0x20325,0,0);
DEFINE_OLEGUID(IID_ISpoolerHook,0x20320,0,0);
DEFINE_OLEGUID(IID_IStream, 0x0000c, 0, 0);
DEFINE_OLEGUID(IID_IStreamDocfile,0x2032C,0,0);
DEFINE_OLEGUID(IID_IStreamTnef,0x20330,0,0);
DEFINE_OLEGUID(IID_ITNEF,0x20319,0,0);
DEFINE_OLEGUID(IID_IXPLogon,0x20315,0,0);
DEFINE_OLEGUID(IID_IXPProvider,0x20312,0,0);
DEFINE_OLEGUID(MUID_PROFILE_INSTANCE,0x20385,0,0);
DEFINE_OLEGUID(PS_MAPI,0x20328,0,0);
DEFINE_OLEGUID(PS_PUBLIC_STRINGS,0x20329,0,0);
DEFINE_OLEGUID(PS_ROUTING_ADDRTYPE,0x20381,0,0);
DEFINE_OLEGUID(PS_ROUTING_DISPLAY_NAME,0x20382,0,0);
DEFINE_OLEGUID(PS_ROUTING_EMAIL_ADDRESSES,0x20380,0,0);
DEFINE_OLEGUID(PS_ROUTING_ENTRYID,0x20383,0,0);
DEFINE_OLEGUID(PS_ROUTING_SEARCH_KEY,0x20384,0,0);


struct _ENTRYID
{
    BYTE abFlags[4];
    BYTE ab[MAPI_DIM];
};
typedef struct _ENTRYID ENTRYID;
typedef struct _ENTRYID *LPENTRYID;


/* The property tag structure. This describes a list of columns */
typedef struct _SPropTagArray
{
    ULONG cValues;              /* Number of elements in aulPropTag */
    ULONG aulPropTag[MAPI_DIM]; /* Property tags */
} SPropTagArray, *LPSPropTagArray;

#define CbNewSPropTagArray(c) \
               (offsetof(SPropTagArray,aulPropTag)+(c)*sizeof(ULONG))
#define CbSPropTagArray(p)    CbNewSPropTagArray((p)->cValues)
#define SizedSPropTagArray(n,id) \
    struct _SPropTagArray_##id { ULONG cValues; ULONG aulPropTag[n]; } id


/* Multi-valued PT_APPTIME property value */
typedef struct _SAppTimeArray
{
    ULONG   cValues; /* Number of doubles in lpat */
    double *lpat;    /* Pointer to double array of length cValues */
} SAppTimeArray;

/* PT_BINARY property value */
typedef struct _SBinary
{
    ULONG  cb;  /* Number of bytes in lpb */
    LPBYTE lpb; /* Pointer to byte array of length cb */
} SBinary, *LPSBinary;

/* Multi-valued PT_BINARY property value */
typedef struct _SBinaryArray
{
    ULONG    cValues; /* Number of SBinarys in lpbin */
    SBinary *lpbin;   /* Pointer to SBinary array of length cValues */
} SBinaryArray;

typedef SBinaryArray ENTRYLIST, *LPENTRYLIST;

/* Multi-valued PT_CY property value */
typedef struct _SCurrencyArray
{
    ULONG  cValues; /* Number of CYs in lpcu */
    CY    *lpcur;   /* Pointer to CY array of length cValues */
} SCurrencyArray;

/* Multi-valued PT_SYSTIME property value */
typedef struct _SDateTimeArray
{
    ULONG     cValues; /* Number of FILETIMEs in lpft */
    FILETIME *lpft;    /* Pointer to FILETIME array of length cValues */
} SDateTimeArray;

/* Multi-valued PT_DOUBLE property value */
typedef struct _SDoubleArray
{
    ULONG   cValues; /* Number of doubles in lpdbl */
    double *lpdbl;   /* Pointer to double array of length cValues */
} SDoubleArray;

/* Multi-valued PT_CLSID property value */
typedef struct _SGuidArray
{
    ULONG cValues; /* Number of GUIDs in lpguid */
    GUID *lpguid;  /* Pointer to GUID array of length cValues */
} SGuidArray;

/* Multi-valued PT_LONGLONG property value */
typedef struct _SLargeIntegerArray
{
    ULONG          cValues; /* Number of long64s in lpli */
    LARGE_INTEGER *lpli;    /* Pointer to long64 array of length cValues */
} SLargeIntegerArray;

/* Multi-valued PT_LONG property value */
typedef struct _SLongArray
{
    ULONG  cValues; /* Number of longs in lpl */
    LONG  *lpl;     /* Pointer to long array of length cValues */
} SLongArray;

/* Multi-valued PT_STRING8 property value */
typedef struct _SLPSTRArray
{
    ULONG  cValues; /* Number of Ascii strings in lppszA */
    LPSTR *lppszA;  /* Pointer to Ascii string array of length cValues */
} SLPSTRArray;

/* Multi-valued PT_FLOAT property value */
typedef struct _SRealArray
{
    ULONG cValues; /* Number of floats in lpflt */
    float *lpflt;  /* Pointer to float array of length cValues */
} SRealArray;

/* Multi-valued PT_SHORT property value */
typedef struct _SShortArray
{
    ULONG      cValues; /* Number of shorts in lpb */
    short int *lpi;     /* Pointer to short array of length cValues */
} SShortArray;

/* Multi-valued PT_UNICODE property value */
typedef struct _SWStringArray
{
    ULONG   cValues; /* Number of Unicode strings in lppszW */
    LPWSTR *lppszW;  /* Pointer to Unicode string array of length cValues */
} SWStringArray;


/* A property value */
union PV_u
{
    short int          i;
    LONG               l;
    ULONG              ul;
    float              flt;
    double             dbl;
    unsigned short     b;
    CY                 cur;
    double             at;
    FILETIME           ft;
    LPSTR              lpszA;
    SBinary            bin;
    LPWSTR             lpszW;
    LPGUID             lpguid;
    LARGE_INTEGER      li;
    SShortArray        MVi;
    SLongArray         MVl;
    SRealArray         MVflt;
    SDoubleArray       MVdbl;
    SCurrencyArray     MVcur;
    SAppTimeArray      MVat;
    SDateTimeArray     MVft;
    SBinaryArray       MVbin;
    SLPSTRArray        MVszA;
    SWStringArray      MVszW;
    SGuidArray         MVguid;
    SLargeIntegerArray MVli;
    SCODE              err;
    LONG               x;
};
typedef union PV_u uPV;

/* Property value structure. This is essentially a mini-Variant. */
struct SPropValue_s
{
  ULONG     ulPropTag;  /* The property type. */
  ULONG     dwAlignPad; /* Alignment, treat as reserved. */
  uPV       Value;      /* The property value. */
};
typedef struct SPropValue_s SPropValue;
typedef struct SPropValue_s *LPSPropValue;

/* Structure describing a table row (a collection of property values). */
struct SRow_s
{
  ULONG        ulAdrEntryPad; /* Padding, treat as reserved. */
  ULONG        cValues;       /* Count of property values in lpProbs. */
  LPSPropValue lpProps;       /* Pointer to an array of property
                                 values of length cValues. */
};
typedef struct SRow_s SRow;
typedef struct SRow_s *LPSRow;


/* Structure describing a set of table rows. */
struct SRowSet_s
{
  ULONG cRows;          /* Count of rows in aRow. */
  SRow  aRow[MAPI_DIM]; /* Array of rows of length cRows. */
};
typedef struct SRowSet_s *LPSRowSet;


/* Structure describing a problem with a property */
typedef struct _SPropProblem
{
    ULONG ulIndex;   /* Index of the property */
    ULONG ulPropTag; /* Proprty tag of the property */
    SCODE scode;     /* Error code of the problem */
} SPropProblem, *LPSPropProblem;

/* A collection of property problems */
typedef struct _SPropProblemArray
{
    ULONG        cProblem;           /* Number of problems in aProblem */
    SPropProblem aProblem[MAPI_DIM]; /* Array of problems of length cProblem */
} SPropProblemArray, *LPSPropProblemArray;


/* Table bookmarks */
typedef ULONG BOOKMARK;

#define BOOKMARK_BEGINNING ((BOOKMARK)0) /* The first row */
#define BOOKMARK_CURRENT   ((BOOKMARK)1) /* The curent table row */
#define BOOKMARK_END       ((BOOKMARK)2) /* The last row */


/* Row restrictions */
typedef struct _SRestriction* LPSRestriction;


/* Errors. */
typedef struct _MAPIERROR
{
    ULONG  ulVersion;       /* Mapi version */
#if defined (UNICODE) || defined (__WINESRC__)
    LPWSTR lpszError;       /* Error and component strings. These are Ascii */
    LPWSTR lpszComponent;   /* unless the MAPI_UNICODE flag is passed in */
#else
    LPSTR  lpszError;
    LPSTR  lpszComponent;
#endif
    ULONG  ulLowLevelError;
    ULONG  ulContext;
} MAPIERROR, *LPMAPIERROR;


/* Sorting */
#define TABLE_SORT_ASCEND  0U
#define TABLE_SORT_DESCEND 1U
#define TABLE_SORT_COMBINE 2U

typedef struct _SSortOrder
{
    ULONG ulPropTag;
    ULONG ulOrder;
} SSortOrder, *LPSSortOrder;

typedef struct _SSortOrderSet
{
    ULONG      cSorts;
    ULONG      cCategories;
    ULONG      cExpanded;
    SSortOrder aSort[MAPI_DIM];
} SSortOrderSet, * LPSSortOrderSet;



typedef struct _MAPINAMEID
{
    LPGUID lpguid;
    ULONG ulKind;
    union
    {
        LONG lID;
        LPWSTR lpwstrName;
    } Kind;
} MAPINAMEID, *LPMAPINAMEID;

#define MNID_ID     0
#define MNID_STRING 1


typedef struct _ADRENTRY
{
    ULONG        ulReserved1;
    ULONG        cValues;
    LPSPropValue rgPropVals;
} ADRENTRY, *LPADRENTRY;

typedef struct _ADRLIST
{
    ULONG    cEntries;
    ADRENTRY aEntries[MAPI_DIM];
} ADRLIST, *LPADRLIST;



/**** Class definitions ****/
typedef const IID *LPCIID;


typedef struct IAttach IAttach;
typedef IAttach *LPATTACH;

typedef struct IMAPIAdviseSink IMAPIAdviseSink;
typedef IMAPIAdviseSink *LPMAPIADVISESINK;

typedef struct IMAPIProgress IMAPIProgress;
typedef IMAPIProgress *LPMAPIPROGRESS;

typedef struct IMAPITable IMAPITable;
typedef IMAPITable *LPMAPITABLE;

typedef struct IMAPIProp IMAPIProp;
typedef IMAPIProp *LPMAPIPROP;

typedef struct IMessage IMessage;
typedef IMessage *LPMESSAGE;

typedef struct IMsgStore IMsgStore;
typedef IMsgStore *LPMDB;



EXTERN_C const IID IID_IMAPIProp;
#undef INTERFACE
#define INTERFACE IMAPIProp
DECLARE_INTERFACE_(IMAPIProp,IUnknown)
{
  /*** IUnknown methods ***/
  STDMETHOD(QueryInterface)(THIS_ REFIID, PVOID*) PURE;
  STDMETHOD_(ULONG,AddRef)(THIS) PURE;
  STDMETHOD_(ULONG,Release)(THIS) PURE;

  /*** IMAPIProp methods ***/
  STDMETHOD(GetLastError)(THIS_ HRESULT, ULONG, LPMAPIERROR FAR*);
  STDMETHOD(SaveChanges)(THIS_ ULONG);
  STDMETHOD(GetProps)(THIS_ LPSPropTagArray, ULONG, ULONG FAR*,
                      LPSPropValue FAR*);
  STDMETHOD(GetPropList)(THIS_ ULONG, LPSPropTagArray FAR *);
  STDMETHOD(OpenProperty)(THIS_ ULONG, LPCIID lpiid, ULONG,
                          ULONG, LPUNKNOWN FAR*);
  STDMETHOD(SetProps)(THIS_ LONG, LPSPropValue, LPSPropProblemArray FAR*);
  STDMETHOD(DeleteProps)(THIS_ LPSPropTagArray, LPSPropProblemArray FAR *);
  STDMETHOD(CopyTo)(THIS_ ULONG, LPCIID, LPSPropTagArray, ULONG,
                    LPMAPIPROGRESS, LPCIID lpInterface, LPVOID,
                    ULONG, LPSPropProblemArray FAR*);
  STDMETHOD(CopyProps)(THIS_ LPSPropTagArray, ULONG, LPMAPIPROGRESS,
                       LPCIID, LPVOID, ULONG, LPSPropProblemArray FAR*);  
  STDMETHOD(GetNamesFromIDs)(THIS_ LPSPropTagArray FAR*, LPGUID,
                             ULONG, ULONG FAR*, LPMAPINAMEID FAR* FAR*);
  STDMETHOD(GetIDsFromNames)(THIS_ ULONG, LPMAPINAMEID FAR*, ULONG,
                             LPSPropTagArray FAR*);

};


EXTERN_C const IID IID_IMessage;
#undef INTERFACE
#define INTERFACE IMessage
DECLARE_INTERFACE_(IMessage,IMAPIProp)
{
  /*** IUnknown methods ***/
  STDMETHOD(QueryInterface)(THIS_ REFIID, PVOID*) PURE;
  STDMETHOD_(ULONG,AddRef)(THIS) PURE;
  STDMETHOD_(ULONG,Release)(THIS) PURE;

  /*** IMAPIProp methods ***/
  STDMETHOD(GetLastError)(THIS_ HRESULT, ULONG, LPMAPIERROR FAR*);
  STDMETHOD(SaveChanges)(THIS_ ULONG);
  STDMETHOD(GetProps)(THIS_ LPSPropTagArray, ULONG,ULONG FAR*,
                      LPSPropValue FAR*);
  STDMETHOD(GetPropList)(THIS_ ULONG, LPSPropTagArray FAR *);
  STDMETHOD(OpenProperty)(THIS_ ULONG, LPCIID lpiid, ULONG,
                          ULONG, LPUNKNOWN FAR*);
  STDMETHOD(SetProps)(THIS_ LONG, LPSPropValue, LPSPropProblemArray FAR*);
  STDMETHOD(DeleteProps)(THIS_ LPSPropTagArray, LPSPropProblemArray FAR *);
  STDMETHOD(CopyTo)(THIS_ ULONG, LPCIID, LPSPropTagArray, ULONG,
                    LPMAPIPROGRESS, LPCIID lpInterface, LPVOID,
                    ULONG, LPSPropProblemArray FAR*);
  STDMETHOD(CopyProps)(THIS_ LPSPropTagArray, ULONG, LPMAPIPROGRESS,
                       LPCIID, LPVOID, ULONG, LPSPropProblemArray FAR*);  
  STDMETHOD(GetNamesFromIDs)(THIS_ LPSPropTagArray FAR*, LPGUID,
                             ULONG, ULONG FAR*, LPMAPINAMEID FAR* FAR*);
  STDMETHOD(GetIDsFromNames)(THIS_ ULONG, LPMAPINAMEID FAR*, ULONG,
                             LPSPropTagArray FAR*);

  /*** IMessage methods ***/
  STDMETHOD(GetAttachmentTable)(THIS_ ULONG, LPMAPITABLE FAR*);
  STDMETHOD(OpenAttach)(THIS_ ULONG, LPCIID, ULONG, LPATTACH FAR*);
  STDMETHOD(CreateAttach)(THIS_ LPCIID, ULONG, ULONG FAR*, LPATTACH FAR*);
  STDMETHOD(DeleteAttach)(THIS_ ULONG, ULONG, LPMAPIPROGRESS, ULONG);
  STDMETHOD(GetRecipientTable)(THIS_ ULONG, LPMAPITABLE FAR*);
  STDMETHOD(ModifyRecipients)(THIS_ ULONG, LPADRLIST);
  STDMETHOD(SubmitMessage)(THIS_ ULONG);
  STDMETHOD(SetReadFlag)(THIS_ ULONG);
};


EXTERN_C const IID IID_IAttachment;
#undef INTERFACE
#define INTERFACE IAttach
DECLARE_INTERFACE_(IAttach, IMAPIProp)
{
  /*** IUnknown methods ***/
  STDMETHOD(QueryInterface)(THIS_ REFIID, PVOID*) PURE;
  STDMETHOD_(ULONG,AddRef)(THIS) PURE;
  STDMETHOD_(ULONG,Release)(THIS) PURE;

  /*** IMAPIProp methods ***/
  STDMETHOD(GetLastError)(THIS_ HRESULT, ULONG, LPMAPIERROR FAR*);
  STDMETHOD(SaveChanges)(THIS_ ULONG);
  STDMETHOD(GetProps)(THIS_ LPSPropTagArray, ULONG,ULONG FAR*,
                      LPSPropValue FAR*);
  STDMETHOD(GetPropList)(THIS_ ULONG, LPSPropTagArray FAR *);
  STDMETHOD(OpenProperty)(THIS_ ULONG, LPCIID lpiid, ULONG,
                          ULONG, LPUNKNOWN FAR*);
  STDMETHOD(SetProps)(THIS_ LONG, LPSPropValue, LPSPropProblemArray FAR*);
  STDMETHOD(DeleteProps)(THIS_ LPSPropTagArray, LPSPropProblemArray FAR *);
  STDMETHOD(CopyTo)(THIS_ ULONG, LPCIID, LPSPropTagArray, ULONG,
                    LPMAPIPROGRESS, LPCIID lpInterface, LPVOID,
                    ULONG, LPSPropProblemArray FAR*);
  STDMETHOD(CopyProps)(THIS_ LPSPropTagArray, ULONG, LPMAPIPROGRESS,
                       LPCIID, LPVOID, ULONG, LPSPropProblemArray FAR*);  
  STDMETHOD(GetNamesFromIDs)(THIS_ LPSPropTagArray FAR*, LPGUID,
                             ULONG, ULONG FAR*, LPMAPINAMEID FAR* FAR*);
  STDMETHOD(GetIDsFromNames)(THIS_ ULONG, LPMAPINAMEID FAR*, ULONG,
                             LPSPropTagArray FAR*);

  /*** IAttach methods ***/
  /* No methods */
};


EXTERN_C const IID IID_IMAPITableData;
#undef INTERFACE
#define INTERFACE IMAPITable
DECLARE_INTERFACE_(IMAPITable,IUnknown)
{
  /*** IUnknown methods ***/
  STDMETHOD(QueryInterface)(THIS_ REFIID, PVOID*) PURE;
  STDMETHOD_(ULONG,AddRef)(THIS) PURE;
  STDMETHOD_(ULONG,Release)(THIS) PURE;

  /*** IMAPITable methods ***/
  STDMETHOD(GetLastError)(THIS_ HRESULT, ULONG, LPMAPIERROR*) PURE;
  STDMETHOD(Advise)(THIS_ ULONG, LPMAPIADVISESINK, ULONG*) PURE;
  STDMETHOD(Unadvise)(THIS_ ULONG ulCxn) PURE;
  STDMETHOD(GetStatus)(THIS_ ULONG*, ULONG*) PURE;
  STDMETHOD(SetColumns)(THIS_ LPSPropTagArray, ULONG) PURE;
  STDMETHOD(QueryColumns)(THIS_ ULONG, LPSPropTagArray*) PURE;
  STDMETHOD(GetRowCount)(THIS_ ULONG, ULONG *) PURE;
  STDMETHOD(SeekRow)(THIS_ BOOKMARK, LONG, LONG*) PURE;
  STDMETHOD(SeekRowApprox)(THIS_ ULONG, ULONG) PURE;
  STDMETHOD(QueryPosition)(THIS_ ULONG*, ULONG*, ULONG*) PURE;
  STDMETHOD(FindRow)(THIS_ LPSRestriction, BOOKMARK, ULONG) PURE;
  STDMETHOD(Restrict)(THIS_ LPSRestriction, ULONG) PURE;
  STDMETHOD(CreateBookmark)(THIS_ BOOKMARK*) PURE;
  STDMETHOD(FreeBookmark)(THIS_ BOOKMARK) PURE;
  STDMETHOD(SortTable)(THIS_ LPSSortOrderSet, ULONG) PURE;
  STDMETHOD(QuerySortOrder)(THIS_ LPSSortOrderSet*) PURE;
  STDMETHOD(QueryRows)(THIS_ LONG, ULONG, LPSRowSet*) PURE;
  STDMETHOD(Abort)(THIS) PURE;
  STDMETHOD(ExpandRow)(THIS_ ULONG, LPBYTE, ULONG, ULONG, 
                       LPSRowSet*, ULONG*) PURE;
  STDMETHOD(CollapseRow)(THIS_ ULONG, LPBYTE, ULONG, ULONG*) PURE;
  STDMETHOD(WaitForCompletion)(THIS_ ULONG, ULONG, ULONG*) PURE;
  STDMETHOD(GetCollapseState)(THIS_ ULONG, ULONG, LPBYTE, ULONG*,LPBYTE*) PURE;
  STDMETHOD(SetCollapseState)(THIS_ ULONG, ULONG, LPBYTE, BOOKMARK*) PURE;
};



/****  Function prototypes. *****/

ULONG   WINAPI UlAddRef(void*);
ULONG   WINAPI UlRelease(void*);
HRESULT WINAPI HrGetOneProp(LPMAPIPROP,ULONG,LPSPropValue*);
HRESULT WINAPI HrSetOneProp(LPMAPIPROP,LPSPropValue);
BOOL    WINAPI FPropExists(LPMAPIPROP,ULONG);
void    WINAPI FreePadrlist(LPADRLIST);
void    WINAPI FreeProws(LPSRowSet);
HRESULT WINAPI HrQueryAllRows(LPMAPITABLE,LPSPropTagArray,LPSRestriction,
                              LPSSortOrderSet,LONG,LPSRowSet*);
LPSPropValue WINAPI PpropFindProp(LPSPropValue,ULONG,ULONG);

HRESULT WINAPI RTFSync (LPMESSAGE, ULONG, BOOL FAR*);


/* Memory allocation routines */
typedef SCODE (WINAPI ALLOCATEBUFFER)(ULONG,LPVOID*);
typedef SCODE (WINAPI ALLOCATEMORE)(ULONG,LPVOID,LPVOID*);
typedef ULONG (WINAPI FREEBUFFER)(LPVOID);
typedef ALLOCATEBUFFER *LPALLOCATEBUFFER;
typedef ALLOCATEMORE *LPALLOCATEMORE;
typedef FREEBUFFER *LPFREEBUFFER;

SCODE WINAPI MAPIAllocateBuffer (ULONG, LPVOID FAR *);
ULONG WINAPI MAPIFreeBuffer (LPVOID);


#if defined (UNICODE)
HRESULT WINAPI OpenStreamOnFile(LPALLOCATEBUFFER,LPFREEBUFFER,
                                ULONG,LPWSTR,LPWSTR,LPSTREAM*);
#else
HRESULT WINAPI OpenStreamOnFile(LPALLOCATEBUFFER,LPFREEBUFFER,
                                ULONG,LPSTR,LPSTR,LPSTREAM*);
#endif

#ifdef __cplusplus
}
#endif
#endif /* MAPI_H */
