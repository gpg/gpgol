/* this ALWAYS GENERATED file contains the definitions for the interfaces */


/* File created by MIDL compiler version 5.01.0164 */
/* at Thu Sep 30 12:18:45 2004
 */
/* Compiler settings for C:\Source_OSS\GData\Gdgpg\GDGPG.idl:
    Oicf (OptLev=i2), W1, Zp8, env=Win32, ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
*/
//@@MIDL_FILE_HEADING(  )


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 440
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __GDGPG_h__
#define __GDGPG_h__

#ifdef __cplusplus
extern "C"{
#endif 

/* Forward Declarations */ 

#ifndef __IGDGPG_FWD_DEFINED__
#define __IGDGPG_FWD_DEFINED__
typedef interface IGDGPG IGDGPG;
#endif 	/* __IGDGPG_FWD_DEFINED__ */


#ifndef __Ig10Code_FWD_DEFINED__
#define __Ig10Code_FWD_DEFINED__
typedef interface Ig10Code Ig10Code;
#endif 	/* __Ig10Code_FWD_DEFINED__ */


#ifndef __GDGPG_FWD_DEFINED__
#define __GDGPG_FWD_DEFINED__

#ifdef __cplusplus
typedef class GDGPG GDGPG;
#else
typedef struct GDGPG GDGPG;
#endif /* __cplusplus */

#endif 	/* __GDGPG_FWD_DEFINED__ */


#ifndef ___Ig10CodeEvents_FWD_DEFINED__
#define ___Ig10CodeEvents_FWD_DEFINED__
typedef interface _Ig10CodeEvents _Ig10CodeEvents;
#endif 	/* ___Ig10CodeEvents_FWD_DEFINED__ */


#ifndef __g10Code_FWD_DEFINED__
#define __g10Code_FWD_DEFINED__

#ifdef __cplusplus
typedef class g10Code g10Code;
#else
typedef struct g10Code g10Code;
#endif /* __cplusplus */

#endif 	/* __g10Code_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

void __RPC_FAR * __RPC_USER MIDL_user_allocate(size_t);
void __RPC_USER MIDL_user_free( void __RPC_FAR * ); 

/* interface __MIDL_itf_GDGPG_0000 */
/* [local] */ 

typedef /* [public] */ 
enum __MIDL___MIDL_itf_GDGPG_0000_0001
    {	g10code_err_Success	= 0,
	g10code_err_InvRecipients	= 1,
	g10code_err_NoRecipients	= 2,
	g10code_err_NoPlaintext	= 3,
	g10code_err_NoCiphertext	= 4,
	g10code_err_NoBinary	= 5,
	g10code_err_NoPassphrase	= 6,
	g10code_err_BadPassphrase	= 7,
	g10code_err_KeyNotFound	= 8
    }	g10code_err_t;



extern RPC_IF_HANDLE __MIDL_itf_GDGPG_0000_v0_0_c_ifspec;
extern RPC_IF_HANDLE __MIDL_itf_GDGPG_0000_v0_0_s_ifspec;

#ifndef __IGDGPG_INTERFACE_DEFINED__
#define __IGDGPG_INTERFACE_DEFINED__

/* interface IGDGPG */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IGDGPG;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("83C42CA3-5D35-4120-8DF6-09D835166594")
    IGDGPG : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE OpenKeyManager( 
            /* [retval][out] */ int __RPC_FAR *pvReturn) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE EncryptAndSignFile( 
            /* [in] */ ULONG hWndParent,
            /* [in] */ BOOL bEncrypt,
            /* [in] */ BOOL bSign,
            /* [in] */ BSTR strFilenameSource,
            /* [in] */ BSTR strFilenameDest,
            /* [in] */ BSTR strRecipient,
            /* [in] */ BOOL bArmor,
            /* [in] */ BOOL bEncryptWithStandardKey,
            /* [retval][out] */ int __RPC_FAR *pvReturn) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE DecryptFile( 
            /* [in] */ ULONG hWndParent,
            /* [in] */ BSTR strFilenameSource,
            /* [in] */ BSTR strFilenameDest,
            /* [retval][out] */ int __RPC_FAR *pvReturn) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ExportStandardKey( 
            /* [in] */ ULONG hWndParent,
            /* [in] */ BSTR strExportFileName,
            /* [retval][out] */ int __RPC_FAR *pvReturn) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ImportKeys( 
            /* [in] */ ULONG hWndParent,
            /* [in] */ BSTR strImportFilename,
            /* [in] */ BOOL bShowMessage,
            /* [out] */ int __RPC_FAR *pvEditCount,
            /* [out] */ int __RPC_FAR *pvImportCount,
            /* [out] */ int __RPC_FAR *pvUnchangeCnt,
            /* [retval][out] */ int __RPC_FAR *pvReturn) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE SetStorePassphraseTime( 
            /* [in] */ int nSeconds) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE InvalidateKeyLists( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Options( 
            /* [in] */ ULONG hWndParent) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE EncryptAndSignNextFile( 
            /* [in] */ ULONG hWndParent,
            /* [in] */ BSTR strFilenameSource,
            /* [in] */ BSTR strFilenameDest,
            /* [in] */ BOOL bArmor,
            /* [retval][out] */ int __RPC_FAR *pvReturn) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE DecryptNextFile( 
            /* [in] */ ULONG hWndParent,
            /* [in] */ BSTR strFilenameSource,
            /* [in] */ BSTR strFilenameDest,
            /* [retval][out] */ int __RPC_FAR *pvReturn) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetGPGOutput( 
            /* [out] */ BSTR __RPC_FAR *hStdErr) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetGPGInfo( 
            /* [in] */ BSTR strFilename,
            /* [out] */ BSTR __RPC_FAR *hInfo) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE VerifyDetachedSignature( 
            /* [in] */ ULONG hWndParent,
            /* [in] */ BSTR strFilenameText,
            /* [in] */ BSTR strFilenameSig,
            /* [retval][out] */ int __RPC_FAR *pvReturn) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE SetLogLevel( 
            /* [in] */ ULONG nLevel) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE SetLogFile( 
            /* [in] */ BSTR strLogFilename,
            /* [retval][out] */ int __RPC_FAR *pvReturn) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IGDGPGVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IGDGPG __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IGDGPG __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IGDGPG __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IGDGPG __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IGDGPG __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IGDGPG __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IGDGPG __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *OpenKeyManager )( 
            IGDGPG __RPC_FAR * This,
            /* [retval][out] */ int __RPC_FAR *pvReturn);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *EncryptAndSignFile )( 
            IGDGPG __RPC_FAR * This,
            /* [in] */ ULONG hWndParent,
            /* [in] */ BOOL bEncrypt,
            /* [in] */ BOOL bSign,
            /* [in] */ BSTR strFilenameSource,
            /* [in] */ BSTR strFilenameDest,
            /* [in] */ BSTR strRecipient,
            /* [in] */ BOOL bArmor,
            /* [in] */ BOOL bEncryptWithStandardKey,
            /* [retval][out] */ int __RPC_FAR *pvReturn);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *DecryptFile )( 
            IGDGPG __RPC_FAR * This,
            /* [in] */ ULONG hWndParent,
            /* [in] */ BSTR strFilenameSource,
            /* [in] */ BSTR strFilenameDest,
            /* [retval][out] */ int __RPC_FAR *pvReturn);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ExportStandardKey )( 
            IGDGPG __RPC_FAR * This,
            /* [in] */ ULONG hWndParent,
            /* [in] */ BSTR strExportFileName,
            /* [retval][out] */ int __RPC_FAR *pvReturn);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ImportKeys )( 
            IGDGPG __RPC_FAR * This,
            /* [in] */ ULONG hWndParent,
            /* [in] */ BSTR strImportFilename,
            /* [in] */ BOOL bShowMessage,
            /* [out] */ int __RPC_FAR *pvEditCount,
            /* [out] */ int __RPC_FAR *pvImportCount,
            /* [out] */ int __RPC_FAR *pvUnchangeCnt,
            /* [retval][out] */ int __RPC_FAR *pvReturn);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetStorePassphraseTime )( 
            IGDGPG __RPC_FAR * This,
            /* [in] */ int nSeconds);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *InvalidateKeyLists )( 
            IGDGPG __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Options )( 
            IGDGPG __RPC_FAR * This,
            /* [in] */ ULONG hWndParent);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *EncryptAndSignNextFile )( 
            IGDGPG __RPC_FAR * This,
            /* [in] */ ULONG hWndParent,
            /* [in] */ BSTR strFilenameSource,
            /* [in] */ BSTR strFilenameDest,
            /* [in] */ BOOL bArmor,
            /* [retval][out] */ int __RPC_FAR *pvReturn);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *DecryptNextFile )( 
            IGDGPG __RPC_FAR * This,
            /* [in] */ ULONG hWndParent,
            /* [in] */ BSTR strFilenameSource,
            /* [in] */ BSTR strFilenameDest,
            /* [retval][out] */ int __RPC_FAR *pvReturn);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetGPGOutput )( 
            IGDGPG __RPC_FAR * This,
            /* [out] */ BSTR __RPC_FAR *hStdErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetGPGInfo )( 
            IGDGPG __RPC_FAR * This,
            /* [in] */ BSTR strFilename,
            /* [out] */ BSTR __RPC_FAR *hInfo);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *VerifyDetachedSignature )( 
            IGDGPG __RPC_FAR * This,
            /* [in] */ ULONG hWndParent,
            /* [in] */ BSTR strFilenameText,
            /* [in] */ BSTR strFilenameSig,
            /* [retval][out] */ int __RPC_FAR *pvReturn);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetLogLevel )( 
            IGDGPG __RPC_FAR * This,
            /* [in] */ ULONG nLevel);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetLogFile )( 
            IGDGPG __RPC_FAR * This,
            /* [in] */ BSTR strLogFilename,
            /* [retval][out] */ int __RPC_FAR *pvReturn);
        
        END_INTERFACE
    } IGDGPGVtbl;

    interface IGDGPG
    {
        CONST_VTBL struct IGDGPGVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IGDGPG_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IGDGPG_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IGDGPG_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IGDGPG_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IGDGPG_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IGDGPG_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IGDGPG_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IGDGPG_OpenKeyManager(This,pvReturn)	\
    (This)->lpVtbl -> OpenKeyManager(This,pvReturn)

#define IGDGPG_EncryptAndSignFile(This,hWndParent,bEncrypt,bSign,strFilenameSource,strFilenameDest,strRecipient,bArmor,bEncryptWithStandardKey,pvReturn)	\
    (This)->lpVtbl -> EncryptAndSignFile(This,hWndParent,bEncrypt,bSign,strFilenameSource,strFilenameDest,strRecipient,bArmor,bEncryptWithStandardKey,pvReturn)

#define IGDGPG_DecryptFile(This,hWndParent,strFilenameSource,strFilenameDest,pvReturn)	\
    (This)->lpVtbl -> DecryptFile(This,hWndParent,strFilenameSource,strFilenameDest,pvReturn)

#define IGDGPG_ExportStandardKey(This,hWndParent,strExportFileName,pvReturn)	\
    (This)->lpVtbl -> ExportStandardKey(This,hWndParent,strExportFileName,pvReturn)

#define IGDGPG_ImportKeys(This,hWndParent,strImportFilename,bShowMessage,pvEditCount,pvImportCount,pvUnchangeCnt,pvReturn)	\
    (This)->lpVtbl -> ImportKeys(This,hWndParent,strImportFilename,bShowMessage,pvEditCount,pvImportCount,pvUnchangeCnt,pvReturn)

#define IGDGPG_SetStorePassphraseTime(This,nSeconds)	\
    (This)->lpVtbl -> SetStorePassphraseTime(This,nSeconds)

#define IGDGPG_InvalidateKeyLists(This)	\
    (This)->lpVtbl -> InvalidateKeyLists(This)

#define IGDGPG_Options(This,hWndParent)	\
    (This)->lpVtbl -> Options(This,hWndParent)

#define IGDGPG_EncryptAndSignNextFile(This,hWndParent,strFilenameSource,strFilenameDest,bArmor,pvReturn)	\
    (This)->lpVtbl -> EncryptAndSignNextFile(This,hWndParent,strFilenameSource,strFilenameDest,bArmor,pvReturn)

#define IGDGPG_DecryptNextFile(This,hWndParent,strFilenameSource,strFilenameDest,pvReturn)	\
    (This)->lpVtbl -> DecryptNextFile(This,hWndParent,strFilenameSource,strFilenameDest,pvReturn)

#define IGDGPG_GetGPGOutput(This,hStdErr)	\
    (This)->lpVtbl -> GetGPGOutput(This,hStdErr)

#define IGDGPG_GetGPGInfo(This,strFilename,hInfo)	\
    (This)->lpVtbl -> GetGPGInfo(This,strFilename,hInfo)

#define IGDGPG_VerifyDetachedSignature(This,hWndParent,strFilenameText,strFilenameSig,pvReturn)	\
    (This)->lpVtbl -> VerifyDetachedSignature(This,hWndParent,strFilenameText,strFilenameSig,pvReturn)

#define IGDGPG_SetLogLevel(This,nLevel)	\
    (This)->lpVtbl -> SetLogLevel(This,nLevel)

#define IGDGPG_SetLogFile(This,strLogFilename,pvReturn)	\
    (This)->lpVtbl -> SetLogFile(This,strLogFilename,pvReturn)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IGDGPG_OpenKeyManager_Proxy( 
    IGDGPG __RPC_FAR * This,
    /* [retval][out] */ int __RPC_FAR *pvReturn);


void __RPC_STUB IGDGPG_OpenKeyManager_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IGDGPG_EncryptAndSignFile_Proxy( 
    IGDGPG __RPC_FAR * This,
    /* [in] */ ULONG hWndParent,
    /* [in] */ BOOL bEncrypt,
    /* [in] */ BOOL bSign,
    /* [in] */ BSTR strFilenameSource,
    /* [in] */ BSTR strFilenameDest,
    /* [in] */ BSTR strRecipient,
    /* [in] */ BOOL bArmor,
    /* [in] */ BOOL bEncryptWithStandardKey,
    /* [retval][out] */ int __RPC_FAR *pvReturn);


void __RPC_STUB IGDGPG_EncryptAndSignFile_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IGDGPG_DecryptFile_Proxy( 
    IGDGPG __RPC_FAR * This,
    /* [in] */ ULONG hWndParent,
    /* [in] */ BSTR strFilenameSource,
    /* [in] */ BSTR strFilenameDest,
    /* [retval][out] */ int __RPC_FAR *pvReturn);


void __RPC_STUB IGDGPG_DecryptFile_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IGDGPG_ExportStandardKey_Proxy( 
    IGDGPG __RPC_FAR * This,
    /* [in] */ ULONG hWndParent,
    /* [in] */ BSTR strExportFileName,
    /* [retval][out] */ int __RPC_FAR *pvReturn);


void __RPC_STUB IGDGPG_ExportStandardKey_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IGDGPG_ImportKeys_Proxy( 
    IGDGPG __RPC_FAR * This,
    /* [in] */ ULONG hWndParent,
    /* [in] */ BSTR strImportFilename,
    /* [in] */ BOOL bShowMessage,
    /* [out] */ int __RPC_FAR *pvEditCount,
    /* [out] */ int __RPC_FAR *pvImportCount,
    /* [out] */ int __RPC_FAR *pvUnchangeCnt,
    /* [retval][out] */ int __RPC_FAR *pvReturn);


void __RPC_STUB IGDGPG_ImportKeys_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IGDGPG_SetStorePassphraseTime_Proxy( 
    IGDGPG __RPC_FAR * This,
    /* [in] */ int nSeconds);


void __RPC_STUB IGDGPG_SetStorePassphraseTime_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IGDGPG_InvalidateKeyLists_Proxy( 
    IGDGPG __RPC_FAR * This);


void __RPC_STUB IGDGPG_InvalidateKeyLists_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IGDGPG_Options_Proxy( 
    IGDGPG __RPC_FAR * This,
    /* [in] */ ULONG hWndParent);


void __RPC_STUB IGDGPG_Options_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IGDGPG_EncryptAndSignNextFile_Proxy( 
    IGDGPG __RPC_FAR * This,
    /* [in] */ ULONG hWndParent,
    /* [in] */ BSTR strFilenameSource,
    /* [in] */ BSTR strFilenameDest,
    /* [in] */ BOOL bArmor,
    /* [retval][out] */ int __RPC_FAR *pvReturn);


void __RPC_STUB IGDGPG_EncryptAndSignNextFile_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IGDGPG_DecryptNextFile_Proxy( 
    IGDGPG __RPC_FAR * This,
    /* [in] */ ULONG hWndParent,
    /* [in] */ BSTR strFilenameSource,
    /* [in] */ BSTR strFilenameDest,
    /* [retval][out] */ int __RPC_FAR *pvReturn);


void __RPC_STUB IGDGPG_DecryptNextFile_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IGDGPG_GetGPGOutput_Proxy( 
    IGDGPG __RPC_FAR * This,
    /* [out] */ BSTR __RPC_FAR *hStdErr);


void __RPC_STUB IGDGPG_GetGPGOutput_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IGDGPG_GetGPGInfo_Proxy( 
    IGDGPG __RPC_FAR * This,
    /* [in] */ BSTR strFilename,
    /* [out] */ BSTR __RPC_FAR *hInfo);


void __RPC_STUB IGDGPG_GetGPGInfo_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IGDGPG_VerifyDetachedSignature_Proxy( 
    IGDGPG __RPC_FAR * This,
    /* [in] */ ULONG hWndParent,
    /* [in] */ BSTR strFilenameText,
    /* [in] */ BSTR strFilenameSig,
    /* [retval][out] */ int __RPC_FAR *pvReturn);


void __RPC_STUB IGDGPG_VerifyDetachedSignature_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IGDGPG_SetLogLevel_Proxy( 
    IGDGPG __RPC_FAR * This,
    /* [in] */ ULONG nLevel);


void __RPC_STUB IGDGPG_SetLogLevel_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IGDGPG_SetLogFile_Proxy( 
    IGDGPG __RPC_FAR * This,
    /* [in] */ BSTR strLogFilename,
    /* [retval][out] */ int __RPC_FAR *pvReturn);


void __RPC_STUB IGDGPG_SetLogFile_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IGDGPG_INTERFACE_DEFINED__ */


#ifndef __Ig10Code_INTERFACE_DEFINED__
#define __Ig10Code_INTERFACE_DEFINED__

/* interface Ig10Code */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_Ig10Code;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("4514D7BD-4BBE-47C1-9552-78E189620636")
    Ig10Code : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Plaintext( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Plaintext( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Ciphertext( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Ciphertext( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Armor( 
            /* [retval][out] */ BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Armor( 
            /* [in] */ BOOL newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_LocalUser( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_LocalUser( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_CompressLevel( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_CompressLevel( 
            /* [in] */ long newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_TextMode( 
            /* [retval][out] */ BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_TextMode( 
            /* [in] */ BOOL newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Expert( 
            /* [retval][out] */ BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Expert( 
            /* [in] */ BOOL newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_PGPMode( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_PGPMode( 
            /* [in] */ long newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_AlwaysTrust( 
            /* [retval][out] */ BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_AlwaysTrust( 
            /* [in] */ BOOL newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE SetLogLevel( 
            long logLevel) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE SetLogFile( 
            BSTR logFile,
            /* [retval][out] */ long __RPC_FAR *pvReturn) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Binary( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Binary( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE AddRecipient( 
            BSTR name,
            /* [retval][out] */ long __RPC_FAR *pvReturn) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Encrypt( 
            /* [retval][out] */ long __RPC_FAR *pvReturn) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ClearRecipient( void) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Output( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Output( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_NoVersion( 
            /* [retval][out] */ BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_NoVersion( 
            /* [in] */ BOOL newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Comment( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Comment( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Passphrase( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Decrypt( 
            /* [retval][out] */ long __RPC_FAR *pvReturn) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Export( 
            BSTR keyNames,
            /* [retval][out] */ long __RPC_FAR *pvReturn) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_EncryptoTo( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_EncryptoTo( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ForceMDC( 
            /* [retval][out] */ BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_ForceMDC( 
            /* [in] */ BOOL newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ForceV3Sig( 
            /* [retval][out] */ BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_ForceV3Sig( 
            /* [in] */ BOOL newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Keyserver( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Keyserver( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_HomeDir( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_HomeDir( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE EncryptFile( 
            /* [in] */ BSTR inFile,
            /* [retval][out] */ long __RPC_FAR *pvReturn) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE DecryptFile( 
            /* [in] */ BSTR inFile,
            /* [retval][out] */ long __RPC_FAR *pvReturn) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct Ig10CodeVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            Ig10Code __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            Ig10Code __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            Ig10Code __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            Ig10Code __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            Ig10Code __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            Ig10Code __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            Ig10Code __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Plaintext )( 
            Ig10Code __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Plaintext )( 
            Ig10Code __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Ciphertext )( 
            Ig10Code __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Ciphertext )( 
            Ig10Code __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Armor )( 
            Ig10Code __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Armor )( 
            Ig10Code __RPC_FAR * This,
            /* [in] */ BOOL newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_LocalUser )( 
            Ig10Code __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_LocalUser )( 
            Ig10Code __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_CompressLevel )( 
            Ig10Code __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_CompressLevel )( 
            Ig10Code __RPC_FAR * This,
            /* [in] */ long newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_TextMode )( 
            Ig10Code __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_TextMode )( 
            Ig10Code __RPC_FAR * This,
            /* [in] */ BOOL newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Expert )( 
            Ig10Code __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Expert )( 
            Ig10Code __RPC_FAR * This,
            /* [in] */ BOOL newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_PGPMode )( 
            Ig10Code __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_PGPMode )( 
            Ig10Code __RPC_FAR * This,
            /* [in] */ long newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_AlwaysTrust )( 
            Ig10Code __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_AlwaysTrust )( 
            Ig10Code __RPC_FAR * This,
            /* [in] */ BOOL newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetLogLevel )( 
            Ig10Code __RPC_FAR * This,
            long logLevel);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetLogFile )( 
            Ig10Code __RPC_FAR * This,
            BSTR logFile,
            /* [retval][out] */ long __RPC_FAR *pvReturn);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Binary )( 
            Ig10Code __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Binary )( 
            Ig10Code __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *AddRecipient )( 
            Ig10Code __RPC_FAR * This,
            BSTR name,
            /* [retval][out] */ long __RPC_FAR *pvReturn);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Encrypt )( 
            Ig10Code __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pvReturn);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ClearRecipient )( 
            Ig10Code __RPC_FAR * This);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Output )( 
            Ig10Code __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Output )( 
            Ig10Code __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_NoVersion )( 
            Ig10Code __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_NoVersion )( 
            Ig10Code __RPC_FAR * This,
            /* [in] */ BOOL newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Comment )( 
            Ig10Code __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Comment )( 
            Ig10Code __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Passphrase )( 
            Ig10Code __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Decrypt )( 
            Ig10Code __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pvReturn);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Export )( 
            Ig10Code __RPC_FAR * This,
            BSTR keyNames,
            /* [retval][out] */ long __RPC_FAR *pvReturn);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EncryptoTo )( 
            Ig10Code __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EncryptoTo )( 
            Ig10Code __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ForceMDC )( 
            Ig10Code __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_ForceMDC )( 
            Ig10Code __RPC_FAR * This,
            /* [in] */ BOOL newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ForceV3Sig )( 
            Ig10Code __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_ForceV3Sig )( 
            Ig10Code __RPC_FAR * This,
            /* [in] */ BOOL newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Keyserver )( 
            Ig10Code __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Keyserver )( 
            Ig10Code __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_HomeDir )( 
            Ig10Code __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_HomeDir )( 
            Ig10Code __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *EncryptFile )( 
            Ig10Code __RPC_FAR * This,
            /* [in] */ BSTR inFile,
            /* [retval][out] */ long __RPC_FAR *pvReturn);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *DecryptFile )( 
            Ig10Code __RPC_FAR * This,
            /* [in] */ BSTR inFile,
            /* [retval][out] */ long __RPC_FAR *pvReturn);
        
        END_INTERFACE
    } Ig10CodeVtbl;

    interface Ig10Code
    {
        CONST_VTBL struct Ig10CodeVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define Ig10Code_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define Ig10Code_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define Ig10Code_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define Ig10Code_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define Ig10Code_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define Ig10Code_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define Ig10Code_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define Ig10Code_get_Plaintext(This,pVal)	\
    (This)->lpVtbl -> get_Plaintext(This,pVal)

#define Ig10Code_put_Plaintext(This,newVal)	\
    (This)->lpVtbl -> put_Plaintext(This,newVal)

#define Ig10Code_get_Ciphertext(This,pVal)	\
    (This)->lpVtbl -> get_Ciphertext(This,pVal)

#define Ig10Code_put_Ciphertext(This,newVal)	\
    (This)->lpVtbl -> put_Ciphertext(This,newVal)

#define Ig10Code_get_Armor(This,pVal)	\
    (This)->lpVtbl -> get_Armor(This,pVal)

#define Ig10Code_put_Armor(This,newVal)	\
    (This)->lpVtbl -> put_Armor(This,newVal)

#define Ig10Code_get_LocalUser(This,pVal)	\
    (This)->lpVtbl -> get_LocalUser(This,pVal)

#define Ig10Code_put_LocalUser(This,newVal)	\
    (This)->lpVtbl -> put_LocalUser(This,newVal)

#define Ig10Code_get_CompressLevel(This,pVal)	\
    (This)->lpVtbl -> get_CompressLevel(This,pVal)

#define Ig10Code_put_CompressLevel(This,newVal)	\
    (This)->lpVtbl -> put_CompressLevel(This,newVal)

#define Ig10Code_get_TextMode(This,pVal)	\
    (This)->lpVtbl -> get_TextMode(This,pVal)

#define Ig10Code_put_TextMode(This,newVal)	\
    (This)->lpVtbl -> put_TextMode(This,newVal)

#define Ig10Code_get_Expert(This,pVal)	\
    (This)->lpVtbl -> get_Expert(This,pVal)

#define Ig10Code_put_Expert(This,newVal)	\
    (This)->lpVtbl -> put_Expert(This,newVal)

#define Ig10Code_get_PGPMode(This,pVal)	\
    (This)->lpVtbl -> get_PGPMode(This,pVal)

#define Ig10Code_put_PGPMode(This,newVal)	\
    (This)->lpVtbl -> put_PGPMode(This,newVal)

#define Ig10Code_get_AlwaysTrust(This,pVal)	\
    (This)->lpVtbl -> get_AlwaysTrust(This,pVal)

#define Ig10Code_put_AlwaysTrust(This,newVal)	\
    (This)->lpVtbl -> put_AlwaysTrust(This,newVal)

#define Ig10Code_SetLogLevel(This,logLevel)	\
    (This)->lpVtbl -> SetLogLevel(This,logLevel)

#define Ig10Code_SetLogFile(This,logFile,pvReturn)	\
    (This)->lpVtbl -> SetLogFile(This,logFile,pvReturn)

#define Ig10Code_get_Binary(This,pVal)	\
    (This)->lpVtbl -> get_Binary(This,pVal)

#define Ig10Code_put_Binary(This,newVal)	\
    (This)->lpVtbl -> put_Binary(This,newVal)

#define Ig10Code_AddRecipient(This,name,pvReturn)	\
    (This)->lpVtbl -> AddRecipient(This,name,pvReturn)

#define Ig10Code_Encrypt(This,pvReturn)	\
    (This)->lpVtbl -> Encrypt(This,pvReturn)

#define Ig10Code_ClearRecipient(This)	\
    (This)->lpVtbl -> ClearRecipient(This)

#define Ig10Code_get_Output(This,pVal)	\
    (This)->lpVtbl -> get_Output(This,pVal)

#define Ig10Code_put_Output(This,newVal)	\
    (This)->lpVtbl -> put_Output(This,newVal)

#define Ig10Code_get_NoVersion(This,pVal)	\
    (This)->lpVtbl -> get_NoVersion(This,pVal)

#define Ig10Code_put_NoVersion(This,newVal)	\
    (This)->lpVtbl -> put_NoVersion(This,newVal)

#define Ig10Code_get_Comment(This,pVal)	\
    (This)->lpVtbl -> get_Comment(This,pVal)

#define Ig10Code_put_Comment(This,newVal)	\
    (This)->lpVtbl -> put_Comment(This,newVal)

#define Ig10Code_put_Passphrase(This,newVal)	\
    (This)->lpVtbl -> put_Passphrase(This,newVal)

#define Ig10Code_Decrypt(This,pvReturn)	\
    (This)->lpVtbl -> Decrypt(This,pvReturn)

#define Ig10Code_Export(This,keyNames,pvReturn)	\
    (This)->lpVtbl -> Export(This,keyNames,pvReturn)

#define Ig10Code_get_EncryptoTo(This,pVal)	\
    (This)->lpVtbl -> get_EncryptoTo(This,pVal)

#define Ig10Code_put_EncryptoTo(This,newVal)	\
    (This)->lpVtbl -> put_EncryptoTo(This,newVal)

#define Ig10Code_get_ForceMDC(This,pVal)	\
    (This)->lpVtbl -> get_ForceMDC(This,pVal)

#define Ig10Code_put_ForceMDC(This,newVal)	\
    (This)->lpVtbl -> put_ForceMDC(This,newVal)

#define Ig10Code_get_ForceV3Sig(This,pVal)	\
    (This)->lpVtbl -> get_ForceV3Sig(This,pVal)

#define Ig10Code_put_ForceV3Sig(This,newVal)	\
    (This)->lpVtbl -> put_ForceV3Sig(This,newVal)

#define Ig10Code_get_Keyserver(This,pVal)	\
    (This)->lpVtbl -> get_Keyserver(This,pVal)

#define Ig10Code_put_Keyserver(This,newVal)	\
    (This)->lpVtbl -> put_Keyserver(This,newVal)

#define Ig10Code_get_HomeDir(This,pVal)	\
    (This)->lpVtbl -> get_HomeDir(This,pVal)

#define Ig10Code_put_HomeDir(This,newVal)	\
    (This)->lpVtbl -> put_HomeDir(This,newVal)

#define Ig10Code_EncryptFile(This,inFile,pvReturn)	\
    (This)->lpVtbl -> EncryptFile(This,inFile,pvReturn)

#define Ig10Code_DecryptFile(This,inFile,pvReturn)	\
    (This)->lpVtbl -> DecryptFile(This,inFile,pvReturn)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE Ig10Code_get_Plaintext_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB Ig10Code_get_Plaintext_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE Ig10Code_put_Plaintext_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB Ig10Code_put_Plaintext_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE Ig10Code_get_Ciphertext_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB Ig10Code_get_Ciphertext_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE Ig10Code_put_Ciphertext_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB Ig10Code_put_Ciphertext_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE Ig10Code_get_Armor_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [retval][out] */ BOOL __RPC_FAR *pVal);


void __RPC_STUB Ig10Code_get_Armor_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE Ig10Code_put_Armor_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [in] */ BOOL newVal);


void __RPC_STUB Ig10Code_put_Armor_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE Ig10Code_get_LocalUser_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB Ig10Code_get_LocalUser_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE Ig10Code_put_LocalUser_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB Ig10Code_put_LocalUser_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE Ig10Code_get_CompressLevel_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB Ig10Code_get_CompressLevel_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE Ig10Code_put_CompressLevel_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [in] */ long newVal);


void __RPC_STUB Ig10Code_put_CompressLevel_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE Ig10Code_get_TextMode_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [retval][out] */ BOOL __RPC_FAR *pVal);


void __RPC_STUB Ig10Code_get_TextMode_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE Ig10Code_put_TextMode_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [in] */ BOOL newVal);


void __RPC_STUB Ig10Code_put_TextMode_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE Ig10Code_get_Expert_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [retval][out] */ BOOL __RPC_FAR *pVal);


void __RPC_STUB Ig10Code_get_Expert_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE Ig10Code_put_Expert_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [in] */ BOOL newVal);


void __RPC_STUB Ig10Code_put_Expert_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE Ig10Code_get_PGPMode_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB Ig10Code_get_PGPMode_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE Ig10Code_put_PGPMode_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [in] */ long newVal);


void __RPC_STUB Ig10Code_put_PGPMode_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE Ig10Code_get_AlwaysTrust_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [retval][out] */ BOOL __RPC_FAR *pVal);


void __RPC_STUB Ig10Code_get_AlwaysTrust_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE Ig10Code_put_AlwaysTrust_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [in] */ BOOL newVal);


void __RPC_STUB Ig10Code_put_AlwaysTrust_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Ig10Code_SetLogLevel_Proxy( 
    Ig10Code __RPC_FAR * This,
    long logLevel);


void __RPC_STUB Ig10Code_SetLogLevel_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Ig10Code_SetLogFile_Proxy( 
    Ig10Code __RPC_FAR * This,
    BSTR logFile,
    /* [retval][out] */ long __RPC_FAR *pvReturn);


void __RPC_STUB Ig10Code_SetLogFile_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE Ig10Code_get_Binary_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB Ig10Code_get_Binary_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE Ig10Code_put_Binary_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB Ig10Code_put_Binary_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Ig10Code_AddRecipient_Proxy( 
    Ig10Code __RPC_FAR * This,
    BSTR name,
    /* [retval][out] */ long __RPC_FAR *pvReturn);


void __RPC_STUB Ig10Code_AddRecipient_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Ig10Code_Encrypt_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pvReturn);


void __RPC_STUB Ig10Code_Encrypt_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Ig10Code_ClearRecipient_Proxy( 
    Ig10Code __RPC_FAR * This);


void __RPC_STUB Ig10Code_ClearRecipient_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE Ig10Code_get_Output_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB Ig10Code_get_Output_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE Ig10Code_put_Output_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB Ig10Code_put_Output_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE Ig10Code_get_NoVersion_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [retval][out] */ BOOL __RPC_FAR *pVal);


void __RPC_STUB Ig10Code_get_NoVersion_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE Ig10Code_put_NoVersion_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [in] */ BOOL newVal);


void __RPC_STUB Ig10Code_put_NoVersion_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE Ig10Code_get_Comment_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB Ig10Code_get_Comment_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE Ig10Code_put_Comment_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB Ig10Code_put_Comment_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE Ig10Code_put_Passphrase_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB Ig10Code_put_Passphrase_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Ig10Code_Decrypt_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pvReturn);


void __RPC_STUB Ig10Code_Decrypt_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Ig10Code_Export_Proxy( 
    Ig10Code __RPC_FAR * This,
    BSTR keyNames,
    /* [retval][out] */ long __RPC_FAR *pvReturn);


void __RPC_STUB Ig10Code_Export_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE Ig10Code_get_EncryptoTo_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB Ig10Code_get_EncryptoTo_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE Ig10Code_put_EncryptoTo_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB Ig10Code_put_EncryptoTo_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE Ig10Code_get_ForceMDC_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [retval][out] */ BOOL __RPC_FAR *pVal);


void __RPC_STUB Ig10Code_get_ForceMDC_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE Ig10Code_put_ForceMDC_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [in] */ BOOL newVal);


void __RPC_STUB Ig10Code_put_ForceMDC_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE Ig10Code_get_ForceV3Sig_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [retval][out] */ BOOL __RPC_FAR *pVal);


void __RPC_STUB Ig10Code_get_ForceV3Sig_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE Ig10Code_put_ForceV3Sig_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [in] */ BOOL newVal);


void __RPC_STUB Ig10Code_put_ForceV3Sig_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE Ig10Code_get_Keyserver_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB Ig10Code_get_Keyserver_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE Ig10Code_put_Keyserver_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB Ig10Code_put_Keyserver_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE Ig10Code_get_HomeDir_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB Ig10Code_get_HomeDir_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE Ig10Code_put_HomeDir_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB Ig10Code_put_HomeDir_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Ig10Code_EncryptFile_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [in] */ BSTR inFile,
    /* [retval][out] */ long __RPC_FAR *pvReturn);


void __RPC_STUB Ig10Code_EncryptFile_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Ig10Code_DecryptFile_Proxy( 
    Ig10Code __RPC_FAR * This,
    /* [in] */ BSTR inFile,
    /* [retval][out] */ long __RPC_FAR *pvReturn);


void __RPC_STUB Ig10Code_DecryptFile_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __Ig10Code_INTERFACE_DEFINED__ */



#ifndef __GDGPGLib_LIBRARY_DEFINED__
#define __GDGPGLib_LIBRARY_DEFINED__

/* library GDGPGLib */
/* [helpstring][version][uuid] */ 


EXTERN_C const IID LIBID_GDGPGLib;

EXTERN_C const CLSID CLSID_GDGPG;

#ifdef __cplusplus

class DECLSPEC_UUID("321F09FC-E2FD-409B-B8D1-60FA7DCDA531")
GDGPG;
#endif

#ifndef ___Ig10CodeEvents_DISPINTERFACE_DEFINED__
#define ___Ig10CodeEvents_DISPINTERFACE_DEFINED__

/* dispinterface _Ig10CodeEvents */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID__Ig10CodeEvents;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("69D0503E-250B-4994-B388-975A604EADDD")
    _Ig10CodeEvents : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _Ig10CodeEventsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            _Ig10CodeEvents __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            _Ig10CodeEvents __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            _Ig10CodeEvents __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            _Ig10CodeEvents __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            _Ig10CodeEvents __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            _Ig10CodeEvents __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            _Ig10CodeEvents __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        END_INTERFACE
    } _Ig10CodeEventsVtbl;

    interface _Ig10CodeEvents
    {
        CONST_VTBL struct _Ig10CodeEventsVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _Ig10CodeEvents_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define _Ig10CodeEvents_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define _Ig10CodeEvents_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define _Ig10CodeEvents_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define _Ig10CodeEvents_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define _Ig10CodeEvents_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define _Ig10CodeEvents_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___Ig10CodeEvents_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_g10Code;

#ifdef __cplusplus

class DECLSPEC_UUID("658BFF7C-83DD-41F8-A712-93728A349A7F")
g10Code;
#endif
#endif /* __GDGPGLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

unsigned long             __RPC_USER  BSTR_UserSize(     unsigned long __RPC_FAR *, unsigned long            , BSTR __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  BSTR_UserMarshal(  unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, BSTR __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  BSTR_UserUnmarshal(unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, BSTR __RPC_FAR * ); 
void                      __RPC_USER  BSTR_UserFree(     unsigned long __RPC_FAR *, BSTR __RPC_FAR * ); 

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif
