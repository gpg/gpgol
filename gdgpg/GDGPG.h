/* this ALWAYS GENERATED file contains the definitions for the interfaces */


/* File created by MIDL compiler version 5.01.0164 */
/* at Mon Sep 06 12:34:18 2004
 */
/* Compiler settings for C:\Source\GData\Gdgpg\GDGPG.idl:
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


#ifndef __GDGPG_FWD_DEFINED__
#define __GDGPG_FWD_DEFINED__

#ifdef __cplusplus
typedef class GDGPG GDGPG;
#else
typedef struct GDGPG GDGPG;
#endif /* __cplusplus */

#endif 	/* __GDGPG_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

void __RPC_FAR * __RPC_USER MIDL_user_allocate(size_t);
void __RPC_USER MIDL_user_free( void __RPC_FAR * ); 

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
