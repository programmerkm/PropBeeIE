

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 6.00.0366 */
/* at Sun Oct 11 10:53:47 2009
 */
/* Compiler settings for .\PropBeeIE.idl:
    Oicf, W1, Zp8, env=Win32 (32b run)
    protocol : dce , ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
//@@MIDL_FILE_HEADING(  )

#pragma warning( disable: 4049 )  /* more than 64k source lines */


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

#ifndef __PropBeeIE_h__
#define __PropBeeIE_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __IPropBee_FWD_DEFINED__
#define __IPropBee_FWD_DEFINED__
typedef interface IPropBee IPropBee;
#endif 	/* __IPropBee_FWD_DEFINED__ */


#ifndef __PropBee_FWD_DEFINED__
#define __PropBee_FWD_DEFINED__

#ifdef __cplusplus
typedef class PropBee PropBee;
#else
typedef struct PropBee PropBee;
#endif /* __cplusplus */

#endif 	/* __PropBee_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

#ifdef __cplusplus
extern "C"{
#endif 

void * __RPC_USER MIDL_user_allocate(size_t);
void __RPC_USER MIDL_user_free( void * ); 

#ifndef __IPropBee_INTERFACE_DEFINED__
#define __IPropBee_INTERFACE_DEFINED__

/* interface IPropBee */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IPropBee;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("738BD3F6-AB09-4DEE-A07D-E55659B64163")
    IPropBee : public IDispatch
    {
    public:
    };
    
#else 	/* C style interface */

    typedef struct IPropBeeVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IPropBee * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IPropBee * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IPropBee * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IPropBee * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IPropBee * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IPropBee * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IPropBee * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        END_INTERFACE
    } IPropBeeVtbl;

    interface IPropBee
    {
        CONST_VTBL struct IPropBeeVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IPropBee_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IPropBee_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IPropBee_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IPropBee_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IPropBee_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IPropBee_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IPropBee_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IPropBee_INTERFACE_DEFINED__ */



#ifndef __PropBeeIELib_LIBRARY_DEFINED__
#define __PropBeeIELib_LIBRARY_DEFINED__

/* library PropBeeIELib */
/* [helpstring][version][uuid] */ 


EXTERN_C const IID LIBID_PropBeeIELib;

EXTERN_C const CLSID CLSID_PropBee;

#ifdef __cplusplus

class DECLSPEC_UUID("ADCC389F-B74A-4881-B5D5-ABDAF58E646F")
PropBee;
#endif
#endif /* __PropBeeIELib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


