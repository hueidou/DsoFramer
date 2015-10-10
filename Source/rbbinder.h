/***************************************************************************
 * RBBINDER.H
 *
 *  DSOFramer: Internet Publishing Provider (MSDAIPP) Compatible Header
 *
 *  Copyright ©1999-2004; Microsoft Corporation. All rights reserved.
 *  Written by Microsoft Developer Support Office Integration (PSS DSOI)
 * 
 *  This code is provided via KB 311765 as a sample. It is not a formal
 *  product and has not been tested with all containers or servers. Use it
 *  for educational purposes only. See the EULA.TXT file included in the
 *  KB download for full terms of use and restrictions.
 *
 *  THIS CODE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
 *  EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED
 *  WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
 *
 ***************************************************************************/
#ifndef DS_RBBINDER_H
#define DS_RBBINDER_H

////////////////////////////////////////////////////////////////////////
// The Microsoft OLEDB Provider for Internet Publishing (MSDAIPP) is now
// standard with MDAC 2.5. However, it will run with MDAC 2.1, so this 
// header will allow us to safely compile without 2.5.
// 
#define OLEDBVER    0x0250

#include <oledb.h>
#include <oledberr.h>

#ifdef __cplusplus
extern "C" {
#endif

#include <pshpack2.h>	// 2-byte structure packing

#ifndef BEGIN_INTERFACE
#define BEGIN_INTERFACE
#endif


////////////////////////////////////////////////////////////////////////
// MSDAIPP Specific GUIDs (not defined in OLEDB.H)
//
DEFINE_GUID(CLSID_MSDAIPP_DSO,      0xAF320921L, 0x9381, 0x11d1, 0x9C, 0x3C, 0x00, 0x00, 0xF8, 0x75, 0xAC, 0x61);
DEFINE_GUID(CLSID_MSDAIPP_BINDER,   0xE1D2BF40L, 0xA96B, 0x11d1, 0x9C, 0x6B, 0x00, 0x00, 0xF8, 0x75, 0xAC, 0x61);
DEFINE_GUID(DBPROPSET_MSDAIPP_INIT, 0x8F1033E3L, 0xB2CD, 0x11d1, 0x9C, 0x74, 0x00, 0x00, 0xF8, 0x75, 0xAC, 0x61);


////////////////////////////////////////////////////////////////////////
// OLEDB Additional defines -- included for machines with MDAC 2.0/2.1
//

////////////////////////////////////////////////////////////////////////
// OLEDB 2.5 GUIDS Redefined for use here; this is to avoid linker errors
//  on machines that have different versions of MDAC libs. 
//
DEFINE_GUID(IIDX_IBindResource, 0x0c733ab1L, 0x2a1c, 0x11ce, 0xad, 0xe5, 0x00, 0xaa, 0x00, 0x44, 0x77, 0x3d);
DEFINE_GUID(IIDX_IDBBinderProperties, 0x0c733ab3L, 0x2a1c, 0x11ce, 0xad, 0xe5, 0x00, 0xaa, 0x00, 0x44, 0x77, 0x3d);
DEFINE_GUID(IIDX_ICreateRow, 0x0c733ab2L, 0x2a1c, 0x11ce, 0xad, 0xe5, 0x00, 0xaa, 0x00, 0x44, 0x77, 0x3d);
DEFINE_GUID(IIDX_IAuthenticate, 0x79eac9d0L, 0xbaf9, 0x11ce, 0x8c, 0x82, 0x00, 0xaa, 0x00, 0x4b, 0xa9, 0x0b);
DEFINE_GUID(DBGUIDX_STREAM, 0xc8b522f9L, 0x5cf3, 0x11ce, 0xad, 0xe5, 0x00, 0xaa, 0x00, 0x44, 0x77, 0x3d);


///////////////////////////////////////////////////////////////////////
// MSDAIPP Binding Interfaces (all this should be standard in MDAC 2.5)
//
#ifndef __IBindResource_FWD_DEFINED__
#define __IBindResource_FWD_DEFINED__

typedef interface IBindResource IBindResource;

typedef DWORD DBBINDURLFLAG;
enum DBBINDURLFLAGENUM
{
	DBBINDURLFLAG_READ	= 0x1L,
	DBBINDURLFLAG_WRITE	= 0x2L,
	DBBINDURLFLAG_READWRITE	= 0x3L,
	DBBINDURLFLAG_SHARE_DENY_READ	= 0x4L,
	DBBINDURLFLAG_SHARE_DENY_WRITE	= 0x8L,
	DBBINDURLFLAG_SHARE_EXCLUSIVE	= 0xcL,
	DBBINDURLFLAG_SHARE_DENY_NONE	= 0x10L,
	DBBINDURLFLAG_ASYNCHRONOUS	= 0x1000L,
	DBBINDURLFLAG_COLLECTION	= 0x2000L,
	DBBINDURLFLAG_DELAYFETCHSTREAM	= 0x4000L,
	DBBINDURLFLAG_DELAYFETCHCOLUMNS	= 0x8000L,
	DBBINDURLFLAG_RECURSIVE	= 0x400000L,
	DBBINDURLFLAG_OUTPUT	= 0x800000L,
	DBBINDURLFLAG_WAITFORINIT	= 0x1000000L,
	DBBINDURLFLAG_OPENIFEXISTS	= 0x2000000L,
	DBBINDURLFLAG_OVERWRITE	= 0x4000000L,
	DBBINDURLFLAG_ISSTRUCTUREDDOCUMENT	= 0x8000000L
};

typedef DWORD DBBINDURLSTATUS;
enum DBBINDURLSTATUSENUM
{
	DBBINDURLSTATUS_S_OK	= 0L,
	DBBINDURLSTATUS_S_DENYNOTSUPPORTED	= 0x1L,
	DBBINDURLSTATUS_S_DENYTYPENOTSUPPORTED	= 0x4L,
	DBBINDURLSTATUS_S_REDIRECTED	= 0x8L
};

enum DBPROP_OLEDB25_RB
{
	DBPROP_INIT_BINDFLAGS	= 0x10eL,
	DBPROP_INIT_LOCKOWNER	= 0x10fL
};

typedef ULONG DBCOUNTITEM;

typedef struct tagDBIMPLICITSESSION
    {
    IUnknown __RPC_FAR *pUnkOuter;
    IID __RPC_FAR *piid;
    IUnknown __RPC_FAR *pSession;
	} DBIMPLICITSESSION;

#endif //__IBindResource_FWD_DEFINED__


#ifndef __IBindResource_INTERFACE_DEFINED__
#define __IBindResource_INTERFACE_DEFINED__

#undef INTERFACE
#define INTERFACE IBindResource
DECLARE_INTERFACE_(IBindResource, IUnknown)
{
BEGIN_INTERFACE
#ifndef NO_BASEINTERFACE_FUNCS
    /* IUnknown methods */
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, LPVOID *ppvObj) PURE;
    STDMETHOD_(ULONG,AddRef)(THIS) PURE;
    STDMETHOD_(ULONG,Release)(THIS) PURE;
#endif
    /* IBindResource methods */
    STDMETHOD(Bind)(THIS_ 
            /* [in] */ IUnknown __RPC_FAR *pUnkOuter,
            /* [in] */ LPCOLESTR pwszURL,
            /* [in] */ DBBINDURLFLAG dwBindURLFlags,
            /* [in] */ REFGUID rguid,
            /* [in] */ REFIID riid,
            /* [in] */ IAuthenticate __RPC_FAR *pAuthenticate,
            /* [unique][out][in] */ DBIMPLICITSESSION __RPC_FAR *pImplSession,
            /* [unique][out][in] */ DBBINDURLSTATUS __RPC_FAR *pdwBindStatus,
            /* [iid_is][out] */ IUnknown __RPC_FAR *__RPC_FAR *ppUnk) PURE;
};

#endif //__IBindResource_INTERFACE_DEFINED__



#ifndef __ICreateRow_INTERFACE_DEFINED__
#define __ICreateRow_INTERFACE_DEFINED__

#undef INTERFACE
#define INTERFACE ICreateRow
DECLARE_INTERFACE_(ICreateRow, IUnknown)
{
BEGIN_INTERFACE
#ifndef NO_BASEINTERFACE_FUNCS
    /* IUnknown methods */
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, LPVOID *ppvObj) PURE;
    STDMETHOD_(ULONG,AddRef)(THIS) PURE;
    STDMETHOD_(ULONG,Release)(THIS) PURE;
#endif
    /* ICreateRow methods */
    STDMETHOD(CreateRow)(THIS_ 
        /* [unique][in] */ IUnknown __RPC_FAR *pUnkOuter,
        /* [in] */ LPCOLESTR pwszURL,
        /* [in] */ DBBINDURLFLAG dwBindURLFlags,
        /* [in] */ REFGUID rguid,
        /* [in] */ REFIID riid,
        /* [unique][in] */ IAuthenticate __RPC_FAR *pAuthenticate,
        /* [unique][out][in] */ IUnknown __RPC_FAR *pImplSession,
        /* [unique][out][in] */ DBBINDURLSTATUS __RPC_FAR *pdwBindStatus,
        /* [out] */ LPOLESTR __RPC_FAR *ppwszNewURL,
        /* [iid_is][out] */ IUnknown __RPC_FAR *__RPC_FAR *ppUnk) PURE;
};

#endif //__ICreateRow_INTERFACE_DEFINED__


#ifndef __IDBBinderProperties_INTERFACE_DEFINED__
#define __IDBBinderProperties_INTERFACE_DEFINED__

#undef INTERFACE
#define INTERFACE IDBBinderProperties
DECLARE_INTERFACE_(IDBBinderProperties, IDBProperties)
{
BEGIN_INTERFACE
#ifndef NO_BASEINTERFACE_FUNCS
    /* IDBProperties methods */
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, LPVOID *ppvObj) PURE;
    STDMETHOD_(ULONG,AddRef)(THIS) PURE;
    STDMETHOD_(ULONG,Release)(THIS) PURE;

    STDMETHOD(GetProperties)(THIS_ ULONG cPropertyIDSets, const DBPROPIDSET __RPC_FAR rgPropertyIDSets[], ULONG __RPC_FAR *pcPropertySets, DBPROPSET __RPC_FAR *__RPC_FAR *prgPropertySets) PURE;
    STDMETHOD(GetPropertyInfo)(THIS_ ULONG cPropertyIDSets, const DBPROPIDSET __RPC_FAR rgPropertyIDSets[], ULONG __RPC_FAR *pcPropertyInfoSets, DBPROPINFOSET __RPC_FAR *__RPC_FAR *prgPropertyInfoSets, OLECHAR __RPC_FAR *__RPC_FAR *ppDescBuffer) PURE;
    STDMETHOD(SetProperties)(THIS_ ULONG cPropertySets, DBPROPSET __RPC_FAR rgPropertySets[]) PURE;
#endif
    /* IDBBinderProperties methods */
    STDMETHOD(Reset)(THIS) PURE;
};

#endif //__IDBBinderProperties_INTERFACE_DEFINED__


////////////////////////////////////////////////////////////////////////
// IAuthenticate (borrowed from urlmon.h to avoid extra includes)
//
#ifndef __IAuthenticate_INTERFACE_DEFINED__
#define __IAuthenticate_INTERFACE_DEFINED__

#undef INTERFACE
#define INTERFACE IAuthenticate
DECLARE_INTERFACE_(IAuthenticate, IUnknown)
{
BEGIN_INTERFACE
#ifndef NO_BASEINTERFACE_FUNCS
    /* IUnknown methods */
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, LPVOID *ppvObj) PURE;
    STDMETHOD_(ULONG,AddRef)(THIS) PURE;
    STDMETHOD_(ULONG,Release)(THIS) PURE;
#endif
    /* IAuthenticate methods */
    STDMETHOD(Authenticate)(THIS_ HWND __RPC_FAR *phwnd, LPWSTR __RPC_FAR *pszUsername, LPWSTR __RPC_FAR *pszPassword) PURE;
};

#endif //__IAuthenticate_INTERFACE_DEFINED__

//OLEDB 2.1 Error values
#ifndef DB_E_READONLY
#define DB_E_READONLY                    ((HRESULT)0x80040E94L)
#define DB_E_RESOURCELOCKED              ((HRESULT)0x80040E92L)
#define DB_E_CANNOTCONNECT               ((HRESULT)0x80040E96L)
#define DB_E_TIMEOUT                     ((HRESULT)0x80040E97L)
#define DB_E_RESOURCEEXISTS              ((HRESULT)0x80040E98L)
#define DB_E_OUTOFSPACE                  ((HRESULT)0x80040E9AL)
#endif

//OLEDB 2.5 Error values
#ifndef DB_SEC_E_SAFEMODE_DENIED
#define DB_SEC_E_SAFEMODE_DENIED         ((HRESULT)0x80040E9BL)
#endif

#ifdef __cplusplus
} //extern "C"
#endif

#endif // DS_RBBINDER_H

