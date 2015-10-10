/***************************************************************************
 * UTILITIES.H
 *
 * DSOFramer: Common Utilities and Macros (Shared)
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
#ifndef DS_UTILITIES_H
#define DS_UTILITIES_H

#include <commdlg.h>
#include <oledlg.h>

////////////////////////////////////////////////////////////////////////
// Fixed Win32 Errors as HRESULTs
//
#define E_WIN32_BUFFERTOOSMALL    0x8007007A   //HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER)
#define E_WIN32_ACCESSVIOLATION   0x800701E7   //HRESULT_FROM_WIN32(ERROR_INVALID_ADDRESS)
#define E_WIN32_LASTERROR        (0x80070000 | GetLastError()) // Assured Error with last Win32 code
#define E_VBA_NOREMOTESERVER      0x800A01CE

////////////////////////////////////////////////////////////////////////
// Heap Allocation
//
STDAPI_(LPVOID) DsoMemAlloc(DWORD cbSize);
STDAPI_(void)   DsoMemFree(LPVOID ptr);

// Override new/delete to use our task allocator
// (removing CRT dependency will improve code performance and size)...
void * _cdecl operator new(size_t size);
void  _cdecl operator delete(void *ptr);

////////////////////////////////////////////////////////////////////////
// String Manipulation Functions
//
STDAPI DsoConvertToUnicodeEx(LPCSTR pszMbcsString, DWORD cbMbcsLen, LPWSTR pwszUnicode, DWORD cbUniLen, UINT uiCodePage);
STDAPI DsoConvertToMBCSEx(LPCWSTR pwszUnicodeString, DWORD cbUniLen, LPSTR pwszMbcsString, DWORD cbMbcsLen, UINT uiCodePage);

STDAPI_(LPWSTR) DsoConvertToLPWSTR(LPCSTR pszMbcsString);
STDAPI_(BSTR)   DsoConvertToBSTR(LPCSTR pszMbcsString);
STDAPI_(LPWSTR) DsoConvertToLPOLESTR(LPCWSTR pwszUnicodeString);
STDAPI_(LPSTR)  DsoConvertToMBCS(LPCWSTR pwszUnicodeString);
STDAPI_(UINT)   DsoCompareStringsEx(LPCWSTR pwsz1, INT cch1, LPCWSTR pwsz2, INT cch2);
STDAPI_(LPWSTR) DsoCopyString(LPCWSTR pwszString);
STDAPI_(LPWSTR) DsoCopyStringCat(LPCWSTR pwszString1, LPCWSTR pwszString2);
STDAPI_(LPWSTR) DsoCopyStringCatEx(LPCWSTR pwszBaseString, UINT cStrs, LPCWSTR *ppwszStrs);
STDAPI_(LPSTR)  DsoCLSIDtoLPSTR(REFCLSID clsid);

////////////////////////////////////////////////////////////////////////
// URL Helpers
//
STDAPI_(BOOL) LooksLikeLocalFile(LPCWSTR pwsz);
STDAPI_(BOOL) LooksLikeUNC(LPCWSTR pwsz);
STDAPI_(BOOL) LooksLikeHTTP(LPCWSTR pwsz);
STDAPI_(BOOL) GetTempPathForURLDownload(WCHAR* pwszURL, WCHAR** ppwszLocalFile);
STDAPI URLDownloadFile(LPUNKNOWN punk, WCHAR* pwszURL, WCHAR* pwszLocalFile);

////////////////////////////////////////////////////////////////////////
// OLE Conversion Functions
//
STDAPI_(void) DsoHimetricToPixels(LONG* px, LONG* py);
STDAPI_(void) DsoPixelsToHimetric(LONG* px, LONG* py);

////////////////////////////////////////////////////////////////////////
// GDI Helper Functions
//
STDAPI_(HBITMAP) DsoGetBitmapFromWindow(HWND hwnd);

////////////////////////////////////////////////////////////////////////
// Windows Helper Functions
//
STDAPI_(BOOL) IsWindowChild(HWND hwndParent, HWND hwndChild);

////////////////////////////////////////////////////////////////////////
// OLE/Typelib Function Wrappers
//
STDAPI DsoGetTypeInfoEx(REFGUID rlibid, LCID lcid, WORD wVerMaj, WORD wVerMin, HMODULE hResource, REFGUID rguid, ITypeInfo** ppti);
STDAPI DsoDispatchInvoke(LPDISPATCH pdisp, LPOLESTR pwszname, DISPID dspid, WORD wflags, DWORD cargs, VARIANT* rgargs, VARIANT* pvtret);
STDAPI DsoReportError(HRESULT hr, LPWSTR pwszCustomMessage, EXCEPINFO* peiDispEx);

////////////////////////////////////////////////////////////////////////
// Unicode Win32 API wrappers (handles thunk down for Win9x)
//
STDAPI_(BOOL) FFileExists(WCHAR* wzPath);
STDAPI_(BOOL) FOpenLocalFile(WCHAR* wzFilePath, DWORD dwAccess, DWORD dwShareMode, DWORD dwCreate, HANDLE* phFile);
STDAPI_(BOOL) FPerformShellOp(DWORD dwOp, WCHAR* wzFrom, WCHAR* wzTo);
STDAPI_(BOOL) FGetModuleFileName(HMODULE hModule, WCHAR** wzFileName);
STDAPI_(BOOL) FIsIECacheFile(LPWSTR pwszFile);
STDAPI_(BOOL) FDrawText(HDC hdc, WCHAR* pwsz, LPRECT prc, UINT fmt);
STDAPI_(BOOL) FSetRegKeyValue(HKEY hk, WCHAR* pwsz);

STDAPI_(BOOL) FOpenPrinter(LPCWSTR pwszPrinter, LPHANDLE phandle);
STDAPI_(BOOL) FGetPrinterSettings(HANDLE hprinter, LPWSTR *ppwszProcessor, LPWSTR *ppwszDevice, LPWSTR *ppwszOutput, LPDEVMODEW *ppdvmode, DWORD *pcbSize);

STDAPI DsoGetFileFromUser(HWND hwndOwner, LPCWSTR pwzTitle, DWORD dwFlags, 
       LPCWSTR pwzFilter, DWORD dwFiltIdx, LPCWSTR pwszDefExt, LPCWSTR pwszCurrentItem, BOOL fShowSave,
       BSTR *pbstrFile, BOOL *pfReadOnly);

STDAPI DsoGetOleInsertObjectFromUser(HWND hwndOwner, LPCWSTR pwzTitle, DWORD dwFlags, 
        BOOL fDocObjectOnly, BOOL fAllowControls, BSTR *pbstrResult, UINT *ptype);

////////////////////////////////////////////////////////////////////////
// Common macros -- Used to make code more readable.
//
#define SEH_TRY           __try {
#define SEH_EXCEPT(hr)    } __except(GetExceptionCode() == EXCEPTION_ACCESS_VIOLATION){hr = E_WIN32_ACCESSVIOLATION;}
#define SEH_EXCEPT_NULL   } __except(GetExceptionCode() == EXCEPTION_ACCESS_VIOLATION){}
#define SEH_START_FINALLY } __finally {
#define SEH_END_FINALLY   }

#define RETURN_ON_FAILURE(x)    if (FAILED(x)) return (x)
#define GOTO_ON_FAILURE(x, lbl) if (FAILED(x)) goto lbl
#define CHECK_NULL_RETURN(v, e) if ((v) == NULL) return (e)

#define SAFE_ADDREF_INTERFACE     if (x) { (x)->AddRef(); }
#define SAFE_RELEASE_INTERFACE(x) if (x) { (x)->Release(); (x) = NULL; }
#define SAFE_SET_INTERFACE(x, y)  if (((x) = (y)) != NULL) ((IUnknown*)(x))->AddRef()
#define SAFE_FREESTRING(s)        if (s) { DsoMemFree(s); (s) = NULL; }
#define SAFE_FREEBSTR(s)          if (s) { SysFreeString(s); (s) = NULL; }

VARIANT*   __fastcall DsoPVarFromPVarRef(VARIANT* px);
BOOL       __fastcall DsoIsVarParamMissing(VARIANT* px);
LPWSTR     __fastcall DsoPVarWStrFromPVar(VARIANT* px);
SAFEARRAY* __fastcall DsoPVarArrayFromPVar(VARIANT* px);
IUnknown*  __fastcall DsoPVarUnkFromPVar(VARIANT* px);
SHORT      __fastcall DsoPVarShortFromPVar(VARIANT* px, SHORT fdef);
LONG       __fastcall DsoPVarLongFromPVar(VARIANT* px, LONG fdef);
BOOL       __fastcall DsoPVarBoolFromPVar(VARIANT* px, BOOL fdef);

#define PARAM_IS_MISSING(x)      DsoIsVarParamMissing(DsoPVarFromPVarRef((x)))
#define LPWSTR_FROM_VARIANT(x)   DsoPVarWStrFromPVar(DsoPVarFromPVarRef(&(x)))
#define LONG_FROM_VARIANT(x, y)  DsoPVarLongFromPVar(DsoPVarFromPVarRef(&(x)), (y))
#define BOOL_FROM_VARIANT(x, y)  DsoPVarBoolFromPVar(DsoPVarFromPVarRef(&(x)), (y))
#define PUNK_FROM_VARIANT(x)     DsoPVarUnkFromPVar(DsoPVarFromPVarRef(&(x)))
#define PSARRAY_FROM_VARIANT(x)  DsoPVarArrayFromPVar(DsoPVarFromPVarRef(&(x)))

#define ASCII_UPPERCASE(x) ((((x) > 96) && ((x) < 123)) ? (x) - 32 : (x))
#define ASCII_LOWERCASE(x) ((((x) > 64) && ((x) <  91)) ? (x) + 32 : (x))

////////////////////////////////////////////////////////////////////////
// Debug macros
//
#ifdef _DEBUG

#define ASSERT(x)  if(!(x)) DebugBreak()
#define ODS(x)	OutputDebugString(x)

#define TRACE1(sz, arg1) { \
	CHAR ach[1024]; \
	wsprintf(ach, (sz), (arg1)); \
	ODS(ach); }

#define TRACE2(sz, arg1, arg2) { \
	CHAR ach[1024]; \
	wsprintf(ach, (sz), (arg1), (arg2)); \
	ODS(ach); }

#define TRACE3(sz, arg1, arg2, arg3) { \
	CHAR ach[1024]; \
	wsprintf(ach, (sz), (arg1), (arg2), (arg3)); \
	ODS(ach); }

#define TRACE_LPRECT(sz, lprc) { \
	CHAR ach[1024]; \
	wsprintf(ach, "RECT %s - left=%d, top=%d, right=%d, bottom=%d\n", \
		(sz), (lprc)->left, (lprc)->top, (lprc)->right, (lprc)->bottom); \
	ODS(ach); }

#else // !defined(_DEBUG)

#define ASSERT(x)
#define ODS(x) 
#define TRACE1(sz, arg1)
#define TRACE2(sz, arg1, arg2)
#define TRACE3(sz, arg1, arg2, arg3)
#define TRACE_LPRECT(sz, lprc)

#endif // (_DEBUG)

////////////////////////////////////////////////////////////////////////
// Macros for Nested COM Interfaces 
//
#ifdef _DEBUG
#define DEFINE_REFCOUNT ULONG    m_cRef;
#define IMPLEMENT_DEBUG_ADDREF 	 m_cRef++;
#define IMPLEMENT_DEBUG_RELEASE(x)  ASSERT(m_cRef > 0); m_cRef--; if (m_cRef == 0){ODS(" > I" #x " released\n");}
#define IMPLEMENT_DEBUG_REFSET   m_cRef = 0;
#define IMPLEMENT_DEBUG_REFCHECK(x) ASSERT(m_cRef == 0); if (m_cRef != 0){ODS(" * I" #x " NOT released!!\n");}
#else
#define DEFINE_REFCOUNT
#define IMPLEMENT_DEBUG_ADDREF
#define IMPLEMENT_DEBUG_RELEASE(x)
#define IMPLEMENT_DEBUG_REFSET
#define IMPLEMENT_DEBUG_REFCHECK(x)
#endif /* !_DEBUG */

#define BEGIN_INTERFACE_PART(localClass, baseClass) \
class X##localClass : public baseClass \
{ public: X##localClass(){IMPLEMENT_DEBUG_REFSET} \
         ~X##localClass(){IMPLEMENT_DEBUG_REFCHECK(##localClass)} \
		STDMETHOD(QueryInterface)(REFIID iid, PVOID* ppvObj); \
		STDMETHOD_(ULONG, AddRef)(); \
		STDMETHOD_(ULONG, Release)(); \
        DEFINE_REFCOUNT

#define END_INTERFACE_PART(localClass) \
} m_x##localClass; \
friend class X##localClass;

#define METHOD_PROLOGUE(theClass, localClass) \
	theClass* pThis = \
		((theClass*)(((BYTE*)this) - (size_t)&(((theClass*)0)->m_x##localClass)));

#define IMPLEMENT_INTERFACE_UNKNOWN(theClass, localClass) \
	ULONG theClass::X##localClass::AddRef() { \
		METHOD_PROLOGUE(theClass, localClass) \
		IMPLEMENT_DEBUG_ADDREF \
		return pThis->AddRef(); \
	} \
	ULONG theClass::X##localClass::Release() { \
		METHOD_PROLOGUE(theClass, localClass) \
		IMPLEMENT_DEBUG_RELEASE(##localClass) \
		return pThis->Release(); \
	} \
	STDMETHODIMP theClass::X##localClass::QueryInterface(REFIID iid, void **ppvObj) { \
		METHOD_PROLOGUE(theClass, localClass) \
		return pThis->QueryInterface(iid, ppvObj); \
	}


#endif //DS_UTILITIES_H