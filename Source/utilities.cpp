/***************************************************************************
 * UTILITIES.CPP
 *
 * Shared helper functions and routines.
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
#include "dsoframer.h"

////////////////////////////////////////////////////////////////////////
// Core Utility Functions
//
////////////////////////////////////////////////////////////////////////
// Heap Allocation (Private Heap)
//
extern HANDLE v_hPrivateHeap;
STDAPI_(LPVOID) DsoMemAlloc(DWORD cbSize)
{
    CHECK_NULL_RETURN(v_hPrivateHeap, NULL);
    return HeapAlloc(v_hPrivateHeap, HEAP_ZERO_MEMORY, cbSize);
}

STDAPI_(void) DsoMemFree(LPVOID ptr)
{
    if ((v_hPrivateHeap) && (ptr))
        HeapFree(v_hPrivateHeap, 0, ptr);
}

void * _cdecl operator new(size_t size){ return DsoMemAlloc(size);}
void  _cdecl operator delete(void *ptr){ DsoMemFree(ptr); }
int __cdecl _purecall(){__asm{int 3}; return 0;}

////////////////////////////////////////////////////////////////////////
// Global String Functions
//
////////////////////////////////////////////////////////////////////////
// DsoConvertToUnicodeEx
//
STDAPI DsoConvertToUnicodeEx(LPCSTR pszMbcsString, DWORD cbMbcsLen, LPWSTR pwszUnicode, DWORD cbUniLen, UINT uiCodePage)
{
	DWORD cbRet;
	UINT iCode = CP_ACP;

	if (IsValidCodePage(uiCodePage))
		iCode = uiCodePage;

	CHECK_NULL_RETURN(pwszUnicode,    E_POINTER);
	pwszUnicode[0] = L'\0';

	CHECK_NULL_RETURN(pszMbcsString,  E_POINTER);
	CHECK_NULL_RETURN(cbMbcsLen,      E_INVALIDARG);
	CHECK_NULL_RETURN(cbUniLen,       E_INVALIDARG);

	cbRet = MultiByteToWideChar(iCode, 0, pszMbcsString, cbMbcsLen, pwszUnicode, cbUniLen);
	if (cbRet == 0)	return E_WIN32_LASTERROR;

	pwszUnicode[cbRet] = L'\0';
	return S_OK;
}

////////////////////////////////////////////////////////////////////////
// DsoConvertToMBCSEx
//
STDAPI DsoConvertToMBCSEx(LPCWSTR pwszUnicodeString, DWORD cbUniLen, LPSTR pszMbcsString, DWORD cbMbcsLen, UINT uiCodePage)
{
	DWORD cbRet;
	UINT iCode = CP_ACP;

	if (IsValidCodePage(uiCodePage))
		iCode = uiCodePage;

	CHECK_NULL_RETURN(pszMbcsString,     E_POINTER);
	pszMbcsString[0] = L'\0';

	CHECK_NULL_RETURN(pwszUnicodeString, E_POINTER);
	CHECK_NULL_RETURN(cbMbcsLen,         E_INVALIDARG);
	CHECK_NULL_RETURN(cbUniLen,          E_INVALIDARG);

	cbRet = WideCharToMultiByte(iCode, 0, pwszUnicodeString, -1, pszMbcsString, cbMbcsLen, NULL, NULL);
	if (cbRet == 0)	return E_WIN32_LASTERROR;

	pszMbcsString[cbRet] = '\0';
	return S_OK;
}

////////////////////////////////////////////////////////////////////////
// DsoConvertToLPWSTR
//
//  Takes a MBCS string and returns a LPWSTR allocated on private heap.
//
STDAPI_(LPWSTR) DsoConvertToLPWSTR(LPCSTR pszMbcsString)
{
	LPWSTR pwsz = NULL;
	UINT cblen, cbnew;

	if ((pszMbcsString) && 
        ((cblen = lstrlen(pszMbcsString)) > 0))
	{
		cbnew = ((cblen + 1) * sizeof(WCHAR));
		if ((pwsz = (LPWSTR)DsoMemAlloc(cbnew)) != NULL) 
		{
			if (FAILED(DsoConvertToUnicodeEx(pszMbcsString, cblen, pwsz, cbnew, GetACP())))
            {
			    DsoMemFree(pwsz);
                pwsz = NULL;
            }
		}
	}

	return pwsz;
}
////////////////////////////////////////////////////////////////////////
// DsoConvertToMBCS
//
//  Takes a WCHAR string and returns a LPSTR on the private heap.
//
STDAPI_(LPSTR) DsoConvertToMBCS(LPCWSTR pwszUnicodeString)
{
	LPSTR psz = NULL;
	UINT cblen, cbnew;

    if ((pwszUnicodeString) && 
        ((cblen = lstrlenW(pwszUnicodeString)) > 0))
	{
		cbnew = ((cblen + 1) * sizeof(WCHAR));
		if ((psz = (LPSTR)DsoMemAlloc(cbnew)) != NULL) 
		{
			if (FAILED(DsoConvertToMBCSEx(pwszUnicodeString, cblen, psz, cbnew, GetACP())))
            {
			    DsoMemFree(psz); psz = NULL;
            }
		}
	}

	return psz;
}

////////////////////////////////////////////////////////////////////////
// DsoConvertToBSTR
//
//  Takes a MBCS string and returns a BSTR. NULL is returned if the 
//  function fails or the string is empty.
//
STDAPI_(BSTR) DsoConvertToBSTR(LPCSTR pszMbcsString)
{
	BSTR bstr = NULL;
    LPWSTR pwsz = DsoConvertToLPWSTR(pszMbcsString);
	if (pwsz)
	{
	    bstr = SysAllocString(pwsz);
	    DsoMemFree(pwsz);
    }
	return bstr;
}

////////////////////////////////////////////////////////////////////////
// DsoConvertToLPOLESTR
//
//  Returns Unicode string in COM Task Memory (CoTaskMemAlloc).
//
STDAPI_(LPWSTR) DsoConvertToLPOLESTR(LPCWSTR pwszUnicodeString)
{
	LPWSTR pwsz;
	UINT cblen;

	CHECK_NULL_RETURN(pwszUnicodeString, NULL);
	cblen = lstrlenW(pwszUnicodeString);

    pwsz = (LPWSTR)CoTaskMemAlloc((cblen * sizeof(WCHAR)) + 2);
    if (pwsz)
    {
        memcpy(pwsz, pwszUnicodeString, (cblen * sizeof(WCHAR)));
        pwsz[cblen] = L'\0'; // Make sure it is NULL terminated.
    }

    return pwsz;
}

////////////////////////////////////////////////////////////////////////
// DsoCopyString
//
//  Duplicates the string into private heap string.
//
STDAPI_(LPWSTR) DsoCopyString(LPCWSTR pwszString)
{
	LPWSTR pwsz;
	UINT cblen;

	CHECK_NULL_RETURN(pwszString, NULL);
	cblen = lstrlenW(pwszString);

    pwsz = (LPWSTR)DsoMemAlloc((cblen * sizeof(WCHAR)) + 2);
    if (pwsz)
    {
        memcpy(pwsz, pwszString, (cblen * sizeof(WCHAR)));
        pwsz[cblen] = L'\0'; // Make sure it is NULL terminated.
    }

    return pwsz;
}

////////////////////////////////////////////////////////////////////////
// DsoCopyStringCat
//
STDAPI_(LPWSTR) DsoCopyStringCat(LPCWSTR pwszString1, LPCWSTR pwszString2)
{return DsoCopyStringCatEx(pwszString1, 1, &pwszString2);}

////////////////////////////////////////////////////////////////////////
// DsoCopyStringCatEx
//
//  Duplicates the string into private heap string and appends one or more
//  strings to the end (concatenation). 
//
STDAPI_(LPWSTR) DsoCopyStringCatEx(LPCWSTR pwszBaseString, UINT cStrs, LPCWSTR *ppwszStrs)
{
	LPWSTR pwsz;
	UINT i, cblenb, cblent;
    UINT *pcblens;

 // We assume you have a base string to start with. If not, we return NULL...
    if ((pwszBaseString == NULL) || 
        ((cblenb = lstrlenW(pwszBaseString)) < 1))
        return NULL;

 // If we have nothing to append, just do a plain copy...
    if ((cStrs == 0) || (ppwszStrs == NULL))
        return DsoCopyString(pwszBaseString);

 // Determine the size of the final string by finding the lengths
 // of each. We create an array of sizes to use later on...
    cblent = cblenb;
    pcblens = new UINT[cStrs];
    CHECK_NULL_RETURN(pcblens,  NULL);

    for (i = 0; i < cStrs; i++)
    {
        pcblens[i] =  lstrlenW(ppwszStrs[i]);
        cblent += pcblens[i];
    }

 // If we have data to append, create the new string and append the
 // data by copying them in place. We expect UTF-16 Unicode strings
 // for this to work, but this should be normal...
	if (cblent > cblenb)
    {
	    pwsz = (LPWSTR)DsoMemAlloc(((cblent + 1) * sizeof(WCHAR)));
	    CHECK_NULL_RETURN(pwsz,   NULL);

	    memcpy(pwsz, pwszBaseString, (cblenb * sizeof(WCHAR)));
        cblent = cblenb;

        for (i = 0; i < cStrs; i++)
        {
		    memcpy((pwsz + cblent), ppwszStrs[i], (pcblens[i] * sizeof(WCHAR)));
            cblent += pcblens[i];
        }
    }
    else pwsz = DsoCopyString(pwszBaseString);

    delete [] pcblens;
	return pwsz;
}

////////////////////////////////////////////////////////////////////////
// DsoCLSIDtoLPSTR
//
STDAPI_(LPSTR) DsoCLSIDtoLPSTR(REFCLSID clsid)
{
	LPSTR psz = NULL;
	LPWSTR pwsz;
	if (SUCCEEDED(StringFromCLSID(clsid, &pwsz)))
	{
		psz = DsoConvertToMBCS(pwsz);
		CoTaskMemFree(pwsz);
	}
    return psz;
}


///////////////////////////////////////////////////////////////////////////////////
// DsoCompareStringsEx
//
//  Calls CompareString API using Unicode version (if available on OS). Otherwise,
//  we have to thunk strings down to MBCS to compare. This is fairly inefficient for
//  Win9x systems that don't handle Unicode, but hey...this is only a sample.
//
STDAPI_(UINT) DsoCompareStringsEx(LPCWSTR pwsz1, INT cch1, LPCWSTR pwsz2, INT cch2)
{
	UINT iret;
	LCID lcid = GetThreadLocale();
	UINT cblen1, cblen2;

 // Check that valid parameters are passed and then contain somethimg...
    if ((pwsz1 == NULL) || (cch1 == 0) || 
        ((cblen1 = ((cch1 > 0) ? cch1 : lstrlenW(pwsz1))) == 0))
		return CSTR_LESS_THAN;

	if ((pwsz2 == NULL) || (cch2 == 0) || 
        ((cblen2 = ((cch2 > 0) ? cch2 : lstrlenW(pwsz2))) == 0))
		return CSTR_GREATER_THAN;

 // If the string is of the same size, then we do quick compare to test for
 // equality (this is slightly faster than calling the API, but only if we
 // expect the calls to find an equal match)...
	if (cblen1 == cblen2)
	{
		for (iret = 0; iret < cblen1; iret++)
		{
			if (pwsz1[iret] == pwsz2[iret])
				continue;

			if (((pwsz1[iret] >= 'A') && (pwsz1[iret] <= 'Z')) &&
				((pwsz1[iret] + ('a' - 'A')) == pwsz2[iret]))
				continue;

			if (((pwsz2[iret] >= 'A') && (pwsz2[iret] <= 'Z')) &&
				((pwsz2[iret] + ('a' - 'A')) == pwsz1[iret]))
				continue;

			break; // don't continue if we can't quickly match...
		}

		// If we made it all the way, then they are equal...
		if (iret == cblen1)
			return CSTR_EQUAL;
	}

 // Now ask the OS to check the strings and give us its read. (We prefer checking
 // in Unicode since this is faster and we may have strings that can't be thunked
 // down to the local ANSI code page)...
	if (v_fUnicodeAPI)
	{
		iret = CompareStringW(lcid, NORM_IGNORECASE | NORM_IGNOREWIDTH, pwsz1, cblen1, pwsz2, cblen2);
	}
	else
	{
	 // If we are on Win9x, we don't have much of choice (thunk the call)...
		LPSTR psz1 = DsoConvertToMBCS(pwsz1);
		LPSTR psz2 = DsoConvertToMBCS(pwsz2);
		iret = CompareStringA(lcid, NORM_IGNORECASE, psz1, -1, psz2, -1);
		DsoMemFree(psz2);
		DsoMemFree(psz1);
	}

	return iret;
}


////////////////////////////////////////////////////////////////////////
// URL Helpers
//
////////////////////////////////////////////////////////////////////////
// General Functions (checks to see if we can recognize type)
//
STDAPI_(BOOL) LooksLikeUNC(LPCWSTR pwsz)
{
	return ((pwsz) && (*pwsz == L'\\') && (*(pwsz + 1) == L'\\') && (*(pwsz + 2) != L'\\'));
}

STDAPI_(BOOL) LooksLikeLocalFile(LPCWSTR pwsz)
{
	return ((pwsz) && 
        (((*pwsz > 64) && (*pwsz < 91)) || ((*pwsz > 96) && (*pwsz < 123))) &&
        (*(pwsz + 1) == L':') && (*(pwsz + 2) == L'\\'));
}

STDAPI_(BOOL) LooksLikeHTTP(LPCWSTR pwsz)
{
	return ((pwsz) && ((*pwsz == L'H') || (*pwsz == L'h')) &&
		((*(pwsz + 1) == L'T') || (*(pwsz + 1) == L't')) &&
		((*(pwsz + 2) == L'T') || (*(pwsz + 2) == L't')) &&
		((*(pwsz + 3) == L'P') || (*(pwsz + 3) == L'p')) &&
		((*(pwsz + 4) == L':') || (((*(pwsz + 4) == L'S') || (*(pwsz + 4) == L's')) && (*(pwsz + 5) == L':'))));
}

////////////////////////////////////////////////////////////////////////
// GetTempPathForURLDownload
//
//  Constructs a temp path for a downloaded file. Note we scan the URL 
//  to find the name of the file so we can use its exention for server 
//  association (in case it is a non-docfile -- like RTF) and also 
//  create our own subfolder to try and avoid name conflicts in temp
//  folder itself. 
//
STDAPI_(BOOL) GetTempPathForURLDownload(WCHAR* pwszURL, WCHAR** ppwszLocalFile)
{
	DWORD  dw;
	LPWSTR pwszTPath = NULL;
	LPWSTR pwszTFile = NULL;
	CHAR   szTmpPath[MAX_PATH];

 // Do a little parameter checking and find length of the URL...
	if (!(pwszURL) || ((dw = lstrlenW(pwszURL)) < 6) ||
		!(LooksLikeHTTP(pwszURL)) || !(ppwszLocalFile))
		return FALSE;

	*ppwszLocalFile = NULL;

	if (GetTempPath(MAX_PATH, szTmpPath))
	{
		DWORD dwtlen = lstrlen(szTmpPath);
		if (dwtlen > 0 && szTmpPath[dwtlen-1] != '\\')
			lstrcat(szTmpPath, "\\");

		lstrcat(szTmpPath, "DsoFramer");

		if (CreateDirectory(szTmpPath, NULL) || GetLastError() == ERROR_ALREADY_EXISTS)
		{
			lstrcat(szTmpPath, "\\");
			pwszTPath = DsoConvertToLPWSTR(szTmpPath);
		}
	}

	if (pwszTPath)
	{
		if (pwszTFile = DsoCopyString(pwszURL))
		{
			LPWSTR pwszT = pwszTFile;
			while (*(++pwszT))
				if (*pwszT == L'?'){*pwszT = L'\0'; break;}

			while (*(--pwszT))
				if (*pwszT == L'/'){++pwszT; break;}

			*ppwszLocalFile = DsoCopyStringCat(pwszTPath, pwszT);

			DsoMemFree(pwszTFile);
		}

		DsoMemFree(pwszTPath);
	}

	return (BOOL)(*ppwszLocalFile);
}

////////////////////////////////////////////////////////////////////////
// URLDownloadFile
//
//  Does a simple URLMON download of file (no save back to server allowed),
//  and no dependent files will be downloaded (just one temp file). 
//
STDAPI URLDownloadFile(LPUNKNOWN punk, WCHAR* pwszURL, WCHAR* pwszLocalFile)
{
    typedef HRESULT (WINAPI *PFN_URLDTFW)(LPUNKNOWN, LPCWSTR, LPCWSTR, DWORD, LPUNKNOWN);
    static PFN_URLDTFW pfnURLDownloadToFileW = NULL;
    HMODULE hUrlmon;

    if (pfnURLDownloadToFileW == NULL)
    {
        if (hUrlmon = LoadLibrary("URLMON.DLL"))
            pfnURLDownloadToFileW = (PFN_URLDTFW)GetProcAddress(hUrlmon, "URLDownloadToFileW");
    }

    if (pfnURLDownloadToFileW == NULL)
        return E_UNEXPECTED;

    return pfnURLDownloadToFileW(punk, pwszURL, pwszLocalFile, 0, NULL);
}


////////////////////////////////////////////////////////////////////////
// OLE Conversion Functions
//
#define HIMETRIC_PER_INCH   2540      // number HIMETRIC units per inch
#define PTS_PER_INCH        72        // number points (font size) per inch

#define MAP_PIX_TO_LOGHIM(x,ppli)   MulDiv(HIMETRIC_PER_INCH, (x), (ppli))
#define MAP_LOGHIM_TO_PIX(x,ppli)   MulDiv((ppli), (x), HIMETRIC_PER_INCH)

////////////////////////////////////////////////////////////////////////
// DsoHimetricToPixels
//
STDAPI_(void) DsoHimetricToPixels(LONG* px, LONG* py)
{
    HDC hdc = GetDC(NULL);
    if (px) *px = MAP_LOGHIM_TO_PIX(*px, GetDeviceCaps(hdc, LOGPIXELSX));
    if (py) *py = MAP_LOGHIM_TO_PIX(*py, GetDeviceCaps(hdc, LOGPIXELSY));
    ReleaseDC(NULL, hdc);
}

////////////////////////////////////////////////////////////////////////
// DsoPixelsToHimetric
//
STDAPI_(void) DsoPixelsToHimetric(LONG* px, LONG* py)
{
    HDC hdc = GetDC(NULL);
    if (px) *px = MAP_PIX_TO_LOGHIM(*px, GetDeviceCaps(hdc, LOGPIXELSX));
    if (py) *py = MAP_PIX_TO_LOGHIM(*py, GetDeviceCaps(hdc, LOGPIXELSY));
    ReleaseDC(NULL, hdc);
}

////////////////////////////////////////////////////////////////////////
// GDI Helper Functions
//
STDAPI_(HBITMAP) DsoGetBitmapFromWindow(HWND hwnd)
{
    HBITMAP hbmpold, hbmp = NULL;
    HDC hdcWin, hdcMem;
    RECT rc;
    INT x, y;

    if (!GetWindowRect(hwnd, &rc))
        return NULL;

    x = (rc.right - rc.left); if (x < 0) x = 1;
    y = (rc.bottom - rc.top); if (y < 0) y = 1;

	hdcWin = GetDC(hwnd);
	hdcMem = CreateCompatibleDC(hdcWin);

	hbmp = CreateCompatibleBitmap(hdcWin, x, y);
	hbmpold = (HBITMAP)SelectObject(hdcMem, hbmp);

	BitBlt(hdcMem, 0,0, x, y, hdcWin, 0,0, SRCCOPY);

	SelectObject(hdcMem, hbmpold);
	DeleteDC(hdcMem);
	ReleaseDC(hwnd, hdcWin);

    return hbmp;
}

////////////////////////////////////////////////////////////////////////
// Windows Helper Functions
//
STDAPI_(BOOL) IsWindowChild(HWND hwndParent, HWND hwndChild)
{
    HWND hwnd;

    if ((hwndChild == NULL) || !IsWindow(hwndChild))
        return FALSE;

    if (hwndParent == NULL)
        return TRUE;

    if (!IsWindow(hwndParent))
        return FALSE;

    hwnd = hwndChild;

    while (hwnd = GetParent(hwnd))
	    if (hwnd == hwndParent)
            return TRUE;

    return FALSE;
}

////////////////////////////////////////////////////////////////////////
// DsoGetTypeInfoEx
//
//  Gets an ITypeInfo from the LIBID specified. Optionally can load and
//  register the typelib from a module resource (if specified). Used to
//  load our typelib on demand.
//
STDAPI DsoGetTypeInfoEx(REFGUID rlibid, LCID lcid, WORD wVerMaj, WORD wVerMin, HMODULE hResource, REFGUID rguid, ITypeInfo** ppti)
{
	HRESULT     hr;
	ITypeLib*   ptlib;

	CHECK_NULL_RETURN(ppti, E_POINTER);
    *ppti = NULL;

 // Try to pull information from the registry...
    hr = LoadRegTypeLib(rlibid, wVerMaj, wVerMin, lcid, &ptlib);

 // If that failed, and we have a resource module to load from,
 // try loading it from the module...
    if (FAILED(hr) && (hResource))
    {
		LPWSTR pwszPath;
        if (FGetModuleFileName(hResource, &pwszPath))
        {
         // Now, load the type library from module resource file...
			hr = LoadTypeLib(pwszPath, &ptlib);

		 // Register library to make things easier next time...
			if (SUCCEEDED(hr))
                RegisterTypeLib(ptlib, pwszPath, NULL);

			DsoMemFree(pwszPath);
		}
    }

 // We have the typelib. Now get the requested typeinfo...
	if (SUCCEEDED(hr))
        hr = ptlib->GetTypeInfoOfGuid(rguid, ppti);

 // Release the type library interface.
    SAFE_RELEASE_INTERFACE(ptlib);
	return hr;
}

////////////////////////////////////////////////////////////////////////
// DsoDispatchInvoke
//
//  Wrapper for IDispatch::Invoke calls to event sinks, or late bound call
//  to embedded object to get ambient property.
//
STDAPI DsoDispatchInvoke(LPDISPATCH pdisp, LPOLESTR pwszname, DISPID dspid, WORD wflags, DWORD cargs, VARIANT* rgargs, VARIANT* pvtret)
{
    HRESULT    hr = S_FALSE;
    DISPID     dspidPut = DISPID_PROPERTYPUT;
    DISPPARAMS dspparm = {NULL, NULL, 0, 0};

	CHECK_NULL_RETURN(pdisp, E_POINTER);

    dspparm.rgvarg = rgargs;
    dspparm.cArgs = cargs;

	if ((wflags & DISPATCH_PROPERTYPUT) || (wflags & DISPATCH_PROPERTYPUTREF))
	{
		dspparm.rgdispidNamedArgs = &dspidPut;
		dspparm.cNamedArgs = 1;
	}

	SEH_TRY

	if (pwszname)
		hr = pdisp->GetIDsOfNames(IID_NULL, &pwszname, 1, LOCALE_USER_DEFAULT, &dspid);

    if (SUCCEEDED(hr))
        hr = pdisp->Invoke(dspid, IID_NULL, LOCALE_USER_DEFAULT, (WORD)(DISPATCH_METHOD | wflags), &dspparm, pvtret, NULL, NULL);

    SEH_EXCEPT(hr)

    return hr;
}

////////////////////////////////////////////////////////////////////////
// DsoReportError -- Report Error for both ComThreadError and DispError.
//
STDAPI DsoReportError(HRESULT hr, LPWSTR pwszCustomMessage, EXCEPINFO* peiDispEx)
{
    BSTR bstrSource, bstrDescription;
    ICreateErrorInfo* pcerrinfo;
    IErrorInfo* perrinfo;
    CHAR szError[MAX_PATH];
    UINT nID = 0;

 // Don't need to do anything unless this is an error.
    if ((hr == S_OK) || SUCCEEDED(hr)) return hr;

 // Is this one of our custom errors (if so we will pull description from resource)...
    if ((hr > DSO_E_ERR_BASE) && (hr < DSO_E_ERR_MAX))
        nID = (hr & 0xFF);

 // Set the source name...
    bstrSource = SysAllocString(L"DsoFramerControl");

 // Set the error description...
    if (pwszCustomMessage)
    {
        bstrDescription = SysAllocString(pwszCustomMessage);
    }
    else if (((nID) && LoadString(v_hModule, nID, szError, sizeof(szError))) || 
             (FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, NULL, hr, 0, szError, sizeof(szError), NULL)))
    {
        bstrDescription = DsoConvertToBSTR(szError);
    }
    else bstrDescription = NULL;
    
 // Set ErrorInfo so that vtable clients can get rich error information...
	if (SUCCEEDED(CreateErrorInfo(&pcerrinfo)))
    {
		pcerrinfo->SetSource(bstrSource);
        pcerrinfo->SetDescription(bstrDescription);

		if (SUCCEEDED(pcerrinfo->QueryInterface(IID_IErrorInfo, (void**) &perrinfo)))
        {
			SetErrorInfo(0, perrinfo);
			perrinfo->Release();
		}
		pcerrinfo->Release();
	}

 // Fill-in DispException Structure for late-boud clients...
    if (peiDispEx)
    {
        peiDispEx->scode = hr;
        peiDispEx->bstrSource = SysAllocString(bstrSource);
        peiDispEx->bstrDescription = SysAllocString(bstrDescription);
    }

 // Free temp strings...
    SAFE_FREEBSTR(bstrDescription);
    SAFE_FREEBSTR(bstrSource);

 // We always return error passed (so caller can chain this in return call).
    return hr;
}

////////////////////////////////////////////////////////////////////////
// Variant Type Helpers (Fast Access to Variant Data)
//
VARIANT* __fastcall DsoPVarFromPVarRef(VARIANT* px)
{return ((px->vt == (VT_VARIANT|VT_BYREF)) ? (px->pvarVal) : px);}

BOOL __fastcall DsoIsVarParamMissing(VARIANT* px)
{return ((px->vt == VT_EMPTY) || ((px->vt & VT_ERROR) == VT_ERROR));}

LPWSTR __fastcall DsoPVarWStrFromPVar(VARIANT* px)
{return ((px->vt == VT_BSTR) ? px->bstrVal : ((px->vt == (VT_BSTR|VT_BYREF)) ? *(px->pbstrVal) : NULL));}

SAFEARRAY* __fastcall DsoPVarArrayFromPVar(VARIANT* px)
{return (((px->vt & (VT_BYREF|VT_ARRAY)) == (VT_BYREF|VT_ARRAY)) ? *(px->pparray) : (((px->vt & VT_ARRAY) == VT_ARRAY) ? px->parray : NULL));}

IUnknown* __fastcall DsoPVarUnkFromPVar(VARIANT* px)
{return (((px->vt == VT_DISPATCH) || (px->vt == VT_UNKNOWN)) ? px->punkVal : (((px->vt == (VT_DISPATCH|VT_BYREF)) || (px->vt == (VT_UNKNOWN|VT_BYREF))) ? *(px->ppunkVal) : NULL));}

SHORT __fastcall DsoPVarShortFromPVar(VARIANT* px, SHORT fdef)
{return (((px->vt & 0xFF) != VT_I2) ? fdef : ((px->vt & VT_BYREF) == VT_BYREF) ? *(px->piVal) : px->iVal);}

LONG __fastcall DsoPVarLongFromPVar(VARIANT* px, LONG fdef)
{return (((px->vt & 0xFF) != VT_I4) ? (LONG)DsoPVarShortFromPVar(px, (SHORT)fdef) : ((px->vt & VT_BYREF) == VT_BYREF) ? *(px->plVal) : px->lVal);}

BOOL __fastcall DsoPVarBoolFromPVar(VARIANT* px, BOOL fdef)
{return (((px->vt & 0xFF) != VT_BOOL) ? (BOOL)DsoPVarLongFromPVar(px, (LONG)fdef) : ((px->vt & VT_BYREF) == VT_BYREF) ? (BOOL)(*(px->pboolVal)) : (BOOL)(px->boolVal));}


////////////////////////////////////////////////////////////////////////
// Win32 Unicode API wrappers
//
//  This project is not compiled to Unicode in order for it to run on Win9x
//  machines. However, in order to try to keep the code language/locale neutral 
//  we use these wrappers to call the Unicode API functions when supported,
//  and thunk down strings to local code page if must run MBCS API.
//

////////////////////////////////////////////////////////////////////////
// FFileExists
//
//  Returns TRUE if the given file exists. Does not handle URLs.
//
STDAPI_(BOOL) FFileExists(WCHAR* wzPath)
{
    DWORD dw = 0xFFFFFFFF;
    if (v_fUnicodeAPI)
    {
        dw = GetFileAttributesW(wzPath);
    }
    else
    {
		LPSTR psz = DsoConvertToMBCS(wzPath);
        if (psz) dw = GetFileAttributesA(psz);
		DsoMemFree(psz);
	}
    return (dw != 0xFFFFFFFF);
}

////////////////////////////////////////////////////////////////////////
// FOpenLocalFile
//
//  Returns TRUE if the file can be opened with the access required.
//  Use the handle for ReadFile/WriteFile operations as normal.
//
STDAPI_(BOOL) FOpenLocalFile(WCHAR* wzFilePath, DWORD dwAccess, DWORD dwShareMode, DWORD dwCreate, HANDLE* phFile)
{
    CHECK_NULL_RETURN(phFile, FALSE);
    *phFile = INVALID_HANDLE_VALUE;
    if (v_fUnicodeAPI)
    {
	    *phFile = CreateFileW(wzFilePath, dwAccess, dwShareMode, NULL, dwCreate, FILE_ATTRIBUTE_NORMAL, NULL);
    }
    else
    {
        LPSTR psz = DsoConvertToMBCS(wzFilePath);
        if (psz) *phFile = CreateFileA(psz, dwAccess, dwShareMode, NULL, dwCreate, FILE_ATTRIBUTE_NORMAL, NULL);
        DsoMemFree(psz);
    }
    return (*phFile != INVALID_HANDLE_VALUE);
}


////////////////////////////////////////////////////////////////////////
// FPerformShellOp
//
//  This function started as a wrapper for SHFileOperation, but that 
//  shell function had an enormous performance hit, especially on Win9x
//  and NT4. To speed things up we removed the shell32 call and are
//  using the kernel32 APIs instead. The only drawback is we can't
//  handle multiple files, but that is not critical for this project.
//
STDAPI_(BOOL) FPerformShellOp(DWORD dwOp, WCHAR* wzFrom, WCHAR* wzTo)
{
	BOOL f = FALSE;
    if (v_fUnicodeAPI)
    {
		switch (dwOp)
		{
		case FO_COPY:		f = CopyFileW(wzFrom, wzTo, FALSE);	break;
		case FO_MOVE: 
		case FO_RENAME:		f = MoveFileW(wzFrom, wzTo);		break;
		case FO_DELETE:		f = DeleteFileW(wzFrom);			break;
		}
	}
    else
    {
	    LPSTR pszFrom = DsoConvertToMBCS(wzFrom);
	    LPSTR pszTo = DsoConvertToMBCS(wzTo);

		switch (dwOp)
		{
		case FO_COPY:		f = CopyFileA(pszFrom, pszTo, FALSE); break;
		case FO_MOVE:
		case FO_RENAME:		f = MoveFileA(pszFrom, pszTo);		break;
		case FO_DELETE:		f = DeleteFileA(pszFrom);			break;
		}

	    if (pszFrom) DsoMemFree(pszFrom);
	    if (pszTo) DsoMemFree(pszTo);
    }

	return f;
}

////////////////////////////////////////////////////////////////////////
// FGetModuleFileName
//
//  Handles Unicode/ANSI paths from a module handle.
//
STDAPI_(BOOL) FGetModuleFileName(HMODULE hModule, WCHAR** wzFileName)
{
    LPWSTR pwsz;
    DWORD dw;

    CHECK_NULL_RETURN(wzFileName, FALSE);
    *wzFileName = NULL;

    pwsz = (LPWSTR)DsoMemAlloc((MAX_PATH * sizeof(WCHAR)));
    CHECK_NULL_RETURN(pwsz, FALSE);

    if (v_fUnicodeAPI)
    {
        dw = GetModuleFileNameW(hModule, pwsz, MAX_PATH);
        if (dw == 0)
        {
            DsoMemFree(pwsz);
            return FALSE;
        }
    }
    else
    {
        dw = GetModuleFileNameA(hModule, (LPSTR)pwsz, MAX_PATH);
        if (dw == 0)
        {
            DsoMemFree(pwsz);
            return FALSE;
        }

        LPWSTR pwsz2 = DsoConvertToLPWSTR((LPSTR)pwsz);
        if (pwsz2 == NULL)
        {
            DsoMemFree(pwsz);
            return FALSE;
        }

        DsoMemFree(pwsz);
        pwsz = pwsz2;
    }

    *wzFileName = pwsz;
    return TRUE;
}


////////////////////////////////////////////////////////////////////////
// FIsIECacheFile
//
//  Determines if file came from IE Cache (treat as read-only).
//
STDAPI_(BOOL) FIsIECacheFile(LPWSTR pwszFile)
{
    BOOL fIsCacheFile = FALSE;
    LPWSTR pwszTmpCache = NULL;
    BYTE rgbuffer[MAX_PATH * sizeof(WCHAR)];
    HKEY hk;

    if (RegOpenKey(HKEY_CURRENT_USER, 
        "Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Folders", &hk) == ERROR_SUCCESS)
    {
        DWORD dwT, dwS; dwT = MAX_PATH; 
        memset(rgbuffer, 0, (MAX_PATH * sizeof(WCHAR)));

        if (v_fUnicodeAPI)
        {
            if ((RegQueryValueExA(hk, "Cache", 0, &dwT, rgbuffer, &dwS) == ERROR_SUCCESS) && 
                (dwT == REG_SZ) && (dwS > 1))
                pwszTmpCache = DsoConvertToLPWSTR((LPSTR)rgbuffer);
        }
        else
        {
            if ((RegQueryValueExW(hk, L"Cache", 0, &dwT, rgbuffer, &dwS) == ERROR_SUCCESS) && 
                (dwT == REG_SZ) && (dwS > 1))
                pwszTmpCache = DsoCopyString((LPWSTR)rgbuffer);
        }

        RegCloseKey(hk);

        if (pwszTmpCache)
        {
            dwS = lstrlenW(pwszTmpCache);
            dwT = lstrlenW(pwszFile);
            fIsCacheFile = ((dwS < dwT) && 
                (DsoCompareStringsEx(pwszTmpCache, dwS, pwszFile, dwS) == CSTR_EQUAL));
            DsoMemFree(pwszTmpCache);
        }
    }

    return fIsCacheFile;
}

////////////////////////////////////////////////////////////////////////
// FDrawText
//
//  This is used by the control for drawing the caption in the titlebar.
//  Since a custom caption could contain Unicode characters only printable
//  on Unicode OS, we should try to use the Unicode version when available.
//
STDAPI_(BOOL) FDrawText(HDC hdc, WCHAR* pwsz, LPRECT prc, UINT fmt)
{
	BOOL f;
    if (v_fUnicodeAPI)
    {
        f = (BOOL)DrawTextW(hdc, pwsz, -1, prc, fmt);
    }
    else
    {
		LPSTR psz = DsoConvertToMBCS(pwsz);
		f = (BOOL)DrawTextA(hdc, psz, -1, prc, fmt);
		DsoMemFree(psz);
    }
	return f;
}


////////////////////////////////////////////////////////////////////////
// FSetRegKeyValue
//
//  We use this for registration when dealing with the file path, since
//  that path may have Unicode characters on some systems. Win9x boxes
//  will have to be converted to ANSI.
//
STDAPI_(BOOL) FSetRegKeyValue(HKEY hk, WCHAR* pwsz)
{
	LONG lret;
    if (v_fUnicodeAPI)
    {
        lret = RegSetValueExW(hk, NULL, 0, REG_SZ, (BYTE*)pwsz, (lstrlenW(pwsz) * sizeof(WCHAR)));
    }
    else
	{
		LPSTR psz = DsoConvertToMBCS(pwsz);
		lret = RegSetValueExA(hk, NULL, 0, REG_SZ, (BYTE*)psz, lstrlen(psz));
		DsoMemFree(psz);
	}
	return (lret == ERROR_SUCCESS);
}

////////////////////////////////////////////////////////////////////////
// FOpenPrinter
//
//  Open the specified printer by name.
//
STDAPI_(BOOL) FOpenPrinter(LPCWSTR pwszPrinter, LPHANDLE phandle)
{
	BOOL fRet = FALSE;
    DWORD dwLastError = 0;
    if (v_fUnicodeAPI)
    {
        PRINTER_DEFAULTSW prtdef; 
        memset(&prtdef, 0, sizeof(PRINTER_DEFAULTSW));
        prtdef.DesiredAccess = PRINTER_ACCESS_USE;
        fRet = OpenPrinterW((LPWSTR)pwszPrinter, phandle, &prtdef);
    }
    else
	{
		LPSTR psz = DsoConvertToMBCS(pwszPrinter);
        fRet = OpenPrinterA(psz, phandle, NULL);
        if (!fRet) dwLastError = GetLastError();
		DsoMemFree(psz);
	}
    if (dwLastError) SetLastError(dwLastError);
	return fRet;
}

////////////////////////////////////////////////////////////////////////
// FGetPrinterSettings
//
//  Returns the default device, port name, and DEVMODE structure for the
//  printer passed. Handles Unicode translation of DEVMODE if on Win9x.
//
STDAPI_(BOOL) FGetPrinterSettings(HANDLE hprinter, LPWSTR *ppwszProcessor, LPWSTR *ppwszDevice, LPWSTR *ppwszOutput, LPDEVMODEW *ppdvmode, DWORD *pcbSize)
{
	BOOL fRet = FALSE;
    DWORD dwLastError = 0;
    DWORD cbNeed, cbAlloc = 0;

    if ((ppwszProcessor == NULL) || (ppwszDevice == NULL) || (ppwszOutput == NULL) ||
        (ppdvmode == NULL) || (pcbSize == NULL))
        return FALSE;

    *ppwszProcessor = NULL; *ppwszDevice = NULL; *ppwszOutput = NULL;
    *ppdvmode = NULL; *pcbSize = 0;

    if (v_fUnicodeAPI) // Use Unicode API if possible (much easier)...
    {
		GetPrinterW(hprinter, 2, NULL, 0, &cbAlloc);
        PRINTER_INFO_2W *pinfo = (PRINTER_INFO_2W*)DsoMemAlloc(++cbAlloc);
        if (pinfo)
        {
            fRet = GetPrinterW(hprinter, 2, (BYTE*)pinfo, cbAlloc, &cbNeed);
            if (fRet)
            {
                *ppwszProcessor = DsoConvertToLPWSTR("winspool");
                *ppwszDevice = DsoCopyString(pinfo->pDriverName);
                *ppwszOutput = DsoCopyString(pinfo->pPortName);

                if (pinfo->pDevMode) // If we have the devmode, just need to copy it...
                {
                    DWORD cbData = (pinfo->pDevMode->dmSize) + (pinfo->pDevMode->dmDriverExtra);
                    *ppdvmode = (DEVMODEW*)DsoMemAlloc(cbData);
                    if (*ppdvmode)
                    {
                        memcpy(*ppdvmode, pinfo->pDevMode, cbData);
                        *pcbSize = cbData;
                    }
                }
            }
            else dwLastError = GetLastError();

            DsoMemFree(pinfo);
        }
        else dwLastError = ERROR_NOT_ENOUGH_MEMORY;   
    }
    else
	{
		GetPrinterA(hprinter, 2, NULL, 0, &cbAlloc);
        PRINTER_INFO_2A *pinfo = (PRINTER_INFO_2A*)DsoMemAlloc(++cbAlloc);
        if (pinfo)
        {
            fRet = GetPrinterA(hprinter, 2, (BYTE*)pinfo, cbAlloc, &cbNeed);
            if (fRet)
            {
                *ppwszProcessor = DsoConvertToLPWSTR("winspool");
                *ppwszDevice = DsoConvertToLPWSTR(pinfo->pDriverName);
                *ppwszOutput = DsoConvertToLPWSTR(pinfo->pPortName);

                if (pinfo->pDevMode) // For Win9x API, we have to convert the DEVMODEA
                {                    // into DEVMODEW so we have Unicode names for TARGETDEVICE...
                    DWORD cbData = sizeof(DEVMODEW) + 
                                ((pinfo->pDevMode->dmSize > sizeof(DEVMODEA)) ? (pinfo->pDevMode->dmSize - sizeof(DEVMODEA)) : 0) +
                                  pinfo->pDevMode->dmDriverExtra;

                    *ppdvmode = (DEVMODEW*)DsoMemAlloc(cbData);
                    if (*ppdvmode)
                    {
                        DsoConvertToUnicodeEx(
                            (LPSTR)(pinfo->pDevMode->dmDeviceName), CCHDEVICENAME,
                            (LPWSTR)((*ppdvmode)->dmDeviceName), CCHDEVICENAME, 0);

                        // The rest of the copy depends on the default size of the DEVMODE.
                        // Just check the size and convert the form name if it exists...
		                if (pinfo->pDevMode->dmSize <= FIELD_OFFSET(DEVMODEA, dmFormName))
                        {
                            memcpy(&((*ppdvmode)->dmSpecVersion), 
                                   &(pinfo->pDevMode->dmSpecVersion),
                                   pinfo->pDevMode->dmSize - CCHDEVICENAME);
		                }
		                else 
                        {
			                memcpy(&((*ppdvmode)->dmSpecVersion), 
                                   &(pinfo->pDevMode->dmSpecVersion),
				                   FIELD_OFFSET(DEVMODEA, dmFormName) -
					               FIELD_OFFSET(DEVMODEA, dmSpecVersion));

                            DsoConvertToUnicodeEx(
                                (LPSTR)(pinfo->pDevMode->dmFormName), CCHFORMNAME,
                                (LPWSTR)((*ppdvmode)->dmFormName), CCHFORMNAME, 0);

			                if (pinfo->pDevMode->dmSize > FIELD_OFFSET(DEVMODEA, dmLogPixels))
				                memcpy(&((*ppdvmode)->dmLogPixels),
                                       &(pinfo->pDevMode->dmLogPixels),
					                   pinfo->pDevMode->dmSize - FIELD_OFFSET(DEVMODEA, dmLogPixels));
		                }
            
                        (*ppdvmode)->dmSize = (WORD)((pinfo->pDevMode->dmSize > sizeof(DEVMODEA)) ?
                                                   (sizeof(DEVMODEW) + (pinfo->pDevMode->dmSize - sizeof(DEVMODEA))) :
                                                    sizeof(DEVMODEW));

                        memcpy((((BYTE*)(*ppdvmode)) + ((*ppdvmode)->dmSize)),
                               (((BYTE*)(pinfo->pDevMode)) + (pinfo->pDevMode->dmSize)),
                               pinfo->pDevMode->dmDriverExtra);

                        *pcbSize = cbData;
                    }
                }
            }
            else dwLastError = GetLastError();

            DsoMemFree(pinfo);
        }
        else dwLastError = ERROR_NOT_ENOUGH_MEMORY;
	}
    if (dwLastError) SetLastError(dwLastError);
	return fRet;
}


///////////////////////////////////////////////////////////////////////////////////
// DsoGetFileFromUser
//
//  Displays the Open/Save dialog using Unicode version if available. Returns the
//  path as a unicode BSTR regardless of OS.
//
STDAPI DsoGetFileFromUser(HWND hwndOwner, LPCWSTR pwzTitle, DWORD dwFlags, 
       LPCWSTR pwzFilter, DWORD dwFiltIdx, LPCWSTR pwszDefExt, LPCWSTR pwszCurrentItem, BOOL fShowSave,
       BSTR *pbstrFile, BOOL *pfReadOnly)
{
	BYTE buffer[MAX_PATH * sizeof(WCHAR)];
	BOOL fSuccess;
	DWORD dw;

 // Make sure they pass a *bstr...
    CHECK_NULL_RETURN(pbstrFile,  E_POINTER);
    *pbstrFile = NULL;

	buffer[0] = 0; buffer[1] = 0;

 // See if we have Unicode function to call. If so, we use OPENFILENAMEW and 
 // get the file path in Unicode, returned as BSTR...
	 if (v_fUnicodeAPI)
	{
		OPENFILENAMEW ofnw;
		memset(&ofnw,  0,   sizeof(OPENFILENAMEW));
	    ofnw.lStructSize  = sizeof(OPENFILENAMEW);
	    ofnw.hwndOwner    = hwndOwner;
	    ofnw.lpstrFilter  = pwzFilter;
	    ofnw.nFilterIndex = dwFiltIdx;
        ofnw.lpstrDefExt  = pwszDefExt;
	    ofnw.lpstrTitle   = pwzTitle;
	    ofnw.lpstrFile    = (LPWSTR)&buffer[0];
	    ofnw.nMaxFile     = MAX_PATH;
	    ofnw.Flags        = dwFlags;

		if (pwszCurrentItem)
		{
			dw = lstrlenW(pwszCurrentItem);
			if ((dw) && (dw < MAX_PATH))
			{
				memcpy(ofnw.lpstrFile, pwszCurrentItem,  dw * sizeof(WCHAR));
				ofnw.lpstrFile[dw] = L'\0';
			}
		}

		if (fShowSave)
			fSuccess = GetSaveFileNameW(&ofnw);
		else
			fSuccess = GetOpenFileNameW(&ofnw);

		if (fSuccess)
        {
			*pbstrFile = SysAllocString((LPWSTR)&buffer[0]);
            if (pfReadOnly) *pfReadOnly = (ofnw.Flags & OFN_READONLY);
        }
	}
	else
	{ // If not, then we use OPENFILENAMEA and thunk down our params to
	  // the MBCS of the system, and then thunk back the Unicode for the return...
		OPENFILENAMEA ofn;
		memset(&ofn,  0,   sizeof(OPENFILENAMEA));
	    ofn.lStructSize  = sizeof(OPENFILENAMEA);
	    ofn.hwndOwner    = hwndOwner;
	    ofn.lpstrFilter  = DsoConvertToMBCS(pwzFilter);
	    ofn.nFilterIndex = dwFiltIdx;
        ofn.lpstrDefExt  = DsoConvertToMBCS(pwszDefExt);
	    ofn.lpstrTitle   = DsoConvertToMBCS(pwzTitle);
	    ofn.lpstrFile    = (LPSTR)&buffer[0];
	    ofn.nMaxFile     = MAX_PATH;
	    ofn.Flags        = dwFlags;

		if (pwszCurrentItem)
			DsoConvertToMBCSEx(pwszCurrentItem, lstrlenW(pwszCurrentItem), (LPSTR)&buffer[0], MAX_PATH, GetACP());

		if (fShowSave)
			fSuccess = GetSaveFileNameA(&ofn);
		else
			fSuccess = GetOpenFileNameA(&ofn);


		if (fSuccess)
        {
			*pbstrFile = DsoConvertToBSTR((LPCSTR)&buffer[0]);
            if (pfReadOnly) *pfReadOnly = (ofn.Flags & OFN_READONLY);
        }

        DsoMemFree((void*)(ofn.lpstrDefExt));
        DsoMemFree((void*)(ofn.lpstrFilter));
        DsoMemFree((void*)(ofn.lpstrTitle));
	}

 // If we got a string, then success. All other errors (even user cancel) should
 // be treated as a general failure (feel free to change this for more full function).
    return ((*pbstrFile == NULL) ? E_FAIL : S_OK);
}

///////////////////////////////////////////////////////////////////////////////////
// DsoGetOleInsertObjectFromUser
//
//  Displays the OLE InsertObject dialog using Unicode version if available.
//
STDAPI DsoGetOleInsertObjectFromUser(HWND hwndOwner, LPCWSTR pwzTitle, DWORD dwFlags, 
        BOOL fDocObjectOnly, BOOL fAllowControls, BSTR *pbstrResult, UINT *ptype)
{
	BYTE buffer[MAX_PATH * sizeof(WCHAR)];
    LPCLSID lpNewExcludeList = NULL;
	int nNewExcludeCount = 0;
	int nNewExcludeLen = 0;

 // Make sure they pass a *bstr...
    CHECK_NULL_RETURN(pbstrResult,  E_POINTER);
    *pbstrResult = NULL;

 // To limit list to just those marked as DocObject servers, you have to enum
 // the registry and create an exclude list for OLE dialog. Exclude all except
 // those that are marked DocObject under their ProgID.
    if (fDocObjectOnly)
    {
	    HKEY hkCLSID;
	    HKEY hkItem;
	    HKEY hkDocObject;
	    DWORD dwIndex = 0;
	    CHAR szName[MAX_PATH+1];

	    if (RegOpenKeyEx(HKEY_CLASSES_ROOT, "CLSID", 0, KEY_READ|KEY_ENUMERATE_SUB_KEYS, &hkCLSID) == ERROR_SUCCESS)
	    {
		    while (RegEnumKey(hkCLSID, dwIndex++, szName, MAX_PATH) == ERROR_SUCCESS)
		    {
			    if (RegOpenKeyEx(hkCLSID, szName, 0, KEY_READ, &hkItem) == ERROR_SUCCESS)
			    {
					if ((RegOpenKeyEx(hkItem, "DocObject", 0, KEY_READ, &hkDocObject) != ERROR_SUCCESS))
					{
						CLSID clsid;
						LPWSTR pwszClsid = DsoConvertToLPWSTR(szName);
						if ((pwszClsid) && SUCCEEDED(CLSIDFromString(pwszClsid, &clsid)))
						{
							if (lpNewExcludeList == NULL)
							{
								nNewExcludeCount = 0;
								nNewExcludeLen = 16;
								lpNewExcludeList = new CLSID[nNewExcludeLen];
							}
							if (nNewExcludeCount == nNewExcludeLen)
							{
								LPCLSID lpOldList = lpNewExcludeList;
								nNewExcludeLen <<= 2;
								lpNewExcludeList = new CLSID[nNewExcludeLen];
								memcpy(lpNewExcludeList, lpOldList, sizeof(CLSID) * nNewExcludeCount);
								delete [] lpOldList;
							}

							lpNewExcludeList[nNewExcludeCount] = clsid;
							nNewExcludeCount++;
						}
						SAFE_FREESTRING(pwszClsid);
                        RegCloseKey(hkDocObject);
					}
				    
				    RegCloseKey(hkItem);
			    }
		    }
		    RegCloseKey(hkCLSID);
	    }
    }
	buffer[0] = 0; buffer[1] = 0;

 // See if we have Unicode function to call...
	if (v_fUnicodeAPI)
	{
	    OLEUIINSERTOBJECTW oidlg = {0};
	    oidlg.cbStruct = sizeof(OLEUIINSERTOBJECTW);
	    oidlg.dwFlags = dwFlags;
	    oidlg.hWndOwner = hwndOwner;
	    oidlg.lpszCaption = pwzTitle;
	    oidlg.lpszFile = (LPWSTR)buffer;
	    oidlg.cchFile = MAX_PATH;
		oidlg.lpClsidExclude = lpNewExcludeList;
		oidlg.cClsidExclude = nNewExcludeCount;

		if (OleUIInsertObjectW(&oidlg) == OLEUI_OK)
		{
			if ((oidlg.dwFlags & IOF_SELECTCREATENEW) && (oidlg.clsid != GUID_NULL))
			{
				LPOLESTR posz;
				if (SUCCEEDED(ProgIDFromCLSID(oidlg.clsid, &posz)))
				{
					*pbstrResult = SysAllocString(posz);
					CoTaskMemFree(posz);
				}
                if (ptype) *ptype = IOF_SELECTCREATENEW;
			}
			else if ((oidlg.dwFlags & IOF_SELECTCREATEFROMFILE) && (buffer[0] != 0))
			{
				*pbstrResult = SysAllocString((LPWSTR)buffer);
                if (ptype) *ptype = IOF_SELECTCREATEFROMFILE;
			}
        }
    }
    else
    {
	    OLEUIINSERTOBJECTA oidlg = {0};
	    oidlg.cbStruct = sizeof(OLEUIINSERTOBJECTA);
	    oidlg.dwFlags = dwFlags;
	    oidlg.hWndOwner = hwndOwner;
	    oidlg.lpszCaption = DsoConvertToMBCS(pwzTitle);
	    oidlg.lpszFile = (LPSTR)buffer;
	    oidlg.cchFile = MAX_PATH;
		oidlg.lpClsidExclude = lpNewExcludeList;
		oidlg.cClsidExclude = nNewExcludeCount;

		if (OleUIInsertObjectA(&oidlg) == OLEUI_OK)
		{
			if ((oidlg.dwFlags & IOF_SELECTCREATENEW) && (oidlg.clsid != GUID_NULL))
			{
				LPOLESTR posz;
				if (SUCCEEDED(ProgIDFromCLSID(oidlg.clsid, &posz)))
				{
					*pbstrResult = SysAllocString(posz);
					CoTaskMemFree(posz);
				}
                if (ptype) *ptype = IOF_SELECTCREATENEW;
			}
			else if ((oidlg.dwFlags & IOF_SELECTCREATEFROMFILE) && (buffer[0] != 0))
			{
				*pbstrResult = DsoConvertToBSTR((LPSTR)buffer);
                if (ptype) *ptype = IOF_SELECTCREATEFROMFILE;
			}
        }

        DsoMemFree((void*)(oidlg.lpszCaption));
    }

    if (lpNewExcludeList)
		delete [] lpNewExcludeList;

 // If we got a string, then success. All other errors (even user cancel) should
 // be treated as a general failure (feel free to change this for more full function).
    return ((*pbstrResult == NULL) ? E_FAIL : S_OK);
}
