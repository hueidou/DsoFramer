/***************************************************************************
 * MAINENTRY.CPP
 *
 * Main DLL Entry and Required COM Entry Points.
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
#define INITGUID // Init the GUIDS for the control...
#include "dsoframer.h"

HINSTANCE        v_hModule        = NULL;   // DLL module handle
HANDLE           v_hPrivateHeap   = NULL;   // Private Memory Heap
ULONG            v_cLocks         = 0;      // Count of server locks
HICON            v_icoOffDocIcon  = NULL;   // Small office icon (for caption bar)
BOOL             v_fUnicodeAPI    = FALSE;  // Flag to determine if we should us Unicode API
BOOL             v_fWindows2KPlus = FALSE;
CRITICAL_SECTION v_csecThreadSynch;

////////////////////////////////////////////////////////////////////////
// DllMain -- OCX Main Entry
//
//
extern "C" BOOL APIENTRY DllMain(HINSTANCE hDllHandle, DWORD dwReason, LPVOID /*lpReserved*/)
{
	switch (dwReason)
	{
	case DLL_PROCESS_ATTACH:
		v_hModule = hDllHandle;
        v_hPrivateHeap = HeapCreate(0, 0x1000, 0);
		v_icoOffDocIcon = (HICON)LoadImage(hDllHandle, MAKEINTRESOURCE(IDI_SMALLOFFDOC), IMAGE_ICON, 16, 16, 0);
		{
			DWORD dwVersion = GetVersion();
			v_fUnicodeAPI = ((dwVersion & 0x80000000) == 0);
			v_fWindows2KPlus = ((v_fUnicodeAPI) && (LOBYTE(LOWORD(dwVersion)) > 4));
		}
		InitializeCriticalSection(&v_csecThreadSynch);
		DisableThreadLibraryCalls(hDllHandle);
		break;

	case DLL_PROCESS_DETACH:
		if (v_icoOffDocIcon) DeleteObject(v_icoOffDocIcon);
        if (v_hPrivateHeap) HeapDestroy(v_hPrivateHeap);
        DeleteCriticalSection(&v_csecThreadSynch);
		break;
	}
	return TRUE;
}

#ifdef DSO_MIN_CRT_STARTUP
extern "C" BOOL APIENTRY _DllMainCRTStartup(HINSTANCE hDllHandle, DWORD dwReason, LPVOID lpReserved)
{return DllMain(hDllHandle, dwReason, lpReserved);}
#endif

////////////////////////////////////////////////////////////////////////
// Standard COM DLL Entry Points
//
//
////////////////////////////////////////////////////////////////////////
// DllCanUnloadNow
//
//
STDAPI DllCanUnloadNow()
{
	return ((v_cLocks == 0) ? S_OK : S_FALSE);
}

////////////////////////////////////////////////////////////////////////
// DllGetClassObject
//
//  Returns IClassFactory instance for FramerControl. We only support
//  this one object for creation.
//
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppv)
{
	HRESULT hr;
	CDsoFramerClassFactory* pcf;

	CHECK_NULL_RETURN(ppv, E_POINTER);
	*ppv = NULL;

 // The only component we can create is the BinderControl...
	if (rclsid != CLSID_FramerControl)
		return CLASS_E_CLASSNOTAVAILABLE;

 // Create the needed class factory...
	pcf = new CDsoFramerClassFactory();
	CHECK_NULL_RETURN(pcf, E_OUTOFMEMORY);

 // Get requested interface.
	if (FAILED(hr = pcf->QueryInterface(riid, ppv)))
	{
		*ppv = NULL; delete pcf;
	}
	else InterlockedIncrement((LPLONG)&v_cLocks);

	return hr;
}

////////////////////////////////////////////////////////////////////////
// DllRegisterServer
//
//  Registration of the OCX.
//
STDAPI DllRegisterServer()
{
    HRESULT hr = S_OK;
    HKEY    hk, hk2;
    DWORD   dwret;
    CHAR    szbuffer[256];
	LPWSTR  pwszModule;
    ITypeInfo* pti;
    
 // If we can't find the path to the DLL, we can't register...
	if (!FGetModuleFileName(v_hModule, &pwszModule))
		return E_FAIL;

 // Setup the CLSID. This is the most important. If there is a critical failure,
 // we will set HR = GetLastError and return...
    if ((dwret = RegCreateKeyEx(HKEY_CLASSES_ROOT, 
		"CLSID\\"DSOFRAMERCTL_CLSIDSTR, 0, NULL, 0, KEY_WRITE, NULL, &hk, NULL)) != ERROR_SUCCESS)
	{
		DsoMemFree(pwszModule);
        return HRESULT_FROM_WIN32(dwret);
	}

    lstrcpy(szbuffer, DSOFRAMERCTL_SHORTNAME);
    RegSetValueEx(hk, NULL, 0, REG_SZ, (BYTE *)szbuffer, lstrlen(szbuffer));

 // Setup the InprocServer32 key...
    dwret = RegCreateKeyEx(hk, "InprocServer32", 0, NULL, 0, KEY_WRITE, NULL, &hk2, NULL);
    if (dwret == ERROR_SUCCESS)
    {
        lstrcpy(szbuffer, "Apartment");
        RegSetValueEx(hk2, "ThreadingModel", 0, REG_SZ, (BYTE *)szbuffer, lstrlen(szbuffer));
        
	 // We call a wrapper function for this setting since the path should be
	 // stored in Unicode to handle non-ANSI file path names on some systems.
	 // This wrapper will convert the path to ANSI if we are running on Win9x.
	 // The rest of the Reg calls should be OK in ANSI since they do not
	 // contain non-ANSI/Unicode-specific characters...
		if (!FSetRegKeyValue(hk2, pwszModule))
            hr = E_ACCESSDENIED;

        RegCloseKey(hk2);

        dwret = RegCreateKeyEx(hk, "ProgID", 0, NULL, 0, KEY_WRITE, NULL, &hk2, NULL);
        if (dwret == ERROR_SUCCESS)
        {
            lstrcpy(szbuffer, DSOFRAMERCTL_PROGID);
            RegSetValueEx(hk2, NULL, 0, REG_SZ, (BYTE *)szbuffer, lstrlen(szbuffer));
            RegCloseKey(hk2);
        }

    }
    else hr = HRESULT_FROM_WIN32(dwret);

	if (SUCCEEDED(hr))
	{
		dwret = RegCreateKeyEx(hk, "Control", 0, NULL, 0, KEY_WRITE, NULL, &hk2, NULL);
		if (dwret == ERROR_SUCCESS)
		{
			RegCloseKey(hk2);
		}
		else hr = HRESULT_FROM_WIN32(dwret);
	}


 // If we succeeded so far, andle the remaining (non-critical) reg keys...
	if (SUCCEEDED(hr))
	{
		dwret = RegCreateKeyEx(hk, "ToolboxBitmap32", 0, NULL, 0, KEY_WRITE, NULL, &hk2, NULL);
		if (dwret == ERROR_SUCCESS)
		{
			LPWSTR pwszT = DsoCopyStringCat(pwszModule, L",102");
			if (pwszT)
			{
				FSetRegKeyValue(hk2, pwszT);
				DsoMemFree(pwszT);
			}
			RegCloseKey(hk2);
		}

		dwret = RegCreateKeyEx(hk, "TypeLib", 0, NULL, 0, KEY_WRITE, NULL, &hk2, NULL);
		if (dwret == ERROR_SUCCESS)
		{
			lstrcpy(szbuffer, DSOFRAMERCTL_TLIBSTR);
			RegSetValueEx(hk2, NULL, 0, REG_SZ, (BYTE *)szbuffer, lstrlen(szbuffer));
			RegCloseKey(hk2);
		}

		dwret = RegCreateKeyEx(hk, "Version", 0, NULL, 0, KEY_WRITE, NULL, &hk2, NULL);
		if (dwret == ERROR_SUCCESS)
		{
			lstrcpy(szbuffer, DSOFRAMERCTL_VERSIONSTR);
			RegSetValueEx(hk2, NULL, 0, REG_SZ, (BYTE *)szbuffer, lstrlen(szbuffer));
			RegCloseKey(hk2);
		}

		dwret = RegCreateKeyEx(hk, "MiscStatus", 0, NULL, 0, KEY_WRITE, NULL, &hk2, NULL);
		if (dwret == ERROR_SUCCESS)
		{
			lstrcpy(szbuffer, "131473");
			RegSetValueEx(hk2, NULL, 0, REG_SZ, (BYTE *)szbuffer, lstrlen(szbuffer));
			RegCloseKey(hk2);
		}

		dwret = RegCreateKeyEx(hk, "DataFormats\\GetSet\\0", 0, NULL, 0, KEY_WRITE, NULL, &hk2, NULL);
		if (dwret == ERROR_SUCCESS)
		{
			lstrcpy(szbuffer, "3,1,32,1");
			RegSetValueEx(hk2, NULL, 0, REG_SZ, (BYTE *)szbuffer, lstrlen(szbuffer));
			RegCloseKey(hk2);
		}
    }

    RegCloseKey(hk);
	DsoMemFree(pwszModule);

 // This should catch any critical failures during setup of CLSID...
	RETURN_ON_FAILURE(hr);

 // Setup the ProgID (non-critical)...
    if (RegCreateKeyEx(HKEY_CLASSES_ROOT, DSOFRAMERCTL_PROGID, 0,
            NULL, 0, KEY_WRITE, NULL, &hk, NULL) == ERROR_SUCCESS)
    {
        lstrcpy(szbuffer, DSOFRAMERCTL_FULLNAME);
        RegSetValueEx(hk, NULL, 0, REG_SZ, (BYTE *)szbuffer, lstrlen(szbuffer));

        if (RegCreateKeyEx(hk, "CLSID", 0,
                NULL, 0, KEY_WRITE, NULL, &hk2, NULL) == ERROR_SUCCESS)
        {
            lstrcpy(szbuffer, DSOFRAMERCTL_CLSIDSTR);
            RegSetValueEx(hk2, NULL, 0, REG_SZ, (BYTE *)szbuffer, lstrlen(szbuffer));
            RegCloseKey(hk2);
        }
        RegCloseKey(hk);
    }

 // Load the type info (this should register the lib once)...
    hr = DsoGetTypeInfoEx(LIBID_DSOFramer, 0,
		DSOFRAMERCTL_VERSION_MAJOR, DSOFRAMERCTL_VERSION_MINOR, v_hModule, CLSID_FramerControl, &pti);
    if (SUCCEEDED(hr)) pti->Release();

	return hr;
}

////////////////////////////////////////////////////////////////////////
// RegRecursiveDeleteKey
//
//  Helper function called by DllUnregisterServer for nested key removal.
//
static HRESULT RegRecursiveDeleteKey(HKEY hkParent, LPCSTR pszSubKey)
{
    HRESULT hr = S_OK;
    HKEY hk;
    DWORD dwret, dwsize;
	FILETIME time ;
    CHAR szbuffer[512];

    dwret = RegOpenKeyEx(hkParent, pszSubKey, 0, KEY_ALL_ACCESS, &hk);
	if (dwret != ERROR_SUCCESS)
		return HRESULT_FROM_WIN32(dwret);

 // Enumerate all of the decendents of this child...
	dwsize = 510 ;
	while (RegEnumKeyEx(hk, 0, szbuffer, &dwsize, NULL, NULL, NULL, &time) == ERROR_SUCCESS)
	{
      // If there are any sub-folders, delete them first (to make NT happy)...
		hr = RegRecursiveDeleteKey(hk, szbuffer);
		if (FAILED(hr)) break;
		dwsize = 510;
	}

 // Close the child...
	RegCloseKey(hk);

	RETURN_ON_FAILURE(hr);

 // Delete this child.
    dwret = RegDeleteKey(hkParent, pszSubKey);
    if (dwret != ERROR_SUCCESS)
        hr = HRESULT_FROM_WIN32(dwret);

	return hr;
}

////////////////////////////////////////////////////////////////////////
// DllUnregisterServer
//
//  Removal code for the OCX.
//
STDAPI DllUnregisterServer()
{
    HRESULT hr;
    hr = RegRecursiveDeleteKey(HKEY_CLASSES_ROOT, "CLSID\\"DSOFRAMERCTL_CLSIDSTR);
    if (SUCCEEDED(hr))
    {
        RegRecursiveDeleteKey(HKEY_CLASSES_ROOT, DSOFRAMERCTL_PROGID);
        RegRecursiveDeleteKey(HKEY_CLASSES_ROOT, "TypeLib\\"DSOFRAMERCTL_TLIBSTR);
    }

  // This means the key does not exist (i.e., the DLL 
  // was alreay unregistered, so return OK)...
    if (hr == 0x80070002) hr = S_OK;

	return hr;
}
