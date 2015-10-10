/***************************************************************************
 * CLASSFACTORY.CPP
 *
 * CDsoFramerClassFactory: The Class Factroy for the control.
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
// CDsoFramerClassFactory - IClassFactory Implementation
//
//  This is a fairly simple CF. We don't provide support for licensing
//  in this sample because it is just a sample. If licensing is important
//  you should change the class to support IClassFactory2.
//

////////////////////////////////////////////////////////////////////////
// CDsoFramerClassFactory::QueryInterface
//
STDMETHODIMP CDsoFramerClassFactory::QueryInterface(REFIID riid, void** ppv)
{
	ODS("CDsoFramerClassFactory::QueryInterface\n");
	CHECK_NULL_RETURN(ppv, E_POINTER);
	
	if ((IID_IUnknown == riid) || (IID_IClassFactory == riid))
	{
        SAFE_SET_INTERFACE(*ppv, (IClassFactory*)this);
		return S_OK;
	}

    *ppv = NULL;
	return E_NOINTERFACE;
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerClassFactory::AddRef
//
STDMETHODIMP_(ULONG) CDsoFramerClassFactory::AddRef(void)
{
	TRACE1("CDsoFramerClassFactory::AddRef - %d\n", m_cRef+1);
    return ++m_cRef;
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerClassFactory::Release
//
STDMETHODIMP_(ULONG) CDsoFramerClassFactory::Release(void)
{
	TRACE1("CDsoFramerClassFactory::Release - %d\n", m_cRef-1);
    if (0 != --m_cRef) return m_cRef;

	ODS("CDsoFramerClassFactory delete\n");
	InterlockedDecrement((LPLONG)&v_cLocks);
    delete this;
    return 0;
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerClassFactory::CreateInstance
//
//  Creates an instance of our control.
//
STDMETHODIMP CDsoFramerClassFactory::CreateInstance(LPUNKNOWN punk, REFIID riid, void** ppv)
{
	HRESULT hr;
	CDsoFramerControl* pocx;
	IUnknown* pnkInternal;

	ODS("CDsoFramerClassFactory::CreateInstance\n");
	CHECK_NULL_RETURN(ppv, E_POINTER);	*ppv = NULL;

 // Aggregation requires you ask for (internal) IUnknown
	if ((punk) && (riid != IID_IUnknown)) 
		return E_INVALIDARG;

 // Create a new instance of the control...
	pocx = new CDsoFramerControl(punk);
	CHECK_NULL_RETURN(pocx, E_OUTOFMEMORY);

 // Grab the internal IUnknown to use for the QI (you don't agg in CF:CreateInstance)...
	pnkInternal = (IUnknown*)&(pocx->m_xInternalUnknown);

 // Initialize the control (windows, etc.) and QI for requested interface...
	if (SUCCEEDED(hr = pocx->InitializeNewInstance()) &&
		SUCCEEDED(hr = pnkInternal->QueryInterface(riid, ppv)))
	{
		InterlockedIncrement((LPLONG)&v_cLocks); // on success, bump up the lock count...
	}
	else {delete pocx; *ppv = NULL;} // else cleanup the object

	return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerClassFactory::LockServer
//
//  Keeps the server loaded in memory.
//
STDMETHODIMP CDsoFramerClassFactory::LockServer(BOOL fLock)
{
	TRACE1("CDsoFramerClassFactory::LockServer - %d\n", fLock);
	if (fLock) InterlockedIncrement((LPLONG)&v_cLocks);
	else InterlockedDecrement((LPLONG)&v_cLocks);
	return S_OK;
}

