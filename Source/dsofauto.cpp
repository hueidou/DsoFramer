/***************************************************************************
 * DSOFAUTO.CPP
 *
 * CDsoFramerControl: Automation interface for Binder Control
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
// CDsoFramerControl - _FramerControl Implementation
//

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::Activate
//
//  Activate the current embedded document (i.e, forward focus).
//
STDMETHODIMP CDsoFramerControl::Activate()
{
    HRESULT hr;
    ODS("CDsoFramerControl::Activate\n");

	if (m_fInControlActivate)
		return S_FALSE;

 // Don't allow recursion of this function or we could get stuck in
 // loop trying to constantly grab focus.
	m_fInControlActivate = TRUE;

 // All we need to do is grab focus. This will tell the host to
 // UI activate our OCX, set focus to our window, and set this control
 // as the active component with the hook manager. 
    hr = UIActivate(TRUE);

 // Invalidate windows to update painting...
    if (SUCCEEDED(hr))
        InvalidateAllChildWindows(m_hwnd);

	m_fInControlActivate = FALSE;

	return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::get_ActiveDocument
//
//  Returns the automation object currently embedded.
//
//  Since we only support a single instance at a time, it might have been
//  better to call this property Object or simply Document, but VB reserves
//  the first name for use by the control extender, and IE reserves the second
//  in its extender, so we decided on the "Office sounding" name. ;-)
//
STDMETHODIMP CDsoFramerControl::get_ActiveDocument(IDispatch** ppdisp)
{
    HRESULT hr = DSO_E_DOCUMENTNOTOPEN;
	IUnknown* punk;

	ODS("CDsoFramerControl::get_ActiveDocument\n");
    CHECK_NULL_RETURN(ppdisp, E_POINTER); *ppdisp = NULL;

 // Get IDispatch from open document object.
	if ((m_pDocObjFrame) && (punk = (IUnknown*)(m_pDocObjFrame->GetActiveObject())))
	{ 
     // Cannot access object if in print preview..
        if (m_pDocObjFrame->InPrintPreview())
            return ProvideErrorInfo(DSO_E_INMODALSTATE);

	  // Ask ip active object for IDispatch interface. If it is not supported on
      // active object interface, try to get it from OLE object iface...
        if (FAILED(hr = punk->QueryInterface(IID_IDispatch, (void**)ppdisp)) && 
            (punk = (IUnknown*)(m_pDocObjFrame->GetOleObject())))
        {
            hr = punk->QueryInterface(IID_IDispatch, (void**)ppdisp);
        }
        ASSERT(SUCCEEDED(hr));
	}

	return ProvideErrorInfo(hr);
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::ExecOleCommand
//
//  Allows caller to pass commands to embedded object via IOleCommandTarget and
//  gives access to a few extra commands to extend functionality for certain object
//  types (not all objects may support all commands).
//
STDMETHODIMP CDsoFramerControl::ExecOleCommand(LONG OLECMDID, VARIANT Options, VARIANT* vInParam, VARIANT* vInOutParam)
{
	HRESULT hr = E_INVALIDARG;
    DWORD dwOptions = (DWORD)LONG_FROM_VARIANT(Options, 0);

	TRACE1("CDsoFramerControl::DoOleCommand(%d)\n", OLECMDID);

 // Cannot access object if in modal condition...
    if ((m_fModalState) || (m_fNoInteractive))
        return ProvideErrorInfo(DSO_E_INMODALSTATE);

 // Handle some custom commands that don't need doc open to call...
	switch (OLECMDID)
	{
	case OLECMDID_RESETFRAMEHOOK:
		{
			LONG lhwnd = (DWORD)LONG_FROM_VARIANT(*vInParam, 0);
			return ResetFrameHook((HWND)lhwnd);
		}
	case OLECMDID_LOCKSERVER:
		if (BOOL_FROM_VARIANT(Options, TRUE) == FALSE)
			return put_LockServer(VARIANT_FALSE);
	}

 // The rest require a doc object loaded first...
    CHECK_NULL_RETURN(m_pDocObjFrame, ProvideErrorInfo(DSO_E_DOCUMENTNOTOPEN));

    switch (OLECMDID)
    {
    case OLECMDID_GETDATAFORMAT:
     // If requesting special data get...
        hr = m_pDocObjFrame->HrGetDataFromObject(vInParam, vInOutParam);
        break;

    case OLECMDID_SETDATAFORMAT:
     // If requesting special data set...
        hr = m_pDocObjFrame->HrSetDataInObject(vInParam, vInOutParam, BOOL_FROM_VARIANT(Options, TRUE));
        break;

    case OLECMDID_LOCKSERVER: // optional lock on server...
        if (BOOL_FROM_VARIANT(Options, FALSE))
            hr = put_LockServer(VARIANT_TRUE);
        break;

	case OLECMDID_NOTIFYACTIVE:
		if (vInParam) // Force notify server that it should be active...
		{
			m_pDocObjFrame->OnNotifyAppActivate(BOOL_FROM_VARIANT(*vInParam, FALSE), 0);
			hr = S_OK;
		}
		break;

    default:
     // Do normal IOleCommandTarget call on object...

     // If options was not passed as long, but as bool, we expect the caller meant to
     // specify if user should be prompted or not, so update the options to allow the
     // assuption to still work as expected (this is for compatibility)...
        if ((dwOptions == 0) && ((DsoPVarFromPVarRef(&Options)->vt & 0xFF) == VT_BOOL))
            dwOptions = (BOOL_FROM_VARIANT(Options, FALSE) ? OLECMDEXECOPT_PROMPTUSER : OLECMDEXECOPT_DODEFAULT);

     // Ask object server to do the command...
	    hr = m_pDocObjFrame->DoOleCommand(OLECMDID, dwOptions, vInParam, vInOutParam);
    }
	return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::CreateNew
//
//  Creates a new document object based on the passed in ProgID or Template.
//  ProgID should map to the "document" object, such as "Word.Document",
//  "Excel.Sheet", "PowerPoint.Slide", etc. A Template name should be full
//  path to the file, whether local, UNC, or HTTP path. No relative paths.
//
STDMETHODIMP CDsoFramerControl::CreateNew(BSTR ProgIdOrTemplate)
{
	HRESULT hr;
	CLSID clsid;
	HCURSOR	hCur;
    IStorage *pstgTemplate = NULL;
    LPWSTR pwszTempFile = NULL;

    TRACE1("CDsoFramerControl::CreateNew(%S)\n", ProgIdOrTemplate);

 // Check the string to make sure a valid item is passed...
	if (!(ProgIdOrTemplate) || (SysStringLen(ProgIdOrTemplate) < 4))
		return E_INVALIDARG;

// Cannot create object if we are not activate yet, or if in modal condition...
    if (!(m_fInPlaceActive) || (m_fModalState)) 
        return ProvideErrorInfo(DSO_E_INMODALSTATE);

 // Make sure any open document is closed first...
	if ((m_pDocObjFrame) && FAILED(hr = Close()))
		return hr;

 // Make sure we are the active component for this process...
    if (FAILED(hr = Activate())) return hr;

 // Let's make a doc frame for ourselves...
	m_pDocObjFrame = CDsoDocObject::CreateInstance((IDsoDocObjectSite*)&m_xDsoDocObjectSite);
	if (!(m_pDocObjFrame)) return E_OUTOFMEMORY;

 // Start a wait operation to notify user...
	hCur = SetCursor(LoadCursor(NULL, IDC_WAIT));
    m_fInDocumentLoad = TRUE;

 // If the string passed looks like a URL, it is a web template. We need
 // to download it to temp location and use for new object...
    if (LooksLikeHTTP(ProgIdOrTemplate) && 
        GetTempPathForURLDownload(ProgIdOrTemplate, &pwszTempFile))
    {
     // Ask URLMON to download the file...
        if (FAILED(hr = URLDownloadFile(NULL, ProgIdOrTemplate, pwszTempFile)))
        {
            DsoMemFree(pwszTempFile); pwszTempFile = NULL;
            goto error_out;
        }

     // If that worked, switch out the name of the template to local file...
        ProgIdOrTemplate = SysAllocString(pwszTempFile);
    }

 // If the string is path to file, then it must be a template. It must be
 // a storage file with CLSID associated with it, like any Office templates 
 // (.dot,.xlt,.pot,.vst,etc.), and path must be fully qualified...
    if (LooksLikeUNC(ProgIdOrTemplate) || LooksLikeLocalFile(ProgIdOrTemplate))
    {
        if ((hr = StgIsStorageFile(ProgIdOrTemplate)) != S_OK)
        {
            hr = (FAILED(hr) ? hr : STG_E_NOTFILEBASEDSTORAGE);
            goto error_out;
        }

     // Open the template for read access only...
        hr = StgOpenStorage(ProgIdOrTemplate, NULL, 
            (STGM_READ | STGM_SHARE_DENY_WRITE | STGM_TRANSACTED),
             NULL, 0, &pstgTemplate);
        GOTO_ON_FAILURE(hr, error_out);

     // We get the CLSID from the template...
        hr = ReadClassStg(pstgTemplate, &clsid);
        if (FAILED(hr) || (clsid == GUID_NULL))
        {
            hr = (FAILED(hr) ? hr : STG_E_OLDFORMAT);
            goto error_out;
        } 
 // Otherwise the string passed is assumed a ProgID...
    }
    else if (FAILED(CLSIDFromProgID(ProgIdOrTemplate, &clsid)))
    {
        hr = DSO_E_INVALIDPROGID;
        goto error_out;
    }

 // If we are here, we must have a valid CLSID for the object...
    ASSERT(clsid != GUID_NULL);

 // If we had delayed the frame hook, we should set it up now...
    if (!(m_pHookManager) && FDelayFrameHookSet())
    {
		m_pHookManager = CDsoFrameHookManager::RegisterFramerControl(m_hwndParent, m_hwnd);
		if (!m_pHookManager) {hr = DSO_E_FRAMEHOOKFAILED; goto error_out;}
    }

	SEH_TRY

 // If we are loading a template, init the storage before the create...
    if (pstgTemplate)
    {
        hr = m_pDocObjFrame->CreateDocObject(pstgTemplate);
    }
    else
    {
      // Create a new doc object and IP activate...
	    hr = m_pDocObjFrame->CreateDocObject(clsid);
    }

 // If the call worked, we can activate it...
	if (SUCCEEDED(hr))
	{
        EnableDropFile(FALSE);

		if (!m_fShowToolbars)
			m_pDocObjFrame->OnNotifyChangeToolState(FALSE);

		hr = m_pDocObjFrame->IPActivateView();
	}

    SEH_EXCEPT(hr)

 // Force a close if an error occurred...
	if (FAILED(hr))
	{
error_out:
		m_fFreezeEvents = TRUE;
		Close();
		m_fFreezeEvents = FALSE;
		hr = ProvideErrorInfo(hr);
	}
	else
	{

	 // Fire the OnDocumentOpened event...
		VARIANT rgargs[2]; 
		rgargs[0].vt = VT_DISPATCH;	get_ActiveDocument(&(rgargs[0].pdispVal));
		rgargs[1].vt = VT_BSTR; rgargs[1].bstrVal = NULL;

        RaiseAutomationEvent(DSOF_DISPID_DOCOPEN, 2, rgargs);
		VariantClear(&rgargs[0]);
    
     // Ensure we are active control...
        Activate();

     // Redraw the caption as needed...
        RedrawCaption();
    }

    m_fInDocumentLoad = FALSE;
	SetCursor(hCur);

    SAFE_RELEASE_INTERFACE(pstgTemplate);

 // Delete the temp file used in the URL download (if any)...
    if (pwszTempFile)
    {
        FPerformShellOp(FO_DELETE, pwszTempFile, NULL);
        DsoMemFree(pwszTempFile);
        SysFreeString(ProgIdOrTemplate);
    }

	return hr;
}


////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::Open
//
//  Creates a document object based on a file or URL. This simulates an
//  "open", but preserves the correct OLE embedding activation required
//  by ActiveX Documents. Opening directly from a file is not recommended.
//  We keep a lock on the original file (unless opened read-only) so the
//  user cannot tell we don't have the file "open".
//
//  The alternate ProgID allows us to open a file that is not associated 
//  with an DocObject server (like *.asp) with the server specified. Also
//  the username/password are for web access (if Document is a URL).
//
STDMETHODIMP CDsoFramerControl::Open(VARIANT Document, VARIANT ReadOnly, VARIANT ProgId, VARIANT WebUsername, VARIANT WebPassword)
{
	HRESULT   hr;
	LPWSTR    pwszDocument  = LPWSTR_FROM_VARIANT(Document);
	LPWSTR    pwszAltProgId = LPWSTR_FROM_VARIANT(ProgId);
	LPWSTR    pwszUserName  = LPWSTR_FROM_VARIANT(WebUsername);
	LPWSTR    pwszPassword  = LPWSTR_FROM_VARIANT(WebPassword);
    BOOL      fOpenReadOnly = BOOL_FROM_VARIANT(ReadOnly, FALSE);
	CLSID     clsidAlt      = GUID_NULL;
	HCURSOR	  hCur;
    IUnknown* punk = NULL;

    BIND_OPTS bopts = {sizeof(BIND_OPTS), BIND_MAYBOTHERUSER, 0, 10000};

    TRACE1("CDsoFramerControl::Open(%S)\n", pwszDocument);

 // We must have either a string (file path or URL) or an object to open from...
	if (!(pwszDocument) || (*pwszDocument == L'\0'))
    {
        if (!(pwszDocument) && ((punk = PUNK_FROM_VARIANT(Document)) == NULL))
		    return E_INVALIDARG;
    }

 // If the user passed the ProgId, find the alternative CLSID for server...
	if ((pwszAltProgId) && FAILED(CLSIDFromProgID(pwszAltProgId, &clsidAlt)))
		return E_INVALIDARG;

// Cannot create object if we are not activate yet, or if in modal condition...
    if (!(m_fInPlaceActive) || (m_fModalState)) 
        return ProvideErrorInfo(DSO_E_INMODALSTATE);

 // OK. If here, all the parameters look good and it is time to try and open
 // the document object. Start by closing any existing document first...
	if ((m_pDocObjFrame) && FAILED(hr = Close()))
		return hr;

 // Make sure we are the active component for this process...
    if (FAILED(hr = Activate())) return hr;

 // Let's make a doc frame for ourselves...
	if (!(m_pDocObjFrame = CDsoDocObject::CreateInstance((IDsoDocObjectSite*)&m_xDsoDocObjectSite)))
		return E_OUTOFMEMORY;

 // If we had delayed the frame hook, we should set it up now...
    if (!(m_pHookManager) && FDelayFrameHookSet())
    {
		m_pHookManager = CDsoFrameHookManager::RegisterFramerControl(m_hwndParent, m_hwnd);
		if (!m_pHookManager) return ProvideErrorInfo(DSO_E_FRAMEHOOKFAILED);
    }

 // Start a wait operation to notify user...
	hCur = SetCursor(LoadCursor(NULL, IDC_WAIT));
    m_fInDocumentLoad = TRUE;

 // Setup the bind options based on read-only flag....
    bopts.grfMode = (STGM_TRANSACTED | STGM_SHARE_DENY_WRITE | (fOpenReadOnly ? STGM_READ : STGM_READWRITE));

	SEH_TRY

 // Normally user gives a string that is path to file...
    if (pwszDocument)
    {
	 // Check if we are opening from a URL, and load that way...
		if (LooksLikeHTTP(pwszDocument))
		{
			hr = m_pDocObjFrame->CreateFromURL(pwszDocument, clsidAlt, &bopts, pwszUserName, pwszPassword);
		}
	 // Otherwise, check if it is a local file and open that...
		else if (FFileExists(pwszDocument))
		{
			hr = m_pDocObjFrame->CreateFromFile(pwszDocument, clsidAlt, &bopts);
		}
		else hr = E_INVALIDARG;
    }
    else if (punk)
    {
     // If we have an object to load from, try loading it direct...
        hr = m_pDocObjFrame->CreateFromRunningObject(punk, NULL, &bopts);
    }
    else hr = E_UNEXPECTED; // Unhandled load type??

 // If successful, we can activate the object...
	if (SUCCEEDED(hr))
	{
        EnableDropFile(FALSE);

		if (!m_fShowToolbars)
			m_pDocObjFrame->OnNotifyChangeToolState(FALSE);

		hr = m_pDocObjFrame->IPActivateView();
	}

    SEH_EXCEPT(hr)

  // Force a close if an error occurred...
	if (FAILED(hr))
	{
		m_fFreezeEvents = TRUE;
		Close();
		m_fFreezeEvents = FALSE;
		hr = ProvideErrorInfo(hr);
	}
	else
	{
	 // Fire the OnDocumentOpened event...
		VARIANT rgargs[2];
		rgargs[0].vt = VT_DISPATCH;	get_ActiveDocument(&(rgargs[0].pdispVal));
		rgargs[1].vt = VT_BSTR;	rgargs[1].bstrVal = SysAllocString(pwszDocument);

        RaiseAutomationEvent(DSOF_DISPID_DOCOPEN, 2, rgargs);
		VariantClear(&rgargs[1]);
		VariantClear(&rgargs[0]);
    
     // Ensure we are active control...
        Activate();

     // Redraw the caption as needed...
        RedrawCaption();
	}

    m_fInDocumentLoad = FALSE;
	SetCursor(hCur);
	return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::Save
//
//  Saves the current object. The optional SaveAs paramter lets the same
//  function act as a SaveAs method.
//
STDMETHODIMP CDsoFramerControl::Save(VARIANT SaveAsDocument, VARIANT OverwriteExisting, VARIANT WebUsername, VARIANT WebPassword)
{
	HCURSOR	  hCur;
	HRESULT	  hr = S_FALSE;
	LPWSTR    pwszDocument = LPWSTR_FROM_VARIANT(SaveAsDocument);
	LPWSTR    pwszUserName = LPWSTR_FROM_VARIANT(WebUsername);
	LPWSTR    pwszPassword = LPWSTR_FROM_VARIANT(WebPassword);
    BOOL      fOverwrite   = BOOL_FROM_VARIANT(OverwriteExisting, FALSE);
    
    TRACE1("CDsoFramerControl::Save(%S)\n", pwszDocument);
    CHECK_NULL_RETURN(m_pDocObjFrame, ProvideErrorInfo(DSO_E_DOCUMENTNOTOPEN));

 // Cannot access object if in modal condition...
    if ((m_fModalState) || (m_fNoInteractive))
        return ProvideErrorInfo(DSO_E_INMODALSTATE);

 // If user passed a value for SaveAs, it must be a valid string...
    if (!(PARAM_IS_MISSING(&SaveAsDocument)) &&
        ((pwszDocument == NULL) || (*pwszDocument == L'\0')))
		return E_INVALIDARG;

 // Raise the BeforeDocumentSaved event to host...
    if (m_dispEvents)
    {
		VARIANT rgargs[3]; VARIANT_BOOL vtboolCancel = VARIANT_FALSE;
		rgargs[2].vt = VT_DISPATCH; get_ActiveDocument(&(rgargs[2].pdispVal));
		rgargs[1].vt = VT_BSTR; rgargs[1].bstrVal = SysAllocString(pwszDocument);
		rgargs[0].vt = (VT_BOOL | VT_BYREF); rgargs[0].pboolVal = &vtboolCancel;

        RaiseAutomationEvent(DSOF_DISPID_BDOCSAVE, 3, rgargs);
		VariantClear(&rgargs[2]);
		VariantClear(&rgargs[1]);

     // Setting Cancel param will abort the save...
        if (vtboolCancel != VARIANT_FALSE)
            return E_ABORT;
    }

 // Now do the save...
	hCur = SetCursor(LoadCursor(NULL, IDC_WAIT));

    SEH_TRY

 // Determine if they want to save to a URL...
	if ((pwszDocument) && LooksLikeHTTP(pwszDocument))
	{
		hr = m_pDocObjFrame->SaveToURL(pwszDocument, fOverwrite, pwszUserName, pwszPassword);
	}
	else if (pwszDocument)
	{
	 // Save to file (local or UNC)...
		hr = m_pDocObjFrame->SaveToFile(pwszDocument, fOverwrite);
	}
	else
	{
	 // Save back to open location...
		hr = m_pDocObjFrame->Save();
	}

    SEH_EXCEPT(hr)

 // If save was successful...
    if (SUCCEEDED(hr))
    {
     // Raise the SaveComplete event to host...
        if (m_dispEvents)
        {
		    VARIANT rgargs[3];
		    rgargs[2].vt = VT_DISPATCH; get_ActiveDocument(&(rgargs[2].pdispVal));
		    rgargs[1].vt = VT_BSTR; rgargs[1].bstrVal = SysAllocString(m_pDocObjFrame->GetSourceDocName());
		    rgargs[0].vt = VT_BSTR; rgargs[0].bstrVal = SysAllocString(m_pDocObjFrame->GetSourceName());
            RaiseAutomationEvent(DSOF_DISPID_SAVECOMPLETE, 3, rgargs);
		    VariantClear(&rgargs[2]);
		    VariantClear(&rgargs[1]);
		    VariantClear(&rgargs[0]);
        }
        
     // Redraw the caption as needed...
        RedrawCaption();
    }

	SetCursor(hCur);
	return ProvideErrorInfo(hr);
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::Close
//
//  Closes the current object.
//
STDMETHODIMP CDsoFramerControl::Close()
{
	ODS("CDsoFramerControl::Close\n");

	CDsoDocObject* pdframe = m_pDocObjFrame;
	if (pdframe)
    {
     // Fire the BeforeDocumentClosed event to give host chance to cancel...
		VARIANT rgargs[2]; VARIANT_BOOL vtboolCancel = VARIANT_FALSE;
		rgargs[1].vt = VT_DISPATCH; get_ActiveDocument(&(rgargs[1].pdispVal));
		rgargs[0].vt = (VT_BOOL | VT_BYREF); rgargs[0].pboolVal = &vtboolCancel;

     // Setting Cancel param should abort the close...
        RaiseAutomationEvent(DSOF_DISPID_BDOCCLOSE, 2, rgargs);
		VariantClear(&rgargs[1]);

        if (vtboolCancel != VARIANT_FALSE)
            return E_ABORT;

     // If not canceled, clear the member variable then call close on doc frame...
	    m_pDocObjFrame = NULL;

		SEH_TRY
		pdframe->Close();
		SEH_EXCEPT_NULL
		delete pdframe;

     // Notify host that item is now closed...
        RaiseAutomationEvent(DSOF_DISPID_DOCCLOSE, 0, NULL);
	}

 // Redraw the caption as needed...
    RedrawCaption();
    EnableDropFile(TRUE);
	return S_OK;
}


////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::(put_Caption/get_Caption)
//
//  Allows you to set a custom cation for the titlebar.
//
STDMETHODIMP CDsoFramerControl::put_Caption(BSTR bstr)
{
 // Free exsiting caption (if any) and save new one (always dirties control)...
    SAFE_FREEBSTR(m_bstrCustomCaption);

 // Set the new one (if provided)...
	if ((bstr) && (SysStringLen(bstr) > 0))
		m_bstrCustomCaption = SysAllocString(bstr);

	ViewChanged();
	m_fDirty = TRUE;
	return S_OK;
}

STDMETHODIMP CDsoFramerControl::get_Caption(BSTR* pbstr)
{
	if (pbstr) *pbstr = (m_bstrCustomCaption ? 
        SysAllocString(m_bstrCustomCaption) : 
        SysAllocString(L"Office Framer Control"));
	return S_OK;
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::(put_Titlebar/get_Titlebar)
//
//  True/False. Should we display the titlebar or not?
//
STDMETHODIMP CDsoFramerControl::put_Titlebar(VARIANT_BOOL vbool)
{
	TRACE1("CDsoFramerControl::put_Titlebar(%d)\n", vbool);

    if (m_fModalState) // Cannot access object if in modal condition...
        return ProvideErrorInfo(DSO_E_INMODALSTATE);

	if (m_fShowTitlebar != (WORD)(BOOL)vbool)
	{
		m_fShowTitlebar = (BOOL)vbool;
		m_fDirty = TRUE;
		OnResize();
		ViewChanged();
	}
	return S_OK;
}

STDMETHODIMP CDsoFramerControl::get_Titlebar(VARIANT_BOOL* pbool)
{
	ODS("CDsoFramerControl::get_Titlebar\n");
	if (pbool) *pbool = (m_fShowTitlebar ? VARIANT_TRUE : VARIANT_FALSE);
	return S_OK;
}


////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::(put_Toolbars/get_Toolbars)
//
//  True/False. Should we display toolbars or not?
//
STDMETHODIMP CDsoFramerControl::put_Toolbars(VARIANT_BOOL vbool)
{
	TRACE1("CDsoFramerControl::put_Toolbars(%d)\n", vbool);

 // If the control is in modal state, we can't do things that
 // will call the server directly, like toggle toolbars...
    if ((m_fModalState) || (m_fNoInteractive))
        return ProvideErrorInfo(DSO_E_INMODALSTATE);

	if (m_fShowToolbars != (WORD)(BOOL)vbool)
	{
		m_fShowToolbars = (BOOL)vbool;
		m_fDirty = TRUE;

		if (m_pDocObjFrame)
			m_pDocObjFrame->OnNotifyChangeToolState(m_fShowToolbars);

		ViewChanged();
        OnResize();
	}
	return S_OK;
}

STDMETHODIMP CDsoFramerControl::get_Toolbars(VARIANT_BOOL* pbool)
{
	ODS("CDsoFramerControl::get_Toolbars\n");
	if (pbool) *pbool = (m_fShowToolbars ? VARIANT_TRUE : VARIANT_FALSE);
	return S_OK;
}


////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::(put_ModalState/get_ModalState)
//
//  True/False. Disables the active object (if any) thereby setting it 
//  up to behave "modal". Any time a dialog or other blocking window 
//  on the same thread is called, the developer should set this to True
//  to let the IP object know it should stay modal in the background. 
//  Set it back to False when the dialog is removed.
//
//  Technically, this should be a counter to allow for nested modal states.
//  However, we thought that might be confusing to some VB/Web developers
//  and since this is only a sample, made it a Boolean property.
//
STDMETHODIMP CDsoFramerControl::put_ModalState(VARIANT_BOOL vbool)
{
	TRACE1("CDsoFramerControl::put_ModalState(%d)\n", vbool);

 // you can't force modal state change unless active...
    if ((m_fNoInteractive) || (!m_fComponentActive))
        return ProvideErrorInfo(E_ACCESSDENIED);

	if (m_fModalState != (WORD)(BOOL)vbool)
		UpdateModalState((vbool != VARIANT_FALSE), TRUE);

	return S_OK;
}

STDMETHODIMP CDsoFramerControl::get_ModalState(VARIANT_BOOL* pbool)
{
	ODS("CDsoFramerControl::get_ModalState\n");
	if (pbool) *pbool = ((m_fModalState) ? VARIANT_TRUE : VARIANT_FALSE);
	return S_OK;
}


////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::ShowDialog
//
//  Uses IOleCommandTarget to get the embedded object to display one of
//  these standard dialogs for the user.
//
STDMETHODIMP CDsoFramerControl::ShowDialog(dsoShowDialogType DlgType)
{
	HRESULT hr = E_ACCESSDENIED;

	TRACE1("CDsoFramerControl::ShowDialog(%d)\n", DlgType);
	if ((DlgType < dsoFileNew) || (DlgType > dsoDialogProperties))
		return E_INVALIDARG;

 // Cannot access object if in modal condition...
    if ((m_fModalState) || (m_fNoInteractive))
        return ProvideErrorInfo(DSO_E_INMODALSTATE);

 // The first three dialog types we handle...
	if (DlgType < dsoDialogSaveCopy)
	{
		hr = DoDialogAction(DlgType);
	}
 // The others are provided by the server via IOleCommandTarget...
	else if (m_pDocObjFrame)
	{
		DWORD dwOleCmd;
		switch (DlgType)
		{
		case dsoDialogSaveCopy:   dwOleCmd = OLECMDID_SAVECOPYAS; break;
		case dsoDialogPageSetup:  dwOleCmd = OLECMDID_PAGESETUP;  break;
		case dsoDialogProperties: dwOleCmd = OLECMDID_PROPERTIES; break;
		default:                  dwOleCmd = OLECMDID_PRINT;
		}
		hr = m_pDocObjFrame->DoOleCommand(dwOleCmd, OLECMDEXECOPT_PROMPTUSER, NULL, NULL);
	}

	return ProvideErrorInfo(hr);
}


////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::(put_EnableFileCommand/get_EnableFileCommand)
//
//  True/False. This allows the developer to disable certain menu/toolbar
//  items that are considered "file-level" -- New, Save, Print, etc.
//
//  We use the Item parameter to set a bit flag which is used when
//  displaying the menu to enable/disable the item. The OnFileCommand
//  event will not fire for disabled commands.
//
STDMETHODIMP CDsoFramerControl::put_EnableFileCommand(dsoFileCommandType Item, VARIANT_BOOL vbool)
{
	TRACE2("CDsoFramerControl::put_EnableFileCommand(%d, %d)\n", Item, vbool);

	if ((Item < dsoFileNew) || (Item > dsoFilePrintPreview))
		return E_INVALIDARG;

 // You cannot access menu when in a modal condition...
    if ((m_fModalState) || (m_fNoInteractive))
        return ProvideErrorInfo(DSO_E_INMODALSTATE);

 // We keep bit flags for menu state. Just set the bit and a update
 // the embedded object as needed. User will see change next time menu is shown...
	UINT code = (1 << Item);
	if (vbool == 0)	m_wFileMenuFlags &= ~(code);
	else 		    m_wFileMenuFlags |= code;

	if (m_pDocObjFrame) // This should update toolbar icon (if server supports it)
		m_pDocObjFrame->DoOleCommand(OLECMDID_UPDATECOMMANDS, 0, NULL, NULL);

	return S_OK;
}

STDMETHODIMP CDsoFramerControl::get_EnableFileCommand(dsoFileCommandType Item, VARIANT_BOOL* pbool)
{
	TRACE1("CDsoFramerControl::get_EnableFileCommand(%d)\n", Item);

	if ((Item < dsoFileNew) || (Item > dsoFilePrintPreview))
		return E_INVALIDARG;

	UINT code = (1 << Item);
	if (pbool) *pbool = ((m_wFileMenuFlags & code) ? VARIANT_TRUE : VARIANT_FALSE);

	return S_OK;
}


////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::(put_BorderStyle/get_BorderStyle)
//
//  Change the border style for the control.
//
STDMETHODIMP CDsoFramerControl::put_BorderStyle(dsoBorderStyle style)
{
	ODS("CDsoFramerControl::put_BorderStyle\n");

	if ((style < dsoBorderNone) || (style > dsoBorder3DThin))
		return E_INVALIDARG;

    if (m_fModalState) // Cannot access object if in modal condition...
        return ProvideErrorInfo(DSO_E_INMODALSTATE);

	if (m_fBorderStyle != (DWORD)style)
	{
		m_fBorderStyle = style;
		m_fDirty = TRUE;
		OnResize();
		ViewChanged();
	}
	return S_OK;
}

STDMETHODIMP CDsoFramerControl::get_BorderStyle(dsoBorderStyle* pstyle)
{
	ODS("CDsoFramerControl::get_BorderStyle\n");
	if (pstyle)	*pstyle = (dsoBorderStyle)m_fBorderStyle;
	return S_OK;
}


////////////////////////////////////////////////////////////////////////
// Control Color Properties...
//
//
STDMETHODIMP CDsoFramerControl::put_BorderColor(OLE_COLOR clr)
{
	ODS("CDsoFramerControl::put_BorderColor\n");
	if (m_clrBorderColor != clr)
	{
		m_clrBorderColor = clr;
		m_fDirty = TRUE;
		ViewChanged();
	}
	return S_OK;
}

STDMETHODIMP CDsoFramerControl::get_BorderColor(OLE_COLOR* pclr)
{if (pclr) *pclr = m_clrBorderColor; return S_OK;}

STDMETHODIMP CDsoFramerControl::put_BackColor(OLE_COLOR clr)
{
	ODS("CDsoFramerControl::put_BackColor\n");
	if (m_clrBackColor != clr)
	{
		m_clrBackColor = clr;
		m_fDirty = TRUE;
		ViewChanged();
	}
	return S_OK;
}

STDMETHODIMP CDsoFramerControl::get_BackColor(OLE_COLOR* pclr)
{if (pclr) *pclr = m_clrBackColor; return S_OK;}

STDMETHODIMP CDsoFramerControl::put_ForeColor(OLE_COLOR clr)
{
	ODS("CDsoFramerControl::put_ForeColor\n");
	if (m_clrForeColor != clr)
	{
		m_clrForeColor = clr;
		m_fDirty = TRUE;
		ViewChanged();
	}
	return S_OK;
}

STDMETHODIMP CDsoFramerControl::get_ForeColor(OLE_COLOR* pclr)
{if (pclr) *pclr = m_clrForeColor; return S_OK;}

STDMETHODIMP CDsoFramerControl::put_TitlebarColor(OLE_COLOR clr)
{
	ODS("CDsoFramerControl::put_TitlebarColor\n");
	if (m_clrTBarColor != clr)
	{
		m_clrTBarColor = clr;
		m_fDirty = TRUE;
		ViewChanged();
	}
	return S_OK;
}

STDMETHODIMP CDsoFramerControl::get_TitlebarColor(OLE_COLOR* pclr)
{if (pclr) *pclr = m_clrTBarColor; return S_OK;}


STDMETHODIMP CDsoFramerControl::put_TitlebarTextColor(OLE_COLOR clr)
{
	ODS("CDsoFramerControl::put_TitlebarTextColor\n");
	if (m_clrTBarTextColor != clr)
	{
		m_clrTBarTextColor = clr;
		m_fDirty = TRUE;
		ViewChanged();
	}
	return S_OK;
}

STDMETHODIMP CDsoFramerControl::get_TitlebarTextColor(OLE_COLOR* pclr)
{if (pclr) *pclr = m_clrTBarTextColor; return S_OK;}


////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::(put_Menubar/get_Menubar)
//
//  True/False. Should we display menu bar?
//
STDMETHODIMP CDsoFramerControl::put_Menubar(VARIANT_BOOL vbool)
{
	TRACE1("CDsoFramerControl::put_Menubar(%d)\n", vbool);

 // If the control is in modal state, we can't do things that
 // will call the server directly, like toggle menu bar...
    if ((m_fModalState) || (m_fNoInteractive))
        return ProvideErrorInfo(DSO_E_INMODALSTATE);

	if (m_fShowMenuBar != (WORD)(BOOL)vbool)
	{
		m_fShowMenuBar = (BOOL)vbool;
		m_fDirty = TRUE;
		ViewChanged();
        OnResize();
	}
	return S_OK;
}

STDMETHODIMP CDsoFramerControl::get_Menubar(VARIANT_BOOL* pbool)
{
	ODS("CDsoFramerControl::get_Menubar\n");
	if (pbool) *pbool = (m_fShowMenuBar ? VARIANT_TRUE : VARIANT_FALSE);
	return S_OK;
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::(put_HostName/get_HostName)
//
//  String setting used for host application name (used in embedding)
//
STDMETHODIMP CDsoFramerControl::put_HostName(BSTR bstr)
{
	TRACE1("CDsoFramerControl::put_HostName(%S)\n", bstr);
    SAFE_FREESTRING(m_pwszHostName);

    if ((bstr) && (SysStringLen(bstr) > 0))
        m_pwszHostName = DsoCopyString(bstr);

    return S_OK;
}

STDMETHODIMP CDsoFramerControl::get_HostName(BSTR* pbstr)
{
	ODS("CDsoFramerControl::get_HostName\n");
    if (pbstr)
        *pbstr = SysAllocString((m_pwszHostName ? m_pwszHostName : L"DsoFramerControl"));
	return S_OK;
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::get_DocumentFullName
//
//  Gets FullName of embedded source file (where default save will save to).
//
STDMETHODIMP CDsoFramerControl::get_DocumentFullName(BSTR* pbstr)
{
	ODS("CDsoFramerControl::get_DocumentFullName\n");
    CHECK_NULL_RETURN(pbstr, E_POINTER);
 // Ask doc object site for the source name...
    *pbstr = (m_pDocObjFrame) ? SysAllocString(m_pDocObjFrame->GetSourceName()) : NULL;
	return S_OK;
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::get_DocumentName
//
//  Gets just the file name of embedded source file.
//
STDMETHODIMP CDsoFramerControl::get_DocumentName(BSTR* pbstr)
{
	ODS("CDsoFramerControl::get_DocumentName\n");
    CHECK_NULL_RETURN(pbstr, E_POINTER);
 // Ask doc object site for the source doc name...
    *pbstr = (m_pDocObjFrame) ? SysAllocString(m_pDocObjFrame->GetSourceDocName()) : NULL;
	return S_OK;
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::get_IsReadOnly
//
//  Returns if file open read-only.
//
STDMETHODIMP CDsoFramerControl::get_IsReadOnly(VARIANT_BOOL* pbool)
{
	ODS("CDsoFramerControl::get_IsReadOnly\n");
	CHECK_NULL_RETURN(pbool, E_POINTER);
    CHECK_NULL_RETURN(m_pDocObjFrame, ProvideErrorInfo(DSO_E_DOCUMENTNOTOPEN));
	*pbool = (m_pDocObjFrame->IsReadOnly() ? VARIANT_TRUE : VARIANT_FALSE);
	return S_OK;
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::get_IsDirty
//
//  Returns TRUE if doc object says it has changes.
//
STDMETHODIMP CDsoFramerControl::get_IsDirty(VARIANT_BOOL* pbool)
{
	ODS("CDsoFramerControl::get_IsDirty\n");
	CHECK_NULL_RETURN(pbool, E_POINTER);
    CHECK_NULL_RETURN(m_pDocObjFrame, ProvideErrorInfo(DSO_E_DOCUMENTNOTOPEN));
	*pbool = (m_pDocObjFrame->IsDirty() ? VARIANT_TRUE : VARIANT_FALSE);
	return S_OK;
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::LockServer
//
//  Allows host to temporarily lock the current document server running, 
//  so that when closing and reopening files using the same server doesn't
//  suffer startup/shutdown penelty of the server going away between calls.
//
STDMETHODIMP CDsoFramerControl::put_LockServer(VARIANT_BOOL vbool)
{
	BOOL fLock = (vbool != VARIANT_FALSE);
	TRACE1("CDsoFramerControl::put_LockServer(%d)\n", (DWORD)vbool);

 // We must have a server open to set a lock...
	if ((fLock) && (m_pDocObjFrame == NULL))
		return ProvideErrorInfo(DSO_E_DOCUMENTNOTOPEN);

	return SetTempServerLock(fLock);
}

STDMETHODIMP CDsoFramerControl::get_LockServer(VARIANT_BOOL* pvbool)
{
	ODS("CDsoFramerControl::get_LockServer\n");
	CHECK_NULL_RETURN(pvbool, E_POINTER);
	*pvbool = (VARIANT_BOOL)((m_pServerLock) ? VARIANT_TRUE : VARIANT_FALSE);
	return S_OK;
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::GetDataObjectContent
//
//  Allows caller to get data from document in clipboard format but without
//  using the clipboard (instead it uses IDataObject). The server has to 
//  support the request format to return the data. Data is returned as byte array.
//
//  NOTE: This only gets the document body content, not the entire document. If
//  you want the entire document, use the Save method with temp filename to save to.
//
STDMETHODIMP CDsoFramerControl::GetDataObjectContent(VARIANT ClipFormatNameOrNumber, VARIANT *pvResults)
{
	ODS("CDsoFramerControl::GetDataObjectContent()\n");
	CHECK_NULL_RETURN(pvResults, E_POINTER); VariantInit(pvResults);
    CHECK_NULL_RETURN(m_pDocObjFrame, ProvideErrorInfo(DSO_E_DOCUMENTNOTOPEN));

 // If the control is in modal state, we can't do anything...
    if ((m_fModalState) || (m_fNoInteractive))
        return ProvideErrorInfo(DSO_E_INMODALSTATE);

	return m_pDocObjFrame->HrGetDataFromObject(&ClipFormatNameOrNumber, pvResults);
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::GetDataObjectContent
//
//  Allows caller to set data into the document body (similar to a paste, but
//  without using the clipboard). You must supply the format name or number and
//  the data in a byte array format.
//
STDMETHODIMP CDsoFramerControl::SetDataObjectContent(VARIANT ClipFormatNameOrNumber, VARIANT DataByteArray)
{
	ODS("CDsoFramerControl::SetDataObjectContent()\n");
    CHECK_NULL_RETURN(m_pDocObjFrame, ProvideErrorInfo(DSO_E_DOCUMENTNOTOPEN));

 // If the control is in modal state, we can't do anything...
    if ((m_fModalState) || (m_fNoInteractive))
        return ProvideErrorInfo(DSO_E_INMODALSTATE);

	return m_pDocObjFrame->HrSetDataInObject(&ClipFormatNameOrNumber, &DataByteArray, TRUE);
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::put_ActivationPolicy
//
//  Allows caller to set policy for how activation changes affect the embedded
//  object. This is actually a bit flag, so user can OR settings.
//
STDMETHODIMP CDsoFramerControl::put_ActivationPolicy(dsoActivationPolicy lPolicy)
{
	TRACE1("CDsoFramerControl::put_ActivationPolicy(%d)\n", lPolicy);
	if (m_pDocObjFrame) return E_ACCESSDENIED;

	if ((lPolicy < dsoDefaultBehavior) || (lPolicy > 0x0F))
		return E_INVALIDARG;

	m_lActivationPolicy = lPolicy;
	return S_OK;
}

STDMETHODIMP CDsoFramerControl::get_ActivationPolicy(dsoActivationPolicy *plPolicy)
{if (plPolicy) *plPolicy = (dsoActivationPolicy)m_lActivationPolicy; return S_OK;}


////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::put_FrameHookPolicy
//
//  Allows caller to set policy for the frame hook.
//
STDMETHODIMP CDsoFramerControl::put_FrameHookPolicy(dsoFrameHookPolicy lPolicy)
{
	HRESULT hr = E_ACCESSDENIED;
	TRACE1("CDsoFramerControl::put_FrameHookPolicy(%d)\n", lPolicy);

	switch (lPolicy)
	{
	case dsoDisableHook:
		if (m_pHookManager) break;
	// else fallthrough...

	case dsoNormalBehavior:
	case dsoSetOnFirstOpen:
		m_lHookPolicy = lPolicy; hr = S_OK;
		break;

	case dsoResetNow:
		if (!FRunningInDesignMode())
			hr = ResetFrameHook(0);
		break;

	default:
		hr = E_INVALIDARG;
	}

	return hr;
}

STDMETHODIMP CDsoFramerControl::get_FrameHookPolicy(dsoFrameHookPolicy *plPolicy)
{if (plPolicy) *plPolicy = (dsoFrameHookPolicy)m_lHookPolicy; return S_OK;}


////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::put_MenuAccelerators
//
//  Sets flag to determine if control should try to handle keys for menu
//  shortcuts when the menus are displayed.
//
STDMETHODIMP CDsoFramerControl::put_MenuAccelerators(VARIANT_BOOL vbool)
{
	ODS("CDsoFramerControl::put_MenuAccelerators\n");
	m_fDisableMenuAccel = (vbool == VARIANT_FALSE);
	return S_OK;
}

STDMETHODIMP CDsoFramerControl::get_MenuAccelerators(VARIANT_BOOL* pvbool)
{
	ODS("CDsoFramerControl::get_MenuAccelerators\n");
	CHECK_NULL_RETURN(pvbool, E_POINTER);
	*pvbool = (m_fDisableMenuAccel ? VARIANT_FALSE : VARIANT_TRUE);
	return S_OK;
}


////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::put_EventsEnabled
//
//  Sets flag to determine if events are frozen or not. 
//
STDMETHODIMP CDsoFramerControl::put_EventsEnabled(VARIANT_BOOL vbool)
{
	ODS("CDsoFramerControl::put_EventsEnabled\n");
	m_fFreezeEvents = (vbool == VARIANT_FALSE);
	return S_OK;
}

STDMETHODIMP CDsoFramerControl::get_EventsEnabled(VARIANT_BOOL* pvbool)
{
	ODS("CDsoFramerControl::get_EventsEnabled\n");
	CHECK_NULL_RETURN(pvbool, E_POINTER);
	*pvbool = (m_fFreezeEvents ? VARIANT_FALSE : VARIANT_TRUE);
	return S_OK;
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl - IDispatch Implementation
//

////////////////////////////////////////////////////////////////////////
// Control's IDispatch Functions
//
//  These are largely standard and just forward calls to the functions
//  above. The only interesting thing here is the "hack" in Invoke to 
//  tell VB/IE that the control is always "Enabled".
//
STDMETHODIMP CDsoFramerControl::GetTypeInfoCount(UINT* pctinfo)
{if (pctinfo) *pctinfo = 1; return S_OK;}

STDMETHODIMP CDsoFramerControl::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo)
{
	HRESULT hr = S_OK;

    ODS("CDsoFramerControl::GetTypeInfo\n");
	CHECK_NULL_RETURN(ppTInfo, E_POINTER); *ppTInfo = NULL;

 // We only support default interface late bound...
	CHECK_NULL_RETURN((iTInfo == 0), DISP_E_BADINDEX);

 // Load the type lib if we don't have the information already.
	if (NULL == m_ptiDispType)
	{
		hr = DsoGetTypeInfoEx(LIBID_DSOFramer, 0, DSOFRAMERCTL_VERSION_MAJOR, DSOFRAMERCTL_VERSION_MINOR,
			v_hModule, IID__FramerControl, &m_ptiDispType);
	}

 // Return interface with ref count (if we have it, otherwise error)...
    SAFE_SET_INTERFACE(*ppTInfo, m_ptiDispType);
    return hr;
}

STDMETHODIMP CDsoFramerControl::GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId)
{
    HRESULT hr;
    ITypeInfo *pti;

	ODS("CDsoFramerControl::GetIDsOfNames\n");
    CHECK_NULL_RETURN((IID_NULL == riid), DISP_E_UNKNOWNINTERFACE);

 // Get the type info for this dispinterface...
    hr = GetTypeInfo(0, lcid, &pti);
    RETURN_ON_FAILURE(hr);

 // Ask OLE to translate the name...
    hr = pti->GetIDsOfNames(rgszNames, cNames, rgDispId);
    pti->Release();
    return hr;
}

STDMETHODIMP CDsoFramerControl::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr)
{
    HRESULT hr;
    ITypeInfo *pti;;

 // VB loves to check for this property (Enabled) during design time.
 // We don't implement it for this control, but we'll return TRUE to
 // to let it know if is enabled and don't bother with call to ITypeInfo...
	if ((dispIdMember == DISPID_ENABLED) && (wFlags & DISPATCH_PROPERTYGET))
	{
		if (pVarResult) // We are always enabled...
			{pVarResult->vt = VT_BOOL;	pVarResult->boolVal = VARIANT_TRUE;	}
		return S_OK;
	}

	TRACE1("CDsoFramerControl::Invoke(dispid = %d)\n", dispIdMember);
    CHECK_NULL_RETURN((IID_NULL == riid), DISP_E_UNKNOWNINTERFACE);

 // Get the type info for this dispinterface...
    hr = GetTypeInfo(0, lcid, &pti);
    RETURN_ON_FAILURE(hr);

 // Store pExcepInfo (to fill-in disp excepinfo if error occurs)...
    m_pDispExcep = pExcepInfo;

 // Call the method using TypeInfo (OLE will call v-table method for us)...
    hr = pti->Invoke((PVOID)this, dispIdMember, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);

    m_pDispExcep = NULL; // Don't need this anymore...
    pti->Release();
    return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::ProvideErrorInfo
//
//  Fills in custom error information (as needed).
//
STDMETHODIMP CDsoFramerControl::ProvideErrorInfo(HRESULT hres)
{
 // Don't need to do anything on success...
	if ((hres == S_OK) || SUCCEEDED(hres))
		return hres;

 // Fill in the error information as needed...
	return DsoReportError(hres, NULL, m_pDispExcep);
}
