/***************************************************************************
 * DSOFCONTROL.CPP
 *
 * CDsoFramerControl: The Base Control
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
// CDsoFramerControl - The Main Control Class
//
//  Serves as a base for the ActiveX Document site object which will do
//  the actual embedding.
//
CDsoFramerControl::CDsoFramerControl(LPUNKNOWN punk)
{
	ODS("CDsoFramerControl::CDsoFramerControl\n");
	m_pOuterUnknown = ((punk) ? punk : (IUnknown*)&m_xInternalUnknown);
}

CDsoFramerControl::~CDsoFramerControl(void)
{
	ODS("CDsoFramerControl::~CDsoFramerControl\n");

	SAFE_RELEASE_INTERFACE(m_ptiDispType);
	SAFE_RELEASE_INTERFACE(m_pClientSite);
	SAFE_RELEASE_INTERFACE(m_pControlSite);
	SAFE_RELEASE_INTERFACE(m_pInPlaceSite);
	SAFE_RELEASE_INTERFACE(m_pInPlaceFrame);
	SAFE_RELEASE_INTERFACE(m_pInPlaceUIWindow);
	SAFE_RELEASE_INTERFACE(m_pViewAdviseSink);
	SAFE_RELEASE_INTERFACE(m_pOleAdviseHolder);
	SAFE_RELEASE_INTERFACE(m_pDataAdviseHolder);
	SAFE_RELEASE_INTERFACE(m_dispEvents);
	SAFE_RELEASE_INTERFACE(m_pOleStorage);
    SAFE_FREESTRING(m_pwszHostName);
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::InitializeNewInstance
//
//  Sets up the inital state for the control. Any other member variable
//  not set here will be NULL/FALSE thanks to zero-init in MemAlloc.
//
STDMETHODIMP CDsoFramerControl::InitializeNewInstance()
{
	m_clrBorderColor	= 0x80000010; // System Button Shadow Color
	m_clrBackColor		= 0x80000005; // System Window Background Color
	m_clrForeColor		= 0x80000008; // System Window Text Color
	m_clrTBarColor		= 0x8000000D; // System Highlight Color
	m_clrTBarTextColor	= 0x8000000E; // System Highlight Text Color
	m_fBorderStyle		= dsoBorderFlat;
	m_fShowTitlebar		= TRUE;
	m_fShowToolbars		= TRUE;
    m_fShowMenuBar      = TRUE;
	m_wFileMenuFlags    = 0xFFFF;     // All items are "enabled" by default
    m_Size.cx           = 240;
    m_Size.cy           = 240;
	return S_OK;
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::InPlaceActivate
//
//  Does all the work in-place activating and/or ui activating the
//  control when asked to do so. Sets up the main control window, the 
//  needed interfaces, and (if not in design) subclasses the main top-level
//  window we can monitor messages that must be forwarded to an ip active
//  DocObject when we have one loaded.
//
STDMETHODIMP CDsoFramerControl::InPlaceActivate(LONG lVerb)
{
    HRESULT hr;
    SIZEL sizel;

	TRACE1("CDsoFramerControl::InPlaceActivate - %d\n", lVerb);

 // if we don't have a client site, then there's not much to do.
    if (!m_pClientSite) return S_OK;

 // get an InPlace site pointer.
    if (NULL == m_pInPlaceSite)
	{
        hr = m_pClientSite->QueryInterface(IID_IOleInPlaceSite, (void **)&m_pInPlaceSite);
        RETURN_ON_FAILURE(hr);
    }

 // if we're not already active, go and do it.
    if (!m_fInPlaceActive)
	{
        OLEINPLACEFRAMEINFO InPlaceFrameInfo;
        RECT rcPos, rcClip;

        hr = m_pInPlaceSite->CanInPlaceActivate();
        if (hr != S_OK){ if (!FAILED(hr)) hr = E_FAIL; goto cleanup;}

     // if we are here, then we have permission to go in-place active.
     // now, announce our intentions to actually go ahead and do this.
        hr = m_pInPlaceSite->OnInPlaceActivate();
        GOTO_ON_FAILURE(hr, cleanup);

     // if we're here, we're ready to go in-place active.  we just need
     // to set up some flags, and then create the window [if we have one]
        m_fInPlaceActive = TRUE;

     // we need to get some information about our location in the parent
     // window, as well as some information about the parent
        hr = m_pInPlaceSite->GetWindow(&m_hwndParent);
		if (FAILED(hr) || !IsWindow(m_hwndParent))
		{
			hr = OLE_E_INVALIDHWND;
			goto cleanup;
		}

        InPlaceFrameInfo.cb = sizeof(OLEINPLACEFRAMEINFO);
        hr = m_pInPlaceSite->GetWindowContext(&m_pInPlaceFrame, &m_pInPlaceUIWindow, &rcPos, &rcClip, &InPlaceFrameInfo);
        GOTO_ON_FAILURE(hr, cleanup);

     // make sure we'll display ourselves in the correct location with the correct size
        sizel.cx = rcPos.right - rcPos.left;
        sizel.cy = rcPos.bottom - rcPos.top;
        m_Size = sizel;

        m_xOleInplaceObject.SetObjectRects(&rcPos, &rcClip);

     // finally, create our window if we have to!
		if (NULL == m_hwnd)
		{
			WNDCLASS wndclass;

			EnterCriticalSection(&v_csecThreadSynch);

         // if class info has not yet been registered, need to do that now...
			if (GetClassInfo(v_hModule, "DSOFramerOCXWnd", &wndclass) == 0)
			{
				memset(&wndclass, 0, sizeof(WNDCLASS));
				wndclass.style          = CS_VREDRAW | CS_HREDRAW | CS_DBLCLKS;
				wndclass.lpfnWndProc    = CDsoFramerControl::ControlWindowProc;
				wndclass.hInstance      = v_hModule;
				wndclass.hCursor        = LoadCursor(NULL, IDC_ARROW);
				wndclass.lpszClassName  = "DSOFramerOCXWnd";
				RegisterClass(&wndclass);
			}

         // create the window with the extent size and position...
			m_hwnd = CreateWindowEx(0, "DSOFramerOCXWnd", NULL,
                            WS_CHILD | WS_CLIPSIBLINGS | WS_VISIBLE,
                            rcPos.left, rcPos.top, m_Size.cx, m_Size.cy,
                            m_hwndParent, NULL,	v_hModule, NULL);

			if (m_hwnd) SetWindowLong(m_hwnd, GWL_USERDATA, (LONG)this);

			LeaveCriticalSection(&v_csecThreadSynch);

			if (!m_hwnd) {hr = E_FAIL; goto cleanup;}
		}

     // finally, tell the host of this
        m_pClientSite->ShowObject();
    }

 // if we're not inplace visible yet, do so now.
    if (!m_fInPlaceVisible)
        SetInPlaceVisible(TRUE);


 // if we were asked to UIActivate, and we currently are not, do so!
    if ((lVerb == OLEIVERB_PRIMARY || 
		 lVerb == OLEIVERB_UIACTIVATE) && (!m_fUIActive))
	{
	 // It is not typical to UI activate ocx during design mode, but if it happens and host
     // has not customized the hook policy, let's switch to delay set so that we don't
	 // accidently hook the VS/VB IDE designer window.
		if (FRunningInDesignMode() && (m_lHookPolicy == 0))
			m_lHookPolicy = dsoSetOnFirstOpen;

     // if we don't have it already, hook top-level parent window when we go
     // ui active; this is needed to track WM_ACTIVATEAPP and other critical messages
     // needed by ole docobjs....
        if ((m_pHookManager == NULL) && FUseFrameHook() && !FDelayFrameHookSet())
        {
			m_pHookManager = CDsoFrameHookManager::RegisterFramerControl(m_hwndParent, m_hwnd);
		    if (!m_pHookManager)
            {
                ODS(" ** UIActivate FAILED (DSO_E_FRAMEHOOKFAILED)\n");
                return ProvideErrorInfo(DSO_E_FRAMEHOOKFAILED);
            }
        }

     // Now UI activate the control for the host. This implies taking focus and
     // setting up the control as the current active object...
        hr = UIActivate(FALSE);

    }

    // Setup drop file support (but only when there is no current embedding)...
    if (m_pDocObjFrame == NULL)
        EnableDropFile(TRUE);

    return S_OK; // be-de-be-de-be-de that's all folks!

cleanup:
 // something catastrophic happened [or, at least something bad].
 // die a horrible fiery mangled painful death...
	TRACE1(" *** InPlaceActivate FAILED (hr=0x%X) ***\n", hr);
    m_fInPlaceActive = FALSE;
    return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::UIActivate
//
//  Make sure our control is UI Active and set as ActiveObject for host.
//
STDMETHODIMP CDsoFramerControl::UIActivate(BOOL fForceUIActive)
{
    HRESULT hr = S_FALSE;
	TRACE1("CDsoFramerControl::UIActivate(fForceUIActive=%d)\n", fForceUIActive);

 // We can't do anything if we aren't inplace active and visible!
    if (!(m_fInPlaceActive) || !(m_fInPlaceVisible))
        return E_UNEXPECTED;

 // If we are not already UI active or are being asked to force it, do the call
    if ((!(m_fUIActive) || (fForceUIActive)) && (m_pInPlaceSite))
    {
        m_fUIActive = TRUE;

     // inform the container of our intent
        hr = m_pInPlaceSite->OnUIActivate();

        if (SUCCEEDED(hr) || (fForceUIActive))
        {
         // take the focus  [which is what UI Activation is all about !]
	        SetFocus(m_hwnd);

         // set ourselves up in the host.
		    if (m_pInPlaceFrame)
	            m_pInPlaceFrame->SetActiveObject((IOleInPlaceActiveObject*)&m_xOleInplaceActiveObject, NULL);

            if (m_pInPlaceUIWindow)
                m_pInPlaceUIWindow->SetActiveObject((IOleInPlaceActiveObject*)&m_xOleInplaceActiveObject, NULL);

         // we have to explicitly say we don't wany any border space.
		    if (m_pInPlaceFrame)
			    m_pInPlaceFrame->SetBorderSpace(NULL);
			    
            if (m_pInPlaceUIWindow)
                m_pInPlaceUIWindow->SetBorderSpace(NULL);

         // Ensure docobj component state is active as well...
            if (!m_fComponentActive)
            {
             // Tell manager we are the active component now...
                if (m_pHookManager) m_pHookManager->SetActiveComponent(m_hwnd);
                else OnComponentActivationChange(TRUE);
            }
        }
    }

    return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::SetInPlaceVisible
//
//  Make sure our control is visible or not (does not actually take ui
//  focus, that should be done in InPlaceActivate).
//
STDMETHODIMP_(void) CDsoFramerControl::SetInPlaceVisible(BOOL fShow)
{
    BOOL fVisible;
    m_fInPlaceVisible = fShow;

    // don't do anything if we don't have a window.  otherwise, set it
    if (m_hwnd)
	{
        fVisible = IsWindowVisible(m_hwnd);

        if (fVisible && !fShow)
			ShowWindow(m_hwnd, SW_HIDE);
        else if (!fVisible && fShow)
			ShowWindow(m_hwnd, SW_SHOWNA);
    }
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::OnDraw
//
//  Drawing is largely limited to border and titlebar, except when in
//  design mode. We don't do any background painting to avoid paint problems
//  when we have an ui active DocObject. It should do any painting for us.
//
STDMETHODIMP_(void) CDsoFramerControl::OnDraw(DWORD dvAspect, HDC hdcDraw, LPRECT prcBounds, LPRECT prcWBounds, HDC hicTargetDev, BOOL fOptimize)
{
	COLORREF clrBackground, clrForeground, clrBorder, clrT;
	HBRUSH   hbshBackground, hbshBorder;
	HFONT    hfontText, hfontOld;
	HPALETTE hpal, hpalOld;
    LOGFONT  lfText = {-11,0,0,0,FW_NORMAL,0,0,0,0,0,0,0,FF_SWISS, "Tahoma"};
	RECT     rc, rcT;
	int      iTxtMode, iMapMode;

 // Copy out the bounding rect...
	CopyRect(&rc, prcBounds);

 // On 256 color systems, we may need a palette...
    if ((GetDeviceCaps(hdcDraw, RASTERCAPS) & RC_PALETTE) &&
		((GetDeviceCaps(hdcDraw, BITSPIXEL) * GetDeviceCaps(hdcDraw, PLANES)) == 8))
    {
        hpal = CreateHalftonePalette(hdcDraw);
		hpalOld = SelectPalette(hdcDraw, hpal, TRUE);
		RealizePalette(hdcDraw);
    }
	else hpal = NULL;

 // Translate the OLE_COLOR values to a matching COLORREF in palette...
	OleTranslateColor(m_clrBackColor,   hpal, &clrBackground);
	OleTranslateColor(m_clrForeColor,   hpal, &clrForeground);
	OleTranslateColor(m_clrBorderColor, hpal, &clrBorder);

 // Create the resources we need for drawing and setting up the DC...
	hbshBackground = CreateSolidBrush(clrBackground);
	hbshBorder     = CreateSolidBrush(clrBorder);

	hfontText = CreateFontIndirect(&lfText);
	hfontOld = (HFONT)SelectObject(hdcDraw, hfontText);
	iTxtMode = SetBkMode(hdcDraw, TRANSPARENT);
    iMapMode = SetMapMode(hdcDraw, MM_TEXT);

 // Based on the choosen border style, draw the border and deflate the
 // drawing rect accordingly...
	switch (m_fBorderStyle)
	{
	case 1:
		{
		HBRUSH hbT = GetSysColorBrush(COLOR_BTNFACE);
		FrameRect(hdcDraw, &rc, hbshBorder);	InflateRect(&rc, -1, -1);
		FrameRect(hdcDraw, &rc, hbT);			InflateRect(&rc, -1, -1);
		FrameRect(hdcDraw, &rc, hbshBorder);	InflateRect(&rc, -1, -1);
		}
		break;

	case 2:
		DrawEdge(hdcDraw, &rc, EDGE_SUNKEN, BF_RECT);	
		InflateRect(&rc, -2, -2);
		break;

	case 3:
		DrawEdge(hdcDraw, &rc, BDR_SUNKENOUTER, BF_RECT);
		InflateRect(&rc, -1, -1);
		break;

	//default: no border...
	}

 // Draw the titlebar if visible...
	if (m_fShowTitlebar)
	{
		COLORREF clrTitlebar, clrTitlebarText;
		HBRUSH   hbshTitlebar, hbshT;
        LPWSTR   pwszDCaption = NULL;
        LPWSTR   pwszROCaption = NULL;
        LPWSTR   pwszCaption = ((m_bstrCustomCaption) ? 
					m_bstrCustomCaption : L"KB311765: Office Framer Control Sample");

		CopyRect(&rcT, &rc);
		rcT.bottom = rcT.top + 21;

		OleTranslateColor(m_clrTBarColor,     hpal, &clrTitlebar);
		OleTranslateColor(m_clrTBarTextColor, hpal, &clrTitlebarText);

		hbshTitlebar = CreateSolidBrush(clrTitlebar);
		FillRect(hdcDraw, &rcT, hbshTitlebar);

		DrawIconEx(hdcDraw, rcT.left+2, rcT.top+2, v_icoOffDocIcon, 16, 16, 0, NULL, DI_NORMAL);

		clrT = SetTextColor(hdcDraw, clrTitlebarText);
		rcT.left += 22;

     // If we have an embedded object, check to see if it is an open doc and
     // append the name to the titlebar. Also append "Read-Only" string if file
     // was open in read-only mode. 
        if (m_pDocObjFrame)
        {
            if ((pwszDCaption = (LPWSTR)m_pDocObjFrame->GetSourceDocName()) != NULL)
            {
                LPCWSTR pwszCat[2] = {L" - ", pwszDCaption};
                pwszDCaption = DsoCopyStringCatEx(pwszCaption, 2, pwszCat);
                pwszCaption = pwszDCaption;
            }

            if (m_pDocObjFrame->IsReadOnly())
            {
                pwszROCaption = DsoCopyStringCat(pwszCaption, L" (Read-Only)");
                pwszCaption = pwszROCaption;
            }
        }
 
	 // Draw out the custom titlebar caption in Unicode if on NT/XP OS, since
	 // a custom caption may have non-ANSI characters...
        FDrawText(hdcDraw, pwszCaption, &rcT, DT_VCENTER|DT_SINGLELINE);

     // Cleanup temp strings...
        if (pwszROCaption) DsoMemFree(pwszROCaption);
        if (pwszDCaption) DsoMemFree(pwszDCaption);

		CopyRect(&rcT, &rc);
		rcT.bottom = rcT.top + 21;
		rcT.top = rcT.bottom - 1;

		hbshT = GetSysColorBrush(COLOR_BTNSHADOW);
		FillRect(hdcDraw, &rcT, hbshT);

		SetTextColor(hdcDraw, clrT);
		DeleteObject(hbshTitlebar);
	}

 // Draw menu bar if visible...
    if (m_fShowMenuBar)
    {
        GetSizeRectForMenuBar(prcBounds, &rcT);
		FillRect(hdcDraw, &rcT, GetSysColorBrush(COLOR_BTNFACE));

     // When in design mode, we just draw placeholder...
        if (FRunningInDesignMode())
        {
		    clrT = SetTextColor(hdcDraw, GetSysColor(COLOR_BTNSHADOW));
		    DrawText(hdcDraw, "Menu Bar", -1, &rcT, DT_SINGLELINE|DT_VCENTER|DT_CENTER);
		    SetTextColor(hdcDraw, clrT);
        }
        else
        {
            HMENU hCurMenu;
            CHAR szbuf[DSO_MAX_MENUNAME_LENGTH];
            UINT i, yTop, yBottom, xLast, xNext;
            SIZE ptMItem; RECT rcMDraw;
            BOOL fSelected;
            DWORD dwDTFlags = (DT_SINGLELINE | DT_VCENTER);

         // Get the current menu and setup to draw the number of items in it...
            if (hCurMenu = GetActivePopupMenu())
            {
                yTop = rcT.top + 3; yBottom = rcT.bottom - 3; xLast = rcT.left + 2;
                if (hCurMenu == m_hmenuFilePopup)
                    { m_cMenuItems = 1; }
                else
                {
                    if ((m_cMenuItems = (USHORT)GetMenuItemCount(hCurMenu)) > DSO_MAX_MENUITEMS)
                        m_cMenuItems = DSO_MAX_MENUITEMS;
                }

             // If we are on an OS that supports it, hide the prefix when we are not UI active
             // or if the host has disabled menu accelerators for this instance...
                if (v_fWindows2KPlus && (!(m_fUIActive) || (m_fDisableMenuAccel)))
                    dwDTFlags |= DT_HIDEPREFIX;

             // For each item in the menu, get the menu name and draw it to the 
             // menu bar. We need to calculate the size taken by the font used and
             // store the information in an array used by mouse move handler to 
             // determine which menu item the user is over...
				for (i = 0; i < m_cMenuItems; i++)
				{
					szbuf[0] = '\0';
                    if (i == 0) lstrcpy(szbuf, "&File");
                    else GetMenuString(hCurMenu, i, szbuf, DSO_MAX_MENUNAME_LENGTH, MF_BYPOSITION);

                    GetTextExtentPoint32(hdcDraw, szbuf, lstrlen(szbuf), &ptMItem);
                    xNext = (xLast + ptMItem.cx + 2);
                    if (xNext > (UINT)(rcT.right)){ m_cMenuItems = (USHORT)i; break; }
                    SetRect(&rcMDraw, xLast, yTop, xNext, yBottom);

                    if (fOptimize)
						CopyRect(&m_rgrcMenuItems[i], &rcMDraw);

                 // If user is over this particular item, we draw it as selected...
                    fSelected = ((m_wSelMenuItem) && (i == (UINT)(m_wSelMenuItem - 1)));
                    if (fSelected)
                    {
						FillRect(hdcDraw, &rcMDraw, GetSysColorBrush(COLOR_HIGHLIGHT));
						clrT = SetTextColor(hdcDraw, GetSysColor(COLOR_HIGHLIGHTTEXT));
                    }
                    rcMDraw.left += 4;
                    DrawText(hdcDraw, szbuf, lstrlen(szbuf), &rcMDraw, dwDTFlags);
                    
					if (fSelected)
						SetTextColor(hdcDraw, clrT);

                    xLast = xNext + 2;
				}
            }
        }
		rcT.top = rcT.bottom - 1; // Draw a line to separate menu from workspace...
		FillRect(hdcDraw, &rcT, GetSysColorBrush(COLOR_BTNSHADOW));
    }


 // Fill in the background (but only if there is no active object)...
	if (!m_pDocObjFrame)
	{
		GetSizeRectForDocument(prcBounds, &rcT);
		FillRect(hdcDraw, &rcT, hbshBackground);

     // When in design mode, we do a little extra stuff to help the 
     // developer visualize some of the control settings (toolbars,
     // foreground color, etc.)...
        if (FRunningInDesignMode())
        {
		    if (m_fShowToolbars)
		    {
			    rcT.bottom = rcT.top + 25;
			    FillRect(hdcDraw, &rcT, GetSysColorBrush(COLOR_BTNFACE));
			    clrT = SetTextColor(hdcDraw, GetSysColor(COLOR_BTNSHADOW));
			    DrawText(hdcDraw, "Toolbar Space", -1, &rcT, DT_SINGLELINE|DT_VCENTER|DT_CENTER);
			    SetTextColor(hdcDraw, clrT);

		        GetSizeRectForDocument(prcBounds, &rcT);
			    rcT.top = rcT.top + 25;
		    }

		    clrT = SetTextColor(hdcDraw, GetSysColor(COLOR_BTNSHADOW));
		    DrawText(hdcDraw, "ActiveX Document Object", -1, &rcT, DT_SINGLELINE|DT_VCENTER|DT_CENTER);
		    SetTextColor(hdcDraw, clrT);
        }
	}
	else if ((m_pDocObjFrame) && (!m_fAppActive || !m_fComponentActive) && (FDrawBitmapOnAppDeactive() || FIPDeactivateOnCompChange()))
	{
	 // In the case we have more than one component, and this component is
	 // has an object but is not ui active we draw a bitmap representation
	 // of the obect to fill the area. This gives the user a sense the object
	 // is there but in background, and can click to reactivate...
		GetSizeRectForDocument(prcBounds, &rcT);

		if (m_hbmDeactive)
		{
			HDC hdcT = CreateCompatibleDC(hdcDraw);
			HBITMAP bmpOld = (HBITMAP)SelectObject(hdcT, m_hbmDeactive);
			BITMAP bm;

			if (GetObject(m_hbmDeactive, sizeof(BITMAP), &bm))
			{
				StretchBlt(hdcDraw, rcT.left, rcT.top, 
					(rcT.right - rcT.left), (rcT.bottom - rcT.top),
					 hdcT, 0, 0, bm.bmWidth, bm.bmHeight, SRCCOPY);
			}
			else
			{
				BitBlt(hdcDraw, rcT.left, rcT.top, 
					(rcT.right - rcT.left), (rcT.bottom - rcT.top), hdcT, 0,0, SRCCOPY);
			}

			SelectObject(hdcT, bmpOld);
			DeleteDC(hdcT);
		}
		else if (!m_fInDocumentLoad)
		{
			FillRect(hdcDraw, &rcT, hbshBackground);
			rcT.top = ((rc.bottom - rc.top) / 2);
			clrT = SetTextColor(hdcDraw, clrForeground);
			DrawText(hdcDraw, "Unable to display inactive document. Click here to reactivate the object.", -1, &rcT, DT_WORDBREAK | DT_CENTER);
			SetTextColor(hdcDraw, clrT);
		}
	} // else The docobj draws the rest...

 // Cleanup the hDC as required...
	DeleteObject(hbshBackground);
	DeleteObject(hbshBorder);

	SelectObject(hdcDraw, hfontOld);
	DeleteObject(hfontText);

	if (hpal)
	{
		SelectPalette(hdcDraw, hpalOld, TRUE);
		DeleteObject(hpal);
	}

	SetBkMode(hdcDraw, iTxtMode);
    SetMapMode(hdcDraw, iMapMode);
	return;
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::OnDestroyWindow
//
//  When the window closes make sure the object closes and release the
//  subclass we have on the parent's top-level window.
//
STDMETHODIMP_(void) CDsoFramerControl::OnDestroyWindow()
{
	ODS("CDsoFramerControl::OnDestroyWindow\n");

	if (m_pDocObjFrame)
        Close();

	if (m_pHookManager)
	{
		m_pHookManager->DetachComponent(m_hwnd);
		m_pHookManager = NULL;
	}

	m_hwnd = NULL;

    if (m_pServerLock)
        SetTempServerLock(FALSE);

	if (m_hbmDeactive)
	{
		DeleteObject(m_hbmDeactive);
		m_hbmDeactive = NULL;
	}

	if (m_hmenuFilePopup)
	{
		DestroyMenu(m_hmenuFilePopup);
		m_hmenuFilePopup = NULL;
	}

}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::OnResize
//
//  Resize the window and notify the doc site.
//
STDMETHODIMP_(void) CDsoFramerControl::OnResize()
{
	RECT rcPlace;
	ODS("CDsoFramerControl::OnResize\n");
	if (m_pDocObjFrame)
	{
        if (FChangeObjActiveOnFocusChange() && (!m_fUIActive))
		    UIActivate(FALSE);

		GetSizeRectForDocument(NULL, &rcPlace);
		m_pDocObjFrame->OnNotifySizeChange(&rcPlace);
	}
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::OnMouseMove
//
//  We watch for mouse movement on menu bar (when visible).
//
STDMETHODIMP_(void) CDsoFramerControl::OnMouseMove(UINT x, UINT y)
{
 // TRACE2("CDsoFramerControl::OnMouseMove(%d, %d)\n", x, y);
 // TRACE3("  fAppActive=%d, fUIActive=%d, fCompActive=%d\n", m_fAppActive, m_fUIActive, m_fComponentActive);
 // TRACE2("  fModalState=%d, fNoInteractive=%d\n", m_fModalState, m_fNoInteractive);

 // When we are in mouse move, watch for menu selection. Only need to
 // do this if bar is visible and we are not in modal state...
    if ((m_fShowMenuBar) && (m_fComponentActive) && (m_fUIActive) && (m_fAppActive) &&
       (!m_fModalState) && (!m_fNoInteractive))
    {
        UINT item;
        POINT pt; pt.x = x; pt.y = y;

     // If we already have item selected, check to see if we are still
     // over it. If so, we can exit now. If not, we sould unselect it.
        if (m_wSelMenuItem)
        {
            item = m_wSelMenuItem - 1;
            if (PtInRect(&m_rgrcMenuItems[item], pt))
                return;

            ReleaseCapture(); m_wSelMenuItem = 0;
		    InvalidateRect(m_hwnd, &m_rgrcMenuItems[item], FALSE);
        }

     // Now loop through the menu items and see if mouse is over them.
     // If it is, we set a mouse capture and draw item as selected...
        for (item = 0; item < m_cMenuItems; item++)
        {
            if (PtInRect(&m_rgrcMenuItems[item], pt))
            {
                SetCapture(m_hwnd);
                m_wSelMenuItem = (USHORT)(item + 1);
		        InvalidateRect(m_hwnd, &m_rgrcMenuItems[item], FALSE);
                break;
            }
        }
    }

    return;
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::OnButtonDown
//
//  We watch for mouse button down events when the titlebar is visible
//  so we can drop down our popup menu if the user clicks around the 
//  area of the file icon. This is our workaround for menu merging.
//
STDMETHODIMP_(void) CDsoFramerControl::OnButtonDown(UINT x, UINT y)
{
	HMENU hCurMenu = NULL;
	RECT rc;

 // Don't do anything if we are in modal state...
    if (m_fModalState) return;

 // We don't handle commands when m_fNoInteractive is set, except for the
 // condition where we are in print preview and want to display custom menu
 // item to end the preview. Office apps have toolbar button to do this, but
 // we will provide this option just in case user needs it to exit preview mode...
    if (m_fNoInteractive)
    {
        if ((m_pDocObjFrame) && (m_pDocObjFrame->InPrintPreview()))
        {
         // If we don't have UI focus, take it...
	        if (!m_fUIActive) UIActivate(TRUE);

         // Create a temp popup menu. We do this on fly to save resources since
         // this is only needed in rare cases...
	        if ((m_fShowTitlebar) && ((x < 22) && (y < 22)))
	        {
                hCurMenu = CreatePopupMenu();
		        if (hCurMenu)
		        {
                 // Add the command to the menu...
                    AppendMenu(hCurMenu, MF_STRING,  MNU_PRINTPV, "E&xit Preview");
                 // Display the menu...
			        GetSizeRectAfterTitlebar(NULL, &rc);
			        MapWindowPoints(m_hwnd, HWND_DESKTOP, (POINT*)&rc, 2);
			        TrackPopupMenu(hCurMenu, 0, rc.left, rc.top - 1, 0, m_hwnd, NULL);
                 // Cleanup...
                    DestroyMenu(hCurMenu);
                    hCurMenu = NULL;
		        }
            }
        }

     // Otherwise we do nothing when m_fNoInteractive == TRUE.
        return;
    }

 // We know we must be app active if user clicked on control...
    if (!m_fAppActive) m_fAppActive = TRUE;

 // If we don't have UI focus, take it...
	if (!m_fUIActive) UIActivate(TRUE);

 // If the caption is showing, we are not modal, and user clicked in the 
 // area around the office doc icon, show the popup menu...
	if ((m_fShowTitlebar) && (!m_fShowMenuBar) && ((x < 22) && (y < 22)))
	{
	 // This will get the File menu or the merged menu based on current state...
		hCurMenu = GetActivePopupMenu();
		if (hCurMenu)
		{
		 // We'll place it right below the titlebar just like sys menu...
			GetSizeRectAfterTitlebar(NULL, &rc);
			MapWindowPoints(m_hwnd, HWND_DESKTOP, (POINT*)&rc, 2);
			TrackPopupMenu(hCurMenu, 0, rc.left, rc.top - 1, 0, m_hwnd, NULL);
		}
	}
    else if ((m_fShowMenuBar) && (m_wSelMenuItem))
    {
        POINT pt;
        UINT item = m_wSelMenuItem - 1;

     // If we are over a visible menu bar item, find the right popup
     // menu for the item using the selitem as an index...
        if (item == 0)
        {   hCurMenu = m_hmenuFilePopup;  }
        else
        {
		    hCurMenu = GetActivePopupMenu();
		    if (hCurMenu) hCurMenu = GetSubMenu(hCurMenu, item);
        }

     // Map the location to display the popup into screen points...
        pt.x = m_rgrcMenuItems[item].left;
        pt.y = m_rgrcMenuItems[item].bottom;
        MapWindowPoints(m_hwnd, HWND_DESKTOP, (POINT*)&pt, 1);

		HWND hwndCurFocus = GetFocus();
		m_fInFocusChange = TRUE;
		if (hwndCurFocus && (hwndCurFocus != m_hwnd))
			SetFocus(m_hwnd);

     // Display the menu...
		TrackPopupMenu(hCurMenu, 0, pt.x, pt.y, 0, m_hwnd, NULL);

		if (hwndCurFocus && (hwndCurFocus != m_hwnd))
			SetFocus(hwndCurFocus);
		m_fInFocusChange = FALSE;

     // When user clicks off, we will resolve the mouse location to
     // our control coordinates and call mousemove to free/reset selection...
        if (GetCursorPos(&pt))
        {
            RECT rcMenu; GetSizeRectForMenuBar(NULL, &rcMenu);
            MapWindowPoints(HWND_DESKTOP, m_hwnd, (POINT*)&pt, 1);
            if (!PtInRect(&rcMenu, pt)) {pt.x = 0; pt.y = 0;}
        }
        else {pt.x = 0; pt.y = 0;}
        OnMouseMove(pt.x, pt.y);
    }
	return;
}

	
////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::OnMenuMessage
//
//  We should get these messages when we are displaying the popup menu.
//  The "File" menu is ours, but the popup may have other menus copied
//  from the merged menu set by the ui active docobject, so we need to
//  forward those messages (as needed) to the right window handler.
//
//  In addition, code is added to enable/disable File menu items based
//  on flags set by developer in EnableFileCommand, and/or call the
//  correct function if one of those commands was selected by user.
//
STDMETHODIMP_(void) CDsoFramerControl::OnMenuMessage(UINT msg, WPARAM wParam, LPARAM lParam)
{
	HMENU hmenu, hmenuMerged = NULL;
	BOOL  fAlwaysSendMessage = FALSE;
	HWND  hwndObjectMenu = NULL;

	if (m_pDocObjFrame)
	{
		hwndObjectMenu = m_pDocObjFrame->GetMenuHWND();
		hmenuMerged = m_pDocObjFrame->GetMergedMenu();
	}

	switch (msg)
	{
	case WM_INITMENU:  m_fObjectMenu = FALSE;      //fall through...
	case WM_ENTERIDLE: fAlwaysSendMessage = TRUE;
		break;

	case WM_MENUSELECT:
		if ((lParam == 0) && (HIWORD(wParam) == 0xFFFF))
			fAlwaysSendMessage = TRUE;
		break;

	case WM_INITMENUPOPUP:
		hmenu = (HMENU)wParam;
		m_fObjectMenu = ((hmenu != m_hmenuFilePopup) && (hmenu != hmenuMerged));

		if ((!m_fObjectMenu) && (hmenu == m_hmenuFilePopup))
		{
			EnableMenuItem(hmenu, MNU_NEW,     MF_BYCOMMAND | ((m_wFileMenuFlags & 1) ? MF_ENABLED : (MF_DISABLED | MF_GRAYED)));
			EnableMenuItem(hmenu, MNU_OPEN,    MF_BYCOMMAND | ((m_wFileMenuFlags & 2) ? MF_ENABLED : (MF_DISABLED | MF_GRAYED)));
			EnableMenuItem(hmenu, MNU_CLOSE,   MF_BYCOMMAND | (((m_wFileMenuFlags &   4) && (m_pDocObjFrame)) ? MF_ENABLED : (MF_DISABLED | MF_GRAYED)));
			EnableMenuItem(hmenu, MNU_SAVE,    MF_BYCOMMAND | (((m_wFileMenuFlags &   8) && (m_pDocObjFrame)) ? MF_ENABLED : (MF_DISABLED | MF_GRAYED)));
			EnableMenuItem(hmenu, MNU_SAVEAS,  MF_BYCOMMAND | (((m_wFileMenuFlags &  16) && (m_pDocObjFrame)) ? MF_ENABLED : (MF_DISABLED | MF_GRAYED)));
			EnableMenuItem(hmenu, MNU_PGSETUP, MF_BYCOMMAND | (((m_wFileMenuFlags &  64) && (m_pDocObjFrame)) ? MF_ENABLED : (MF_DISABLED | MF_GRAYED)));
			EnableMenuItem(hmenu, MNU_PRINTPV, MF_BYCOMMAND | (((m_wFileMenuFlags & 256) && (m_pDocObjFrame)) ? MF_ENABLED : (MF_DISABLED | MF_GRAYED)));
			EnableMenuItem(hmenu, MNU_PRINT,   MF_BYCOMMAND | (((m_wFileMenuFlags &  32) && (m_pDocObjFrame)) ? MF_ENABLED : (MF_DISABLED | MF_GRAYED)));
			EnableMenuItem(hmenu, MNU_PROPS,   MF_BYCOMMAND | (((m_wFileMenuFlags & 128) && (m_pDocObjFrame)) ? MF_ENABLED : (MF_DISABLED | MF_GRAYED)));
		}
		break;

	case WM_COMMAND:
     // We handle one special case for print preview exit, but otherwise
     // we need to either forward to the merged menu handler (if enabled),
     // or handle command through an asynchronous callback...
        if ((m_fNoInteractive) && (LOWORD(wParam) == MNU_PRINTPV))
        {
            PrintPreviewExit();
        }
		else if ((m_fObjectMenu) && (hwndObjectMenu))
        {
			PostMessage(hwndObjectMenu, msg, wParam, lParam);
        }
		else
		{
			DWORD dwCmd = 0;
			switch ((int)LOWORD(wParam))
			{
			case MNU_NEW:     dwCmd = OLECMDID_NEW;  break;
			case MNU_OPEN:    dwCmd = OLECMDID_OPEN;  break;
			case MNU_SAVE:    dwCmd = OLECMDID_SAVE;  break;
			case MNU_SAVEAS:  dwCmd = OLECMDID_SAVEAS;  break;
			case MNU_PGSETUP: dwCmd = OLECMDID_PAGESETUP;  break;
			case MNU_PRINT:   dwCmd = OLECMDID_PRINT;  break;
			case MNU_PROPS:   dwCmd = OLECMDID_PROPERTIES;  break;
            case MNU_PRINTPV: dwCmd = OLECMDID_PRINTPREVIEW; break;
			}
			PostMessage(m_hwnd, DSO_WM_ASYNCH_OLECOMMAND, dwCmd, 0);
		}

		return; // we just return
	}

	if ((hwndObjectMenu) && ((m_fObjectMenu) || (fAlwaysSendMessage)))
		SendMessage(hwndObjectMenu, msg, wParam, lParam);

	return;
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::OnToolbarAction
//
//  Does an associated OLECMDID action and raises the OnFileCommand event
//  for developer to override as needed. The name comes from the fact that
//  most Office servers will call this (via IOleCommandTarget) when user
//  clicks on a toolbar button for file-command item. It also serves for
//  our menu commands to keep the code consistant. Both menu and toolbar
//  call this method asynchronously (via PostMessage) to avoid blocking the
//  remote docobject server.
//
STDMETHODIMP_(void) CDsoFramerControl::OnToolbarAction(DWORD cmd)
{
	ODS("CDsoFramerControl::OnToolbarAction\n");
	HRESULT hr = S_OK;
	VARIANT_BOOL wCancelAction = VARIANT_FALSE;
	VARIANT rgargs[2];
	VARIANT vtT = {0};

 // We don't reenter if in a modal state already...
    if (m_fModalState) return;

 // If not interactive, we cannot run command...
    if (m_fNoInteractive)
    {
     // unless we are in preview, in which case we can force exit 
     // and then run command after preview completes...
        if ((m_pDocObjFrame) && (m_pDocObjFrame->InPrintPreview()))
        {
            PrintPreviewExit();
			if (cmd != OLECMDID_PRINTPREVIEW)
				PostMessage(m_hwnd, DSO_WM_ASYNCH_OLECOMMAND, cmd, 0);
        }
        return;
    }

 // We will do a default action by default, but give developer a chance
 // to override in OnFileCommand event...
	rgargs[0].vt = VT_BOOL|VT_BYREF;
	rgargs[0].pboolVal = &wCancelAction;

 // Map the OLECMDID to one of our values...
	rgargs[1].vt = VT_I4;
	switch (cmd)
	{
	case OLECMDID_NEW:           rgargs[1].lVal = dsoFileNew; break;
	case OLECMDID_OPEN:          rgargs[1].lVal = dsoFileOpen; break;
	case OLECMDID_SAVE:          rgargs[1].lVal = dsoFileSave; break;
	case OLECMDID_SAVEAS:        rgargs[1].lVal = dsoFileSaveAs; break;
	case OLECMDID_PRINT:         rgargs[1].lVal = dsoFilePrint; break;
	case OLECMDID_PAGESETUP:     rgargs[1].lVal = dsoFilePageSetup; break;
	case OLECMDID_PROPERTIES:    rgargs[1].lVal = dsoFileProperties; break;
    case OLECMDID_PRINTPREVIEW:  rgargs[1].lVal = dsoFilePrintPreview; break;
	default:    rgargs[1].lVal = dsoFileClose;
	}
	
 // Only do action if item is enabled...
	if ((1 << rgargs[1].lVal) & m_wFileMenuFlags)
	{
	 // Let control developer handle the event first...
		hr = RaiseAutomationEvent(DSOF_DISPID_FILECMD,  2, rgargs);
		TRACE1("Disp event returned 0x%X \n", hr);

     // If the event was canceled (or event handler threw an
     // unhandled error from user code), bail out now...
        if ((wCancelAction) || (hr == DISP_E_EXCEPTION))
            return;

     // Do the action based on the event...
        switch (rgargs[1].lVal)
        {
        case dsoFileClose:		hr = Close();	break;
        case dsoFilePrint:		hr = ShowDialog(dsoDialogPrint); break;
        case dsoFilePageSetup:	hr = ShowDialog(dsoDialogPageSetup); break;
        case dsoFileProperties:	hr = ShowDialog(dsoDialogProperties); break;
        case dsoFilePrintPreview: hr = PrintPreview(); break;

        case dsoFileSave:
	        hr = Save(vtT, vtT, vtT, vtT);
	        if ((hr != DSO_E_DOCUMENTREADONLY) && 
				(hr != DSO_E_NOTBEENSAVED))
		        break; // fall through to SaveAs if file is read-only or never been saved...

        default:
	        hr = DoDialogAction(
		        (rgargs[1].lVal == dsoFileNew) ? dsoDialogNew : 
		        (rgargs[1].lVal == dsoFileOpen) ? dsoDialogOpen : dsoDialogSave);
        }

     // Display error information to user if we failed...
        if (FAILED(hr) && (hr != E_ABORT))
            FAlertUser(hr, NULL);
	}

	return;
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::FAlertUser
//
//  Display an error alert to the user based on failed HRESULT.
//
STDMETHODIMP_(BOOL) CDsoFramerControl::FAlertUser(HRESULT hr, LPWSTR pwsFileName)
{
    LPSTR pszError = NULL;
    LPSTR pszTitle = NULL;
    DWORD dwFlags  = MB_ICONINFORMATION|MB_SETFOREGROUND;

    switch (hr)
    {
    case STG_E_ACCESSDENIED:
    case STG_E_SHAREVIOLATION:
    case STG_E_LOCKVIOLATION:
    case MK_E_CANTOPENFILE:
    case E_ACCESSDENIED:
        pszError = "Unable to Open/Save the file selected due to security permissions or a share violation.";
        pszTitle = "Access Denied";
        break;

    case DSO_E_DOCUMENTNOTOPEN:
        pszError = "The operation could not be performed because no document is currently open.";
        pszTitle = "Unable to Perform Command";
        break;

    case DSO_E_DOCUMENTREADONLY:
        pszError = "The operation could not be performed because the document was opened read-only.";
        pszTitle = "Unable to Update Document";
        break;

    case CO_E_MSI_ERROR:
    case DSO_E_INVALIDSERVER:
    case DSO_E_INVALIDPROGID:
    case STG_E_NOTFILEBASEDSTORAGE:
        pszError = "There is no ActiveX Document server currently registered or installed for the item selected. "
                   "It cannot be loaded into the control.";
        pszTitle = "Unable to Insert Document";
        break;

    case DSO_E_INMODALSTATE:
        MessageBeep(0);
        return FALSE; // Unable to alert user if in a modal condition!

    case DSO_E_COMMANDNOTSUPPORTED:
        pszError = "This operation could not be performed because the document server "
		           "does not support the command, or it is disabled when embedded.";
        pszTitle = "Command Not Supported";
        break;

    case STG_E_TOOMANYOPENFILES:
        pszError = "You can only open one document at a time. Please choose a single file to open.";
        pszTitle = "Command Not Supported";
        break;

    case DSO_E_REQUIRESMSDAIPP:
        pszError = "You cannot open/save to a web folder unless MDAC 2.5 is installed. The operation cannot be performed.";
        pszTitle = "Unable to Gain Write Access for URL";
        break;

    default:
        pszError = "The program encountered an error trying to perform the command. "					
		           "If the problem persists, contact the program vendor for a possible solution.";
        pszTitle = "Command Error";
        dwFlags  = MB_ICONHAND|MB_SETFOREGROUND;
    }

    MessageBox(m_hwnd, pszError, pszTitle, dwFlags);
    return TRUE;
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::EnableDropFile
//
//  Decide when to allow file drop for open.
//
STDMETHODIMP_(void) CDsoFramerControl::EnableDropFile(BOOL fEnable)
{
    TRACE1("CDsoFramerControl::EnableDropFile(%d)\n", fEnable);
    DragAcceptFiles(m_hwnd, fEnable);
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::OnDropFile
//
//  Open file dropped by user on the control.
//
STDMETHODIMP_(void) CDsoFramerControl::OnDropFile(HDROP hdrpFile)
{
    DWORD cbItems;
    CHAR szFileDrop[MAX_PATH];
    VARIANT vtFile;

    ODS("CDsoFramerControl::OnDropFile()\n");
    cbItems = DragQueryFile(hdrpFile, 0xFFFFFFFF, NULL, 0);
    if (cbItems == 0)
        return;

    if (cbItems > 1)
    {
        FAlertUser(STG_E_TOOMANYOPENFILES, NULL);
        return;
    }

    szFileDrop[0] = 0;
    vtFile.vt = VT_BSTR;

    if (DragQueryFile(hdrpFile, 0, szFileDrop, MAX_PATH) &&
        (vtFile.bstrVal = DsoConvertToBSTR(szFileDrop)))
    {
        HRESULT hr;
        VARIANT vtMissing;
        vtMissing.vt = VT_ERROR;
        vtMissing.scode = DISP_E_PARAMNOTFOUND;

        hr = Open(vtFile, vtMissing, vtMissing, vtMissing, vtMissing);
        if (FAILED(hr)) FAlertUser(hr, vtFile.bstrVal);
        
        VariantClear(&vtFile);
    }

    return;
}


////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::OnTimer
//
//  Timer event fired.
//
STDMETHODIMP_(void) CDsoFramerControl::OnTimer(UINT id)
{
 // TRACE1("CDsoFramerControl::OnTimer(%d)\n", id);
    switch (id)
    {
    case SYNCPAINT_TIMER_ID:
        {
            DWORD dwTick = GetTickCount();
            if ((dwTick - m_uiSyncPaint) > 300)
            {
                ODS("CDsoFramerControl::OnSyncPaint() Invalidate\n");
                KillTimer(m_hwnd, SYNCPAINT_TIMER_ID);
                m_fSyncPaintTimer = FALSE;
				InvalidateAllChildWindows(m_hwnd);
            }
        }
        break;
    }
    return;
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::OnSyncPaint
//
//  Notifies redraw is needed b/c of external window movement.
//
STDMETHODIMP_(void) CDsoFramerControl::OnSyncPaint()
{
 // ODS("CDsoFramerControl::OnSyncPaint()\n");
    m_uiSyncPaint = GetTickCount();
    if ((!m_fSyncPaintTimer) && (!m_fAppActive))
    {
        SetTimer(m_hwnd, SYNCPAINT_TIMER_ID, 50, NULL);
        m_fSyncPaintTimer = TRUE;
    }
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::OnForegroundCompChange
//
//  Shows/Hide embedding when window goes foreground/background active. 
//
STDMETHODIMP_(void) CDsoFramerControl::OnForegroundCompChange(BOOL fCompActive)
{
	ODS("CDsoFramerControl::OnForegroundCompChange\n");
    if (m_pDocObjFrame)
    {
        HWND hwndEmbed = m_pDocObjFrame->GetDocWindow();

        if (fCompActive)
        {
         // Make sure embed window is visible when active...
            if (!IsWindowVisible(hwndEmbed))
			    ShowWindow(hwndEmbed, SW_SHOW);

            if (m_hbmDeactive) DeleteObject(m_hbmDeactive);
            m_hbmDeactive = NULL;
        }
        else if ((FDrawBitmapOnAppDeactive() || FIPDeactivateOnCompChange()) && (m_hbmDeactive == NULL))
        {
         // The default behavior on app deactive is to grap bitmap of active object before
         // telling it we lost app activation. This allows us to draw the last state of the 
         // object while it is deactive. There is now an option to avoid this and keep object
         // UI active while app deactive, but it may not work for all DocObj servers.
		    m_hbmDeactive = DsoGetBitmapFromWindow(hwndEmbed);
		    ShowWindow(hwndEmbed, SW_HIDE);
        }
    }
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::OnAppActivationChange
//
//  Notifes embedded object when the host app gains/loses foreground status
//  as the active application. This notification is critical to DocObjects since
//  it is the primary means for the host and server to synch which top-level 
//  frame window (the host or the server's native main window) should be on the
//  z-order stack for the shared thread input state. Thus allowing dialogs and tool
//  windows to appear the correct order when displayed.
//
STDMETHODIMP_(void) CDsoFramerControl::OnAppActivationChange(BOOL fAppActive, DWORD dwThreadID)
{
	TRACE2("CDsoFramerControl::OnAppActivationChange(fAppActive=%d, tid=0x%X)\n", fAppActive, dwThreadID);
    m_fAppActive = fAppActive;

    if ((m_pDocObjFrame) && (m_fComponentActive))
    {
     // Create/Free bitmap for background appearance..
        OnForegroundCompChange(fAppActive);

     // Notify it that it should be app active ofr deactive...
		m_pDocObjFrame->OnNotifyAppActivate(fAppActive, (dwThreadID ? dwThreadID : GetCurrentThreadId()));
    }
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::OnComponentActivationChange
//
//  Handles showing/hiding doc object during activation state changes.
//  This is needed to properly handle focus changes when more than one
//  instance of the control is running on the same top-level window.
//
STDMETHODIMP_(void) CDsoFramerControl::OnComponentActivationChange(BOOL fActivate)
{
	TRACE1("CDsoFramerControl::OnComponentActivationChange(fActive=%d)\n", fActivate);

	if (fActivate)
	{
	 // Set us active if we currently are not and raise activation change event...
		if (!m_fComponentActive)
		{
			m_fComponentActive = TRUE;
            OnForegroundCompChange(TRUE);
		 // Raise event to component host on state change...
			PostMessage(m_hwnd, DSO_WM_ASYNCH_STATECHANGE, DSO_STATE_ACTIVATION, (LPARAM)fActivate);
		}

     // Notify the embedded object...
        if (m_pDocObjFrame)
		{
         // Determine if we should restore from IP deactivation... 
			if (FIPDeactivateOnCompChange() && !(m_pDocObjFrame->IsIPActive()))
            {
               m_pDocObjFrame->IPActivateView();
 		       OnForegroundCompChange(TRUE);
            }
		}
	}
	else
	{

     // Inform embedded object we lost activation...
		if (m_pDocObjFrame)
		{
         // If we are in print preview, we need to exit in case other control will
         // end up talking to same OLE server...
		    if (m_pDocObjFrame->InPrintPreview())
            {
                PrintPreviewExit();
                UpdateWindow(m_hwnd);
            }

         // Provide control host option to IP deactive view when another framer control 
         // takes primary focus. This helps certain resource problems if you a lot of controls
         // all IP active at the same time.  This should not normally be used.
			if (FIPDeactivateOnCompChange() && m_pDocObjFrame->IsIPActive())
            {
		        OnForegroundCompChange(FALSE);
                m_pDocObjFrame->IPDeactivateView();
            }
			
		}

	 // Change the current activation state and raise event...
		if (m_fComponentActive)
		{
			m_fComponentActive = FALSE;
            OnForegroundCompChange(FALSE);
		 // Raise event to component host on state change...
			PostMessage(m_hwnd, DSO_WM_ASYNCH_STATECHANGE, DSO_STATE_ACTIVATION, (LPARAM)fActivate);
		}

	}
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::OnCtrlFocusChange
//
//  Checks if window (or one of its children) is getting or losing focus so
//  that can notify both the host and embed server using OnUIFocusChange.
//
STDMETHODIMP_(void) CDsoFramerControl::OnCtrlFocusChange(BOOL fCtlGotFocus, HWND hFocusWnd)
{
	TRACE2("CDsoFramerControl::OnCtrlFocusChange(fCtlGotFocus=%d, hwndCurFocus=%x)\n", fCtlGotFocus, hFocusWnd);
    if (!(m_fInFocusChange) && ((fCtlGotFocus) || !IsWindowChild(m_hwnd, hFocusWnd)))
    {
        OnUIFocusChange(fCtlGotFocus);
    }
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::OnUIFocusChange
//
//  Selectively handle focus changes to make sure both host and ui 
//  active docobject are in synch.
//
STDMETHODIMP_(void) CDsoFramerControl::OnUIFocusChange(BOOL fUIActive)
{
    BOOL fFocusUpdated = FALSE;
	TRACE1("CDsoFramerControl::OnUIFocusChange(fUIActive=%d)\n", fUIActive);

 // We shouldn't try to do focus change when modal...
    if ((m_fModalState) || (m_fInFocusChange))
    {
        ODS(" -- Got focus while modal or in change (do nothing) --\n");
        return;
    }

  // When IP active, notify the control host of focus change...
	if ((m_fInPlaceActive) && (m_pControlSite))
	{

     // Set flag to prevent recursion during focus changes...
		m_fInFocusChange = TRUE;

		if (fUIActive) // If we got focus...
		{
		 // Notify host we have focus (we only need
		 // to do this if we are not already UI Active)...
			if (!(m_fHasFocus))
			{
				ODS(" -- Tell host app we should have focus --\n");
                m_fHasFocus = (fFocusUpdated = TRUE);
				m_pControlSite->OnFocus(TRUE); 
			}

		 // Let the hook manager know that this control is the active object...
			if (m_pHookManager)
				m_pHookManager->SetActiveComponent(m_hwnd);

         // If so configured, force component to show up active...
		    if (FChangeObjActiveOnFocusChange())
		    {
                ODS(" -- Do component active on focus gain --\n");
			    OnComponentActivationChange(TRUE);
		    }

		 // If we have an active document, forward the focus...
			if (m_pDocObjFrame)
				m_pDocObjFrame->OnNotifyControlFocus(TRUE);

		}
		else // else we lost focus...
		{

	     // When we lose focus, only notify host if we lost to window
	     // that does not belong to us...
			if (m_pDocObjFrame)
				m_pDocObjFrame->OnNotifyControlFocus(FALSE);

         // If so configured, force component loss on focus loss...
		    if (FChangeObjActiveOnFocusChange())
		    {
                ODS(" -- Do component deactive on focus loss --\n");
			    OnComponentActivationChange(FALSE);
			    m_fActivateOnStatus = TRUE;
		    }

         // If we are UI active still, we need to tell host that we lost 
         // focus and are now UI deactive.
            if (m_fHasFocus)
            {
			    ODS(" -- Notify host of lost focus (ie, UI deactivate us) --\n");
                m_fHasFocus = FALSE; fFocusUpdated = TRUE;
			    m_pControlSite->OnFocus(FALSE);
            }
		}

     // Clear focus change flag...
		m_fInFocusChange = FALSE;

     // Redraw menu bar in case of state change
        if ((fFocusUpdated) && (m_fShowMenuBar))
        {
            RECT rc; GetSizeRectForMenuBar(NULL, &rc);
            InvalidateRect(m_hwnd, &rc, TRUE);
        }
	}

	return;
}
////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::UpdateModalState
//
//  Called either by developer (ModalState property) or by the embedded
//  DocObject to notify host (VB/IE) that we are modal.
//
STDMETHODIMP_(void) CDsoFramerControl::UpdateModalState(BOOL fModeless, BOOL fNotifyIPObject)
{
	TRACE2("CDsoFramerControl::UpdateModalState(AllowModeless=%d, NotifyIP=%d)\n", fModeless, fNotifyIPObject);
	if (fModeless == (int)m_fModalState)
	{
		IOleInPlaceActiveObject* pipao;

		m_fModalState = !(fModeless);
		ODS("Modal state changed\n");

	  // Excel doesn't like us to notify the host of changes in modality
	  // if it is the one who initialied the call. So, we check the
	  // NotifyIPObj flag and only notify host when the IPObj is not the caller...
		if ((fNotifyIPObject) && (m_pInPlaceFrame))
			m_pInPlaceFrame->EnableModeless(fModeless);

	  // Again, if IPObj is not the caller and we have Ipobj, let it know 
	  // of the change in modal state...
		if ((fNotifyIPObject) && (m_pDocObjFrame) && 
			(pipao = m_pDocObjFrame->GetActiveObject()))
			pipao->EnableModeless(fModeless);

	 // Post notification of return from modal...
		if (fModeless)
			PostMessage(m_hwnd, DSO_WM_ASYNCH_STATECHANGE, DSO_STATE_RETURNFROMMODAL, (LPARAM)TRUE);
	}
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::UpdateInteractiveState
//
//  Called when interactivity changes (mostly due to print preview).
//
STDMETHODIMP_(void) CDsoFramerControl::UpdateInteractiveState(BOOL fActive)
{
    if (fActive)
    {
        m_fNoInteractive = FALSE;
        if (m_fShowMenuPrev) put_Menubar(VARIANT_TRUE);
        if (!m_fShowToolsPrev) put_Toolbars(VARIANT_FALSE);

	 // Notify control host that preview mode has ended...
    	RaiseAutomationEvent(DSOF_DISPID_ENDPREVIEW, 0, NULL);
    }
    else
    {
        m_fShowMenuPrev = m_fShowMenuBar;
		m_fShowToolsPrev = m_fShowToolbars;
        if (m_fShowMenuBar) put_Menubar(VARIANT_FALSE);
        if (!m_fShowToolbars) put_Toolbars(VARIANT_TRUE);
        m_fNoInteractive = TRUE;
    }
    return;
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::OnPaletteChanged
//
//  This method is also called from the subclassed parent. Per docobject
//  spec, a ui active object should get first chance of any palette updates.
//
STDMETHODIMP_(void) CDsoFramerControl::OnPaletteChanged(HWND hwndPalChg)
{
	ODS("CDsoFramerControl::OnPaletteChanged\n");
	if (m_pDocObjFrame)	m_pDocObjFrame->OnNotifyPaletteChanged(hwndPalChg);
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::OnSysCommandMenu
//
//  Checks for menu shortcut keys when ui active.
//
STDMETHODIMP_(BOOL) CDsoFramerControl::OnSysCommandMenu(CHAR ch)
{
    if (m_fUIActive)
    {
        TRACE1("CDsoFramerControl::OnSysCommandMenu(%d)\n", (DWORD)ch);
        return TRUE;
    }
    return FALSE;
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::GetActivePopupMenu
//
//  Constructs our popup menu and returns the handle. The menu can be
//  either the File menu by itself (no docobj is open) or the merged
//  menu as a popup.
//
STDMETHODIMP_(HMENU) CDsoFramerControl::GetActivePopupMenu()
{
	HMENU hPopup, hMergedMenu, hServerMenu;

 // If we haven't made a File menu yet, make it. Here is where you might
 // want to add custom items if you have other file-level commands you would
 // the control to support. These are the basics...
	if (!m_hmenuFilePopup)
	{
		m_hmenuFilePopup = CreatePopupMenu(); m_cMenuItems = 0;
		if (!m_hmenuFilePopup) return NULL;

		AppendMenu(m_hmenuFilePopup, MF_STRING,    MNU_NEW,     "&New...\tCtrl+N");
		AppendMenu(m_hmenuFilePopup, MF_STRING,    MNU_OPEN,    "&Open...\tCtrl+O");
		AppendMenu(m_hmenuFilePopup, MF_STRING,    MNU_CLOSE,   "&Close");
		AppendMenu(m_hmenuFilePopup, MF_SEPARATOR, 0,            NULL);
		AppendMenu(m_hmenuFilePopup, MF_STRING,    MNU_SAVE,    "&Save\tCtrl+S");
		AppendMenu(m_hmenuFilePopup, MF_STRING,    MNU_SAVEAS,  "Save &As...");
		AppendMenu(m_hmenuFilePopup, MF_SEPARATOR, 0,            NULL);
		AppendMenu(m_hmenuFilePopup, MF_STRING,    MNU_PGSETUP, "Page Set&up...");
		AppendMenu(m_hmenuFilePopup, MF_STRING,    MNU_PRINTPV, "Print Pre&view");
		AppendMenu(m_hmenuFilePopup, MF_STRING,    MNU_PRINT,   "&Print...");
		AppendMenu(m_hmenuFilePopup, MF_SEPARATOR, 0,            NULL);
		AppendMenu(m_hmenuFilePopup, MF_STRING,    MNU_PROPS,   "Propert&ies");
	}

 // If we have a docobj and it has a merged menu, then lets return a 
 // merged menu between the docobj and our file menu...
	if ((m_pDocObjFrame) && (m_pDocObjFrame->GetMenuHWND()))
	{
	 // We only need to create the merged copy once.
		if (!(hMergedMenu = m_pDocObjFrame->GetMergedMenu()))
		{
			hMergedMenu = CreatePopupMenu(); m_cMenuItems = 0;
			if ((hMergedMenu) && (hServerMenu = (m_pDocObjFrame->GetActiveMenu())))
			{
				CHAR szbuf[MAX_PATH];
				HMENU hT;
                CHAR *pch;

				InsertMenu(hMergedMenu, 0, MF_BYPOSITION|MF_POPUP, (UINT)m_hmenuFilePopup, "&File");
				m_rgchMenuAccel[0] = 'f';

				int cbMenuCnt = GetMenuItemCount(hServerMenu);
				if (cbMenuCnt > DSO_MAX_MENUITEMS) cbMenuCnt = DSO_MAX_MENUITEMS;

				for (int i = 0; i < cbMenuCnt; i++)
				{
					hT = GetSubMenu(hServerMenu, i);
					if (hT)
					{
						szbuf[0] = '\0';
						GetMenuString(hServerMenu, i, szbuf, MAX_PATH, MF_BYPOSITION);
						InsertMenu(hMergedMenu, (i + 1), MF_BYPOSITION|MF_POPUP, (UINT)hT, szbuf);

                        pch = szbuf;
                        while (*pch && (*pch++ != '&'))
                            ;
                        m_rgchMenuAccel[i + 1] = (CHAR)((*pch) ? ASCII_LOWERCASE(*pch) : 0xFF);
					}
				}

				if (cbMenuCnt < DSO_MAX_MENUITEMS)
					m_rgchMenuAccel[cbMenuCnt + 1] = 0x00;

				m_pDocObjFrame->SetMergedMenu(hMergedMenu);                
			}
		}

		hPopup = hMergedMenu;
	}
	else
    {
        hPopup = m_hmenuFilePopup;
        m_rgchMenuAccel[0] = 'f';
        m_rgchMenuAccel[1] = 0;
    }

	return hPopup;
}


////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::FRunningInDesignMode
//
//  Determines if we are in design mode. Called when painting the control
//  and when deciding whether to subclass the parent.
//
STDMETHODIMP_(BOOL) CDsoFramerControl::FRunningInDesignMode()
{
	IDispatch *pdisp;

 // We must have a control site.
	CHECK_NULL_RETURN(m_pClientSite, FALSE);

 // If we have done this before, we don't need to keep doing it unless
 // the host has notified us of the state change (see our code in 
 // XOleControl::OnAmbientPropertyChange)...
	if ((!m_fModeFlagValid) && 
		SUCCEEDED(m_pClientSite->QueryInterface(IID_IDispatch, (void **)&pdisp)))
	{
		VARIANT vtUserMode;
		m_fDesignMode = FALSE; // assume run mode

		if (SUCCEEDED(DsoDispatchInvoke(pdisp, NULL, 
				DISPID_AMBIENT_USERMODE, DISPATCH_PROPERTYGET, 0, NULL, &vtUserMode)))
		{
		 // UserMode is True when control is in Run mode, False when in design.
		 // We assume run mode, so we only care to set the flag if in design.
			m_fDesignMode = !(BOOL_FROM_VARIANT(vtUserMode, TRUE));
			VariantClear(&vtUserMode);
		}

		m_fModeFlagValid = TRUE;
     // Release the IDispatch pointer...
        pdisp->Release();
	}

	return m_fDesignMode;
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::InvalidateAllChildWindows
//
//  Invalidate all child windows attached to the window passed.
//
STDMETHODIMP_(BOOL) CDsoFramerControl::InvalidateAllChildWindows(HWND hwnd)
{
	ODS("CDsoFramerControl::InvalidateAllChildWindows()\n");
	return EnumChildWindows(hwnd, InvalidateAllChildWindowsCallback, 0);
}

STDMETHODIMP_(BOOL) CDsoFramerControl::InvalidateAllChildWindowsCallback(HWND hwnd, LPARAM lParam)
{
	RECT rc; 
	GetClientRect(hwnd, &rc);
	InvalidateRect(hwnd, &rc, TRUE);
	return TRUE;
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::RaiseAutomationEvent
//
//  Raises automation events to host application.
//
STDMETHODIMP CDsoFramerControl::RaiseAutomationEvent(DISPID did, ULONG cargs, VARIANT *pvtargs)
{
    HRESULT hr = ((m_fModalState) ? DISP_E_EXCEPTION : S_FALSE);
    TRACE1("CDsoFramerControl::RaiseAutomationEvent(%d)\n", did);

 // We should only raise event if we have sink, not already in a event, and 
 // we are not in modal condition with respect to the remote server...
    if ((m_dispEvents) && (!m_fFreezeEvents) && (!m_fModalState))
    {
        m_fFreezeEvents = TRUE;
        hr = DsoDispatchInvoke(m_dispEvents, NULL, did, 0, cargs, pvtargs, NULL);
        m_fFreezeEvents = FALSE;
    }

    return hr;
}

////////////////////////////////////////////////////////////////////////
//
// CDsoFramerControl::DoDialogAction
//
//  A very simple implementation of a New/Open/Save dialogs to serve as
//  default actions in case developer does not override these "File"
//  commands in OnFileCommand.
//
STDMETHODIMP CDsoFramerControl::DoDialogAction(dsoShowDialogType item)
{
	HRESULT    hr = S_FALSE;
    BSTR       bstr;
    BOOL       fReadOnly;
	VARIANT    vT[3];

    static const WCHAR v_wszFileFilter[] =
        L"Microsoft Office Files\0*.doc;*.docx;*.docm;*.rtf;*.xls;*.xlsx;*.xlsm;*.csv;*.ppt;*.pptx;*.pptm;*.mpp;*.vsd;*.vdx\0"
		L"Microsoft Word Files\0*.doc;*.docx;*.docm;*.rtf\0"
		L"Microsoft Excel Files\0*.xls;*.xlsx;*.xlsm;*.csv\0"
		L"Microsoft PowerPoint Files\0*.ppt;*.pptx;*.pptm\0"
		L"Microsoft Project Files\0*.mpp\0"
		L"Microsoft Visio Files\0*.vsd;*.vdx\0"
		L"All Files (*.*)\0*.*\0\0";

 // Switch on the action type... 
    switch (item)
    {
    case dsoDialogOpen: // Get file from user and call Open...
        if (SUCCEEDED(DsoGetFileFromUser(m_hwnd, NULL, (OFN_FILEMUSTEXIST|OFN_EXPLORER),
            v_wszFileFilter, 1, NULL, NULL, FALSE, &bstr, &fReadOnly)))
        {
			vT[0].vt = VT_BSTR;  vT[0].bstrVal = bstr;
            vT[1].vt = VT_BOOL;	 vT[1].boolVal = (fReadOnly ? VARIANT_TRUE : VARIANT_FALSE);
            vT[2].vt = VT_ERROR; vT[2].scode = DISP_E_PARAMNOTFOUND;

			hr = Open(vT[0], vT[1], vT[2], vT[2], vT[2]);
			SysFreeString(bstr);
        }
        break;

    case dsoDialogSave: // Get file from user and call Save...
        {
            LPWSTR pwszSaveType = NULL;
            LPWSTR pwszDefExt = NULL;

            if (m_pDocObjFrame)
                m_pDocObjFrame->GetDocumentTypeAndFileExtension(&pwszSaveType, &pwszDefExt);

            if (SUCCEEDED(DsoGetFileFromUser(m_hwnd, NULL, (OFN_HIDEREADONLY|OFN_OVERWRITEPROMPT|OFN_EXPLORER),
                ((pwszSaveType) ? pwszSaveType : v_wszFileFilter), 1, pwszDefExt, NULL, TRUE, &bstr, NULL)))
            {
			    vT[0].vt = VT_BSTR;  vT[0].bstrVal = bstr;
			    vT[1].vt = VT_BOOL;	 vT[1].boolVal = VARIANT_TRUE;
                vT[2].vt = VT_ERROR; vT[2].scode = DISP_E_PARAMNOTFOUND;

			    hr = Save(vT[0], vT[1], vT[2], vT[2]);
			    SysFreeString(bstr);
            }

            if (pwszSaveType) DsoMemFree(pwszSaveType);
            if (pwszDefExt) DsoMemFree(pwszDefExt);
        }
        break;

    case dsoDialogNew: // Get either ProgID or File to create new object from...
        if (SUCCEEDED(DsoGetOleInsertObjectFromUser(m_hwnd, L"Insert New Document Object",
            (IOF_SELECTCREATENEW | IOF_DISABLELINK | IOF_DISABLEDISPLAYASICON | IOF_HIDECHANGEICON),
            TRUE, FALSE, &bstr, NULL)))
        {
			hr = CreateNew(bstr);
			SysFreeString(bstr);
        }
        break;

    default: MessageBeep(0);
    }

	return hr;
}
////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::SetTempServerLock
//
//  Sets up lock of OLE server to make unload and new load of files faster
//  and to avoid race conditions on reload of file.
//
STDMETHODIMP CDsoFramerControl::SetTempServerLock(BOOL fLock)
{
    HRESULT hr = S_FALSE;
    TRACE1("CDsoFramerControl::SetTempServerLock(%d)\n", fLock);

    if (fLock)
    {
        CLSID *pclsid;

     // First get the the CLSID of the current loaded object...
        pclsid = m_pDocObjFrame->GetServerCLSID();

     // If we already have a lock, determine if user is asking us to lock
     // the same server (we only lock once for a given server type)...
		if (m_pServerLock)
		{
            if (*(m_pServerLock->GetServerCLSID()) == *pclsid)
                return S_FALSE;

         // If the current document is from a different server, free the old lock...
            hr = SetTempServerLock(FALSE);
			ASSERT(SUCCEEDED(hr)); // Sanity check in debug...
		}

     // To lock the server, just make a dummy OLE object and don't ever site it...
        m_pServerLock = CDsoDocObject::CreateInstance((IDsoDocObjectSite*)&m_xDsoDocObjectSite);
        hr = ((m_pServerLock) ? m_pServerLock->CreateDocObject(*pclsid) : E_OUTOFMEMORY);
        if (SUCCEEDED(hr))
        {
			hr = m_pServerLock->SetRunningServerLock(TRUE);

          // If we lock Excel, it steals focus from this host to the lock object, so we
          // need to take it back so the real object isn't mixed up state...
            if (SUCCEEDED(hr) && (m_pServerLock->IsExcelObject()) && (m_fAppActive))
            {
                OnAppActivationChange(TRUE, 0);
            }
		}
		else
		{
            delete m_pServerLock;
            m_pServerLock = NULL;
        }

	}
	else if (m_pServerLock)
	{
        SEH_TRY
			hr = m_pServerLock->SetRunningServerLock(FALSE);
            hr = m_pServerLock->Close();
        SEH_EXCEPT(hr)

		ASSERT(SUCCEEDED(hr)); // Sanity check in debug...
        delete m_pServerLock;
        m_pServerLock = NULL;
	}

    return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::ResetFrameHook
//
//  Allows host to reset the frame hook. 
//
STDMETHODIMP CDsoFramerControl::ResetFrameHook(HWND hwndFrameWindow)
{
	DWORD dwProcessId;
	TRACE1("CDsoFramerControl::ResetFrameHook(%d)\n", hwndFrameWindow);

 // If the control is in modal state, we can't do anything...
    if ((m_fModalState) || (m_fNoInteractive))
        return ProvideErrorInfo(DSO_E_INMODALSTATE);

 // Verify that hwnd passed is a valid window handle and belogs to same thread
 // as the control for parenting purposes...
	if ((hwndFrameWindow) && (!IsWindow(hwndFrameWindow) || 
		 (GetWindowThreadProcessId(hwndFrameWindow, &dwProcessId), dwProcessId != GetCurrentProcessId())))
			return E_INVALIDARG;

 // We must have a window already created for this to work...
	if ((m_hwnd == NULL) || !IsWindow(m_hwnd) ||
		(m_hwndParent == NULL) || !IsWindow(m_hwndParent))
		return E_ACCESSDENIED;

 // Detach from the current window (removing that hook if this is the last component)...
	if (m_pHookManager)
	{
		HRESULT hr = m_pHookManager->DetachComponent(m_hwnd);
		if (FAILED(hr)) return hr;
		m_pHookManager = NULL;
	}

 // Now attach to the new parent window (or reattach based on the parent from our window)...
	m_pHookManager = CDsoFrameHookManager::RegisterFramerControl((hwndFrameWindow ? hwndFrameWindow : m_hwndParent), m_hwnd);
	if (m_pHookManager == NULL) return E_FAIL;

	return S_OK;
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::RaiseActivationEvent
//
//  Called to raise component change event on activation switch.
//
STDMETHODIMP_(void) CDsoFramerControl::RaiseActivationEvent(BOOL fActive)
{
    VARIANT rgargs[1]; rgargs[0].vt = VT_BOOL;
    rgargs[0].boolVal = (fActive ? VARIANT_TRUE : VARIANT_FALSE);
    RaiseAutomationEvent(DSOF_DISPID_ACTIVATE, 1, rgargs);
    return;
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::ControlWindowProc
//
//  The window proc for our control.
//
STDMETHODIMP_(LRESULT) CDsoFramerControl::ControlWindowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	CDsoFramerControl* pCtl = (CDsoFramerControl*)GetWindowLong(hwnd, GWL_USERDATA);
	if (pCtl)
	{
		switch (msg)
		{
			case WM_PAINT:
			{
				// call the OnDraw routine.
				PAINTSTRUCT ps;
				RECT        rc;
				HDC         hdc;
				hdc = BeginPaint(hwnd, &ps);
				GetClientRect(hwnd, &rc);
				pCtl->OnDraw(DVASPECT_CONTENT, hdc, (RECT*)&rc, NULL, NULL, TRUE);
				EndPaint(hwnd, &ps);
			}
			break;

			case WM_SIZE:
				pCtl->OnResize();
				break;

			case WM_DESTROY:
				pCtl->OnDestroyWindow();
				SetWindowLong(hwnd, GWL_USERDATA, 0);
				break;

			case WM_SETFOCUS:
			case WM_KILLFOCUS:
                pCtl->OnCtrlFocusChange((BOOL)(msg == WM_SETFOCUS), (HWND)wParam);
				break;

			case WM_MOUSEMOVE:
				pCtl->OnMouseMove(LOWORD(lParam), HIWORD(lParam));
				break;

			case WM_LBUTTONDOWN:
				pCtl->OnButtonDown(LOWORD(lParam), HIWORD(lParam));
				break;

			case DSO_WM_ASYNCH_OLECOMMAND:
                ODS(" -- Got DSO_WM_ASYNCH_OLECOMMAND -- \n");
				pCtl->OnToolbarAction((DWORD)wParam);
				break;

			case DSO_WM_ASYNCH_STATECHANGE:
                TRACE1(" -- Got DSO_WM_ASYNCH_STATECHANGE %d --\n", wParam);
                {
                    BOOL fCondition = (lParam != 0);
                    switch (wParam)
                    {
                    case DSO_STATE_MODAL:
                        pCtl->UpdateModalState(fCondition, FALSE);
                        break;
                    case DSO_STATE_ACTIVATION:
                        pCtl->RaiseActivationEvent(fCondition);
                        break;
                    case DSO_STATE_INTERACTIVE:
                        pCtl->UpdateInteractiveState(fCondition);
                        break;
					case DSO_STATE_RETURNFROMMODAL:
						{
							if ((pCtl->m_fModalState == FALSE) && 
                                (pCtl->m_fAppActive))
							{
                                pCtl->OnUIFocusChange(TRUE);
							}
						}
						break;

					default:
						ODS(" -- Unknown State Change!! --\n");
						break;
                    }
                }
                break;

			case DSO_WM_HOOK_NOTIFY_COMPACTIVE:
				pCtl->OnComponentActivationChange((BOOL)wParam);
                wParam = lParam; lParam = 0;
                // falls through with wParam swapped...

			case DSO_WM_HOOK_NOTIFY_APPACTIVATE:
				pCtl->OnAppActivationChange((BOOL)wParam, (DWORD)lParam);
				break;

			case DSO_WM_HOOK_NOTIFY_SYNCPAINT:
				pCtl->OnSyncPaint();
				break;

			case DSO_WM_HOOK_NOTIFY_PALETTECHANGE:
				pCtl->OnPaletteChanged((HWND)wParam);
				break;

            case WM_ENABLE:
                pCtl->OnWindowEnable((BOOL)wParam);
                break;

			case WM_MENUSELECT:
			case WM_DRAWITEM:
			case WM_MEASUREITEM:
			case WM_ENTERIDLE:
			case WM_INITMENU:
			case WM_INITMENUPOPUP:
			case WM_COMMAND:
				pCtl->OnMenuMessage(msg, wParam, lParam);
				break;

            case WM_DROPFILES:
                pCtl->OnDropFile((HDROP)wParam);
                break;

            case WM_TIMER:
                pCtl->OnTimer((UINT)wParam);
                break;

            case WM_SYSCOMMAND:
                if ((wParam & 0xFFF0) == SC_KEYMENU)
                {
                    if (pCtl->m_xDsoDocObjectSite.SysMenuCommand((UINT)lParam) == S_OK)
                    {
                        ODS(" -- Eat SC_KeyMenu message --\n");
                        return 1;
                    }
                }
                // fall through to defwinproc...

            default:
                return DefWindowProc(hwnd, msg, wParam, lParam);

		}

        return (LRESULT)0;
	}

	return DefWindowProc(hwnd, msg, wParam, lParam);
}

////////////////////////////////////////////////////////////////////////
//
// OCX Interfaces for CDsoFramerControl
//
//  NOTE: The automation interfaces are implemented in dsofauto.cpp.
//  The interfaces that follow are largely the interfaces used by the
//  OCX host for normal control containment.
//

////////////////////////////////////////////////////////////////////////
//
// CDsoFramerControl::XInternalUnknown -- Internal IUnknown interface
//     !!! Controls lifetime of CDsoFramerControl !!!
//
//   STDMETHODIMP         QueryInterface(REFIID riid, void ** ppv);
//   STDMETHODIMP_(ULONG) AddRef(void);
//   STDMETHODIMP_(ULONG) Release(void);
//
STDMETHODIMP CDsoFramerControl::XInternalUnknown::QueryInterface(REFIID riid, void** ppv)
{
 // ODS("CDsoFramerControl::InternalQueryInterface\n");
	CHECK_NULL_RETURN(ppv, E_POINTER);
	
	HRESULT hr = S_OK;
	METHOD_PROLOGUE(CDsoFramerControl, InternalUnknown);

	if (IID_IUnknown == riid)
	{
		*ppv = (IUnknown*)this;
	}
	else if ((IID_IDispatch == riid) || (IID__FramerControl == riid))
	{
		*ppv = (_FramerControl*)pThis;
	}
	else if (IID_IOleObject == riid)
	{
		*ppv = (IOleObject*)&(pThis->m_xOleObject);
	}
	else if (IID_IOleControl == riid)
	{
		*ppv = (IOleControl*)&(pThis->m_xOleControl);
	}
	else if (IID_IPersistPropertyBag == riid)
	{
		*ppv = (IPersistPropertyBag*)&(pThis->m_xPersistPropertyBag);
	}
	else if ((IID_IPersistStreamInit == riid) || (IID_IPersistStream == riid) || (IID_IPersist == riid))
	{
		*ppv = (IPersistStreamInit*)&(pThis->m_xPersistStreamInit);
	}
	else if ((IID_IOleInPlaceObject == riid) || (IID_IOleWindow == riid))
	{
		*ppv = (IOleInPlaceObject*)&(pThis->m_xOleInplaceObject);
	}	
	else if (IID_IOleInPlaceActiveObject == riid)
	{
		*ppv = (IOleInPlaceActiveObject*)&(pThis->m_xOleInplaceActiveObject);
	}
	else if ((IID_IViewObjectEx == riid) || (IID_IViewObject == riid) || (IID_IViewObject2 == riid))
	{
		*ppv = (IViewObjectEx*)&(pThis->m_xViewObjectEx);
	}
    else if (IID_IDataObject == riid)
	{
		*ppv = (IDataObject*)&(pThis->m_xDataObject);
	}
	else if (IID_ISupportErrorInfo == riid)
	{
		*ppv = (ISupportErrorInfo*)&(pThis->m_xSupportErrorInfo);
	}
	else if (IID_IProvideClassInfo == riid)
	{
		*ppv = (IProvideClassInfo*)&(pThis->m_xProvideClassInfo);
	}
	else if (IID_IConnectionPointContainer == riid)
	{
		*ppv = (IConnectionPointContainer*)&(pThis->m_xConnectionPointContainer);
	}
	else if (IID_IConnectionPoint == riid)
	{
		*ppv = (IConnectionPoint*)&(pThis->m_xConnectionPoint);
	}
	else if (IID_IEnumConnectionPoints == riid)
	{
		*ppv = (IEnumConnectionPoints*)&(pThis->m_xEnumConnectionPoints);
	}
    else if (IID_IPersistStorage == riid)
	{
		*ppv = (IPersistStorage*)&(pThis->m_xPersistStorage);
	}
	else if (IID_IObjectSafety == riid)
	{
		*ppv = (IObjectSafety*)&(pThis->m_xObjectSafety);
	}
    else if (IID_IOleCommandTarget == riid)
	{
		*ppv = (IOleCommandTarget*)&(pThis->m_xOleCommandTarget);
	}
	else if ((IID_IDsoDocObjectSite == riid) || (IID_IServiceProvider == riid))
	{
		*ppv = (IDsoDocObjectSite*)&(pThis->m_xDsoDocObjectSite);
	}
    else 
	{
		*ppv = NULL;
		hr = E_NOINTERFACE;
	}

	if (NULL != *ppv) // AddRef those interfaces we will return...
        ((IUnknown*)(*ppv))->AddRef();

	return hr;
}

STDMETHODIMP_(ULONG) CDsoFramerControl::XInternalUnknown::AddRef(void)
{
	METHOD_PROLOGUE(CDsoFramerControl, InternalUnknown);
	//TRACE1("CDsoFramerControl::InternalAddRef - %d\n", (pThis->m_cRef + 1));
    return ++(pThis->m_cRef);
}

STDMETHODIMP_(ULONG) CDsoFramerControl::XInternalUnknown::Release(void)
{
	METHOD_PROLOGUE(CDsoFramerControl, InternalUnknown);
	//TRACE1("CDsoFramerControl::InternalRelease - %d\n", (pThis->m_cRef - 1));
    if (0 != --(pThis->m_cRef)) return (pThis->m_cRef);
	ODS("CDsoFramerControl delete\n");
	InterlockedDecrement((LPLONG)&v_cLocks);
	delete pThis; return 0;
}


////////////////////////////////////////////////////////////////////////
//
// CDsoFramerControl::XPersistStreamInit -- IPersistStreamInit Implementation
//
//	 STDMETHODIMP GetClassID(CLSID *pClassID);
//	 STDMETHODIMP IsDirty(void);
//	 STDMETHODIMP Load(LPSTREAM pStm);
//	 STDMETHODIMP Save(LPSTREAM pStm, BOOL fClearDirty);
//	 STDMETHODIMP GetSizeMax(ULARGE_INTEGER* pcbSize);
//	 STDMETHODIMP InitNew(void);
//
IMPLEMENT_INTERFACE_UNKNOWN(CDsoFramerControl, PersistStreamInit)

STDMETHODIMP CDsoFramerControl::XPersistStreamInit::GetClassID(CLSID *pClassID)
{
	ODS("CDsoFramerControl::XPersistStreamInit::GetClassID\n");
	if (pClassID) *pClassID = CLSID_FramerControl;
	return S_OK;
}

STDMETHODIMP CDsoFramerControl::XPersistStreamInit::IsDirty(void)
{
	METHOD_PROLOGUE(CDsoFramerControl, PersistStreamInit);
    ODS("CDsoFramerControl::XPersistStreamInit::IsDirty\n");
    return (pThis->m_fDirty) ? S_OK : S_FALSE;
}


#define STREAMHDR_SIGNATURE_OLD 0x1234ABCD  // Signature to identify our format (avoid crashes!)
#define STREAMHDR_SIGNATURE_NEW 0xD501D501  // Signature to identify our format (avoid crashes!)

STDMETHODIMP CDsoFramerControl::XPersistStreamInit::Load(LPSTREAM pStm)
{
    HRESULT hr;
	DWORD dwSig, dwT;

	METHOD_PROLOGUE(CDsoFramerControl, PersistStreamInit);
    ODS("CDsoFramerControl::XPersistStreamInit::Load\n");

 // look for our header structure, so we can verify stream validity.
    hr = pStm->Read(&dwSig, sizeof(DWORD), NULL);
    RETURN_ON_FAILURE(hr);

    if ((dwSig != STREAMHDR_SIGNATURE_NEW) && 
		(dwSig != STREAMHDR_SIGNATURE_OLD))
        return E_UNEXPECTED;

 // we like the stream.  let's go load in our two properties.
    hr = pStm->Read(&(pThis->m_Size), sizeof(SIZEL), NULL);
    RETURN_ON_FAILURE(hr);

	DsoHimetricToPixels(&(pThis->m_Size.cx), &(pThis->m_Size.cy));

    hr = pStm->Read(&(pThis->m_clrBorderColor), sizeof(OLE_COLOR), NULL);
    RETURN_ON_FAILURE(hr);

    hr = pStm->Read(&(pThis->m_clrBackColor), sizeof(OLE_COLOR), NULL);
    RETURN_ON_FAILURE(hr);

    hr = pStm->Read(&(pThis->m_clrForeColor), sizeof(OLE_COLOR), NULL);
    RETURN_ON_FAILURE(hr);

    hr = pStm->Read(&(pThis->m_clrTBarColor), sizeof(OLE_COLOR), NULL);
    RETURN_ON_FAILURE(hr);

    hr = pStm->Read(&(pThis->m_clrTBarTextColor), sizeof(OLE_COLOR), NULL);
    RETURN_ON_FAILURE(hr);

    hr = pStm->Read(&dwT, sizeof(DWORD), NULL);
    RETURN_ON_FAILURE(hr);

	pThis->m_fBorderStyle  = LOBYTE(LOWORD(dwT));
    pThis->m_fShowMenuBar  = HIBYTE(LOWORD(dwT));
	pThis->m_fShowToolbars = LOBYTE(HIWORD(dwT));
	pThis->m_fShowTitlebar = HIBYTE(HIWORD(dwT));

    hr = pStm->Read(&dwT, sizeof(DWORD), NULL);
    RETURN_ON_FAILURE(hr);

	if (dwT)
	{
		LPWSTR pwsz = (LPWSTR)DsoMemAlloc(dwT+2);
		if (pwsz)
		{
			hr = pStm->Read(pwsz, dwT, NULL);
			if (SUCCEEDED(hr))
				pThis->m_bstrCustomCaption = SysAllocString(pwsz);
		}
		DsoMemFree(pwsz);
	}

 // Read in new properties for 1.3 version control...
	if (dwSig == STREAMHDR_SIGNATURE_NEW)
	{
		DWORD dwValue;

		hr = pStm->Read(&dwValue, sizeof(DWORD), NULL);
		RETURN_ON_FAILURE(hr);

		pThis->m_lActivationPolicy  = LOWORD(dwValue);
		pThis->m_lHookPolicy		= HIWORD(dwValue);

		hr = pStm->Read(&dwValue, sizeof(DWORD), NULL);
		RETURN_ON_FAILURE(hr);

		pThis->m_fDisableMenuAccel = LOBYTE(LOWORD(dwT));
	}

    return S_OK;
}

STDMETHODIMP CDsoFramerControl::XPersistStreamInit::Save(LPSTREAM pStm, BOOL fClearDirty)
{
    HRESULT hr;
	DWORD   dwT, dwSig = STREAMHDR_SIGNATURE_NEW;
	SIZEL   slSize;

	METHOD_PROLOGUE(CDsoFramerControl, PersistStreamInit);
    ODS("CDsoFramerControl::XPersistStreamInit::Save\n");

 // first thing to do is write out our stream sig...
    hr = pStm->Write(&dwSig, sizeof(DWORD), NULL);
    RETURN_ON_FAILURE(hr);

 // the only properties we're currently persisting here are the size
 // properties for this control.  make sure we do that in HiMetric
	slSize = pThis->m_Size;
	DsoPixelsToHimetric(&(slSize.cx), &(slSize.cy));

    hr = pStm->Write(&slSize, sizeof(SIZEL), NULL);
    RETURN_ON_FAILURE(hr);

    hr = pStm->Write(&(pThis->m_clrBorderColor), sizeof(OLE_COLOR), NULL);
    RETURN_ON_FAILURE(hr);

    hr = pStm->Write(&(pThis->m_clrBackColor), sizeof(OLE_COLOR), NULL);
    RETURN_ON_FAILURE(hr);

    hr = pStm->Write(&(pThis->m_clrForeColor), sizeof(OLE_COLOR), NULL);
    RETURN_ON_FAILURE(hr);

    hr = pStm->Write(&(pThis->m_clrTBarColor), sizeof(OLE_COLOR), NULL);
    RETURN_ON_FAILURE(hr);

    hr = pStm->Write(&(pThis->m_clrTBarTextColor), sizeof(OLE_COLOR), NULL);
    RETURN_ON_FAILURE(hr);

	dwT = (DWORD)MAKELONG(
        MAKEWORD((BYTE)(pThis->m_fBorderStyle), (BYTE)(pThis->m_fShowMenuBar)),
		MAKEWORD((BYTE)(pThis->m_fShowToolbars), (BYTE)(pThis->m_fShowTitlebar)));

    hr = pStm->Write(&dwT, sizeof(DWORD), NULL);
    RETURN_ON_FAILURE(hr);

	dwT = 0;
	if (pThis->m_bstrCustomCaption)
		dwT = SysStringByteLen(pThis->m_bstrCustomCaption);

	hr = pStm->Write(&dwT, sizeof(DWORD), NULL);
	RETURN_ON_FAILURE(hr);

	if (dwT)
	{
		hr = pStm->Write(pThis->m_bstrCustomCaption, dwT, NULL);
		RETURN_ON_FAILURE(hr);
	}

 // Version 1.3 added properties...
	dwT = (DWORD)MAKELONG((WORD)(pThis->m_lActivationPolicy), (WORD)(pThis->m_lHookPolicy));
    hr = pStm->Write(&dwT, sizeof(DWORD), NULL);
    RETURN_ON_FAILURE(hr);

	dwT = (DWORD)MAKELONG(MAKEWORD((BYTE)(pThis->m_fDisableMenuAccel), 0), 0);
    hr = pStm->Write(&dwT, sizeof(DWORD), NULL);
    RETURN_ON_FAILURE(hr);

 // clear out dirty flag and notify that we're done with save.
    if (fClearDirty)
        pThis->m_fDirty = FALSE;

    if (pThis->m_pOleAdviseHolder)
        pThis->m_pOleAdviseHolder->SendOnSave();

    return S_OK;
}

STDMETHODIMP CDsoFramerControl::XPersistStreamInit::GetSizeMax(ULARGE_INTEGER *pcbSize)
{
    ODS("CDsoFramerControl::XPersistStreamInit::GetSizeMax\n");
    return E_NOTIMPL;
}

STDMETHODIMP CDsoFramerControl::XPersistStreamInit::InitNew(void)
{
    METHOD_PROLOGUE(CDsoFramerControl, PersistStreamInit);
    ODS("CDsoFramerControl::XPersistStreamInit::InitNew\n");
 // If we are new control, we are dirty by default...
	pThis->m_fDirty	= TRUE;  
    return S_OK;
}

////////////////////////////////////////////////////////////////////////
//
// CDsoFramerControl::XPersistPropertyBag -- IPersistPropertyBag Implementation
//
//	 STDMETHODIMP GetClassID(CLSID *pClassID);
//	 STDMETHODIMP InitNew();
//	 STDMETHODIMP Load(IPropertyBag* pPropBag, IErrorLog* pErrorLog);
//	 STDMETHODIMP Save(IPropertyBag* pPropBag, BOOL fClearDirty, BOOL fSaveAllProperties);
//
IMPLEMENT_INTERFACE_UNKNOWN(CDsoFramerControl, PersistPropertyBag)

STDMETHODIMP CDsoFramerControl::XPersistPropertyBag::GetClassID(CLSID *pClassID)
{
    ODS("CDsoFramerControl::XPersistPropertyBag::GetClassID\n");
	if (pClassID) *pClassID = CLSID_FramerControl;
	return S_OK;
}

STDMETHODIMP CDsoFramerControl::XPersistPropertyBag::InitNew()
{
	METHOD_PROLOGUE(CDsoFramerControl, PersistPropertyBag);
    ODS("CDsoFramerControl::XPersistPropertyBag::InitNew\n");
	return pThis->m_xPersistStreamInit.InitNew();
}

STDMETHODIMP CDsoFramerControl::XPersistPropertyBag::Load(IPropertyBag* pPropBag, IErrorLog* pErrorLog)
{
    VARIANT v;
    HRESULT hr;
	SIZEL sl = {5000, 3000};

	METHOD_PROLOGUE(CDsoFramerControl, PersistPropertyBag);
    ODS("CDsoFramerControl::XPersistPropertyBag::Load\n");

    v.vt = VT_I4;
	v.lVal = 0;

    hr = pPropBag->Read(L"_ExtentX", &v, pErrorLog);
    if (SUCCEEDED(hr)){sl.cx = v.lVal;}
    
    v.lVal = 0;
    hr = pPropBag->Read(L"_ExtentY", &v, pErrorLog);
    if (SUCCEEDED(hr)){sl.cy = v.lVal;}

	DsoHimetricToPixels(&(sl.cx), &(sl.cy));
	pThis->m_Size = sl;

    hr = pPropBag->Read(L"Titlebar", &v, pErrorLog);
	if (SUCCEEDED(hr)){pThis->m_fShowTitlebar = v.boolVal;}

    hr = pPropBag->Read(L"Toolbars", &v, pErrorLog);
	if (SUCCEEDED(hr)){pThis->m_fShowToolbars = v.boolVal;}

    hr = pPropBag->Read(L"BorderStyle", &v, pErrorLog);
	if (SUCCEEDED(hr)){pThis->m_fBorderStyle = v.lVal;}

    hr = pPropBag->Read(L"BorderColor", &v, pErrorLog);
	if (SUCCEEDED(hr)){pThis->m_clrBorderColor = v.lVal;}

    hr = pPropBag->Read(L"BackColor", &v, pErrorLog);
	if (SUCCEEDED(hr)){pThis->m_clrBackColor = v.lVal;}

    hr = pPropBag->Read(L"ForeColor", &v, pErrorLog);
	if (SUCCEEDED(hr)){pThis->m_clrForeColor = v.lVal;}

    hr = pPropBag->Read(L"TitlebarColor", &v, pErrorLog);
	if (SUCCEEDED(hr)){pThis->m_clrTBarColor = v.lVal;}

    hr = pPropBag->Read(L"TitlebarTextColor", &v, pErrorLog);
	if (SUCCEEDED(hr)){pThis->m_clrTBarTextColor = v.lVal;}

    hr = pPropBag->Read(L"Caption", &v, pErrorLog);
	if (SUCCEEDED(hr))
	{
		LPWSTR pwsz = LPWSTR_FROM_VARIANT(v);
		if (pwsz) pThis->m_bstrCustomCaption = SysAllocString(pwsz);
		VariantClear(&v);
	}

    pThis->m_fShowMenuBar = FALSE;
    hr = pPropBag->Read(L"Menubar", &v, pErrorLog);
	if (SUCCEEDED(hr)){pThis->m_fShowMenuBar = BOOL_FROM_VARIANT(v, FALSE);}

    hr = pPropBag->Read(L"ActivationPolicy", &v, pErrorLog);
	if (SUCCEEDED(hr)){pThis->m_lActivationPolicy = LONG_FROM_VARIANT(v, 0);}

    hr = pPropBag->Read(L"FrameHookPolicy", &v, pErrorLog);
	if (SUCCEEDED(hr)){pThis->m_lHookPolicy = LONG_FROM_VARIANT(v, 0);}

    hr = pPropBag->Read(L"MenuAccelerators", &v, pErrorLog);
	if (SUCCEEDED(hr)){pThis->m_fDisableMenuAccel = !BOOL_FROM_VARIANT(v, TRUE);}

    return S_OK;
}

STDMETHODIMP CDsoFramerControl::XPersistPropertyBag::Save(IPropertyBag* pPropBag, BOOL fClearDirty, BOOL fSaveAllProperties)
{
    HRESULT hr;
    VARIANT v;
    SIZEL   sl;

	METHOD_PROLOGUE(CDsoFramerControl, PersistPropertyBag);
    ODS("CDsoFramerControl::XPersistPropertyBag::Save\n");

	sl = pThis->m_Size;
    DsoPixelsToHimetric(&(sl.cx), &(sl.cy));

    v.vt = VT_I4; v.lVal = sl.cx;
    hr = pPropBag->Write(L"_ExtentX", &v);
    RETURN_ON_FAILURE(hr);

    v.lVal = sl.cy;
    hr = pPropBag->Write(L"_ExtentY", &v);
    RETURN_ON_FAILURE(hr);

    v.lVal = pThis->m_clrBorderColor;
    hr = pPropBag->Write(L"BorderColor", &v);
    RETURN_ON_FAILURE(hr);

    v.lVal = pThis->m_clrBackColor;
    hr = pPropBag->Write(L"BackColor", &v);
    RETURN_ON_FAILURE(hr);

    v.lVal = pThis->m_clrForeColor;
    hr = pPropBag->Write(L"ForeColor", &v);
    RETURN_ON_FAILURE(hr);

    v.lVal = pThis->m_clrTBarColor;
    hr = pPropBag->Write(L"TitlebarColor", &v);
    RETURN_ON_FAILURE(hr);

    v.lVal = pThis->m_clrTBarTextColor;
    hr = pPropBag->Write(L"TitlebarTextColor", &v);
    RETURN_ON_FAILURE(hr);

    v.lVal = pThis->m_fBorderStyle;
    hr = pPropBag->Write(L"BorderStyle", &v);
    RETURN_ON_FAILURE(hr);

    v.vt = VT_BOOL;	v.boolVal = pThis->m_fShowTitlebar;
    hr = pPropBag->Write(L"Titlebar", &v);
    RETURN_ON_FAILURE(hr);

	v.boolVal = pThis->m_fShowToolbars;
    hr = pPropBag->Write(L"Toolbars", &v);
    RETURN_ON_FAILURE(hr);

	if (pThis->m_bstrCustomCaption)
	{
		v.vt = VT_BSTR;	v.bstrVal = pThis->m_bstrCustomCaption;
		hr = pPropBag->Write(L"Caption", &v);
		RETURN_ON_FAILURE(hr);
	}

    v.vt = VT_BOOL; v.boolVal = pThis->m_fShowMenuBar;
    hr = pPropBag->Write(L"Menubar", &v);
    RETURN_ON_FAILURE(hr);

    v.vt = VT_I4; v.lVal = pThis->m_lActivationPolicy;
    hr = pPropBag->Write(L"ActivationPolicy", &v);
    RETURN_ON_FAILURE(hr);

    v.vt = VT_I4; v.lVal = pThis->m_lHookPolicy;
    hr = pPropBag->Write(L"FrameHookPolicy", &v);
    RETURN_ON_FAILURE(hr);

    v.vt = VT_BOOL;	v.boolVal = !(pThis->m_fDisableMenuAccel);
    hr = pPropBag->Write(L"MenuAccelerators", &v);
    RETURN_ON_FAILURE(hr);

 // now clear the dirty flag and send out notification 
 // that we're done.  
    if (fClearDirty)
        pThis->m_fDirty = FALSE;

    if (pThis->m_pOleAdviseHolder)
        pThis->m_pOleAdviseHolder->SendOnSave();

    return S_OK;
}

////////////////////////////////////////////////////////////////////////
//
// CDsoFramerControl::XPersistStorage -- IPersistStorage Implementation
//
//	 STDMETHODIMP GetClassID(CLSID *pClassID);
//	 STDMETHODIMP IsDirty(void);
//	 STDMETHODIMP InitNew(LPSTORAGE pStg);
//	 STDMETHODIMP Load(LPSTORAGE pStg);
//	 STDMETHODIMP Save(LPSTORAGE pStg, BOOL fSameAsLoad);
//	 STDMETHODIMP SaveCompleted(LPSTORAGE pStg);
//	 STDMETHODIMP HandsOffStorage(void);
//	 
IMPLEMENT_INTERFACE_UNKNOWN(CDsoFramerControl, PersistStorage)

STDMETHODIMP CDsoFramerControl::XPersistStorage::GetClassID(CLSID *pClassID)
{
    ODS("CDsoFramerControl::XPersistStorage::GetClassID\n");
	if (pClassID) *pClassID = CLSID_FramerControl;
	return S_OK;
}

STDMETHODIMP CDsoFramerControl::XPersistStorage::IsDirty(void)
{
	METHOD_PROLOGUE(CDsoFramerControl, PersistStorage);
    ODS("CDsoFramerControl::XPersistStorage::IsDirty\n");
    return (pThis->m_fDirty) ? S_OK : S_FALSE;
}

STDMETHODIMP CDsoFramerControl::XPersistStorage::InitNew(LPSTORAGE pStg)
{
    HRESULT hr;
	METHOD_PROLOGUE(CDsoFramerControl, PersistStorage);
    ODS("CDsoFramerControl::XPersistStorage::InitNew\n");

    CHECK_NULL_RETURN(pStg, E_POINTER);
    SAFE_RELEASE_INTERFACE(pThis->m_pOleStorage);

    hr = pThis->m_xPersistStreamInit.InitNew();
    if (SUCCEEDED(hr))
        { SAFE_SET_INTERFACE(pThis->m_pOleStorage, pStg); };

	return hr;
}

STDMETHODIMP CDsoFramerControl::XPersistStorage::Load(LPSTORAGE pStg)
{
    HRESULT hr;
    IStream* pstm;
    DWORD dwStmMode = (STGM_DIRECT | STGM_READWRITE | STGM_SHARE_EXCLUSIVE);

	METHOD_PROLOGUE(CDsoFramerControl, PersistStorage);
    ODS("CDsoFramerControl::XPersistStorage::Load\n");

    CHECK_NULL_RETURN(pStg, E_POINTER);
    SAFE_RELEASE_INTERFACE(pThis->m_pOleStorage);

    hr = pStg->OpenStream(L"DsoFrameCtlContents", 0, dwStmMode, 0, &pstm);
    if (SUCCEEDED(hr) && (pstm))
    {
        hr = pThis->m_xPersistStreamInit.Load(pstm);
        if (SUCCEEDED(hr))
            { SAFE_SET_INTERFACE(pThis->m_pOleStorage, pStg); };

        pstm->Release();
    }

	return hr;
}

STDMETHODIMP CDsoFramerControl::XPersistStorage::Save(LPSTORAGE pStg, BOOL fSameAsLoad)
{
    HRESULT hr;
    IStream* pstm;
    DWORD dwStmMode = (STGM_DIRECT | STGM_READWRITE | STGM_SHARE_EXCLUSIVE);

	METHOD_PROLOGUE(CDsoFramerControl, PersistStorage);
    ODS("CDsoFramerControl::XPersistStorage::Save\n");

	if (fSameAsLoad)
	{
		if ((pThis->m_pOleStorage == NULL) || 
			((pStg) && (pStg != pThis->m_pOleStorage)))
		{
			SAFE_RELEASE_INTERFACE(pThis->m_pOleStorage);
			SAFE_SET_INTERFACE(pThis->m_pOleStorage, pStg);
		}

		CHECK_NULL_RETURN(pThis->m_pOleStorage, E_POINTER);
		pStg = pThis->m_pOleStorage;
	}

    CHECK_NULL_RETURN(pStg, E_POINTER);

    hr = pStg->OpenStream(L"DsoFrameCtlContents", 0, dwStmMode, 0, &pstm);
    if (hr == STG_E_FILENOTFOUND)
        hr = pStg->CreateStream(L"DsoFrameCtlContents", dwStmMode, 0, 0, &pstm);

    if (SUCCEEDED(hr) && (pstm))
    {
        hr = pThis->m_xPersistStreamInit.Save(pstm, FALSE);
        pstm->Release();
    }

	return hr;
}

STDMETHODIMP CDsoFramerControl::XPersistStorage::SaveCompleted(LPSTORAGE pStg)
{
	METHOD_PROLOGUE(CDsoFramerControl, PersistStorage);
    ODS("CDsoFramerControl::XPersistStorage::SaveCompleted\n");
    pThis->m_fDirty = FALSE;
    return S_OK;
}

STDMETHODIMP CDsoFramerControl::XPersistStorage::HandsOffStorage(void)
{
	METHOD_PROLOGUE(CDsoFramerControl, PersistStorage);
    ODS("CDsoFramerControl::XPersistStorage::HandsOffStorage\n");
    SAFE_RELEASE_INTERFACE(pThis->m_pOleStorage);
    return S_OK;
}

////////////////////////////////////////////////////////////////////////
//
// CDsoFramerControl::XOleObject -- IOleObject Implementation
//
//	 STDMETHODIMP SetClientSite(IOleClientSite *pClientSite);
//	 STDMETHODIMP GetClientSite(IOleClientSite **ppClientSite);
//	 STDMETHODIMP SetHostNames(LPCOLESTR szContainerApp, LPCOLESTR szContainerObj);
//	 STDMETHODIMP Close(DWORD dwSaveOption);
//	 STDMETHODIMP SetMoniker(DWORD dwWhichMoniker, IMoniker *pmk);
//	 STDMETHODIMP GetMoniker(DWORD dwAssign, DWORD dwWhichMoniker, IMoniker **ppmk);
//	 STDMETHODIMP InitFromData(IDataObject *pDataObject, BOOL fCreation, DWORD dwReserved);
//	 STDMETHODIMP GetClipboardData(DWORD dwReserved, IDataObject **ppDataObject);
//	 STDMETHODIMP DoVerb(LONG iVerb, LPMSG lpmsg, IOleClientSite *pActiveSite, LONG lindex, HWND hwndParent, LPCRECT lprcPosRect);
//	 STDMETHODIMP EnumVerbs(IEnumOleVerb **ppEnumOleVerb);
//	 STDMETHODIMP Update();
//	 STDMETHODIMP IsUpToDate();
//	 STDMETHODIMP GetUserClassID(CLSID *pClsid);
//	 STDMETHODIMP GetUserType(DWORD dwFormOfType, LPOLESTR *pszUserType);
//	 STDMETHODIMP SetExtent(DWORD dwDrawAspect, SIZEL *psizel);
//	 STDMETHODIMP GetExtent(DWORD dwDrawAspect, SIZEL *psizel);
//	 STDMETHODIMP Advise(IAdviseSink *pAdvSink, DWORD *pdwConnection);
//	 STDMETHODIMP Unadvise(DWORD dwConnection);
//	 STDMETHODIMP EnumAdvise(IEnumSTATDATA **ppenumAdvise);
//	 STDMETHODIMP GetMiscStatus(DWORD dwAspect, DWORD *pdwStatus);
//	 STDMETHODIMP SetColorScheme(LOGPALETTE *pLogpal);
//
IMPLEMENT_INTERFACE_UNKNOWN(CDsoFramerControl, OleObject)

STDMETHODIMP CDsoFramerControl::XOleObject::SetClientSite(IOleClientSite *pClientSite)
{

	METHOD_PROLOGUE(CDsoFramerControl, OleObject);
    ODS("CDsoFramerControl::XOleObject::SetClientSite\n");

    SAFE_RELEASE_INTERFACE(pThis->m_pClientSite);
    SAFE_RELEASE_INTERFACE(pThis->m_pControlSite);

 // store away the new client site
    pThis->m_pClientSite = pClientSite;

  // if we've actually got one, then get some other interfaces we want to keep
  // around, and keep a handle on it
    if (pClientSite)
	{
        pClientSite->AddRef();
        pClientSite->QueryInterface(IID_IOleControlSite, (void **)&(pThis->m_pControlSite));
    }
    return S_OK;
}

STDMETHODIMP CDsoFramerControl::XOleObject::GetClientSite(IOleClientSite **ppClientSite)
{
	METHOD_PROLOGUE(CDsoFramerControl, OleObject);
    ODS("CDsoFramerControl::XOleObject::GetClientSite\n");
    if (ppClientSite) { SAFE_SET_INTERFACE(*ppClientSite, pThis->m_pClientSite); }
	return S_OK;
}

STDMETHODIMP CDsoFramerControl::XOleObject::SetHostNames(LPCOLESTR /*szContainerApp*/, LPCOLESTR /*szContainerObj*/)
{
    ODS("CDsoFramerControl::XOleObject::SetHostNames\n");
	return S_OK;
}

STDMETHODIMP CDsoFramerControl::XOleObject::Close(DWORD dwSaveOption)
{
	HRESULT hr;

	METHOD_PROLOGUE(CDsoFramerControl, OleObject);
    ODS("CDsoFramerControl::XOleObject::Close\n");

    if (pThis->m_fInPlaceActive)
	{
        hr = pThis->m_xOleInplaceObject.InPlaceDeactivate();
        RETURN_ON_FAILURE(hr);
    }

 // handle the save flag.
    if ((pThis->m_fDirty) && (pThis->m_pClientSite) &&
		(dwSaveOption == OLECLOSE_SAVEIFDIRTY || 
		 dwSaveOption == OLECLOSE_PROMPTSAVE))
	{
		pThis->m_pClientSite->SaveObject();
    }

    if (pThis->m_pOleAdviseHolder)
        pThis->m_pOleAdviseHolder->SendOnClose();

    return S_OK;
}

STDMETHODIMP CDsoFramerControl::XOleObject::SetMoniker(DWORD dwWhichMoniker, IMoniker *pmk)
{
    ODS("CDsoFramerControl::XOleObject::SetMoniker\n");
    return E_NOTIMPL;
}

STDMETHODIMP CDsoFramerControl::XOleObject::GetMoniker(DWORD dwAssign, DWORD dwWhichMoniker, IMoniker **ppmk)
{
    ODS("CDsoFramerControl::XOleObject::GetMoniker\n");
	if (ppmk) *ppmk = NULL;
    return E_NOTIMPL;
}

STDMETHODIMP CDsoFramerControl::XOleObject::InitFromData(IDataObject *pDataObject, BOOL fCreation, DWORD dwReserved)
{
    ODS("CDsoFramerControl::XOleObject::InitFromData\n");
    return E_NOTIMPL;
}

STDMETHODIMP CDsoFramerControl::XOleObject::GetClipboardData(DWORD /*dwReserved*/, IDataObject **ppDataObject)
{
    ODS("CDsoFramerControl::XOleObject::GetClipboardData\n");
	if (ppDataObject) *ppDataObject = NULL;
    return E_NOTIMPL;
}

STDMETHODIMP CDsoFramerControl::XOleObject::DoVerb(LONG iVerb, LPMSG lpmsg, IOleClientSite *pActiveSite, LONG lindex, HWND hwndParent, LPCRECT lprcPosRect)
{
	METHOD_PROLOGUE(CDsoFramerControl, OleObject);
    ODS("CDsoFramerControl::XOleObject::DoVerb\n");

    switch (iVerb)
	{
    case OLEIVERB_PRIMARY:
    case OLEIVERB_SHOW:
    case OLEIVERB_INPLACEACTIVATE:
    case OLEIVERB_UIACTIVATE:
        return pThis->InPlaceActivate(iVerb);

    case OLEIVERB_HIDE:
        pThis->m_xOleInplaceObject.UIDeactivate();
        if (pThis->m_fInPlaceVisible)
            pThis->SetInPlaceVisible(FALSE);
        return S_OK;
    }

    return OLEOBJ_S_INVALIDVERB;
}

STDMETHODIMP CDsoFramerControl::XOleObject::EnumVerbs(IEnumOLEVERB **ppEnumOleVerb)
{
    ODS("CDsoFramerControl::XOleObject::EnumVerbs\n");
	return OleRegEnumVerbs(CLSID_FramerControl, ppEnumOleVerb);
}

STDMETHODIMP CDsoFramerControl::XOleObject::Update()
{
    ODS("CDsoFramerControl::XOleObject::Update\n");
    return S_OK;
}

STDMETHODIMP CDsoFramerControl::XOleObject::IsUpToDate()
{
    ODS("CDsoFramerControl::XOleObject::IsUpToDate\n");
    return S_OK;
}

STDMETHODIMP CDsoFramerControl::XOleObject::GetUserClassID(CLSID *pClsid)
{
    ODS("CDsoFramerControl::XOleObject::GetUserClassID\n");
	if (pClsid) *pClsid = CLSID_FramerControl;
    return S_OK;
}

STDMETHODIMP CDsoFramerControl::XOleObject::GetUserType(DWORD dwFormOfType, LPOLESTR *pszUserType)
{
    ODS("CDsoFramerControl::XOleObject::GetUserType\n");
    return OleRegGetUserType(CLSID_FramerControl, dwFormOfType, pszUserType);
}

STDMETHODIMP CDsoFramerControl::XOleObject::SetExtent(DWORD dwDrawAspect, SIZEL *psizel)
{
    SIZEL sl;

	METHOD_PROLOGUE(CDsoFramerControl, OleObject);
    ODS("CDsoFramerControl::XOleObject::SetExtent\n");

	CHECK_NULL_RETURN(psizel, E_POINTER);
	sl = *psizel;

 // we don't support any other aspects.
    if (!(dwDrawAspect & DVASPECT_CONTENT))
        return DV_E_DVASPECT;

 // change the units to pixels, and resize the control.
    DsoHimetricToPixels(&(sl.cx), &(sl.cy));

 // if the size already matches the extents, then don't do anything
 // and don't mark the control dirty.
	if (sl.cx != pThis->m_Size.cx || sl.cy != pThis->m_Size.cy)
	{
		pThis->m_Size = sl;
		pThis->m_fDirty = TRUE; // changing size should dirty

		if (pThis->m_fInPlaceActive)
		{
			RECT rect;  

		 // Use OnPosRectChange for older OCX containters and Access
		 // who just refuses to get with the times (like me <g>)...
			GetWindowRect(pThis->m_hwnd, &rect);
			MapWindowPoints(NULL, pThis->m_hwndParent, (LPPOINT)&rect, 2);
			rect.right = rect.left + sl.cx;
			rect.bottom = rect.top + sl.cy;
			pThis->m_pInPlaceSite->OnPosRectChange(&rect);

		 // Resize the window based on the new size...
			SetWindowPos(pThis->m_hwnd, NULL, 0, 0, sl.cx, sl.cy,
					SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
		}
		else pThis->ViewChanged();
	}
    
    return S_OK;
}

STDMETHODIMP CDsoFramerControl::XOleObject::GetExtent(DWORD dwDrawAspect, SIZEL *psizel)
{
	METHOD_PROLOGUE(CDsoFramerControl, OleObject);
    ODS("CDsoFramerControl::XOleObject::GetExtent\n");

	CHECK_NULL_RETURN(psizel, E_POINTER);

    if (!(dwDrawAspect & DVASPECT_CONTENT))
        return DV_E_DVASPECT;

	*psizel = pThis->m_Size;
	DsoPixelsToHimetric(&(psizel->cx), &(psizel->cy));
    return S_OK;
}

STDMETHODIMP CDsoFramerControl::XOleObject::Advise(IAdviseSink *pAdvSink, DWORD *pdwConnection)
{
	METHOD_PROLOGUE(CDsoFramerControl, OleObject);
    ODS("CDsoFramerControl::XOleObject::Advise\n");

	if (NULL == pThis->m_pOleAdviseHolder)
	{
		if (FAILED(CreateOleAdviseHolder(&(pThis->m_pOleAdviseHolder))))
			return E_FAIL;
	}
	return pThis->m_pOleAdviseHolder->Advise(pAdvSink, pdwConnection);
}

STDMETHODIMP CDsoFramerControl::XOleObject::Unadvise(DWORD dwConnection)
{
	METHOD_PROLOGUE(CDsoFramerControl, OleObject);
    ODS("CDsoFramerControl::XOleObject::Unadvise\n");

	CHECK_NULL_RETURN(pThis->m_pOleAdviseHolder, CONNECT_E_NOCONNECTION);
    return pThis->m_pOleAdviseHolder->Unadvise(dwConnection);
}

STDMETHODIMP CDsoFramerControl::XOleObject::EnumAdvise(IEnumSTATDATA **ppenumAdvise)
{
	METHOD_PROLOGUE(CDsoFramerControl, OleObject);
    ODS("CDsoFramerControl::XOleObject::EnumAdvise\n");

	CHECK_NULL_RETURN(ppenumAdvise, E_POINTER);
	*ppenumAdvise = NULL;

	CHECK_NULL_RETURN(pThis->m_pOleAdviseHolder, E_FAIL);
    return pThis->m_pOleAdviseHolder->EnumAdvise(ppenumAdvise);
}

STDMETHODIMP CDsoFramerControl::XOleObject::GetMiscStatus(DWORD dwAspect, DWORD *pdwStatus)
{
    ODS("CDsoFramerControl::XOleObject::GetMiscStatus\n");

    if (dwAspect != DVASPECT_CONTENT)
        return DV_E_DVASPECT;
    
    if (pdwStatus)
		*pdwStatus = 	OLEMISC_SETCLIENTSITEFIRST |
						OLEMISC_ACTIVATEWHENVISIBLE | 
						OLEMISC_RECOMPOSEONRESIZE |
						OLEMISC_CANTLINKINSIDE | 
						OLEMISC_INSIDEOUT;
    return S_OK;
}

STDMETHODIMP CDsoFramerControl::XOleObject::SetColorScheme(LOGPALETTE *pLogpal)
{
    ODS("CDsoFramerControl::XOleObject::SetColorScheme\n");
    return S_OK;
}

////////////////////////////////////////////////////////////////////////
//
// CDsoFramerControl::XOleControl -- IOleControl Implementation
//
//	 STDMETHODIMP GetControlInfo(CONTROLINFO* pCI);
//	 STDMETHODIMP OnMnemonic(LPMSG pMsg);
//	 STDMETHODIMP OnAmbientPropertyChange(DISPID dispID);
//	 STDMETHODIMP FreezeEvents(BOOL bFreeze);
//
IMPLEMENT_INTERFACE_UNKNOWN(CDsoFramerControl, OleControl)

STDMETHODIMP CDsoFramerControl::XOleControl::GetControlInfo(CONTROLINFO* pCI)
{
	ODS("CDsoFramerControl::XOleControl::GetControlInfo\n");
	CHECK_NULL_RETURN(pCI, E_POINTER);
    pCI->hAccel = NULL; pCI->cAccel = 0; pCI->dwFlags = 0;
    return S_OK;
}

STDMETHODIMP CDsoFramerControl::XOleControl::OnMnemonic(LPMSG pMsg)
{
	METHOD_PROLOGUE(CDsoFramerControl, OleControl);
	ODS("CDsoFramerControl::XOleControl::OnMnemonic\n");
	return pThis->InPlaceActivate(OLEIVERB_UIACTIVATE);
}

STDMETHODIMP CDsoFramerControl::XOleControl::OnAmbientPropertyChange(DISPID dispID)
{
	METHOD_PROLOGUE(CDsoFramerControl, OleControl);
	ODS("CDsoFramerControl::XOleControl::OnAmbientPropertyChange\n");

 // We keep track of state changes in design/run modes...
    if (dispID == DISPID_AMBIENT_USERMODE || dispID == DISPID_UNKNOWN)
        pThis->m_fModeFlagValid = FALSE;

	return S_OK;
}

STDMETHODIMP CDsoFramerControl::XOleControl::FreezeEvents(BOOL bFreeze)
{
	METHOD_PROLOGUE(CDsoFramerControl, OleControl);
	TRACE1("CDsoFramerControl::XOleControl::FreezeEvents = %s\n", (bFreeze ? "TRUE" : "FALSE"));
    pThis->m_fFreezeEvents = bFreeze;
	return S_OK;
}


////////////////////////////////////////////////////////////////////////
//
// CDsoFramerControl::XOleInplaceObject -- IOleInplaceObject Implementation
//
//	 STDMETHODIMP GetWindow(HWND *phwnd);
//	 STDMETHODIMP ContextSensitiveHelp(BOOL fEnterMode);
//	 STDMETHODIMP InPlaceDeactivate();
//	 STDMETHODIMP UIDeactivate();
//	 STDMETHODIMP SetObjectRects(LPCRECT lprcPosRect, LPCRECT lprcClipRect);
//	 STDMETHODIMP ReactivateAndUndo();
//
IMPLEMENT_INTERFACE_UNKNOWN(CDsoFramerControl, OleInplaceObject)

STDMETHODIMP CDsoFramerControl::XOleInplaceObject::GetWindow(HWND *phwnd)
{
	METHOD_PROLOGUE(CDsoFramerControl, OleInplaceObject);
	ODS("CDsoFramerControl::XOleInplaceObject::GetWindow\n");
	if (!(pThis->m_fInPlaceActive)) return E_FAIL;
	if (phwnd) *phwnd = pThis->m_hwnd;
	return S_OK;
}

STDMETHODIMP CDsoFramerControl::XOleInplaceObject::ContextSensitiveHelp(BOOL /*fEnterMode*/)
{
	ODS("CDsoFramerControl::XOleInplaceObject::ContextSensitiveHelp\n");
    return E_NOTIMPL;
}

STDMETHODIMP CDsoFramerControl::XOleInplaceObject::InPlaceDeactivate()
{
	METHOD_PROLOGUE(CDsoFramerControl, OleInplaceObject);
	ODS("CDsoFramerControl::XOleInplaceObject::InPlaceDeactivate\n");

 // if we're not in-place active yet, then this is easy.
    if (!(pThis->m_fInPlaceActive))
        return S_OK;

 // transition from UIActive back to active
    if (pThis->m_fUIActive)
        UIDeactivate();

 // tell the host we're going away
    pThis->m_pInPlaceSite->OnInPlaceDeactivate();

    pThis->m_fInPlaceActive = FALSE;
    pThis->m_fInPlaceVisible = FALSE;

 // if we have a window, tell it to go away.
    if (pThis->m_hwnd)
		{DestroyWindow(pThis->m_hwnd); pThis->m_hwnd = NULL;}
    
    SAFE_RELEASE_INTERFACE(pThis->m_pInPlaceFrame);
    SAFE_RELEASE_INTERFACE(pThis->m_pInPlaceUIWindow);

    return S_OK;
}

STDMETHODIMP CDsoFramerControl::XOleInplaceObject::UIDeactivate()
{
	METHOD_PROLOGUE(CDsoFramerControl, OleInplaceObject);
	ODS("CDsoFramerControl::XOleInplaceObject::UIDeactivate\n");

 // if we're not UIActive, not much to do.
    if (!(pThis->m_fUIActive))
        return S_OK;

    pThis->m_fUIActive = FALSE;

 // notify frame windows, if appropriate, that we're no longer ui-active.
    if (pThis->m_pInPlaceUIWindow)
		pThis->m_pInPlaceUIWindow->SetActiveObject(NULL, NULL);

    if (pThis->m_pInPlaceFrame)
		pThis->m_pInPlaceFrame->SetActiveObject(NULL, NULL);

	pThis->OnUIFocusChange(FALSE);

 // we don't need to explicitly release the focus here since somebody
 // else grabbing the focus is what is likely to cause us to get lose it
    pThis->m_pInPlaceSite->OnUIDeactivate(FALSE);

    return S_OK;
}

STDMETHODIMP CDsoFramerControl::XOleInplaceObject::SetObjectRects(LPCRECT lprcPosRect, LPCRECT lprcClipRect)
{
    RECT rcIXect;

	METHOD_PROLOGUE(CDsoFramerControl, OleInplaceObject);
	ODS("CDsoFramerControl::XOleInplaceObject::SetObjectRects\n");

  // move our window to the new location and handle clipping.
    if (pThis->m_hwnd)
	{
	 // if asked to clip the window, we'll use a clip region...
		if ((lprcClipRect) && (IntersectRect(&rcIXect, lprcPosRect, lprcClipRect)) &&
			(!EqualRect(&rcIXect, lprcPosRect)))
		{
			OffsetRect(&rcIXect, -(lprcPosRect->left), -(lprcPosRect->top));
			SetWindowRgn(pThis->m_hwnd, CreateRectRgnIndirect(&rcIXect), TRUE);
			pThis->m_fUsingWindowRgn = TRUE;
		}
		else if (pThis->m_fUsingWindowRgn)
		{
            SetWindowRgn(pThis->m_hwnd, NULL, TRUE);
            pThis->m_fUsingWindowRgn = FALSE;
		}

        // set our control's location, but don't change it's size at all
        // [people for whom zooming is important should set that up here]
        SetWindowPos(pThis->m_hwnd, NULL, lprcPosRect->left, lprcPosRect->top, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
    }

 // save out our current location.
    pThis->m_rcLocation = *lprcPosRect;
	return S_OK;
}

STDMETHODIMP CDsoFramerControl::XOleInplaceObject::ReactivateAndUndo()
{
	ODS("CDsoFramerControl::XOleInplaceObject::ReactivateAndUndo\n");
	return E_NOTIMPL;
}


////////////////////////////////////////////////////////////////////////
//
// CDsoFramerControl::XOleInplaceActiveObject -- IOleInplaceActiveObject Implementation
//
//	 STDMETHODIMP GetWindow(HWND *phwnd);
//	 STDMETHODIMP ContextSensitiveHelp(BOOL fEnterMode);
//	 STDMETHODIMP TranslateAccelerator(LPMSG lpmsg);
//	 STDMETHODIMP OnFrameWindowActivate(BOOL fActivate);
//	 STDMETHODIMP OnDocWindowActivate(BOOL fActivate);
//	 STDMETHODIMP ResizeBorder(LPCRECT prcBorder, IOleInPlaceUIWindow *pUIWindow, BOOL fFrameWindow);
//	 STDMETHODIMP EnableModeless(BOOL fEnable);
//
IMPLEMENT_INTERFACE_UNKNOWN(CDsoFramerControl, OleInplaceActiveObject)

STDMETHODIMP CDsoFramerControl::XOleInplaceActiveObject::GetWindow(HWND *phwnd)
{
	METHOD_PROLOGUE(CDsoFramerControl, OleInplaceActiveObject);
	ODS("CDsoFramerControl::XOleInplaceActiveObject::GetWindow\n");
	if (phwnd) *phwnd = pThis->m_hwnd;
	return S_OK;
}

STDMETHODIMP CDsoFramerControl::XOleInplaceActiveObject::ContextSensitiveHelp(BOOL /*fEnterMode*/)
{
	ODS("CDsoFramerControl::XOleInplaceActiveObject::ContextSensitiveHelp\n");
    return E_NOTIMPL;
}


////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::XOleInplaceActiveObject::TranslateAccelerator
//
//  We handle certain keys as accelerators, but only if host does not
//  implement it itself, and only when no document is open. The reason
//  why is due to two factors:
//
//  (1) When an out-of-proc AxDoc server is embedded and has focus, 
//      its thread (not ours) gets the accel keys. Word/PPT will treat
//      these for us and contact us through IOleCommandTarget, so
//      we don't want to handle them if AxDoc is handling them for us.
//
//  (2) We cannot know if a given host uses the same accel keys for 
//      other tasks, and we don't want to impose ours on a host. By
//      explicitly asking host to handle it first, we should be able 
//      to avoid any conflicts, and give developer a change to override
//      our handler as needed.
//
static short _SpecialKeyState()
{
    BOOL bShift = (GetKeyState(VK_SHIFT) < 0);
    BOOL bCtrl  = (GetKeyState(VK_CONTROL) < 0);
    BOOL bAlt   = (GetKeyState(VK_MENU) < 0);
    return (short)(bShift + (bCtrl << 1) + (bAlt << 2));
}

STDMETHODIMP CDsoFramerControl::XOleInplaceActiveObject::TranslateAccelerator(LPMSG lpmsg)
{
    HRESULT hr = S_FALSE;
	METHOD_PROLOGUE(CDsoFramerControl, OleInplaceActiveObject);
    CHECK_NULL_RETURN(lpmsg, E_POINTER);
	TRACE2("CDsoFramerControl::XOleInplaceActiveObject::TranslateAccelerator(msg=%d, wparam=%d)\n", lpmsg->message, lpmsg->wParam);

 // Forward it back to the site for processing on its side...
    if (pThis->m_pControlSite)
    {
        hr = pThis->m_pControlSite->TranslateAccelerator(lpmsg, _SpecialKeyState());
        if (hr == S_OK) return S_OK; // If site handled it, nothing else we need to do..
    }

    if ((GetKeyState(VK_CONTROL) < 0) && (!(pThis->m_fModalState)) && 
        (lpmsg->message == WM_KEYUP) && (!(pThis->m_pDocObjFrame)))
    {
        TRACE1(" Handle Accelerator? (%d)\n", lpmsg->wParam);
        switch (lpmsg->wParam) // Looking for Ctrl+N and Ctrl+O keys...
        {
        case 0x4E: PostMessage(pThis->m_hwnd, DSO_WM_ASYNCH_OLECOMMAND, OLECMDID_NEW,  0); hr = S_OK; break;
        case 0x4F: PostMessage(pThis->m_hwnd, DSO_WM_ASYNCH_OLECOMMAND, OLECMDID_OPEN, 0); hr = S_OK; break;
        case 0x53: PostMessage(pThis->m_hwnd, DSO_WM_ASYNCH_OLECOMMAND, OLECMDID_SAVE, 0); hr = S_OK; break;
        }
    }

 // we didn't want it.
    return hr;
}

STDMETHODIMP CDsoFramerControl::XOleInplaceActiveObject::OnFrameWindowActivate(BOOL fActivate)
{
	TRACE1("CDsoFramerControl::XOleInplaceActiveObject::OnFrameWindowActivate - %d\n", fActivate);
    return S_OK;
}

STDMETHODIMP CDsoFramerControl::XOleInplaceActiveObject::OnDocWindowActivate(BOOL fActivate)
{
	METHOD_PROLOGUE(CDsoFramerControl, OleInplaceActiveObject);
	TRACE1("CDsoFramerControl::XOleInplaceActiveObject::OnDocWindowActivate - %d\n", fActivate);
    
	if (pThis->m_fUIActive && fActivate && pThis->m_pInPlaceFrame)
        pThis->m_pInPlaceFrame->SetBorderSpace(NULL);

	return S_OK;
}

STDMETHODIMP CDsoFramerControl::XOleInplaceActiveObject::ResizeBorder(LPCRECT prcBorder, IOleInPlaceUIWindow *pUIWindow, BOOL fFrameWindow)
{
	ODS("CDsoFramerControl::XOleInplaceActiveObject::ResizeBorder\n");
	return S_OK;
}

STDMETHODIMP CDsoFramerControl::XOleInplaceActiveObject::EnableModeless(BOOL fEnable)
{
	TRACE1("CDsoFramerControl::XOleInplaceActiveObject::EnableModeless(%d)\n", fEnable);
	return S_OK;
}

////////////////////////////////////////////////////////////////////////
//
// CDsoFramerControl::XViewObjectEx -- IViewObjectEx Implementation
//
//	 STDMETHODIMP Draw(DWORD dwAspect, LONG lindex, void *pvAspect, DVTARGETDEVICE * ptd, HDC hicTargetDev, HDC hdcDraw, const LPRECTL lprcBounds, const LPRECTL lprcWBounds, BOOL(*)(DWORD)pfnContinue, DWORD dwContinue); 
//	 STDMETHODIMP GetColorSet(DWORD dwAspect, LONG lindex, void* pvAspect, DVTARGETDEVICE *ptd, HDC hicTargetDev, LOGPALETTE** ppColorSet);
//	 STDMETHODIMP Freeze(DWORD dwAspect, LONG lindex, void* pvAspect, DWORD* pdwFreeze);
//	 STDMETHODIMP Unfreeze(DWORD dwFreeze);
//	 STDMETHODIMP SetAdvise(DWORD dwAspect, DWORD advf, IAdviseSink* pAdvSink);
//	 STDMETHODIMP GetAdvise(DWORD* pdwAspect, DWORD* padvf, IAdviseSink** ppAdvSink);
//	 STDMETHODIMP GetExtent(DWORD dwAspect, DWORD lindex, DVTARGETDEVICE ptd, LPSIZEL lpsizel);
//	 STDMETHODIMP GetRect(DWORD dwAspect, LPRECTL pRect);
//	 STDMETHODIMP GetViewStatus(DWORD* pdwStatus);
//	 STDMETHODIMP QueryHitPoint(DWORD dwAspect, LPCRECT pRectBounds, POINT ptlLoc, LONG lCloseHint, DWORD *pHitResult);
//	 STDMETHODIMP QueryHitRect(DWORD dwAspect, LPCRECT pRectBounds, LPCRECT pRectLoc, LONG lCloseHint, DWORD *pHitResult);
//	 STDMETHODIMP GetNaturalExtent(DWORD dwAspect, LONG lindex, DVTARGETDEVICE *ptd, HDC hicTargetDev, DVEXTENTINFO *pExtentInfo, LPSIZEL pSizel);
//
IMPLEMENT_INTERFACE_UNKNOWN(CDsoFramerControl, ViewObjectEx)

STDMETHODIMP CDsoFramerControl::XViewObjectEx::Draw(DWORD dwDrawAspect, LONG lIndex, void *pvAspect, DVTARGETDEVICE *ptd, HDC hicTargetDevice, HDC hdcDraw, LPCRECTL prcBounds, LPCRECTL prcWBounds, BOOL (__stdcall *pfnContinue)(DWORD dwContinue), DWORD dwContinue)
{
    RECT rc;
    int iMode = 0;
    POINT pVp = {0,0};
	POINT pW = {0,0};
    BOOL fOptimize = FALSE;
    BOOL fMetafile = FALSE;

	METHOD_PROLOGUE(CDsoFramerControl, ViewObjectEx);
	ODS("CDsoFramerControl::XViewObjectEx::Draw\n");

	if (!((dwDrawAspect == DVASPECT_CONTENT) || 
		  (dwDrawAspect == DVASPECT_OPAQUE)))
            return DV_E_DVASPECT;

	SaveDC(hdcDraw);
	
 // first, have to do a little bit to support printing.
    if (GetDeviceCaps(hdcDraw, TECHNOLOGY) == DT_METAFILE)
	{
     // We are dealing with a metafile.
        fMetafile = TRUE;
    }
/* 
  // check to see if we have any flags passed in the pvAspect parameter.
   if (pvAspect && ((DVASPECTINFO *)pvAspect)->cb == sizeof(DVASPECTINFO))
        fOptimize = (((DVASPECTINFO *)pvAspect)->dwFlags & DVASPECTINFOFLAG_CANOPTIMIZE) ? TRUE : FALSE;
*/
  // Get the draw rect, if they didn't give us one, just copy over ours
	memcpy(&rc, ((prcBounds) ? (RECT*)prcBounds : &(pThis->m_rcLocation)), sizeof(RECT));

    if (!fMetafile)
	{
        LPtoDP(hdcDraw, (POINT*)&rc, 2);
        SetViewportOrgEx(hdcDraw, 0, 0, &pVp);
        SetWindowOrgEx(hdcDraw, 0, 0, &pW);
        iMode = SetMapMode(hdcDraw, MM_TEXT);
    }

  // prcWBounds is NULL and not used if we are not dealing with a metafile.
  // For metafiles, we pass on rc as *prcBounds, we should also include prcWBounds
    pThis->OnDraw(dwDrawAspect, hdcDraw, &rc, (RECT*)prcWBounds, hicTargetDevice, fOptimize);

  // clean up the DC when we're done with it, if appropriate.
    if (!fMetafile)
	{
        SetViewportOrgEx(hdcDraw, pVp.x, pVp.y, NULL);
        SetWindowOrgEx(hdcDraw, pW.x, pW.y, NULL);
        SetMapMode(hdcDraw, iMode);
    }

	RestoreDC(hdcDraw, -1);
    return S_OK;
}

STDMETHODIMP CDsoFramerControl::XViewObjectEx::GetColorSet(DWORD dwAspect, LONG lindex, void* pvAspect, DVTARGETDEVICE *ptd, HDC hicTargetDev, LOGPALETTE** ppColorSet)
{
	ODS("CDsoFramerControl::XViewObjectEx::GetColorSet\n");
    return E_NOTIMPL;
}

STDMETHODIMP CDsoFramerControl::XViewObjectEx::Freeze(DWORD dwAspect, LONG lindex, void* pvAspect, DWORD* pdwFreeze)
{
	ODS("CDsoFramerControl::XViewObjectEx::Freeze\n");
    return E_NOTIMPL;
}

STDMETHODIMP CDsoFramerControl::XViewObjectEx::Unfreeze(DWORD dwFreeze)
{
	ODS("CDsoFramerControl::XViewObjectEx::Unfreeze\n");
    return E_NOTIMPL;
}

STDMETHODIMP CDsoFramerControl::XViewObjectEx::SetAdvise(DWORD dwAspect, DWORD advf, IAdviseSink* pAdviseSink)
{
	METHOD_PROLOGUE(CDsoFramerControl, ViewObjectEx);
	ODS("CDsoFramerControl::XViewObjectEx::SetAdvise\n");

 // if it's not a content aspect, we don't support it.
    if (!(dwAspect & DVASPECT_CONTENT)) 
        return DV_E_DVASPECT;

 // set up some flags  [we gotta stash for GetAdvise ...]
    pThis->m_fViewAdvisePrimeFirst = (advf & ADVF_PRIMEFIRST) ? TRUE : FALSE;
    pThis->m_fViewAdviseOnlyOnce = (advf & ADVF_ONLYONCE) ? TRUE : FALSE;

    SAFE_RELEASE_INTERFACE(pThis->m_pViewAdviseSink);
    SAFE_SET_INTERFACE(pThis->m_pViewAdviseSink, pAdviseSink);

 // prime them if they want it [we need to store this so they can get flags later]
    if (pThis->m_fViewAdvisePrimeFirst)
        pThis->ViewChanged();

    return S_OK;
}

STDMETHODIMP CDsoFramerControl::XViewObjectEx::GetAdvise(DWORD* pdwAspect, DWORD* padvf, IAdviseSink** ppAdviseSink)
{
	METHOD_PROLOGUE(CDsoFramerControl, ViewObjectEx);
	ODS("CDsoFramerControl::XViewObjectEx::GetAdvise\n");

    if (pdwAspect)
        *pdwAspect = DVASPECT_CONTENT;

    if (padvf)
	{
        *padvf = 0;
        if (pThis->m_fViewAdviseOnlyOnce) *padvf |= ADVF_ONLYONCE;
        if (pThis->m_fViewAdvisePrimeFirst) *padvf |= ADVF_PRIMEFIRST;
    }

    if (ppAdviseSink)
	    { SAFE_SET_INTERFACE(*ppAdviseSink, pThis->m_pViewAdviseSink); }

	return S_OK;
}

STDMETHODIMP CDsoFramerControl::XViewObjectEx::GetExtent(DWORD dwDrawAspect, LONG lindex, DVTARGETDEVICE *ptd, LPSIZEL psizel)
{
	METHOD_PROLOGUE(CDsoFramerControl, ViewObjectEx);
	ODS("CDsoFramerControl::XViewObjectEx::GetExtent\n");
	return pThis->m_xOleObject.GetExtent(dwDrawAspect, psizel);
}

STDMETHODIMP CDsoFramerControl::XViewObjectEx::GetRect(DWORD dwAspect, LPRECTL pRect)
{
	METHOD_PROLOGUE(CDsoFramerControl, ViewObjectEx);
	ODS("CDsoFramerControl::XViewObjectEx::GetRect\n");

    if (dwAspect != DVASPECT_CONTENT)
        return DV_E_DVASPECT;

	if (pRect)
	{
		CopyRect((LPRECT)pRect, &(pThis->m_rcLocation));
	 // Convert to himetric (according to docs, this is required)
		DsoPixelsToHimetric(&(pRect->left), &(pRect->top));
		DsoPixelsToHimetric(&(pRect->right), &(pRect->bottom));
	}

	return S_OK;
}

STDMETHODIMP CDsoFramerControl::XViewObjectEx::GetViewStatus(DWORD* pdwStatus)
{
	ODS("CDsoFramerControl::XViewObjectEx::GetViewStatus\n");
	if (pdwStatus) *pdwStatus = VIEWSTATUS_OPAQUE;
	return S_OK;
}

STDMETHODIMP CDsoFramerControl::XViewObjectEx::QueryHitPoint(DWORD dwAspect, LPCRECT pRectBounds, POINT ptlLoc, LONG lCloseHint, DWORD *pHitResult)
{
 // ODS("CDsoFramerControl::XViewObjectEx::QueryHitPoint\n");
    if (dwAspect != DVASPECT_CONTENT) return DV_E_DVASPECT;
    CHECK_NULL_RETURN(pHitResult, E_POINTER);
    *pHitResult = PtInRect(pRectBounds, ptlLoc) ? HITRESULT_HIT : HITRESULT_OUTSIDE;
    return S_OK;
}

STDMETHODIMP CDsoFramerControl::XViewObjectEx::QueryHitRect(DWORD dwAspect, LPCRECT pRectBounds, LPCRECT pRectLoc, LONG lCloseHint, DWORD *pHitResult)
{
    RECT rc;
	//ODS("CDsoFramerControl::XViewObjectEx::QueryHitRect\n");
    if (dwAspect != DVASPECT_CONTENT) return DV_E_DVASPECT;
    *pHitResult = IntersectRect(&rc, pRectBounds, pRectLoc) ? HITRESULT_HIT : HITRESULT_OUTSIDE;
    return S_OK;
}

STDMETHODIMP CDsoFramerControl::XViewObjectEx::GetNaturalExtent(DWORD dwAspect, LONG lindex, DVTARGETDEVICE *ptd, HDC hicTargetDev, DVEXTENTINFO *pExtentInfo, LPSIZEL pSizel)
{
	ODS("CDsoFramerControl::XViewObjectEx::GetNaturalExtent\n");
	return E_NOTIMPL;
}

////////////////////////////////////////////////////////////////////////
//
// CDsoFramerControl::XDataObject -- IDataObject Implementation
//
//	 STDMETHODIMP Draw(DWORD dwAspect, LONG lindex, void *pvAspect, DVTARGETDEVICE * ptd, HDC hicTargetDev, HDC hdcDraw, const LPRECTL lprcBounds, const LPRECTL lprcWBounds, BOOL(*)(DWORD)pfnContinue, DWORD dwContinue); 
//   STDMETHODIMP GetData(FORMATETC *pfmtc,  STGMEDIUM *pstgm);
//   STDMETHODIMP GetDataHere(FORMATETC *pfmtc, STGMEDIUM *pstgm);
//   STDMETHODIMP QueryGetData(FORMATETC *pfmtc);
//   STDMETHODIMP GetCanonicalFormatEtc(FORMATETC *pfmtcIn, FORMATETC *pfmtcOut);
//   STDMETHODIMP SetData(FORMATETC *pfmtc, STGMEDIUM *pstgm, BOOL fRelease);
//   STDMETHODIMP EnumFormatEtc(DWORD dwDirection, IEnumFORMATETC **ppenum);
//   STDMETHODIMP DAdvise(FORMATETC *pfmtc, DWORD advf, IAdviseSink *psink, DWORD *pdwConnection);
//   STDMETHODIMP DUnadvise(DWORD dwConnection);
//   STDMETHODIMP EnumDAdvise(IEnumSTATDATA **ppenum);
//
IMPLEMENT_INTERFACE_UNKNOWN(CDsoFramerControl, DataObject)

STDMETHODIMP CDsoFramerControl::XDataObject::GetData(FORMATETC *pfmtc,  STGMEDIUM *pstgm)
{
    HRESULT hr;
    HMETAFILE hwmf;
    LPMETAFILEPICT pwmf;
    HGLOBAL hmem;
    SIZE sizeMetric;
    RECT rc;
    HDC hdc;

	METHOD_PROLOGUE(CDsoFramerControl, DataObject);
    TRACE2("CDsoFramerControl::XDataObject::GetData(%d,%d)\n", pfmtc->cfFormat, pfmtc->tymed);
    CHECK_NULL_RETURN(pfmtc, E_POINTER); CHECK_NULL_RETURN(pstgm, E_POINTER);

 // Confirm data requested is a supportted type...
    if (FAILED(hr = QueryGetData(pfmtc)))
        return hr;

 // If caller wants a metafile of the current object, draw it...
    hdc = CreateMetaFileW(NULL);

    sizeMetric.cx = pThis->m_Size.cx;
    sizeMetric.cy = pThis->m_Size.cy;

    SetRect(&rc, 0, 0, sizeMetric.cx, sizeMetric.cy);
    DsoPixelsToHimetric(&(sizeMetric.cx), &(sizeMetric.cy));
    sizeMetric.cy = -(sizeMetric.cy);

    SetMapMode(hdc, MM_ANISOTROPIC);
	SetWindowOrgEx(hdc, 0, 0, NULL);
	SetWindowExtEx(hdc, rc.right, rc.bottom, NULL);

 // Delegate to same IViewObject::Draw used by most containers for metafile...
    hr = pThis->m_xViewObjectEx.Draw(DVASPECT_CONTENT, -1, NULL, NULL, NULL, hdc, (RECTL*)&rc, NULL, NULL, 0);
    ASSERT(SUCCEEDED(hr));

	if ((hwmf = CloseMetaFile(hdc)) == NULL)
		return E_UNEXPECTED;

 // Now allocate the transfer medium and set the metafile...
    hmem = GlobalAlloc((GMEM_SHARE|GMEM_MOVEABLE), sizeof(METAFILEPICT));
	if (hmem == NULL) { DeleteMetaFile(hwmf); return STG_E_MEDIUMFULL; }

	pwmf = (LPMETAFILEPICT)GlobalLock(hmem);
    if (pwmf)
    {
	    pwmf->hMF  = hwmf;
	    pwmf->mm   = MM_ANISOTROPIC;
	    pwmf->xExt = sizeMetric.cx;
	    pwmf->yExt = sizeMetric.cy;
	    GlobalUnlock(hmem);
    }

	pstgm->tymed = TYMED_MFPICT;
	pstgm->hGlobal = hmem;
	pstgm->pUnkForRelease = NULL;
    return S_OK;
}

STDMETHODIMP CDsoFramerControl::XDataObject::GetDataHere(FORMATETC *pfmtc, STGMEDIUM *pstgm)
{
    ODS("CDsoFramerControl::XDataObject::GetDataHere\n");
    return DV_E_FORMATETC;
}

STDMETHODIMP CDsoFramerControl::XDataObject::QueryGetData(FORMATETC *pfmtc)
{
    TRACE2("CDsoFramerControl::XDataObject::QueryGetData(%d,%d)\n", pfmtc->cfFormat, pfmtc->tymed);
    CHECK_NULL_RETURN(pfmtc, E_POINTER);

    if (pfmtc->cfFormat != CF_METAFILEPICT)
        return DV_E_FORMATETC; // We only support WMF for IDataObj presentation.

    if (pfmtc->tymed != TYMED_MFPICT)
        return DV_E_TYMED; // It better be in in right format...

    if (pfmtc->dwAspect != DVASPECT_CONTENT)
        return DV_E_DVASPECT; // We only do content (not thumbnail or partial)...

    return S_OK; // If here, it looks fine to us...
}

STDMETHODIMP CDsoFramerControl::XDataObject::GetCanonicalFormatEtc(FORMATETC *pfmtcIn, FORMATETC *pfmtcOut)
{
    ODS("CDsoFramerControl::XDataObject::GetCanonicalFormatEtc\n");
    return E_NOTIMPL;
}

STDMETHODIMP CDsoFramerControl::XDataObject::SetData(FORMATETC *pfmtc, STGMEDIUM *pstgm, BOOL fRelease)
{
    ODS("CDsoFramerControl::XDataObject::SetData\n");
    return E_NOTIMPL;
}

STDMETHODIMP CDsoFramerControl::XDataObject::EnumFormatEtc(DWORD dwDirection, IEnumFORMATETC **ppenum)
{
    ODS("CDsoFramerControl::XDataObject::EnumFormatEtc\n");
	return OleRegEnumFormatEtc(CLSID_FramerControl, dwDirection, ppenum);
}

STDMETHODIMP CDsoFramerControl::XDataObject::DAdvise(FORMATETC *pfmtc, DWORD advf, IAdviseSink *psink, DWORD *pdwConnection)
{
	METHOD_PROLOGUE(CDsoFramerControl, DataObject);
    ODS("CDsoFramerControl::XDataObject::DAdvise\n");

 // Create OLE Data advise holder to keep sink for us...
	if ((NULL == pThis->m_pDataAdviseHolder) && 
        FAILED(CreateDataAdviseHolder(&(pThis->m_pDataAdviseHolder))))
			return E_FAIL;

 // Add this advise sink for the given format to the collection...
	return pThis->m_pDataAdviseHolder->Advise((IDataObject*)&(pThis->m_xDataObject), pfmtc, advf, psink, pdwConnection);
}

STDMETHODIMP CDsoFramerControl::XDataObject::DUnadvise(DWORD dwConnection)
{
	METHOD_PROLOGUE(CDsoFramerControl, DataObject);
    ODS("CDsoFramerControl::XDataObject::DUnadvise\n");
	CHECK_NULL_RETURN(pThis->m_pDataAdviseHolder, CONNECT_E_NOCONNECTION);
    return pThis->m_pDataAdviseHolder->Unadvise(dwConnection);
}

STDMETHODIMP CDsoFramerControl::XDataObject::EnumDAdvise(IEnumSTATDATA **ppenum)
{
	METHOD_PROLOGUE(CDsoFramerControl, DataObject);
    ODS("CDsoFramerControl::XDataObject::EnumDAdvise\n");
	CHECK_NULL_RETURN(pThis->m_pDataAdviseHolder, CONNECT_E_NOCONNECTION);
    return pThis->m_pDataAdviseHolder->EnumAdvise(ppenum);
}

////////////////////////////////////////////////////////////////////////
//
// CDsoFramerControl::XProvideClassInfo -- IProvideClassInfo Implementation
//
//	 STDMETHODIMP GetClassInfo(ITypeInfo** ppTI);
//
IMPLEMENT_INTERFACE_UNKNOWN(CDsoFramerControl, ProvideClassInfo)

STDMETHODIMP CDsoFramerControl::XProvideClassInfo::GetClassInfo(ITypeInfo** ppTI)
{
	ODS("CDsoFramerControl::XProvideClassInfo::GetClassInfo\n");
	return DsoGetTypeInfoEx(LIBID_DSOFramer, 0,
		DSOFRAMERCTL_VERSION_MAJOR, DSOFRAMERCTL_VERSION_MINOR, v_hModule, CLSID_FramerControl, ppTI);
}

////////////////////////////////////////////////////////////////////////
//
// CDsoFramerControl::XConnectionPointContainer -- IConnectionPointContainer Implementation
//
//	 STDMETHODIMP EnumConnectionPoints(IEnumConnectionPoints **ppEnum);
//	 STDMETHODIMP FindConnectionPoint(REFIID riid, IConnectionPoint **ppCP);
//
IMPLEMENT_INTERFACE_UNKNOWN(CDsoFramerControl, ConnectionPointContainer)

STDMETHODIMP CDsoFramerControl::XConnectionPointContainer::EnumConnectionPoints(IEnumConnectionPoints **ppEnum)
{
	METHOD_PROLOGUE(CDsoFramerControl, ConnectionPointContainer);
    ODS("CDsoFramerControl::XConnectionPointContainer::EnumConnectionPoints\n");
    if (ppEnum) { SAFE_SET_INTERFACE(*ppEnum, &(pThis->m_xEnumConnectionPoints)); };
	return S_OK;
}

STDMETHODIMP CDsoFramerControl::XConnectionPointContainer::FindConnectionPoint(REFIID riid, IConnectionPoint **ppCP)
{
	METHOD_PROLOGUE(CDsoFramerControl, ConnectionPointContainer);
    ODS("CDsoFramerControl::XConnectionPointContainer::FindConnectionPoint\n");
	if (riid != DIID__DFramerCtlEvents) return CONNECT_E_NOCONNECTION;
	if (ppCP) { SAFE_SET_INTERFACE(*ppCP, &(pThis->m_xConnectionPoint)); };
	return S_OK;
}


////////////////////////////////////////////////////////////////////////
//
// CDsoFramerControl::XEnumConnectionPoints -- IEnumConnectionPoints Implementation
//
//	 STDMETHODIMP Next(ULONG cConnections, IConnectionPoint **rgpcn, ULONG *pcFetched);
//	 STDMETHODIMP Skip(ULONG cConnections);
//	 STDMETHODIMP Reset(void);
//	 STDMETHODIMP Clone(IEnumConnectionPoints **ppEnum);
//
IMPLEMENT_INTERFACE_UNKNOWN(CDsoFramerControl, EnumConnectionPoints)

STDMETHODIMP CDsoFramerControl::XEnumConnectionPoints::Next(ULONG cConnections, IConnectionPoint **rgpcn, ULONG *pcFetched)
{
	HRESULT hr = S_FALSE;

	METHOD_PROLOGUE(CDsoFramerControl, EnumConnectionPoints);
    ODS("CDsoFramerControl::XEnumConnectionPoints::Next\n");

	CHECK_NULL_RETURN(rgpcn, E_POINTER);
	*rgpcn = NULL;

	CHECK_NULL_RETURN(pcFetched, E_POINTER);
	*pcFetched = 0;

	if (pThis->m_fConCntDone)
		pThis->m_fConCntDone = FALSE;
	else
	{
        SAFE_SET_INTERFACE(*rgpcn, &(pThis->m_xConnectionPoint));
		*pcFetched = 1;
		pThis->m_fConCntDone = TRUE; hr = S_OK;
	}

	return hr;
}

STDMETHODIMP CDsoFramerControl::XEnumConnectionPoints::Skip(ULONG cConnections)
{return S_OK;}

STDMETHODIMP CDsoFramerControl::XEnumConnectionPoints::Reset(void)
{return S_OK;}

STDMETHODIMP CDsoFramerControl::XEnumConnectionPoints::Clone(IEnumConnectionPoints **ppEnum)
{if (ppEnum) *ppEnum = this; return S_OK;}


////////////////////////////////////////////////////////////////////////
//
// CDsoFramerControl::XConnectionPoint -- IConnectionPoint Implementation
//
//	 STDMETHODIMP GetConnectionInterface(IID *pIID);
//	 STDMETHODIMP GetConnectionPointContainer(IConnectionPointContainer **ppCPC);
//	 STDMETHODIMP Advise(IUnknown *pUnk, DWORD *pdwCookie);
//	 STDMETHODIMP Unadvise(DWORD dwCookie);
//	 STDMETHODIMP EnumConnections(IEnumConnections **ppEnum);
//
IMPLEMENT_INTERFACE_UNKNOWN(CDsoFramerControl, ConnectionPoint)

STDMETHODIMP CDsoFramerControl::XConnectionPoint::GetConnectionInterface(IID *pIID)
{
    ODS("CDsoFramerControl::XConnectionPoint::GetConnectionInterface\n");
	if (pIID) *pIID = DIID__DFramerCtlEvents;
	return S_OK;
}

STDMETHODIMP CDsoFramerControl::XConnectionPoint::GetConnectionPointContainer(IConnectionPointContainer **ppCPC)
{
	METHOD_PROLOGUE(CDsoFramerControl, ConnectionPoint);
    ODS("CDsoFramerControl::XConnectionPoint::GetConnectionPointContainer\n");
    if (ppCPC) { SAFE_SET_INTERFACE(*ppCPC, &(pThis->m_xConnectionPointContainer)); };
	return S_OK;
}

STDMETHODIMP CDsoFramerControl::XConnectionPoint::Advise(IUnknown *pUnk, DWORD *pdwCookie)
{
	METHOD_PROLOGUE(CDsoFramerControl, ConnectionPoint);
    ODS("CDsoFramerControl::XConnectionPoint::Advise\n");

    CHECK_NULL_RETURN(pUnk, E_POINTER);
    CHECK_NULL_RETURN(pdwCookie, E_POINTER);

	SAFE_RELEASE_INTERFACE(pThis->m_dispEvents);

 // FIX: The COM Interop layer in .NET CLR returns a different RCW for IDispatch than for a specific
 // dispinterface GUID, so for events to work in .NET WinForm applications, we should explicitly use
 // the RCW for that dispinterface. Other event source clients may fail QI for this GUID, so if it 
 // does return an error, we will QI again for IDispatch. The event issue is change in behavior from 
 // how VB6 implemented event sinks, and how VB7/C# does them...
	HRESULT hr = pUnk->QueryInterface(DIID__DFramerCtlEvents, (void**)&(pThis->m_dispEvents));
    if (FAILED(hr)) hr = pUnk->QueryInterface(IID_IDispatch, (void**)&(pThis->m_dispEvents));

    if (SUCCEEDED(hr)) *pdwCookie = (DWORD)(pThis);
	return hr;
}

STDMETHODIMP CDsoFramerControl::XConnectionPoint::Unadvise(DWORD dwCookie)
{
	METHOD_PROLOGUE(CDsoFramerControl, ConnectionPoint);
    ODS("CDsoFramerControl::XConnectionPoint::Unadvise\n");
	if (dwCookie == (DWORD)(pThis))
	{
		SAFE_RELEASE_INTERFACE(pThis->m_dispEvents);
	}
	return (pThis->m_dispEvents == NULL) ? S_OK : CONNECT_E_NOCONNECTION;
}

STDMETHODIMP CDsoFramerControl::XConnectionPoint::EnumConnections(IEnumConnections **ppEnum)
{if (ppEnum) *ppEnum = NULL; return E_NOTIMPL;}


////////////////////////////////////////////////////////////////////////
//
// CDsoFramerControl::XOleCommandTarget -- IOleCommandTarget Implementation
//
//   STDMETHODIMP QueryStatus(const GUID *pguidCmdGroup, ULONG cCmds, OLECMD prgCmds[], OLECMDTEXT *pCmdText);
//   STDMETHODIMP Exec(const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANTARG *pvaIn, VARIANTARG *pvaOut);            
//
IMPLEMENT_INTERFACE_UNKNOWN(CDsoFramerControl, OleCommandTarget)

STDMETHODIMP CDsoFramerControl::XOleCommandTarget::QueryStatus(const GUID *pguidCmdGroup, ULONG cCmds, OLECMD prgCmds[], OLECMDTEXT *pCmdText)
{
	METHOD_PROLOGUE(CDsoFramerControl, OleCommandTarget);
    //ODS("CDsoFramerControl::XOleCommandTarget::QueryStatus\n");
	if (pguidCmdGroup != NULL)
		return OLECMDERR_E_UNKNOWNGROUP;
			
	if (prgCmds == NULL)
		return E_INVALIDARG;

    for (int i = 0; i < (int)cCmds; i++) 
	{
		switch (prgCmds[i].cmdID) 
		{
		case OLECMDID_NEW:			prgCmds[i].cmdf = (((pThis->m_wFileMenuFlags &   1) ? OLECMDF_ENABLED : 0) | OLECMDF_SUPPORTED); break;
		case OLECMDID_OPEN:			prgCmds[i].cmdf = (((pThis->m_wFileMenuFlags &   2) ? OLECMDF_ENABLED : 0) | OLECMDF_SUPPORTED); break;
		case OLECMDID_SAVE:			prgCmds[i].cmdf = (((pThis->m_wFileMenuFlags &   8) ? OLECMDF_ENABLED : 0) | OLECMDF_SUPPORTED); break;
		case OLECMDID_SAVEAS:		prgCmds[i].cmdf = (((pThis->m_wFileMenuFlags &  16) ? OLECMDF_ENABLED : 0) | OLECMDF_SUPPORTED); break;
		case OLECMDID_PRINT:		prgCmds[i].cmdf = (((pThis->m_wFileMenuFlags &  32) ? OLECMDF_ENABLED : 0) | OLECMDF_SUPPORTED); break;
		case OLECMDID_PAGESETUP:	prgCmds[i].cmdf = (((pThis->m_wFileMenuFlags &  64) ? OLECMDF_ENABLED : 0) | OLECMDF_SUPPORTED); break;
		case OLECMDID_PROPERTIES:	prgCmds[i].cmdf = (((pThis->m_wFileMenuFlags & 128) ? OLECMDF_ENABLED : 0) | OLECMDF_SUPPORTED); break;
        case OLECMDID_PRINTPREVIEW: prgCmds[i].cmdf = (((pThis->m_wFileMenuFlags & 256) ? OLECMDF_ENABLED : 0) | OLECMDF_SUPPORTED); break;
		default: prgCmds[i].cmdf = 0;
		}
	}
	
	if (pCmdText)
    {
        pCmdText->cmdtextf = MSOCMDTEXTF_NONE;
        pCmdText->cwActual = 0;
    }

    return S_OK;
}

STDMETHODIMP CDsoFramerControl::XOleCommandTarget::Exec(const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANTARG *pvaIn, VARIANTARG *pvaOut)
{
	METHOD_PROLOGUE(CDsoFramerControl, OleCommandTarget);
    ODS("CDsoFramerControl::XOleCommandTarget::Exec\n");

    if ((pguidCmdGroup != NULL) && (*pguidCmdGroup != GUID_NULL))
        return OLECMDERR_E_UNKNOWNGROUP;
	    
    if (nCmdexecopt == MSOCMDEXECOPT_SHOWHELP)
        return OLECMDERR_E_NOHELP;

	switch (nCmdID)
	{
	case OLECMDID_NEW:
	case OLECMDID_OPEN:
	case OLECMDID_SAVE:
	case OLECMDID_SAVEAS:
	case OLECMDID_PRINT:
	case OLECMDID_PAGESETUP:
	case OLECMDID_PROPERTIES:
    case OLECMDID_PRINTPREVIEW:
		PostMessage(pThis->m_hwnd, DSO_WM_ASYNCH_OLECOMMAND, nCmdID, nCmdexecopt);
		return S_OK;
	}

	return OLECMDERR_E_NOTSUPPORTED;
}

////////////////////////////////////////////////////////////////////////
//
// CDsoFramerControl::XSupportErrorInfo -- ISupportErrorInfo Implementation
//
//	 STDMETHODIMP InterfaceSupportsErrorInfo(REFIID riid);
//
IMPLEMENT_INTERFACE_UNKNOWN(CDsoFramerControl, SupportErrorInfo)

STDMETHODIMP CDsoFramerControl::XSupportErrorInfo::InterfaceSupportsErrorInfo(REFIID riid)
{
	ODS("CDsoFramerControl::XSupportErrorInfo::InterfaceSupportsErrorInfo\n");
	return ((riid == IID__FramerControl) ? S_OK : E_FAIL);
}

////////////////////////////////////////////////////////////////////////
//
// CDsoFramerControl::XObjectSafety -- IObjectSafety Implementation
//
//  STDMETHODIMP GetInterfaceSafetyOptions(REFIID riid, DWORD *pdwSupportedOptions,DWORD *pdwEnabledOptions);
//  STDMETHODIMP SetInterfaceSafetyOptions(REFIID riid, DWORD dwOptionSetMask, DWORD dwEnabledOptions);
//
IMPLEMENT_INTERFACE_UNKNOWN(CDsoFramerControl, ObjectSafety)

STDMETHODIMP CDsoFramerControl::XObjectSafety::GetInterfaceSafetyOptions(REFIID riid, DWORD *pdwSupportedOptions,DWORD *pdwEnabledOptions)
{
	ODS("CDsoFramerControl::XObjectSafety::GetInterfaceSafetyOptions\n");
	CHECK_NULL_RETURN(pdwEnabledOptions, E_POINTER);

	if (pdwSupportedOptions)
		*pdwSupportedOptions = (INTERFACESAFE_FOR_UNTRUSTED_DATA|INTERFACESAFE_FOR_UNTRUSTED_CALLER);

 // We are safe for initialization, but not necessarily for scripting. Since script
 // can open documents with macros, and get access to Office OM via ActiveDocument,
 // we cannot be "safe" for malicious script. But this allows the control to load
 // without prompt (just not scripted without prompt)...
	*pdwEnabledOptions = INTERFACESAFE_FOR_UNTRUSTED_DATA;
	return S_OK;
}

STDMETHODIMP CDsoFramerControl::XObjectSafety::SetInterfaceSafetyOptions(REFIID riid, DWORD dwOptionSetMask, DWORD dwEnabledOptions)
{
	ODS("CDsoFramerControl::XObjectSafety::SetInterfaceSafetyOptions\n");
	return ((((riid == IID_IPersist) || (riid == IID_IPersistPropertyBag)) &&
		(dwEnabledOptions == INTERFACESAFE_FOR_UNTRUSTED_DATA)) ? S_OK : E_FAIL);
}

////////////////////////////////////////////////////////////////////////
//
// CDsoFramerControl::XDsoDocObjectSite -- IDsoDocObjectSite Implementation
//
//  STDMETHODIMP GetWindow(HWND* phWnd);
//  STDMETHODIMP GetBorder(LPRECT prcBorder);
//  STDMETHODIMP SetStatusText(LPCOLESTR pszText);
//  STDMETHODIMP GetHostName(LPWSTR *ppwszHostName);
//
IMPLEMENT_INTERFACE_UNKNOWN(CDsoFramerControl, DsoDocObjectSite)

STDMETHODIMP CDsoFramerControl::XDsoDocObjectSite::QueryService(REFGUID guidService, REFIID riid, void **ppv)
{
	HRESULT hr = E_NOINTERFACE;
    ODS("CDsoFramerControl::XDsoDocObjectSite::QueryService\n");

    METHOD_PROLOGUE(CDsoFramerControl, DsoDocObjectSite);
    CHECK_NULL_RETURN(ppv, E_POINTER);

 // We will return control IDispatch for cross container automation...
    if (guidService == SID_SContainerDispatch)
    {
		ODS(" -- Container Dispatch Given --\n");
        hr = pThis->QueryInterface(riid, ppv);
    }
 // TODO: Support IHlinkFrame services for cross hyperlinking inside control??

 // If we don't handle the call, and our host has a provider, pass the call along to see
 // if the top-level host offers the service requested (like hyperlinking to new web page)...
    if (FAILED(hr) && (pThis->m_pClientSite))
    {
        IServiceProvider *pprov = NULL;
        if (SUCCEEDED(pThis->m_pClientSite->QueryInterface(IID_IServiceProvider, (void**)&pprov)) && (pprov))
        {
            if (SUCCEEDED(hr = pprov->QueryService(guidService, riid, ppv)))
                { ODS(" -- Service provided by top-level host --\n"); }
            pprov->Release();
        }
    }

    return hr;
}

STDMETHODIMP CDsoFramerControl::XDsoDocObjectSite::GetWindow(HWND* phWnd)
{
	METHOD_PROLOGUE(CDsoFramerControl, DsoDocObjectSite);
    ODS("CDsoFramerControl::XDsoDocObjectSite::GetWindow\n");
    if (phWnd) *phWnd = pThis->m_hwnd;
    return S_OK;
}

STDMETHODIMP CDsoFramerControl::XDsoDocObjectSite::GetBorder(LPRECT prcBorder)
{
	METHOD_PROLOGUE(CDsoFramerControl, DsoDocObjectSite);
    ODS("CDsoFramerControl::XDsoDocObjectSite::GetBorder\n");
    pThis->GetSizeRectForDocument(NULL, prcBorder);
    return S_OK;
}

STDMETHODIMP CDsoFramerControl::XDsoDocObjectSite::GetHostName(LPWSTR *ppwszHostName)
{
	METHOD_PROLOGUE(CDsoFramerControl, DsoDocObjectSite);
    ODS("CDsoFramerControl::XDsoDocObjectSite::GetHostName\n");
    if (ppwszHostName) *ppwszHostName = DsoCopyString(pThis->m_pwszHostName);
    return S_OK;
}

STDMETHODIMP CDsoFramerControl::XDsoDocObjectSite::SysMenuCommand(UINT uiCharCode)
{
 	METHOD_PROLOGUE(CDsoFramerControl, DsoDocObjectSite);
    TRACE1("CDsoFramerControl::XDsoDocObjectSite::SysMenuCommand(%x)\n", uiCharCode);
	
 // If we are processing a sys menu command, see if it is a menu accel for our
 // displayed menu, and if so we handle this with a fake click of the menu button...
    if ((uiCharCode) && (pThis->m_fShowMenuBar) && !(pThis->m_fDisableMenuAccel))
    {
        UINT i;
        CHAR *pch;

	 // Loop all the menu items and check against their accelerators...
		for (i = 0; i < DSO_MAX_MENUITEMS; i++)
		{
			pch = &(pThis->m_rgchMenuAccel[0]);
			if (*pch)
			{
				if ((*pch != 0xFF) && ((UINT)(*pch) == ASCII_LOWERCASE(uiCharCode)))
				{
				 // Found one that match, so fake mouse move and button down/up to run click code...
					ODS("-- Trigger Menu Accelerator --\n");

					PostMessage(pThis->m_hwnd, WM_MOUSEMOVE, (WPARAM)0, 
						MAKELPARAM(pThis->m_rgrcMenuItems[i].left + 1, pThis->m_rgrcMenuItems[i].top + 1));

					PostMessage(pThis->m_hwnd, WM_LBUTTONDOWN, (WPARAM)MK_LBUTTON, 
						MAKELPARAM(pThis->m_rgrcMenuItems[i].left + 1, pThis->m_rgrcMenuItems[i].top + 1));

					PostMessage(pThis->m_hwnd, WM_LBUTTONUP, (WPARAM)0, 
						MAKELPARAM(pThis->m_rgrcMenuItems[i].left + 1, pThis->m_rgrcMenuItems[i].top + 1));

				 // Return OK to caller to know to eat this message.
					return S_OK;
				}
			}
			else break; 
        }
   }

   return S_FALSE;
}

STDMETHODIMP CDsoFramerControl::XDsoDocObjectSite::SetStatusText(LPCOLESTR pszText)
{
	METHOD_PROLOGUE(CDsoFramerControl, DsoDocObjectSite);

 // If m_fActivateOnStatus flag is set, see if we are not UI active. If not, try to activate...
    if ((pThis->m_fActivateOnStatus) && (pThis->m_fComponentActive) && (pThis->m_fAppActive) && 
		!(pThis->m_fUIActive) && !(pThis->m_fModalState))
    {
        ODS(" -- ForceUIActiveFromSetStatusText --\n");
		pThis->m_fActivateOnStatus = FALSE;
        pThis->Activate();
    }

    return S_OK;
}


////////////////////////////////////////////////////////////////////////
// CDsoFrameHookManager
//
CDsoFrameHookManager* vgGlobalManager = NULL;

////////////////////////////////////////////////////////////////////////
// CDsoFrameHookManager::RegisterFramerControl
//
//  Creates a frame hook and registers it with the active manager. If 
//  this is the first time called, the global manager will be created
//  now.  The manager can handle up to DSOF_MAX_CONTROLS controls (10).
//
STDMETHODIMP_(CDsoFrameHookManager*) CDsoFrameHookManager::RegisterFramerControl(HWND hwndParent, HWND hwndControl)
{
	CDsoFrameHookManager* pManager = NULL;
	ODS("CDsoFrameHookManager::RegisterFramerControl\n");
    CHECK_NULL_RETURN(hwndControl, NULL);

	EnterCriticalSection(&v_csecThreadSynch);
	pManager = vgGlobalManager;
	if (pManager == NULL)
		pManager = (vgGlobalManager = new CDsoFrameHookManager());
	LeaveCriticalSection(&v_csecThreadSynch);

	if (pManager)
	{
        if (hwndParent == NULL)
            hwndParent = hwndControl;

		if (FAILED(pManager->AddComponent(hwndParent, hwndControl)))
			return NULL;
	}

	return pManager;
}

////////////////////////////////////////////////////////////////////////
// CDsoFrameHookManager::AddComponent
//
//  Register the component and create the actual frame hook (if needed).
//
STDMETHODIMP CDsoFrameHookManager::AddComponent(HWND hwndParent, HWND hwndControl)
{
	DWORD dwHostProcessID;
	CDsoFrameWindowHook* phook;
	HRESULT hr = E_OUTOFMEMORY;
	HWND hwndNext, hwndHost = hwndParent;
	ODS("CDsoFrameHookManager::AddComponent\n");

 // Find the top-level window this control is sited on...
	while (hwndNext = GetParent(hwndHost))
		hwndHost = hwndNext;

 // We have to get valid window that is in-thread for a subclass to work...
	if (!IsWindow(hwndHost) || 
		(GetWindowThreadProcessId(hwndHost, &dwHostProcessID), 
			dwHostProcessID != GetCurrentProcessId()))
		return E_ACCESSDENIED;

 // Next if we have room in the array, add the component to the list and
 // hook its parent window so we get messages for it.
	EnterCriticalSection(&v_csecThreadSynch);
	if ((vgGlobalManager) && (m_cComponents < DSOF_MAX_CONTROLS))
	{
	 // Hook the host window for this control...
		phook = CDsoFrameWindowHook::AttachToFrameWindow(hwndHost);
		if (phook) 
		{// Add the component to the list if hooked...
			m_pComponents[m_cComponents].hwndControl = hwndControl;
			m_pComponents[m_cComponents].phookFrame  = phook;
			m_cComponents++; hr = S_OK; 
		}
	}
	LeaveCriticalSection(&v_csecThreadSynch);

	if (SUCCEEDED(hr)) // Make sure item is the new active component
		hr = SetActiveComponent(hwndControl);

	return hr;
}

STDMETHODIMP CDsoFrameHookManager::DetachComponent(HWND hwndControl)
{
	DWORD dwIndex;
	CHECK_NULL_RETURN(m_cComponents, E_FAIL);

 // Find this control in the list...
	for (dwIndex = 0; dwIndex < m_cComponents; dwIndex++)
		if (m_pComponents[dwIndex].hwndControl == hwndControl) break;

	if (dwIndex > m_cComponents)
		return E_INVALIDARG;

 // If we found the index, remove the item and shift remaining
 // items down the list. Change active index as needed...
	EnterCriticalSection(&v_csecThreadSynch);

	m_pComponents[dwIndex].hwndControl = NULL;
	m_pComponents[dwIndex].phookFrame->Detach();

 // Compact the list...
	while (++dwIndex < m_cComponents)
	{
		m_pComponents[dwIndex-1].hwndControl = m_pComponents[dwIndex].hwndControl;
		m_pComponents[dwIndex-1].phookFrame = m_pComponents[dwIndex].phookFrame;
		m_pComponents[dwIndex].hwndControl = NULL;
	}

 // Decrement the count. If count still exists, forward activation
 // to next component in the list...
	if (--m_cComponents)
	{
		while (m_idxActive >= m_cComponents)
			--m_idxActive;

		HWND hwndActive = GetActiveComponentWindow();
		if (hwndActive) 
			SendNotifyMessage(hwndActive, DSO_WM_HOOK_NOTIFY_COMPACTIVE, (WPARAM)TRUE, (LPARAM)(m_fAppActive));
	}
	else
	{
		vgGlobalManager = NULL;
	 // If this is the last control, we can remove the manger!
		delete this;
	}

	LeaveCriticalSection(&v_csecThreadSynch);

	return S_OK;
}

STDMETHODIMP CDsoFrameHookManager::SetActiveComponent(HWND hwndControl)
{
	ODS(" -- CDsoFrameHookManager::SetActiveComponent -- \n");
	CHECK_NULL_RETURN(m_cComponents, E_FAIL);
	DWORD dwIndex;

 // Find the index of the control...
	for (dwIndex = 0; dwIndex < m_cComponents; dwIndex++)
		if (m_pComponents[dwIndex].hwndControl == hwndControl) break;

	if (dwIndex > m_cComponents)
		return E_INVALIDARG;

 // If it is not already the active item, notify old component it is
 // losing activation, and notify new component that it is gaining...
	EnterCriticalSection(&v_csecThreadSynch);
	if (dwIndex != m_idxActive)
	{
		HWND hwndActive = GetActiveComponentWindow();
		if (hwndActive) 
			SendNotifyMessage(hwndActive, DSO_WM_HOOK_NOTIFY_COMPACTIVE, (WPARAM)FALSE, (LPARAM)0);

		m_idxActive = dwIndex;

		hwndActive = GetActiveComponentWindow();
		if (hwndActive)
			SendNotifyMessage(hwndActive, DSO_WM_HOOK_NOTIFY_COMPACTIVE, (WPARAM)TRUE, (LPARAM)(m_fAppActive));
	}
	LeaveCriticalSection(&v_csecThreadSynch);

	return S_OK;
}

STDMETHODIMP CDsoFrameHookManager::OnComponentNotify(DWORD msg, WPARAM wParam, LPARAM lParam)
{
	TRACE3("CDsoFrameHookManager::OnComponentNotify(%d, %d, %d)\n", msg, wParam, lParam);
	HWND hwndActiveComp = GetActiveComponentWindow();
	
 // The manager needs to keep track of AppActive state...
	if (msg == DSO_WM_HOOK_NOTIFY_APPACTIVATE)
	{
		EnterCriticalSection(&v_csecThreadSynch);
		m_fAppActive = (BOOL)wParam;
		LeaveCriticalSection(&v_csecThreadSynch);
	}

 // Send the message to the active component
    SendNotifyMessage(hwndActiveComp, msg, wParam, lParam);
	return S_OK;
}

STDMETHODIMP_(BOOL) CDsoFrameHookManager::SendNotifyMessage(HWND hwnd, DWORD msg, WPARAM wParam, LPARAM lParam)
{
    BOOL fResult;

 // Check if the caller is on the same thread as the active component. If so we can
 // forward the message directly for processing. If not, we have to post and wait for
 // that thread to process message at a later time...
	if (GetWindowThreadProcessId(hwnd, NULL) == GetCurrentThreadId())
	{ // Notify now...
		fResult = (BOOL)SendMessage(hwnd, msg, wParam, lParam);
	}
	else
	{ // Notify later...
		fResult = PostMessage(hwnd, msg, wParam, lParam);
	}

    return fResult;
}


////////////////////////////////////////////////////////////////////////
// CDsoFrameWindowHook
//
////////////////////////////////////////////////////////////////////////
// CDsoFrameWindowHook::AttachToFrameWindow
//
//  Creates the frame hook for the host window.
//
STDMETHODIMP_(CDsoFrameWindowHook*) CDsoFrameWindowHook::AttachToFrameWindow(HWND hwndParent)
{
	CDsoFrameWindowHook* phook = NULL;
	BOOL fHookSuccess = FALSE;

 // We have to get valid window ...
	if (!IsWindow(hwndParent))
		return NULL;

	EnterCriticalSection(&v_csecThreadSynch);
	phook = CDsoFrameWindowHook::GetHookFromWindow(hwndParent);
	if (phook)
	{ // Already have this window hooked, so just addref the count.
		phook->AddRef();
	}
	else
	{ // Need to create a new hook object for this window...
		if ((phook = new CDsoFrameWindowHook()))
		{
			phook->m_hwndTopLevelHost = hwndParent;
			phook->m_fHostUnicodeWindow = (v_fUnicodeAPI && IsWindowUnicode(hwndParent));

		 // Setup the subclass using Unicode or Ansi methods depending on window 
		 // create state. Note if either property assoc or window subclass fail...
			if (phook->m_fHostUnicodeWindow)
			{
				fHookSuccess = SetPropW(hwndParent, L"DSOFramerWndHook", (HANDLE)phook);
				if (fHookSuccess)
				{
					phook->m_pfnOrigWndProc = (WNDPROC)SetWindowLongW(hwndParent, GWL_WNDPROC, (LONG)(WNDPROC)HostWindowProcHook);
					fHookSuccess = (phook->m_pfnOrigWndProc != NULL);
				}
			}
			else
			{
				fHookSuccess = SetPropA(hwndParent, "DSOFramerWndHook", (HANDLE)phook);
				if (fHookSuccess)
				{
					phook->m_pfnOrigWndProc = (WNDPROC)SetWindowLongA(hwndParent, GWL_WNDPROC, (LONG)(WNDPROC)HostWindowProcHook);
					fHookSuccess = (phook->m_pfnOrigWndProc != NULL);
				}
			}

		 // If we failed the subclass, kill this object and return error...
			if (!fHookSuccess)
			{
				delete phook;
				phook = NULL;
			}
			else phook->m_cHookCount++;
		}
	}
	LeaveCriticalSection(&v_csecThreadSynch);

	return phook;
}

////////////////////////////////////////////////////////////////////////
// CDsoFrameWindowHook::Detach
//
//  Removes the hook from the frame window if this is the last control
//  to reference the hook.
//
STDMETHODIMP CDsoFrameWindowHook::Detach()
{
 // Only need to call this if item is attached.
	EnterCriticalSection(&v_csecThreadSynch);
	if (m_cHookCount)
	{
		if (--m_cHookCount == 0)
		{ // We can remove the hook!
			if (m_fHostUnicodeWindow)
			{
				SetWindowLongW(m_hwndTopLevelHost, GWL_WNDPROC, (LONG)(m_pfnOrigWndProc));
				RemovePropW(m_hwndTopLevelHost, L"DSOFramerWndHook");
			}
			else
			{
				SetWindowLongA(m_hwndTopLevelHost, GWL_WNDPROC, (LONG)(m_pfnOrigWndProc));
				RemovePropA(m_hwndTopLevelHost, "DSOFramerWndHook");
			}
			m_hwndTopLevelHost = NULL;

		// we can kill this object now...
			delete this;
		}
	}
	LeaveCriticalSection(&v_csecThreadSynch);
	return S_OK;
}

////////////////////////////////////////////////////////////////////////
// CDsoFrameWindowHook::GetHookFromWindow
//
//  Returns the hook control associated with this object.
//
STDMETHODIMP_(CDsoFrameWindowHook*) CDsoFrameWindowHook::GetHookFromWindow(HWND hwnd)
{
	CDsoFrameWindowHook* phook = (CDsoFrameWindowHook*)GetPropW(hwnd, L"DSOFramerWndHook");
	if (phook == NULL) phook = (CDsoFrameWindowHook*)GetPropA(hwnd, "DSOFramerWndHook");
	return phook;
}

////////////////////////////////////////////////////////////////////////
// CDsoFrameWindowHook::HostWindowProcHook
//
//  The window proc for the subclassed host window.
//
STDMETHODIMP_(LRESULT) CDsoFrameWindowHook::HostWindowProcHook(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	CDsoFrameWindowHook* phook = CDsoFrameWindowHook::GetHookFromWindow(hwnd);
	if (phook) // We better get our prop for this window!
	{
	 // We handle the important ActiveDocument frame-level events like task 
	 // window activation/deactivation, palette change notfication, and key/mouse
	 // focus notification on host app focus...
		switch (msg)
		{
		case WM_ACTIVATEAPP:
			if (vgGlobalManager)
				vgGlobalManager->OnComponentNotify(DSO_WM_HOOK_NOTIFY_APPACTIVATE, wParam, lParam);
			break;

		case WM_PALETTECHANGED:
			if (vgGlobalManager)
				vgGlobalManager->OnComponentNotify(DSO_WM_HOOK_NOTIFY_PALETTECHANGE, wParam, lParam);
			break;

        case WM_SYNCPAINT:
			if (vgGlobalManager)
				vgGlobalManager->OnComponentNotify(DSO_WM_HOOK_NOTIFY_SYNCPAINT, wParam, lParam);
			break;

		}

	 // After processing, allow calls to fall through to host app. We need to call
	 // the appropriate handler for Unicode or non-Unicode windows...
		if (phook->m_fHostUnicodeWindow)
			return CallWindowProcW(phook->m_pfnOrigWndProc, hwnd, msg, wParam, lParam);
		else
			return CallWindowProcA(phook->m_pfnOrigWndProc, hwnd, msg, wParam, lParam);
	}

 // Should not be reached, but just in case call default proc...
	return (IsWindowUnicode(hwnd) ? 
			DefWindowProcW(hwnd, msg, wParam, lParam) : 
			DefWindowProcA(hwnd, msg, wParam, lParam));
}