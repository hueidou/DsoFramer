/***************************************************************************
 * DSOFDOCOBJ.H
 *
 * DSOFramer: OLE DocObject Site component (used by the control)
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
#ifndef DS_DSOFDOCOBJ_H 
#define DS_DSOFDOCOBJ_H

////////////////////////////////////////////////////////////////////
// Declarations for Interfaces used in DocObject Containment
//
#include <docobj.h>    // Standard DocObjects (common to all AxDocs)
#include "ipprevw.h"   // PrintPreview (for select Office apps)
#include "rbbinder.h"  // Internet Publishing (for Web Folder write access) 

////////////////////////////////////////////////////////////////////////
// Microsoft Office 97-2003 Document Object GUIDs
//
DEFINE_GUID(CLSID_WORD_DOCUMENT_DOC,    0x00020906, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);
DEFINE_GUID(CLSID_EXCEL_WORKBOOK_XLS,   0x00020820, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);
DEFINE_GUID(CLSID_EXCEL_CHART_XLS,      0x00020821, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);
DEFINE_GUID(CLSID_PPT_PRESENTATION_PPT, 0x64818D10, 0x4F9B, 0x11CF, 0x86, 0xEA, 0x00, 0xAA, 0x00, 0xB9, 0x29, 0xE8);
DEFINE_GUID(CLSID_VISIO_DRAWING_VSD,    0x00021A13, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);
DEFINE_GUID(CLSID_PROJECT_DOCUMENT_MPP, 0x74B78F3A, 0xC8C8, 0x11D1, 0xBE, 0x11, 0x00, 0xC0, 0x4F, 0xB6, 0xFA, 0xF1);
DEFINE_GUID(CLSID_MSHTML_DOCUMENT,      0x25336920, 0x03F9, 0x11CF, 0x8F, 0xD0, 0x00, 0xAA, 0x00, 0x68, 0x6F, 0x13);

////////////////////////////////////////////////////////////////////////
// Microsoft Office 2007 Document GUIDs
//
DEFINE_GUID(CLSID_WORD_DOCUMENT_DOCX,  0xF4754C9B, 0x64F5, 0x4B40, 0x8A, 0xF4, 0x67, 0x97, 0x32, 0xAC, 0x06, 0x07);
DEFINE_GUID(CLSID_WORD_DOCUMENT_DOCM,  0x18A06B6B, 0x2F3F, 0x4E2B, 0xA6, 0x11, 0x52, 0xBE, 0x63, 0x1B, 0x2D, 0x22);
DEFINE_GUID(CLSID_EXCEL_WORKBOOK_XLSX, 0x00020830, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);
DEFINE_GUID(CLSID_EXCEL_WORKBOOK_XLSM, 0x00020832, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);
DEFINE_GUID(CLSID_EXCEL_WORKBOOK_XLSB, 0x00020833, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);
DEFINE_GUID(CLSID_PPT_PRESENTATION_PPTX, 0xCF4F55F4, 0x8F87, 0x4D47, 0x80, 0xBB, 0x58, 0x08, 0x16, 0x4B, 0xB3, 0xF8);
DEFINE_GUID(CLSID_PPT_PRESENTATION_PPTM, 0xDC020317, 0xE6E2, 0x4A62, 0xB9, 0xFA, 0xB3, 0xEF, 0xE1, 0x66, 0x26, 0xF4);

////////////////////////////////////////////////////////////////////
// IDsoDocObjectSite -- {444CA1F7-B405-4002-95C3-A455BC9F4F55}
//
// Implemented by control host for callbacks. Required interface.
//
interface IDsoDocObjectSite : public IServiceProvider
{
    STDMETHOD(GetWindow)(HWND* phWnd) PURE;
    STDMETHOD(GetBorder)(LPRECT prcBorder) PURE;
    STDMETHOD(SetStatusText)(LPCOLESTR pszText) PURE;
    STDMETHOD(GetHostName)(LPWSTR *ppwszHostName) PURE;
    STDMETHOD(SysMenuCommand)(UINT uiCharCode) PURE;
};
DEFINE_GUID(IID_IDsoDocObjectSite, 0x444CA1F7, 0xB405, 0x4002, 0x95, 0xC3, 0xA4, 0x55, 0xBC, 0x9F, 0x4F, 0x55);


////////////////////////////////////////////////////////////////////
// CDsoDocObject -- ActiveDocument Container Site Object
//
//  The CDsoDocObject object handles all the DocObject embedding for the 
//  control and os fairly self-contained. Like the control it has its 
//  own window, but it merely acts as a parent for the embedded object
//  window(s) which it activates. 
//
//  CDsoDocObject works by taking a file (or automation object) and
//  copying out the OLE storage used for its persistent data. It then
//  creates a new embedding based on the data. If a storage is not
//  avaiable, it will attempt to oad the file directly, but the results 
//  are less predictable using this manner since DocObjects are embeddings
//  and not links and this component has limited support for links. As a
//  result, we will attempt to keep our own storage copy in most cases.
//
//  You should note that this approach is different than one taken by the
//  web browser control, which is basically a link container which will
//  try to embed (ip activate) if allowed, but if not it opens the file 
//  externally and keeps the link. If CDsoDocObject cannot embed the object,
//  it returns an error. It will not open the object external.
//  
//  Like the control, this object also uses nested classes for the OLE 
//  interfaces used in the embedding. They are easier to track and easier
//  to debug if a specific interface is over/under released. Again this was
//  a design decision to make the sample easier to break apart, but not required.
//
//  Because the object is not tied to the top-level window, it constructs
//  the OLE merged menu as a set of popup menus which the control then displays
//  in whatever form it wants. You would need to customize this if you used
//  the control in a host and wanted the menus to merge with the actual host
//  menu bar (on the top-level window or form).
// 
class CDsoDocObject : public IUnknown
{
public:
    CDsoDocObject();
    ~CDsoDocObject();

 // Static Create Method (Host Provides Site Interface)
	static STDMETHODIMP_(CDsoDocObject*) CreateInstance(IDsoDocObjectSite* phost);

 // IUnknown Implementation
    STDMETHODIMP QueryInterface(REFIID riid, void** ppv);
    STDMETHODIMP_(ULONG) AddRef(void);
    STDMETHODIMP_(ULONG) Release(void);

 // IOleClientSite Implementation
    BEGIN_INTERFACE_PART(OleClientSite, IOleClientSite)
        STDMETHODIMP SaveObject(void);
        STDMETHODIMP GetMoniker(DWORD dwAssign, DWORD dwWhichMoniker, IMoniker** ppmk);
        STDMETHODIMP GetContainer(IOleContainer** ppContainer);
        STDMETHODIMP ShowObject(void);
        STDMETHODIMP OnShowWindow(BOOL fShow);
        STDMETHODIMP RequestNewObjectLayout(void);
    END_INTERFACE_PART(OleClientSite)

 // IOleInPlaceSite Implementation
    BEGIN_INTERFACE_PART(OleInPlaceSite, IOleInPlaceSite)
        STDMETHODIMP GetWindow(HWND* phwnd);
        STDMETHODIMP ContextSensitiveHelp(BOOL fEnterMode);
        STDMETHODIMP CanInPlaceActivate(void);
        STDMETHODIMP OnInPlaceActivate(void);
        STDMETHODIMP OnUIActivate(void);
        STDMETHODIMP GetWindowContext(IOleInPlaceFrame** ppFrame, IOleInPlaceUIWindow** ppDoc, LPRECT lprcPosRect, LPRECT lprcClipRect, LPOLEINPLACEFRAMEINFO lpFrameInfo);
        STDMETHODIMP Scroll(SIZE sz);
        STDMETHODIMP OnUIDeactivate(BOOL fUndoable);
        STDMETHODIMP OnInPlaceDeactivate(void);
        STDMETHODIMP DiscardUndoState(void);
        STDMETHODIMP DeactivateAndUndo(void);
        STDMETHODIMP OnPosRectChange(LPCRECT lprcPosRect);
    END_INTERFACE_PART(OleInPlaceSite)

 // IOleDocumentSite Implementation
    BEGIN_INTERFACE_PART(OleDocumentSite, IOleDocumentSite)
        STDMETHODIMP ActivateMe(IOleDocumentView* pView);
    END_INTERFACE_PART(OleDocumentSite)

 // IOleInPlaceFrame Implementation
    BEGIN_INTERFACE_PART(OleInPlaceFrame, IOleInPlaceFrame)
        STDMETHODIMP GetWindow(HWND* phWnd);
        STDMETHODIMP ContextSensitiveHelp(BOOL fEnterMode);
        STDMETHODIMP GetBorder(LPRECT prcBorder);
        STDMETHODIMP RequestBorderSpace(LPCBORDERWIDTHS pBW);
        STDMETHODIMP SetBorderSpace(LPCBORDERWIDTHS pBW);
        STDMETHODIMP SetActiveObject(LPOLEINPLACEACTIVEOBJECT pIIPActiveObj, LPCOLESTR pszObj);
        STDMETHODIMP InsertMenus(HMENU hMenu, LPOLEMENUGROUPWIDTHS pMGW);
        STDMETHODIMP SetMenu(HMENU hMenu, HOLEMENU hOLEMenu, HWND hWndObj);
        STDMETHODIMP RemoveMenus(HMENU hMenu);
        STDMETHODIMP SetStatusText(LPCOLESTR pszText);
        STDMETHODIMP EnableModeless(BOOL fEnable);
        STDMETHODIMP TranslateAccelerator(LPMSG pMSG, WORD wID);
    END_INTERFACE_PART(OleInPlaceFrame)

 // IOleCommandTarget  Implementation
    BEGIN_INTERFACE_PART(OleCommandTarget , IOleCommandTarget)
        STDMETHODIMP QueryStatus(const GUID *pguidCmdGroup, ULONG cCmds, OLECMD prgCmds[], OLECMDTEXT *pCmdText);
        STDMETHODIMP Exec(const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANTARG *pvaIn, VARIANTARG *pvaOut);            
    END_INTERFACE_PART(OleCommandTarget)

 // IServiceProvider Implementation
    BEGIN_INTERFACE_PART(ServiceProvider , IServiceProvider)
        STDMETHODIMP QueryService(REFGUID guidService, REFIID riid, void **ppv);
    END_INTERFACE_PART(ServiceProvider)

 // IAuthenticate  Implementation
    BEGIN_INTERFACE_PART(Authenticate , IAuthenticate)
        STDMETHODIMP Authenticate(HWND *phwnd, LPWSTR *pszUsername, LPWSTR *pszPassword);
    END_INTERFACE_PART(Authenticate)

 // IContinueCallback Implementation
    BEGIN_INTERFACE_PART(ContinueCallback , IContinueCallback)
        STDMETHODIMP FContinue(void);
        STDMETHODIMP FContinuePrinting(LONG cPagesPrinted, LONG nCurrentPage, LPOLESTR pwszPrintStatus);
    END_INTERFACE_PART(ContinueCallback)

 // IOlePreviewCallback Implementation
    BEGIN_INTERFACE_PART(PreviewCallback , IOlePreviewCallback)
        STDMETHODIMP Notify(DWORD wStatus, LONG nLastPage, LPOLESTR pwszPreviewStatus);
    END_INTERFACE_PART(PreviewCallback)

 // DocObject Class Methods IDsoDocObjectSite
    STDMETHODIMP  InitializeNewInstance(IDsoDocObjectSite* phost);
    STDMETHODIMP  CreateDocObject(REFCLSID rclsid);
    STDMETHODIMP  CreateDocObject(IStorage *pstg);
    STDMETHODIMP  CreateFromFile(LPWSTR pwszFile, REFCLSID rclsid, LPBIND_OPTS pbndopts);
    STDMETHODIMP  CreateFromURL(LPWSTR pwszUrlFile, REFCLSID rclsid, LPBIND_OPTS pbndopts, LPWSTR pwszUserName, LPWSTR pwszPassword);
    STDMETHODIMP  CreateFromRunningObject(LPUNKNOWN punkObj, LPWSTR pwszObjectName, LPBIND_OPTS pbndopts);
    STDMETHODIMP  IPActivateView();
    STDMETHODIMP  IPDeactivateView();
    STDMETHODIMP  UIActivateView();
    STDMETHODIMP  UIDeactivateView();
	STDMETHODIMP_(BOOL) IsDirty();
    STDMETHODIMP  Save();
    STDMETHODIMP  SaveToFile(LPWSTR pwszFile, BOOL fOverwriteFile);
    STDMETHODIMP  SaveToURL(LPWSTR pwszURL, BOOL fOverwriteFile, LPWSTR pwszUserName, LPWSTR pwszPassword);
    STDMETHODIMP  PrintDocument(LPCWSTR pwszPrinter, LPCWSTR pwszOutput, UINT cCopies, UINT nFrom, UINT nTo, BOOL fPromptUser);
    STDMETHODIMP  StartPrintPreview();
    STDMETHODIMP  ExitPrintPreview(BOOL fForceExit);
    STDMETHODIMP  DoOleCommand(DWORD dwOleCmdId, DWORD dwOptions, VARIANT* vInParam, VARIANT* vInOutParam);
    STDMETHODIMP  Close();

 // Control should notify us on these conditions (so we can pass to IP object)...
    STDMETHODIMP_(void) OnNotifySizeChange(LPRECT prc);
    STDMETHODIMP_(void) OnNotifyAppActivate(BOOL fActive, DWORD dwThreadID);
    STDMETHODIMP_(void) OnNotifyPaletteChanged(HWND hwndPalChg);
    STDMETHODIMP_(void) OnNotifyChangeToolState(BOOL fShowTools);
    STDMETHODIMP_(void) OnNotifyControlFocus(BOOL fGotFocus);
   
    STDMETHODIMP  HrGetDataFromObject(VARIANT *pvtType, VARIANT *pvtOutput);
    STDMETHODIMP  HrSetDataInObject(VARIANT *pvtType, VARIANT *pvtInput, BOOL fMbcsString);

    STDMETHODIMP_(BOOL) GetDocumentTypeAndFileExtension(WCHAR** ppwszFileType, WCHAR** ppwszFileExt);

 // Inline accessors for control to get IP object info...
    inline IOleInPlaceActiveObject*  GetActiveObject(){return m_pipactive;}
    inline IOleObject*               GetOleObject(){return m_pole;}
	inline HWND         GetDocWindow(){return m_hwnd;}
    inline HWND         GetActiveWindow(){return m_hwndUIActiveObj;}
    inline BOOL         IsReadOnly(){return m_fOpenReadOnly;}
    inline BOOL         InPrintPreview(){return ((m_pprtprv != NULL) || (m_fInPptSlideShow));}
    inline HWND         GetMenuHWND(){return m_hwndMenuObj;}
    inline HMENU        GetActiveMenu(){return m_hMenuActive;}
	inline HMENU        GetMergedMenu(){return m_hMenuMerged;}
	inline void         SetMergedMenu(HMENU h){m_hMenuMerged = h;}
    inline LPCWSTR      GetSourceName(){return ((m_pwszWebResource) ? m_pwszWebResource : m_pwszSourceFile);}
    inline LPCWSTR      GetSourceDocName(){return ((m_pwszSourceFile) ? &m_pwszSourceFile[m_idxSourceName] : NULL);}
    inline CLSID*       GetServerCLSID(){return &m_clsidObject;}
    inline BOOL         IsIPActive(){return (m_pipobj != NULL);}

	BOOL IsWordObject()
	{return ((m_clsidObject == CLSID_WORD_DOCUMENT_DOC) || 
			 (m_clsidObject == CLSID_WORD_DOCUMENT_DOCX) || 
			 (m_clsidObject == CLSID_WORD_DOCUMENT_DOCM));
	}
	BOOL IsExcelObject()
	{return ((m_clsidObject == CLSID_EXCEL_WORKBOOK_XLS) || 
			 (m_clsidObject == CLSID_EXCEL_WORKBOOK_XLSX) ||
			 (m_clsidObject == CLSID_EXCEL_WORKBOOK_XLSM) ||
			 (m_clsidObject == CLSID_EXCEL_WORKBOOK_XLSB) ||
			 (m_clsidObject == CLSID_EXCEL_CHART_XLS));
	}
	BOOL IsPPTObject()
	{return ((m_clsidObject == CLSID_PPT_PRESENTATION_PPT) || 
			 (m_clsidObject == CLSID_PPT_PRESENTATION_PPTX) || 
			 (m_clsidObject == CLSID_PPT_PRESENTATION_PPTM));
	}
	BOOL IsVisioObject()
	{return (m_clsidObject == CLSID_VISIO_DRAWING_VSD);}

    STDMETHODIMP SetRunningServerLock(BOOL fLock);

protected:
 // Internal helper methods...
    STDMETHODIMP             InstantiateDocObjectServer(REFCLSID rclsid, IOleObject **ppole);
    STDMETHODIMP             CreateObjectStorage(REFCLSID rclsid);
    STDMETHODIMP             SaveObjectStorage();
    STDMETHODIMP             SaveDocToMoniker(IMoniker *pmk, IBindCtx *pbc, BOOL fKeepLock);
    STDMETHODIMP             SaveDocToFile(LPWSTR pwszFullName, BOOL fKeepLock);
    STDMETHODIMP             ValidateDocObjectServer(REFCLSID rclsid);
    STDMETHODIMP_(BOOL)      ValidateFileExtension(WCHAR* pwszFile, WCHAR** ppwszOut);

    STDMETHODIMP_(void)      OnDraw(DWORD dvAspect, HDC hdcDraw, LPRECT prcBounds, LPRECT prcWBounds, HDC hicTargetDev, BOOL fOptimize);

    STDMETHODIMP             EnsureOleServerRunning(BOOL fLockRunning);
    STDMETHODIMP_(void)      FreeRunningLock();
    STDMETHODIMP_(void)      TurnOffWebToolbar(BOOL fTurnedOff);
    STDMETHODIMP_(void)      ClearMergedMenu();
    STDMETHODIMP_(DWORD)     CalcDocNameIndex(LPCWSTR pwszPath);
    STDMETHODIMP_(void)      CheckForPPTPreviewChange();

 // These functions allow the component to access files in a Web Folder for 
 // write access using the Microsoft Provider for Internet Publishing (MSDAIPP), 
 // which is installed by Office and comes standard in Windows 2000/ME/XP/2003. The
 // provider is not required to use the component, only if you wish to save to 
 // an FPSE or DAV Web Folder (URL). 
    STDMETHODIMP_(IUnknown*) CreateIPPBindResource();
    STDMETHODIMP             IPPDownloadWebResource(LPWSTR pwszURL, LPWSTR pwszFile, IStream** ppstmKeepForSave);
    STDMETHODIMP             IPPUploadWebResource(LPWSTR pwszFile, IStream** ppstmSave, LPWSTR pwszURLSaveTo, BOOL fOverwriteFile);

    static STDMETHODIMP_(LRESULT)  FrameWindowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);

 // The private class variables...
private:
    ULONG                m_cRef;				// Reference count
    HWND                 m_hwnd;                // our window
    HWND                 m_hwndCtl;             // The control's window
    RECT                 m_rcViewRect;          // Viewable area (set by host)
    IDsoDocObjectSite   *m_psiteCtl;            // The control's site interface
    IOleCommandTarget   *m_pcmdCtl;             // IOCT of host (for frame msgs)

    LPWSTR               m_pwszHostName;        // Ole Host Name for container
    LPWSTR               m_pwszSourceFile;      // Path to Source File (on Open)
    IMoniker            *m_pmkSourceFile;       // Moniker to original source file 
    IBindCtx            *m_pbctxSourceFile;     // Bind context used to original source file
    IStorage            *m_pstgSourceFile;      // Original File Storage (if open/save file)
    DWORD                m_idxSourceName;       // Index to doc name in m_pwszSourceFile

    CLSID                m_clsidObject;         // CLSID of the embedded object
    IStorage            *m_pstgroot;            // Root temp storage
    IStorage            *m_pstgfile;            // In-memory file storage
    IStream             *m_pstmview;            // In-memory view info stream

    LPWSTR               m_pwszWebResource;     // The full URL to the web resource
    IStream             *m_pstmWebResource;     // Original Download Stream (if open/save URL)
    IUnknown            *m_punkIPPResource;     // MSDAIPP provider resource (for URL authoring)
    LPWSTR               m_pwszUsername;        // Username and password used by MSDAIPP
    LPWSTR               m_pwszPassword;        // for Authentication (see IAuthenticate)

    IOleObject              *m_pole;            // Embedded OLE Object (OLE)
    IOleInPlaceObject       *m_pipobj;          // The IP object methods (OLE)
    IOleInPlaceActiveObject *m_pipactive;       // The UI Active object methods (OLE)
    IOleDocumentView        *m_pdocv;           // MSO Document View (DocObj)
    IOleCommandTarget       *m_pcmdt;           // MSO Command Target (DocObj)
    IOleInplacePrintPreview *m_pprtprv;         // MSO Print Preview (DocObj)

    HMENU                m_hMenuActive;         // The menu supplied by embedded object
    HMENU                m_hMenuMerged;         // The merged menu (set by control host)
    HOLEMENU             m_holeMenu;            // The OLE Menu Descriptor
    HWND                 m_hwndMenuObj;         // The window for menu commands
    HWND                 m_hwndIPObject;        // IP active object window
    HWND                 m_hwndUIActiveObj;     // UI Active object window
    DWORD                m_dwObjectThreadID;    // Thread Id of UI server
    BORDERWIDTHS         m_bwToolSpace;         // Toolspace...

 // Bitflags (state info)...
    unsigned int         m_fDisplayTools:1;
    unsigned int         m_fDisconnectOnQuit:1;
    unsigned int         m_fAppWindowActive:1;
    unsigned int         m_fOpenReadOnly:1;
    unsigned int         m_fObjectInModalCondition:1;
    unsigned int         m_fObjectIPActive:1;
    unsigned int         m_fObjectUIActive:1;
    unsigned int         m_fObjectActivateComplete:1;
	unsigned int         m_fLockedServerRunning:1;
	unsigned int         m_fLoadedFromAuto:1;
	unsigned int         m_fInClose:1;
	unsigned int         m_fAttemptPptPreview:1;
	unsigned int         m_fInPptSlideShow:1;

};


#endif //DS_DSOFDOCOBJ_H