/***************************************************************************
 * IPPREVIEW.H
 *
 *  DSOFramer: IOleInplacePrintPreview/IOlePreviewCallback
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
#ifndef __IPPREVIEW_H
#define __IPPREVIEW_H

////////////////////////////////////////////////////////////////////////
//
// IOlePreviewCallback
//
// Implemented by host to receive notifaction messages
// from ip object while displaying a print preview.
//
DEFINE_GUID(IID_IOlePreviewCallback, 0xB722BCD5, 0x4E68, 0x101B, 0xA2, 0xBC, 0x00, 0xAA, 0x00, 0x40, 0x47, 0x70);

#undef INTERFACE
#define INTERFACE IOlePreviewCallback
DECLARE_INTERFACE_(IOlePreviewCallback, IUnknown)
{
BEGIN_INTERFACE
#ifndef NO_BASEINTERFACE_FUNCS
    // IUnknown methods
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, LPVOID *ppvObj) PURE;
    STDMETHOD_(ULONG,AddRef)(THIS) PURE;
    STDMETHOD_(ULONG,Release)(THIS) PURE;
#endif
    // IOlePreviewCallback methods
    STDMETHOD(Notify)(THIS_ DWORD wStatus, LONG nLastPage, LPOLESTR pwszPreviewStatus) PURE;
};

#define NOTIFY_FINISHED             1
#define NOTIFY_BUSY                 2
#define NOTIFY_IDLE                 4
#define NOTIFY_DISABLERESIZE        8
#define NOTIFY_QUERYCLOSEPREVIEW    16
#define NOTIFY_FORCECLOSEPREVIEW    32
#define NOTIFY_UIACTIVE             64
#define NOTIFY_UNABLETOPREVIEW      128

////////////////////////////////////////////////////////////////////////
//
// IOleInplacePrintPreview
//
// Implemented by server to start/stop print preview. Hosts should
// call QueryStatus to make sure server is able to enter preview mode
// before calling StartPrintPreview.
//
DEFINE_GUID(IID_IOleInplacePrintPreview, 0xB722BCD4, 0x4E68, 0x101B, 0xA2, 0xBC, 0x00, 0xAA, 0x00, 0x40, 0x47, 0x70);

#undef INTERFACE
#define INTERFACE IOleInplacePrintPreview
DECLARE_INTERFACE_(IOleInplacePrintPreview, IUnknown)
{
BEGIN_INTERFACE
#ifndef NO_BASEINTERFACE_FUNCS
    // IUnknown methods
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, LPVOID *ppvObj) PURE;
    STDMETHOD_(ULONG,AddRef)(THIS) PURE;
    STDMETHOD_(ULONG,Release)(THIS) PURE;
#endif
    // IOleInplacePrintPreview methods
    STDMETHOD(StartPrintPreview)(THIS_ DWORD grfFlags, DVTARGETDEVICE *ptd, IOlePreviewCallback *ppCallback, LONG nFirstPage) PURE;
    STDMETHOD(EndPrintPreview)(THIS_ BOOL fForceClose) PURE;
    STDMETHOD(QueryStatus)(THIS_ void) PURE;
};

#define PREVIEWFLAG_MAYBOTHERUSER           1
#define PREVIEWFLAG_PROMPTUSER	            2
#define PREVIEWFLAG_USERMAYCHANGEPRINTER    4
#define PREVIEWFLAG_RECOMPOSETODEVICE	    8


#endif //__IPPREVIEW_H
