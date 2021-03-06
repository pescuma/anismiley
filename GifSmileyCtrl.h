#ifndef GifSmileyCtrl_h__
#define GifSmileyCtrl_h__

// GifSmileyCtrl.h : Declaration of the CGifSmileyCtrl
#pragma once
#include "resource.h"       // main symbols
#include <atlctl.h>

#include "gdiplus.h"
#ifdef _MSC_VER
#include <delayimp.h>
#endif

#include <map>
#include <list>

#pragma comment(lib, "delayimp")
#pragma comment(lib, "gdiplus")

using namespace Gdiplus;
// IGifSmileyCtrl
[
	object,
	uuid(CB64102B-8CE4-4A55-B050-131C435A3A3F),
	dual,
	helpstring("IGifSmileyCtrl Interface"),
	pointer_default(unique)
]
__interface IGifSmileyCtrl : public IDispatch
{
public:
	[propput, bindable, requestedit, id(DISPID_BACKCOLOR)]
	HRESULT BackColor([in]OLE_COLOR clr);
	[propget, bindable, requestedit, id(DISPID_BACKCOLOR)]
	HRESULT BackColor([out,retval]OLE_COLOR* pclr);
	[propget, bindable, requestedit, id(DISPID_HWND)]
	HRESULT HWND([out, retval]long* pHWND);  
    //Methods
    [id(1)]	HRESULT LoadFromFile( [in] BSTR bstrFileName );
    [id(2)]	HRESULT LoadFromFileSized( [in] BSTR bstrFileName, [in] INT nHeight );
	[id(3)]	HRESULT SetHostWindow( [in] long hwndHostWindow, [in] INT nNotyfyMode );
	[id(4)] HRESULT ShowHint( );
	[id(5)] HRESULT OnMsg( [in] HWND hwnd, [in] UINT msg, [in] WPARAM wParam, [in] LPARAM lParam, [out] LRESULT *res );
	[id(6)] HRESULT GetRECT( [out] RECT *rc );
	[id(7)]	HRESULT LoadFlash( [in] BSTR bstrFileName, [in] BSTR bstrFlashVars, [in] INT nWidth, [in] INT nHeight );
};

// ITooltipData
[
	object,
	uuid(58B32D03-1BD2-4840-992E-9AE799FD4ADE),
	//dual,
	helpstring("ITooltipData Interface"),
	pointer_default(unique)
]
__interface ITooltipData : public IUnknown
{
public:
	[id(1)] HRESULT SetTooltip( [in] BSTR bstrHint);
	[id(2)] HRESULT GetTooltip( [out, retval] BSTR * bstrHint);
};

class FlashWrapper;


//CGifSmileyCtrl
[
	coclass,
	control,
	default(IGifSmileyCtrl),
	threading(apartment),
	vi_progid("GifSmiley.GifSmileyCtrl"),
	progid("GifSmiley.GifSmileyCtrl.1"),
	version(1.0),
	uuid("DB35DD77-55E2-4905-8075-EB351BB5CBC1"),
	helpstring("GifSmileyCtrl Class"),
	registration_script("control.rgs")
]

class ATL_NO_VTABLE CGifSmileyCtrl :
	public ITooltipData,
    public CComControl<CGifSmileyCtrl>,
    public IOleObjectImpl<CGifSmileyCtrl>,
    public IOleInPlaceObjectWindowlessImpl<CGifSmileyCtrl>,
    public CStockPropImpl<CGifSmileyCtrl, IGifSmileyCtrl>,
    public IViewObjectExImpl<CGifSmileyCtrl>,
    public IOleControlImpl<CGifSmileyCtrl>
{
public:

DECLARE_OLEMISC_STATUS(OLEMISC_RECOMPOSEONRESIZE |
	OLEMISC_CANTLINKINSIDE |
	OLEMISC_INSIDEOUT |
	OLEMISC_ACTIVATEWHENVISIBLE |
	OLEMISC_SETCLIENTSITEFIRST
)

BEGIN_COM_MAP(CGifSmileyCtrl)
	COM_INTERFACE_ENTRY(IOleInPlaceObjectWindowless)
    COM_INTERFACE_ENTRY(IGifSmileyCtrl)
	COM_INTERFACE_ENTRY(ITooltipData)
	COM_INTERFACE_ENTRY(IOleControl)
	COM_INTERFACE_ENTRY(IOleObject)	
END_COM_MAP()


BEGIN_PROP_MAP(CGifSmileyCtrl)
	PROP_ENTRY("BackColor", DISPID_BACKCOLOR, CLSID_StockColorPage)
END_PROP_MAP()

BEGIN_MSG_MAP(CGifSmileyCtrl)
	CHAIN_MSG_MAP(CComControl<CGifSmileyCtrl>)
	DEFAULT_REFLECTION_HANDLER()
END_MSG_MAP()

    DECLARE_PROTECT_FINAL_CONSTRUCT()

public:

	static HRESULT _InitModule();
	static HRESULT _UninitModule();


	// methods
	CGifSmileyCtrl();
	~CGifSmileyCtrl();

	HRESULT FireViewChange();
	HRESULT OnDrawAdvanced(ATL_DRAWINFO& di);
	HRESULT LoadFromFile( BSTR bstrFileName );
	HRESULT LoadFromFileSized( BSTR bstrFileName, INT nHeight );
	HRESULT LoadFlash( BSTR bstrFileName, BSTR bstrFlashVars, INT nWidth, INT nHeight );
	HRESULT SetHostWindow (long hwndHostWindow, INT nNotyfyMode );
	void	OnBackColorChanged();
	HRESULT FinalConstruct() {	return S_OK;}
	void	FinalRelease(){}
    HRESULT SetTooltip( BSTR bstrHint)  { m_strHint=bstrHint; return S_OK; }
    HRESULT GetTooltip( BSTR * bstrHint) { *bstrHint = m_strHint.AllocSysString(); return S_OK; }
	HRESULT ShowHint( )
	{
		ShowSmileyTooltip();
		return S_OK;
	};
	HRESULT OnMsg(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, LRESULT *res);
	HRESULT GetRECT(RECT *rc);

public:
    //  properties:    
	OLE_COLOR	m_clrBackColor;
//	SIZEL		m_sizeExtent;

private:
	typedef std::map<UINT_PTR, CGifSmileyCtrl*> MAPTIMER;
	typedef std::list<class ImageItem* > LISTIMAGES;
    typedef std::list<class CGifSmileyCtrl* > LISTSMILEYS;
	typedef std::map<HWND, struct HostWindowData> MAPHOSTINFO;

	static int         _refCount;
    static ULONG_PTR   _gdiPlusToken;
    static bool        _gdiPlusFail;
    static int         _gdiPlusRefCount;
    static MAPTIMER    _mapTimers;
    static LISTIMAGES  _listImages;
    static LISTSMILEYS _listSmileys;
    static HWND        _hwndToolTips;
    static CGifSmileyCtrl * _pCurrentTip;
    static UINT_PTR    _tipTimerID;
	static UINT_PTR	   _purgeImageListTimerId;
	static MAPHOSTINFO _mapHostInfo;
	
    ATL::CString m_strHint;
	Bitmap * m_pGifImage;
	UINT   * m_pDelays;
	UINT     m_nFrameCount;
	Size     m_nFrameSize;
	BOOL     m_bTransparent;
	UINT     m_nCurrentFrame;
	UINT_PTR m_nTimerId;
	bool     m_nAnimated;
	HWND	 m_hwndParent;
	RECT	 m_rectPos;
	bool	 m_bPaintValid;
	INT		 m_nNotifyMode;		//0 - none, 1 - empty, 2 - before, 3 - after
    DWORD	 m_dwFlags;
	FlashWrapper *m_pFlash;

private:  
    
	static bool _InitGdiPlus();
 	static void _DestroyGdiPlus();
    static VOID CALLBACK TimerProc( HWND, UINT, UINT_PTR, DWORD );  
    static VOID CALLBACK ToolTipTimerProc( HWND, UINT, UINT_PTR, DWORD );  
	static VOID CALLBACK PurgeImageListTimerProc( HWND, UINT, UINT_PTR, DWORD );  
	static LRESULT CALLBACK HostWindowSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
	static void HandleMsg( IGifSmileyCtrl * iGifSmlCtrl, HWND hwnd, POINT &pt, UINT msg, WPARAM wParam, LPARAM lParam );

	STDMETHOD (Close) (DWORD dwSaveOption)
	{
		StopAnimation();
		return __super::Close(dwSaveOption);
	}
    BOOL        m_bTipShow;
    HWND        ShowSmileyTooltip();
    void        HideSmileyTip();
	BOOL		IsVisible(ATL_DRAWINFO& di);
	BOOL		SelectSmileyClipRgn(HDC hDC, RECT& SmileyPos, HRGN& hOldRegion, HRGN& hNewRegion,  BOOL bTuneBorder);
	void		ResetClip(HDC hDC, HRGN& hOldRgn, HRGN& hNewRgn );

    HRESULT		UpdateSmileyPos( long left, long bottom, long zoom, long flags );
	int			GetObjectPos( HWND hWnd );
	HRESULT		OnDrawSmiley(ATL_DRAWINFO& di, bool bCustom);
	void		DoDrawSmiley(HDC ,  RECT& , int , int , int , int );
	HRESULT		UnloadImage( );
	void		AdvanceFrame();
	void		OnTimer();    
	HRESULT		OnDraw(ATL_DRAWINFO& di);
	void		StopAnimation();
	void		StartAnimation();
};

class ImageItem
{  
private:
    ATL::CString m_strFilename;
    int          m_nHeight;
    Bitmap*      m_pBitmap;
    int          m_nRef;  
    UINT*        m_pFrameDelays;
    int          m_nFrameCount;
    Size         m_FrameSize;

public:
    ImageItem();
    ~ImageItem();
    Bitmap * LoadImageFromFile(ATL::CString& strFilename, int nHeight);
    bool UnloadImage();
    bool IsEqual(ATL::CString& strFilename, int nHeight) 
        {return ( !m_strFilename.CompareNoCase(strFilename) && m_nHeight==nHeight ); }
    bool IsEqual(Bitmap * pBitmap) 
        { return (pBitmap==m_pBitmap); }
    int  GetRefCount() 
        { return m_nRef; }
    UINT * GetFrameDelays() 
        { return m_pFrameDelays;} 
    UINT GetFrameCount()    
        { return m_nFrameCount; }
    Size GetFrameSize() 
        { return m_FrameSize; }
};

struct HostWindowData
{
	HWND hwnd;
	WNDPROC pOldProc;
	HostWindowData(): hwnd(NULL), pOldProc(NULL) {};
};

#endif // GifSmileyCtrl_h__