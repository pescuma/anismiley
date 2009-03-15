// GifSmileyCtrl.cpp : Implementation of CGifSmileyCtrl
#include "stdafx.h"
#include "GifSmileyCtrl.h"
#include "richedit.h"
#include "richole.h"
#include "tom.h"
#include "flash9e.tlh"

#define DEFINE_GUIDXXX(name, l, w1, w2, b1, b2, b3, b4, b5, b6, b7, b8) \
    EXTERN_C const GUID CDECL name \
    = { l, w1, w2, { b1, b2,  b3,  b4,  b5,  b6,  b7,  b8 } }

DEFINE_GUIDXXX(IID_ITextDocument,0x8CC497C0,0xA1DF,0x11CE,0x80,0x98,
               0x00,0xAA,0x00,0x47,0xBE,0x5D);

DEFINE_GUIDXXX(IID_IGifSmileyCtrl, 0xCB64102B, 0x8CE4, 0x4A55,0xB0,0x50,
			   0x13,0x1C,0x43,0x5A,0x3A,0x3F);

// {58B32D03-1BD2-4840-992E-9AE799FD4ADE}
DEFINE_GUIDXXX(IID_ITooltipData, 0x58b32d03, 0x1bd2, 0x4840,  0x99, 0x2e,
			   0x9a, 0xe7, 0x99, 0xfd, 0x4a, 0xde );


//DEFINE_GUIDXXX(IID_IGifSmileyCtrl2, 0x0418FB4B, 0xE1AF, 0x4e32,0x94,0xAD,
//  			   0xFF,0x32,0x2C,0x62,0x2A,0xD3);


#include "m_anismiley.h"

#define EM_GETSCROLLPOS         (WM_USER + 221)
#define EM_GETZOOM				(WM_USER + 224)

#define RELEASE(_X_) if (_X_ != NULL) { _X_->Release(); _X_ = NULL; }


/////////////////////////////////////////////////////////////////////////////
// STATIC 
//////////////////////////////////////////////////////////////////////////
int         CGifSmileyCtrl::_refCount        = 0;
ULONG_PTR   CGifSmileyCtrl::_gdiPlusToken    = 0;
bool        CGifSmileyCtrl::_gdiPlusFail     = true;
int         CGifSmileyCtrl::_gdiPlusRefCount = 0;
HWND        CGifSmileyCtrl::_hwndToolTips    = NULL;
//HWND        CGifSmileyCtrl::_hwndTimerWnd    = NULL;
CGifSmileyCtrl *CGifSmileyCtrl::_pCurrentTip = NULL;
UINT_PTR    CGifSmileyCtrl::_tipTimerID      = 0;
UINT_PTR    CGifSmileyCtrl::_purgeImageListTimerId      = 0;



CGifSmileyCtrl::MAPTIMER CGifSmileyCtrl::_mapTimers;
CGifSmileyCtrl::LISTIMAGES CGifSmileyCtrl::_listImages;
CGifSmileyCtrl::LISTSMILEYS CGifSmileyCtrl::_listSmileys;
CGifSmileyCtrl::MAPHOSTINFO CGifSmileyCtrl::_mapHostInfo;

extern "C" BOOL g_bGdiPlusFail=FALSE;

extern "C" BOOL InitModule()
{    
	return CGifSmileyCtrl::_InitModule();
}

extern "C" HRESULT UninitModule()
{
    return CGifSmileyCtrl::_UninitModule();
}


class FlashWrapper : virtual public IOleClientSite,
					 virtual public IOleInPlaceSiteWindowless,
					 virtual public IOleInPlaceFrame,
					 virtual public IStorage
{
public:
	int refCount;
	SIZE size;
	RECT pos;
	CGifSmileyCtrl *owner;
	IShockwaveFlash *flash;
	IOleObject *flashOleObject;
	IViewObjectEx *flashViewObject;
	IOleInPlaceObjectWindowless *flashInPlaceObjWindowless;

	FlashWrapper(CGifSmileyCtrl *anOwner, const TCHAR *filename, const TCHAR *flashVars,
				 int width = -1, int height = -1)
	{
		refCount = 0;
		flash = NULL;
		flashOleObject = NULL;
		flashViewObject = NULL;
		flashInPlaceObjWindowless = NULL;

		owner = anOwner;

		HRESULT hr;
		long readyState;
		double val;
		
		hr = OleCreate(__uuidof(ShockwaveFlash), IID_IOleObject, OLERENDER_DRAW, 0, 
						(IOleClientSite *) this, (IStorage *) this, (void **) &flashOleObject);
		if (FAILED(hr)) goto err;

		hr = OleSetContainedObject(flashOleObject, TRUE);
		if (FAILED(hr)) goto err;

		hr = flashOleObject->QueryInterface(__uuidof(IShockwaveFlash), (void **) &flash);
		if (FAILED(hr)) goto err;

		hr = flashOleObject->QueryInterface(__uuidof(IViewObjectEx), (void **) &flashViewObject);
		if (FAILED(hr)) goto err;

		hr = flashOleObject->QueryInterface(__uuidof(IOleInPlaceObjectWindowless), (void **) &flashInPlaceObjWindowless);
		if (FAILED(hr)) goto err;

		flash->put_WMode(L"transparent");
		//flash->put_Scale(L"showAll");
		flash->put_ScaleMode(0);
		flash->put_BackgroundColor(0x00000000);
		flash->put_EmbedMovie(TRUE);
		flash->put_Loop(TRUE);
		flash->put_FlashVars(_bstr_t(flashVars));

		hr = flash->LoadMovie(0, _bstr_t(filename));
		if (FAILED(hr)) goto err;

		hr = flash->get_ReadyState(&readyState);
		if (FAILED(hr)) goto err;
		if (readyState != 3 && readyState != 4) goto err;

		if (width > 0)
		{
			size.cx = width;
		}
		else
		{
			hr = flash->TGetPropertyAsNumber(_bstr_t("/"), 8, &val);
			if (FAILED(hr)) 
				size.cx = 200;
			else
				size.cx = (long)(val + 0.5);
		}

		if (height > 0)
		{
			size.cy = height;
		}
		else
		{
			hr = flash->TGetPropertyAsNumber(_bstr_t("/"), 9, &val);
			if (FAILED(hr)) 
				size.cy = 150;
			else
				size.cy = (long)(val + 0.5);
		}

		pos.left = 0;
		pos.top = 0;
		pos.right = size.cx;
		pos.bottom = size.cy;

		flashInPlaceObjWindowless->SetObjectRects(&pos, &pos);

		hr = flash->Play();
		if (FAILED(hr)) goto err;

		hr = flashOleObject->DoVerb(OLEIVERB_SHOW, NULL, (IOleClientSite *) this, 0, NULL, NULL);
		if (FAILED(hr)) goto err;

		return;

err: 
		Destroy();
	}

	virtual ~FlashWrapper()
	{
		Destroy();
	}

	void Destroy()
	{
		if (flashOleObject != NULL)
			flashOleObject->Close(OLECLOSE_NOSAVE);
		RELEASE(flashViewObject)
		RELEASE(flashInPlaceObjWindowless)
		RELEASE(flashOleObject)
		RELEASE(flash)
	}

	BOOL isValid() 
	{
		return flash != NULL;
	}

	void SetPos(const RECT &aPos) 
	{
		if (!isValid())
			return;

		pos = aPos;
		flashInPlaceObjWindowless->SetObjectRects(&pos, &pos);
	}

	void Draw(HDC hdc)
	{
		if (!isValid())
			return;

		OleDraw(flashViewObject, DVASPECT_TRANSPARENT, hdc, &pos);
	}

	LRESULT OnMsg(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
	{
		LRESULT res;
		flashInPlaceObjWindowless->OnWindowMessage(msg, wParam, lParam, &res);
		return res;
	}


	//interface methods

	//IUnknown 
	HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void ** ppvObject)
	{
		if (IsEqualGUID(riid, IID_IUnknown))
			*ppvObject = (void*)(this);
		else if (IsEqualGUID(riid, IID_IOleInPlaceSite))
			*ppvObject = (void*)dynamic_cast<IOleInPlaceSite *>(this);
		else if (IsEqualGUID(riid, IID_IOleInPlaceSiteEx))
			*ppvObject = (void*)dynamic_cast<IOleInPlaceSiteEx *>(this);
		else if (IsEqualGUID(riid, IID_IOleInPlaceSiteWindowless))
			*ppvObject = (void*)dynamic_cast<IOleInPlaceSiteWindowless *>(this);
		else if (IsEqualGUID(riid, IID_IStorage))
			*ppvObject = (void*)dynamic_cast<IStorage *>(this);
		else if (IsEqualGUID(riid, IID_IOleInPlaceFrame))
			*ppvObject = (void*)dynamic_cast<IOleInPlaceFrame *>(this);
		else
			*ppvObject = 0;
		if (!(*ppvObject))
			return E_NOINTERFACE; //if dynamic_cast returned 0
		refCount++;
		return S_OK;
	}
	ULONG STDMETHODCALLTYPE AddRef() { return ++refCount; }
	ULONG STDMETHODCALLTYPE Release() { return --refCount; }


	//IOleClientSite
	virtual HRESULT STDMETHODCALLTYPE SaveObject() { return E_NOTIMPL; }
	virtual HRESULT STDMETHODCALLTYPE GetMoniker(DWORD dwAssign, DWORD dwWhichMoniker, IMoniker ** ppmk) { return E_NOTIMPL; }
	virtual HRESULT STDMETHODCALLTYPE GetContainer(LPOLECONTAINER FAR* ppContainer) { *ppContainer = 0; return E_NOINTERFACE; }
	virtual HRESULT STDMETHODCALLTYPE ShowObject() { return S_OK; }
	virtual HRESULT STDMETHODCALLTYPE OnShowWindow(BOOL fShow) { return E_NOTIMPL; }
	virtual HRESULT STDMETHODCALLTYPE RequestNewObjectLayout() { return E_NOTIMPL; }

	//IOleInPlaceSite
	virtual HRESULT STDMETHODCALLTYPE GetWindow(HWND FAR* lphwnd){ *lphwnd = 0; return S_OK; }
	virtual HRESULT STDMETHODCALLTYPE ContextSensitiveHelp(BOOL fEnterMode) { return E_NOTIMPL; }
	virtual HRESULT STDMETHODCALLTYPE CanInPlaceActivate() { return S_OK; }
	virtual HRESULT STDMETHODCALLTYPE OnInPlaceActivate() { return S_OK; }
	virtual HRESULT STDMETHODCALLTYPE OnUIActivate() { return S_OK; }
	virtual HRESULT STDMETHODCALLTYPE GetWindowContext(LPOLEINPLACEFRAME FAR* lplpFrame,LPOLEINPLACEUIWINDOW FAR* lplpDoc,LPRECT lprcPosRect,LPRECT lprcClipRect,LPOLEINPLACEFRAMEINFO lpFrameInfo)
	{
		*lplpFrame = (LPOLEINPLACEFRAME)this;

		*lplpDoc = 0;

		lpFrameInfo->fMDIApp = FALSE;
		lpFrameInfo->hwndFrame = 0;
		lpFrameInfo->haccel = 0;
		lpFrameInfo->cAccelEntries = 0;
		
		*lprcPosRect = pos;
		*lprcClipRect = pos;
		return S_OK;
	}
	virtual HRESULT STDMETHODCALLTYPE Scroll(SIZE scrollExtent) { return E_NOTIMPL; }
	virtual HRESULT STDMETHODCALLTYPE OnUIDeactivate(BOOL fUndoable) { return S_OK; }
	virtual HRESULT STDMETHODCALLTYPE OnInPlaceDeactivate() { return S_OK; }
	virtual HRESULT STDMETHODCALLTYPE DiscardUndoState() { return E_NOTIMPL; }
	virtual HRESULT STDMETHODCALLTYPE DeactivateAndUndo() { return E_NOTIMPL; }
	virtual HRESULT STDMETHODCALLTYPE OnPosRectChange(LPCRECT lprcPosRect) { return S_OK; }

	//IOleInPlaceSiteEx
	virtual HRESULT STDMETHODCALLTYPE OnInPlaceActivateEx(BOOL __RPC_FAR *pfNoRedraw, DWORD dwFlags) { if (pfNoRedraw) *pfNoRedraw = FALSE; return S_OK; }
	virtual HRESULT STDMETHODCALLTYPE OnInPlaceDeactivateEx(BOOL fNoRedraw) { return S_FALSE; }
	virtual HRESULT STDMETHODCALLTYPE RequestUIActivate(void) { return S_FALSE; }

	//IOleInPlaceSiteWindowless
    virtual HRESULT STDMETHODCALLTYPE CanWindowlessActivate(void) { return S_OK; }
    virtual HRESULT STDMETHODCALLTYPE GetCapture(void) { return S_FALSE; }
    virtual HRESULT STDMETHODCALLTYPE SetCapture(BOOL fCapture) { return S_FALSE; }
    virtual HRESULT STDMETHODCALLTYPE GetFocus(void) { return S_OK; }
    virtual HRESULT STDMETHODCALLTYPE SetFocus(BOOL fFocus) { return S_OK; }
    virtual HRESULT STDMETHODCALLTYPE GetDC(LPCRECT pRect, DWORD grfFlags, HDC __RPC_FAR *phDC) { return S_FALSE; }
    virtual HRESULT STDMETHODCALLTYPE ReleaseDC(HDC hDC) { return S_FALSE; }
    virtual HRESULT STDMETHODCALLTYPE InvalidateRect(LPCRECT pRect,BOOL fErase)
	{
		owner->FireViewChange();
		return S_OK;
	}
    virtual HRESULT STDMETHODCALLTYPE InvalidateRgn(HRGN hRGN, BOOL fErase) { return S_OK; }
    virtual HRESULT STDMETHODCALLTYPE ScrollRect(INT dx, INT dy, LPCRECT pRectScroll, LPCRECT pRectClip) { return E_NOTIMPL; }
    virtual HRESULT STDMETHODCALLTYPE AdjustRect(LPRECT prc) { return S_FALSE; }
    virtual HRESULT STDMETHODCALLTYPE OnDefWindowMessage(UINT msg, WPARAM wParam, LPARAM lParam, LRESULT __RPC_FAR *plResult) 
	{
		return S_FALSE; 
	}

	//IOleInPlaceFrame
	virtual HRESULT STDMETHODCALLTYPE GetBorder(LPRECT lprectBorder) { return E_NOTIMPL; }
	virtual HRESULT STDMETHODCALLTYPE RequestBorderSpace(LPCBORDERWIDTHS pborderwidths) { return E_NOTIMPL; }
	virtual HRESULT STDMETHODCALLTYPE SetBorderSpace(LPCBORDERWIDTHS pborderwidths) { return E_NOTIMPL; }
	virtual HRESULT STDMETHODCALLTYPE SetActiveObject(IOleInPlaceActiveObject *pActiveObject, LPCOLESTR pszObjName) { return S_OK; }
	virtual HRESULT STDMETHODCALLTYPE InsertMenus(HMENU hmenuShared, LPOLEMENUGROUPWIDTHS lpMenuWidths) { return E_NOTIMPL; }
	virtual HRESULT STDMETHODCALLTYPE SetMenu(HMENU hmenuShared, HOLEMENU holemenu, HWND hwndActiveObject) { return S_OK; }
	virtual HRESULT STDMETHODCALLTYPE RemoveMenus(HMENU hmenuShared) { return E_NOTIMPL; }
	virtual HRESULT STDMETHODCALLTYPE SetStatusText(LPCOLESTR pszStatusText) { return S_OK; }
	virtual HRESULT STDMETHODCALLTYPE EnableModeless(BOOL fEnable) { return S_OK; }
	virtual HRESULT STDMETHODCALLTYPE TranslateAccelerator(LPMSG lpmsg, WORD wID) { return E_NOTIMPL; }

	//IStorage
	virtual HRESULT STDMETHODCALLTYPE CreateStream(const WCHAR *pwcsName, DWORD grfMode, DWORD reserved1, DWORD reserved2, IStream **ppstm) { return E_NOTIMPL; }
	virtual HRESULT STDMETHODCALLTYPE OpenStream(const WCHAR * pwcsName, void *reserved1, DWORD grfMode, DWORD reserved2, IStream **ppstm) { return E_NOTIMPL; }
	virtual HRESULT STDMETHODCALLTYPE CreateStorage(const WCHAR *pwcsName, DWORD grfMode, DWORD reserved1, DWORD reserved2, IStorage **ppstg) { return E_NOTIMPL; }
	virtual HRESULT STDMETHODCALLTYPE OpenStorage(const WCHAR * pwcsName, IStorage * pstgPriority, DWORD grfMode, SNB snbExclude, DWORD reserved, IStorage **ppstg) { return E_NOTIMPL; }
	virtual HRESULT STDMETHODCALLTYPE CopyTo(DWORD ciidExclude, IID const *rgiidExclude, SNB snbExclude,IStorage *pstgDest) { return E_NOTIMPL; }
	virtual HRESULT STDMETHODCALLTYPE MoveElementTo(const OLECHAR *pwcsName,IStorage * pstgDest, const OLECHAR *pwcsNewName, DWORD grfFlags) { return E_NOTIMPL; }
	virtual HRESULT STDMETHODCALLTYPE Commit(DWORD grfCommitFlags) { return E_NOTIMPL; }
	virtual HRESULT STDMETHODCALLTYPE Revert() { return E_NOTIMPL; }
	virtual HRESULT STDMETHODCALLTYPE EnumElements(DWORD reserved1, void * reserved2, DWORD reserved3, IEnumSTATSTG ** ppenum) { return E_NOTIMPL; }
	virtual HRESULT STDMETHODCALLTYPE DestroyElement(const OLECHAR *pwcsName) { return E_NOTIMPL; }
	virtual HRESULT STDMETHODCALLTYPE RenameElement(const WCHAR *pwcsOldName, const WCHAR *pwcsNewName) { return E_NOTIMPL; }
	virtual HRESULT STDMETHODCALLTYPE SetElementTimes(const WCHAR *pwcsName, FILETIME const *pctime, FILETIME const *patime, FILETIME const *pmtime) { return E_NOTIMPL; }
	virtual HRESULT STDMETHODCALLTYPE SetClass(REFCLSID clsid) { return S_OK; }
	virtual HRESULT STDMETHODCALLTYPE SetStateBits(DWORD grfStateBits, DWORD grfMask) { return E_NOTIMPL; }
	virtual HRESULT STDMETHODCALLTYPE Stat(STATSTG * pstatstg, DWORD grfStatFlag) { return E_NOTIMPL; }


};


bool CGifSmileyCtrl::_InitGdiPlus(void)
{
    Gdiplus::GdiplusStartupInput gdiplusStartupInput;
    if (!g_bGdiPlusFail )
        CGifSmileyCtrl::_gdiPlusFail=false;
    
    __try 
    {
        if (_gdiPlusToken == 0 && !_gdiPlusFail)
        {
            Gdiplus::GdiplusStartup(&_gdiPlusToken, &gdiplusStartupInput, NULL);
        }
    }
    __except ( EXCEPTION_EXECUTE_HANDLER ) 
    {
        _gdiPlusFail = true;
        g_bGdiPlusFail=TRUE;
    }
    if (!_gdiPlusFail) _gdiPlusRefCount++;
    return !_gdiPlusFail;
}

void CGifSmileyCtrl::_DestroyGdiPlus(void)
{
    _gdiPlusRefCount--;
    if (_gdiPlusToken != 0 && _gdiPlusRefCount!=0)
    {
        Gdiplus::GdiplusShutdown(_gdiPlusToken);
        __FUnloadDelayLoadedDLL2("gdiplus.dll");
        _gdiPlusToken = 0;
    }
}


HRESULT  CGifSmileyCtrl::_InitModule()
{    
	_InitGdiPlus();
    return S_OK;

}
HRESULT  CGifSmileyCtrl::_UninitModule()
{
    if (_hwndToolTips) ::DestroyWindow(_hwndToolTips);
    _hwndToolTips=NULL;

    _mapTimers.clear();

	LISTIMAGES::iterator it=_listImages.begin();   
	while (it!=_listImages.end())
	{
		delete (*it);
		it++;
	}
	_listImages.clear();
	_DestroyGdiPlus();
    
    return S_OK;
}
// CGifSmileyCtrl
CGifSmileyCtrl::CGifSmileyCtrl()
   :m_pGifImage( NULL ),
    m_pDelays( NULL ),
    m_nFrameCount( 0 ),
    m_nFrameSize( 0, 0 ),
    m_bTransparent( true ),
    m_nCurrentFrame( 0 ),
    m_nTimerId( 0 ),
    m_hwndParent( NULL ),
    m_bPaintValid( false ),
    m_bTipShow( FALSE ),
	m_pFlash( NULL )
{
    memset(&m_rectPos,0,sizeof(RECT));
    if (_refCount == 0)
    {
        _InitModule();
    }
    _refCount++;
    _listSmileys.push_back(this);
    //if (_listSmileys.size()==1)
    //    _tipTimerID=::SetTimer(NULL, (UINT_PTR) this, 500, ToolTipTimerProc);
}
CGifSmileyCtrl::~CGifSmileyCtrl()
{
	// Desubclass
	{
		MAPHOSTINFO::iterator it = _mapHostInfo.find(m_hwndParent);
		if (it != _mapHostInfo.end())
		{
			HostWindowData hwdat = it->second;
			if (hwdat.pOldProc != NULL)
				::SetWindowLong(m_hwndParent, GWL_WNDPROC, (LONG) hwdat.pOldProc);
			_mapHostInfo.erase(it);
		}
	}

    UnloadImage();
 
    LISTSMILEYS::iterator it=_listSmileys.begin();
    for (it=_listSmileys.begin(); it!=_listSmileys.end(); it++)
    {
        if (*it!=this) continue;
        _listSmileys.erase(it); 
        break;
    }
   
    --_refCount;
    if (_refCount == 0)  
    {
        if (_listSmileys.size()==0) ::KillTimer(NULL, _tipTimerID);
        _UninitModule();
    }
}

HRESULT CGifSmileyCtrl::UnloadImage( )
{
    if ( m_nTimerId ) 
        ::KillTimer( NULL, m_nTimerId );
	if (m_pFlash)
		delete m_pFlash;
    if (m_pGifImage)
    {
        LISTIMAGES::iterator it=_listImages.begin();
        while (it!=_listImages.end())
        {
            if ((*it) && (*it)->IsEqual(m_pGifImage))
            {
                if ( (*it)->UnloadImage() && _purgeImageListTimerId==0 ) 
				{
					_purgeImageListTimerId=::SetTimer(NULL,0, 10000, PurgeImageListTimerProc);
				}
					//_listImages.erase(it);
                break;
            }
            it++;
        }
    }
    m_nTimerId=NULL;   
    m_pGifImage=NULL;
	m_pFlash=NULL;
    m_pDelays=NULL;
    m_nFrameCount=0;
    m_nCurrentFrame=0;
    m_nFrameSize=Size(0,0);
    m_sizeExtent.cx=0;
    m_sizeExtent.cy=0;

    return S_OK;
}
HRESULT CGifSmileyCtrl::LoadFromFile( BSTR bstrFileName )
{
    return CGifSmileyCtrl::LoadFromFileSized( bstrFileName, (INT) 0 );
}

HRESULT CGifSmileyCtrl::LoadFlash( BSTR bstrFileName, BSTR bstrFlashVars, INT nWidth, INT nHeight )
{
    if ( _gdiPlusFail ) return E_FAIL;
    UnloadImage();

	if (nWidth <= 0 || nHeight <= 0)
		nWidth = nHeight = 0;

    ATL::CString sFilename(bstrFileName);
    ATL::CString sFlashVars(bstrFlashVars);

	m_pFlash = new FlashWrapper(this, sFilename.GetBuffer(), sFlashVars.GetBuffer(), nWidth, nHeight);
	if (!m_pFlash->isValid())
	{
		delete m_pFlash;
		m_pFlash = NULL;
		return E_FAIL;
	}

	m_nFrameSize.Width = m_pFlash->size.cx;
	m_nFrameSize.Height = m_pFlash->size.cy;
	m_nFrameCount = 1;
    SIZEL size;
    size.cx=m_nFrameSize.Width+2;
    size.cy=m_nFrameSize.Height;
    AtlPixelToHiMetric(&size,&m_sizeExtent);
	FireViewChange();
    return S_OK;
}

HRESULT CGifSmileyCtrl::LoadFromFileSized( BSTR bstrFileName, INT nHeight )
{
    if ( _gdiPlusFail ) return E_FAIL;
    UnloadImage();

    ATL::CString sFilename(bstrFileName);

    ImageItem * foundImage=NULL;
	LISTIMAGES::iterator it=_listImages.begin();
	while (it!=_listImages.end())
	{
		if ((*it) && (*it)->IsEqual(sFilename,nHeight))
		{
			foundImage=(*it);
			break;
		}
		it++;
	}
    if (!foundImage)
    {
        foundImage=new ImageItem;
        _listImages.push_back(foundImage);
    }

    if (foundImage)
    {
        m_pGifImage=foundImage->LoadImageFromFile(sFilename,nHeight);
        m_pDelays=foundImage->GetFrameDelays();
        m_nFrameCount=foundImage->GetFrameCount();
        m_nFrameSize=foundImage->GetFrameSize();
        if (m_nFrameSize.Width==0 || m_nFrameSize.Height==0)
        {
            UnloadImage();
            return E_FAIL;
        }
        SIZEL size;
        size.cx=m_nFrameSize.Width+2;
        size.cy=m_nFrameSize.Height;
        AtlPixelToHiMetric(&size,&m_sizeExtent);
		FireViewChange();
        return S_OK;
    }
    return E_FAIL;    
}
void  CGifSmileyCtrl::HideSmileyTip( )
{
	if (this==NULL) return;
    if (!m_bTipShow) return;
    if (_pCurrentTip!=this) return;
    m_bTipShow=FALSE;
    if (_hwndToolTips)
        ::DestroyWindow(_hwndToolTips);
    _hwndToolTips=NULL;
    _pCurrentTip=NULL;
    
}
HWND CGifSmileyCtrl::ShowSmileyTooltip( )
{
    TOOLINFO ti;
	{
		POINT pt;
		GetCursorPos(&pt);
		::ScreenToClient(m_hwndParent,&pt);
		if (!PtInRect(&m_rectPos,pt))
		{
			if (_pCurrentTip)
				_pCurrentTip->HideSmileyTip();
			_pCurrentTip=NULL;
			return NULL;
		}
	}

    if (m_bTipShow && _pCurrentTip==this)
        return _hwndToolTips;
    //check tip
    if (_hwndToolTips && _pCurrentTip && _pCurrentTip!=this)
    {
        _pCurrentTip->HideSmileyTip();
    }
	//check if mouse other this smiley

    _pCurrentTip=this;
    if (!_hwndToolTips) //::DestroyWindow(_hwndToolTips);
    {
        //  _hwndToolTips = CreateWindowEx(WS_EX_TOPMOST, TOOLTIPS_CLASS, TEXT(""), WS_POPUP, 0, 0, 0, 0, NULL, NULL, GetModuleHandle(NULL), NULL);

        _hwndToolTips = CreateWindowEx(0, TOOLTIPS_CLASS, NULL,
            WS_POPUP | TTS_NOPREFIX | TTS_ALWAYSTIP,
            CW_USEDEFAULT, CW_USEDEFAULT,
            CW_USEDEFAULT, CW_USEDEFAULT,
            m_hwndParent, NULL, GetModuleHandle(NULL),
            NULL);

        ::SetWindowPos(_hwndToolTips, HWND_TOPMOST,0, 0, 0, 0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
    }

    ZeroMemory(&ti, sizeof(ti));
    ti.cbSize = sizeof(ti);
    ti.uFlags = TTF_IDISHWND;
    ti.hwnd = m_hwndParent;
    ti.uId = (UINT)m_hwndParent;
    if (::SendMessage(_hwndToolTips, TTM_GETTOOLINFO, 0, (LPARAM)&ti)) {
        ::SendMessage(_hwndToolTips, TTM_DELTOOL, 0, (LPARAM)&ti);
    }
    ti.uFlags = TTF_IDISHWND|TTF_SUBCLASS;
    ti.rect=m_rectPos;
    ti.uId = (UINT)m_hwndParent;
    ti.lpszText=(TCHAR*)m_strHint.GetString();
    ::SendMessage(_hwndToolTips,TTM_ADDTOOL,0,(LPARAM)&ti);
	//::SendMessage(_hwndToolTips, TTM_ACTIVATE, TRUE, 0);
    ::SendMessage(_hwndToolTips,TTM_SETDELAYTIME,TTDT_INITIAL,(LPARAM) MAKELONG(100, 0));
    //::SendMessage(_hwndToolTips,TTM_POPUP,0,(LPARAM)&ti);
    m_bTipShow=TRUE;

    return _hwndToolTips;
}

HRESULT CGifSmileyCtrl::UpdateSmileyPos( long left, long bottom, long zoom, long flags )
{
	int zoomnom=HIWORD(zoom);
	int zoomden=LOWORD(zoom);
	if (!zoomden) zoomden=1;
	m_rectPos.bottom=bottom;
	m_rectPos.left=left;
	m_rectPos.top=bottom-(m_nFrameSize.Height*zoomnom/zoomden);
	m_rectPos.right=left+(m_nFrameSize.Width*zoomnom/zoomden)+2;
	if (flags==1)
		m_bPaintValid=false;
	else
		m_bPaintValid=true;
	m_dwFlags=flags;
	return S_OK;
} 

BOOL CGifSmileyCtrl::SelectSmileyClipRgn(HDC hDC, RECT& SmileyPos, HRGN& hOldRegion,  HRGN& hNewRegion, BOOL bTuneBorder)
{
	hNewRegion=NULL;
	hOldRegion=NULL;

	if (!bTuneBorder)
	{
		hNewRegion=CreateRectRgn(SmileyPos.left, SmileyPos.top, SmileyPos.right, SmileyPos.bottom);
	}
	else if (m_hwndParent)
	{
		RECT rcParent;
		RECT rcVis;
		RECT wrc;
		POINT pt={0};

		::ClientToScreen(m_hwndParent, &pt);
		::GetClientRect(m_hwndParent, &rcParent);
		::GetWindowRect(m_hwndParent,&wrc);
		OffsetRect(&rcParent, pt.x-wrc.left, pt.y-wrc.top);
		
		IntersectRect(&rcVis, &rcParent, &SmileyPos);
		if (IsRectEmpty(&rcVis)) return FALSE;
		hNewRegion=CreateRectRgn(rcVis.left,rcVis.top,rcVis.right,rcVis.bottom);
	}
	else return FALSE;
	
	if (GetClipRgn(hDC,hOldRegion)!=1)
	{
		DeleteObject(hOldRegion);
		hOldRegion=NULL;
	}

	SelectClipRgn(hDC, hNewRegion);
	return TRUE;
}
void CGifSmileyCtrl::ResetClip(HDC hDC, HRGN& hOldRgn, HRGN& hNewRgn ) 
{
	if ( hOldRgn) SelectClipRgn( hDC, hOldRgn );
	if ( hOldRgn!=NULL ) DeleteObject( hOldRgn );
	if ( hNewRgn!=NULL ) DeleteObject( hNewRgn );
}

HRESULT CGifSmileyCtrl::OnDrawAdvanced(ATL_DRAWINFO& di)
{
    RECT& rc = *(RECT*)di.prcBounds;
	HRGN hOldRgn, hNewRgn;
	SelectSmileyClipRgn(di.hdcDraw, rc, hOldRgn, hNewRgn, FALSE);
    
	OnDraw(di);
	
	ResetClip(di.hdcDraw, hOldRgn, hNewRgn);
    
	return S_OK;
}

HRESULT CGifSmileyCtrl::OnDraw(ATL_DRAWINFO& di)
{
	MAPHOSTINFO::iterator it = _mapHostInfo.find(m_hwndParent);
	if (it != _mapHostInfo.end())
	{
		HostWindowData &hwd = it->second;
		if (hwd.hwnd != NULL && hwd.pOldProc == NULL)
			hwd.pOldProc = (WNDPROC)::SetWindowLong(m_hwndParent, GWL_WNDPROC, (LONG) HostWindowSubclassProc);
	}

	return OnDrawSmiley(di,false);
}

HRESULT CGifSmileyCtrl::SetHostWindow( long hwndHostWindow, INT nNotyfyMode )
{
	m_hwndParent=(HWND)hwndHostWindow;
	m_nNotifyMode=nNotyfyMode;

	// Subclass parent to show tooltip
	WNDPROC oldProc = (WNDPROC)::GetWindowLong(m_hwndParent, GWL_WNDPROC);
	HostWindowData hwd = _mapHostInfo[m_hwndParent];
	if (hwd.hwnd==NULL)
	{
		hwd.hwnd=m_hwndParent;
		hwd.pOldProc=NULL;
//		hwd.pOldProc=(WNDPROC)::SetWindowLong(m_hwndParent, GWL_WNDPROC, (LONG) HostWindowSubclassProc);
		_mapHostInfo[m_hwndParent]=hwd;		
	}	
	return S_OK;

}

void CGifSmileyCtrl::DoDrawSmiley(HDC hdc, RECT& rc, int ExtentWidth, int ExtentHeight, int frameWidth, int frameHeight)
{
	Rect rect(0,0,ExtentWidth, ExtentHeight);    
	Bitmap bmp(ExtentWidth, ExtentHeight, PixelFormat32bppARGB );
	Graphics * mem=Graphics::FromImage(&bmp);        
	if (!m_bTransparent)
	{
		COLORREF col=(COLORREF)(m_clrBackColor);
		SolidBrush brush(Color(GetRValue(col),GetGValue(col),GetBValue(col)));
		mem->FillRectangle(&brush, rect);
	}

	if (m_pFlash)
	{
		RECT r = { 0, 0, ExtentWidth, ExtentHeight };
		m_pFlash->SetPos(r);

		HDC memHDC = mem->GetHDC();
		m_pFlash->Draw(memHDC);
		mem->ReleaseHDC(memHDC);
	}
	else
	{
		mem->DrawImage(m_pGifImage, rect, m_nCurrentFrame*frameWidth, 0, frameWidth, frameHeight, UnitPixel);
	}

	Graphics g(hdc);
	g.DrawImage(&bmp, rc.left, rc.top, ExtentWidth, ExtentHeight);
	delete mem;
}
BOOL CGifSmileyCtrl::IsVisible(ATL_DRAWINFO& di)
{
	//TO DO: check if smiley portion of parent window is really visible in case of obscured window
	//if (!::IsWindowVisible(m_hwndParent)) 
	//	return FALSE;
	if (!::RectVisible(di.hdcDraw,(RECT*)di.prcBounds)) 
		return FALSE;
	return TRUE;
}
HRESULT CGifSmileyCtrl::OnDrawSmiley(ATL_DRAWINFO& di, bool bCustom=false)
{
    USES_CONVERSION;
    if (di.dwDrawAspect != DVASPECT_CONTENT)  return E_FAIL;
    if (!m_pGifImage && !m_pFlash)  return E_FAIL;
	if ( bCustom&&!IsVisible(di) ) return S_OK;
	RECT& rc = *(RECT*)di.prcBounds;

	HRGN hOldRgn, hNewRgn;

	if (!IsRectEmpty(&m_rectPos))
	{   //strange workaround for drawing zoom out smileys (look MS calculate it one pix larger than exactly)
		if (rc.bottom-rc.top-1 == m_rectPos.bottom-m_rectPos.top 
			&& rc.right-rc.left== m_rectPos.right-m_rectPos.left)
			rc.top+=1;
	}

	if ( bCustom )SelectSmileyClipRgn(di.hdcDraw, rc, hOldRgn, hNewRgn, TRUE);
	
	InflateRect(&rc,-1,0); //border offset to fix blinked cursor painting
    if ( (m_dwFlags&REO_INVERTEDSELECT) == 0 || !bCustom || m_bTransparent)
        DoDrawSmiley(di.hdcDraw, rc, rc.right-rc.left,rc.bottom-rc.top, m_nFrameSize.Width, m_nFrameSize.Height);
    else
    {
        Bitmap bmp(rc.right-rc.left,rc.bottom-rc.top, PixelFormat32bppARGB);
        Graphics g(&bmp);
        COLORREF col=(COLORREF)(m_clrBackColor);
        SolidBrush brush(Color(GetRValue(col),GetGValue(col),GetBValue(col)));
        g.FillRectangle( &brush, 0 ,0, rc.right-rc.left, rc.bottom-rc.top);
        HDC hdc=g.GetHDC();
        RECT mrc={0};
        mrc.right=rc.right-rc.left;
        mrc.bottom=rc.bottom-rc.top;
        DoDrawSmiley(hdc, mrc, mrc.right-mrc.left,mrc.bottom-mrc.top, m_nFrameSize.Width, m_nFrameSize.Height);
        InvertRect(hdc, &mrc);
        BitBlt(di.hdcDraw, rc.left, rc.top, rc.right-rc.left, rc.bottom-rc.top, hdc, 0, 0, SRCCOPY );
        g.ReleaseHDC(hdc);       
    }
	if ((m_dwFlags&REO_SELECTED) == REO_SELECTED && bCustom)
	{
		//Draw frame around
		HBRUSH oldBrush=(HBRUSH)SelectObject(di.hdcDraw, GetStockObject(NULL_BRUSH)); 
		HPEN oldPen=(HPEN)SelectObject(di.hdcDraw, GetStockObject(BLACK_PEN));
		::Rectangle(di.hdcDraw, rc.left, rc.top, rc.right, rc.bottom );
		SelectObject(di.hdcDraw, oldBrush);
		SelectObject(di.hdcDraw, oldPen);
	}
    AdvanceFrame();
	if (!bCustom) 
        m_bPaintValid=false;
	ResetClip(di.hdcDraw, hOldRgn, hNewRgn);

    return S_OK;
}


void CGifSmileyCtrl::OnBackColorChanged()
{
    if ( (m_clrBackColor&0xFF000000)==0xFF000000) 
        m_bTransparent=true;
    else
        m_bTransparent=false;
    FireViewChange();
}

int CGifSmileyCtrl::GetObjectPos( HWND hWnd )
{
    if ( !hWnd ) return 0;
    IRichEditOle * ole=NULL;
    if (!::SendMessage(hWnd, EM_GETOLEINTERFACE, 0, (LPARAM)&ole)) return 0;
    int nCount=ole->GetObjectCount();
    for (int i=nCount-1; i>=0; i--)
    {
        REOBJECT reobj={0};
        reobj.cbStruct=sizeof(REOBJECT);
        ole->GetObject(i,&reobj,REO_GETOBJ_NO_INTERFACES);
        if (reobj.clsid==__uuidof(CGifSmileyCtrl) && ((CGifSmileyCtrl*)reobj.dwUser)==this)
        {
            long left, bottom;
            HRESULT res;
            ITextDocument * iDoc=NULL;
            ITextRange *iRange=NULL;
            
            ole->QueryInterface(IID_ITextDocument,(void**)&iDoc);
            if (!iDoc) break;

            iDoc->Range(reobj.cp, reobj.cp, &iRange);
			if (reobj.dwFlags&REO_BELOWBASELINE)
				res=iRange->GetPoint(TA_BOTTOM|TA_LEFT, &left, &bottom);
			else
				res=iRange->GetPoint(TA_BASELINE|TA_LEFT, &left, &bottom);

            iRange->Release();
            iDoc->Release();

            if (res!=S_OK) //object is out of screen let do normal fireview change
            {
                UpdateSmileyPos(-100, -100, 1, 1);
                break;
            }

            DWORD nom=1, den=1;
            int zoom=MAKELONG(den,nom);
            if (::SendMessage(hWnd, EM_GETZOOM, (WPARAM)&nom, (LPARAM)&den))   
                zoom=(den && nom) ? MAKELONG(den,nom) : zoom;

            RECT windowOffset={0};
            ::GetWindowRect(hWnd,&windowOffset);
            CHARRANGE chr={0};
            ::SendMessage(hWnd,EM_EXGETSEL, 0, (LPARAM)&chr);
            DWORD flag=0;//(reobj.dwFlags & (REO_INVERTEDSELECT|REO_SELECTED*/));
            if  ( chr.cpMin!=chr.cpMax &&
                 (  (reobj.cp<=chr.cpMax && reobj.cp>=chr.cpMin) ||
                    (reobj.cp>=chr.cpMax && reobj.cp<=chr.cpMin)  )  )
			{
                 flag|=REO_SELECTED;
				 if((chr.cpMax-chr.cpMin)!= 1)
					 flag|=REO_INVERTEDSELECT;
			}
            UpdateSmileyPos(left-windowOffset.left, bottom-windowOffset.top, zoom, flag);

            break;            
        }
    }
    ole->Release();
    return 0;
}

inline HRESULT CGifSmileyCtrl::FireViewChange()
{
    if (m_bInPlaceActive)
    {
        // Active
        if (m_hWndCD != NULL)
            ::InvalidateRect(m_hWndCD, NULL, m_bTransparent); // Window based
        else if (m_bWndLess && m_spInPlaceSite != NULL)
            m_spInPlaceSite->InvalidateRect(NULL, m_bTransparent); // Windowless
    }
    else if (!m_hwndParent)
        SendOnViewChange(DVASPECT_CONTENT);
    else
    {
        GetObjectPos( m_hwndParent );

        COLORREF oldColor=m_clrBackColor;
        BOOL     fOldTransparent= m_bTransparent;
        HWND     hwndOld=m_hwndParent;
        HWND     hwndGrand=::GetParent(m_hwndParent);
        FVCNDATA_NMHDR nmhdr={0};
        
        nmhdr.cbSize=sizeof(FVCNDATA_NMHDR);
        nmhdr.hwndFrom=hwndOld;
        nmhdr.code=NM_FIREVIEWCHANGE;
        nmhdr.clrBackground=oldColor;
        nmhdr.fTransparent=fOldTransparent;
        nmhdr.bAction= m_bPaintValid ? (fOldTransparent ? FVCA_INVALIDATE : FVCA_DRAW ) : FVCA_SENDVIEWCHANGE;
        nmhdr.rcRect=m_rectPos;
        //SendPrePaintEvent
        nmhdr.bEvent=FVCN_PREFIRE;
        if (hwndGrand) ::SendMessage(hwndGrand,WM_NOTIFY,(WPARAM)nmhdr.hwndFrom,(LPARAM)&nmhdr);
        //set members
        m_clrBackColor=nmhdr.clrBackground;
        m_bTransparent=nmhdr.fTransparent;
        m_hwndParent=nmhdr.hwndFrom;
        switch (nmhdr.bAction)
        {
			case FVCA_SKIPDRAW:
				AdvanceFrame();
				break;
            case FVCA_SENDVIEWCHANGE:
                SendOnViewChange(DVASPECT_CONTENT);
                break;
            case FVCA_INVALIDATE:
                ::InvalidateRect(m_hwndParent, &(nmhdr.rcRect), FALSE);
                break;
            case FVCA_DRAW:
            case FVCA_CUSTOMDRAW:
                ATL_DRAWINFO di;
                memset(&di,0,sizeof(ATL_DRAWINFO));
                di.cbSize=sizeof(ATL_DRAWINFO);
                di.dwDrawAspect=DVASPECT_CONTENT;
                if (nmhdr.bAction==FVCA_DRAW)
                    di.hdcDraw=::GetWindowDC(m_hwndParent);
                else
                    di.hdcDraw=nmhdr.hDC;
                RECTL rcl={nmhdr.rcRect.left, nmhdr.rcRect.top, nmhdr.rcRect.right, nmhdr.rcRect.bottom};
                di.prcBounds=&rcl;
                OnDrawSmiley(di,nmhdr.bAction==FVCA_DRAW);
                if (nmhdr.bAction==FVCA_DRAW)
                    ::ReleaseDC(m_hwndParent, di.hdcDraw);
                break;
        }
        //SendPostPaintEvent
        nmhdr.bEvent=FVCN_POSTFIRE;
        if (hwndGrand) ::SendMessage(hwndGrand,WM_NOTIFY,(WPARAM)nmhdr.hwndFrom,(LPARAM)&nmhdr);
        //restore members
        m_clrBackColor=oldColor;
        m_bTransparent=fOldTransparent;
        m_hwndParent=hwndOld;
    }

    return S_OK;
}


static IGifSmileyCtrl *GetSmileyCtrl(HWND hwnd, POINT pt)
{
	IRichEditOle * ole=NULL;
	ITextDocument* textDoc=NULL;
	ITextRange* range=NULL;	
	IUnknown *iObject = NULL;
	IGifSmileyCtrl *iGifSmlCtrl=NULL;
	RECT rc = {0};

	if (!::SendMessage(hwnd, EM_GETOLEINTERFACE, 0, (LPARAM)&ole)) goto err;
	if (ole->QueryInterface(IID_ITextDocument, (void**)&textDoc) != S_OK) goto err;
	if (textDoc->RangeFromPoint(pt.x, pt.y, &range) != S_OK) goto err;

	if (range->GetEmbeddedObject(&iObject) != S_OK) 
	{
		LONG cp;
		range->GetStart(&cp);
		range->SetStart(cp-1);
		range->SetEnd(cp);
		if (range->GetEmbeddedObject(&iObject) != S_OK) goto err;
	}

	if (iObject->QueryInterface(IID_IGifSmileyCtrl, (void**) &iGifSmlCtrl) != S_OK) goto err;

	iGifSmlCtrl->GetRECT(&rc);

	::ScreenToClient(hwnd, &pt);

	if (!PtInRect(&rc, pt))
	{
		iGifSmlCtrl->Release();
		iGifSmlCtrl = NULL;
	}

err:
	if (iObject) iObject->Release();
	if (range) range->Release();
	if (textDoc) textDoc->Release();
	if (ole) ole->Release();

	return iGifSmlCtrl;
}


LRESULT CALLBACK CGifSmileyCtrl::HostWindowSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	MAPHOSTINFO::iterator it = _mapHostInfo.find(hwnd);
	if (it == _mapHostInfo.end())
		return FALSE;
	HostWindowData hwdat = it->second;
	if	(hwdat.hwnd != hwnd) 
		return FALSE;

	/*
	if (msg != EM_GETOLEINTERFACE)
	{
		LPARAM pos = GetMessagePos();
		POINT pt={LOWORD(pos), HIWORD(pos)};
		IGifSmileyCtrl *iGifSmlCtrl = GetSmileyCtrl(hwnd, pt);
		if (iGifSmlCtrl)
		{
			switch (msg)
			{
				case WM_LBUTTONDOWN: OutputDebugStringA(" ** MSG: WM_LBUTTONDOWN\n"); break;
				case WM_LBUTTONUP: OutputDebugStringA(" ** MSG: WM_LBUTTONUP\n"); break;
				case WM_LBUTTONDBLCLK: OutputDebugStringA(" ** MSG: WM_LBUTTONDBLCLK\n"); break;
				case WM_RBUTTONDOWN: OutputDebugStringA(" ** MSG: WM_RBUTTONDOWN\n"); break;
				case WM_RBUTTONUP: OutputDebugStringA(" ** MSG: WM_RBUTTONUP\n"); break;
				case WM_RBUTTONDBLCLK: OutputDebugStringA(" ** MSG: WM_RBUTTONDBLCLK\n"); break;
				case WM_MBUTTONDOWN: OutputDebugStringA(" ** MSG: WM_MBUTTONDOWN\n"); break;
				case WM_MBUTTONUP: OutputDebugStringA(" ** MSG: WM_MBUTTONUP\n"); break;
				case WM_MBUTTONDBLCLK: OutputDebugStringA(" ** MSG: WM_MBUTTONDBLCLK\n"); break;
				case WM_MOUSEWHEEL: OutputDebugStringA(" ** MSG: WM_MOUSEWHEEL\n"); break;
				case WM_MOUSEMOVE: OutputDebugStringA(" ** MSG: WM_MOUSEMOVE\n"); break;
				case WM_SETCURSOR: OutputDebugStringA(" ** MSG: WM_SETCURSOR\n"); break;
				case WM_NCHITTEST: OutputDebugStringA(" ** MSG: WM_NCHITTEST\n"); break;
				case WM_DESTROY: OutputDebugStringA(" ** MSG: WM_DESTROY\n"); break;
				case EM_EXGETSEL: OutputDebugStringA(" ** MSG: EM_EXSETSEL\n"); break;
				case EM_GETZOOM: OutputDebugStringA(" ** MSG: EM_GETZOOM\n"); break;
				default:
				{
					TCHAR tmp[1024];
					_sntprintf(tmp, 1024, _T(" ** MSG: %x\n"), msg);
					OutputDebugString(tmp);
					break;
				}
			}

			iGifSmlCtrl->Release();
		}
	}
	*/

	switch (msg)
	{
	case EM_GETOLEINTERFACE:
		break;

	case WM_LBUTTONDOWN:
	case WM_LBUTTONUP:
	case WM_LBUTTONDBLCLK:
	case WM_RBUTTONDOWN:
	case WM_RBUTTONUP:
	case WM_RBUTTONDBLCLK:
	case WM_MBUTTONDOWN: 
	case WM_MBUTTONUP:
	case WM_MBUTTONDBLCLK:
	case WM_MOUSEWHEEL:
	case WM_MOUSEMOVE:
	{
		POINT pt = { LOWORD(lParam), HIWORD(lParam) };
		::ClientToScreen(hwnd, &pt);

		IGifSmileyCtrl *iGifSmlCtrl = GetSmileyCtrl(hwnd, pt);
		if (iGifSmlCtrl)
		{
			LRESULT ret = 0;

			iGifSmlCtrl->OnMsg(hwnd, msg, wParam, lParam, &ret);
			iGifSmlCtrl->Release();

			if (ret)
				return ret;
		}
		else
			_pCurrentTip->HideSmileyTip();

		break;
	}

	case WM_SETCURSOR:
	case WM_NCHITTEST:
	{
		LPARAM pos = GetMessagePos();
		POINT pt={LOWORD(pos), HIWORD(pos)};

		IGifSmileyCtrl *iGifSmlCtrl = GetSmileyCtrl(hwnd, pt);
		if (iGifSmlCtrl)
		{
			LRESULT ret = 0;
			
			iGifSmlCtrl->OnMsg(hwnd, msg, wParam, lParam, &ret);
			iGifSmlCtrl->Release();

			return ret;
		}
		break;
	}
	case WM_DESTROY:
		{
			//Desubclass
			if (hwdat.pOldProc != NULL)
				::SetWindowLong(hwnd, GWL_WNDPROC, (LONG) hwdat.pOldProc);
			_mapHostInfo.erase(it);
			break;
		}
	}
	return CallWindowProc(hwdat.pOldProc, hwnd, msg, wParam, lParam);
}
void CGifSmileyCtrl::StopAnimation()
{
	_mapTimers.erase(m_nTimerId);
	::KillTimer(NULL,m_nTimerId);
	m_nTimerId=NULL;
}
VOID CALLBACK CGifSmileyCtrl::TimerProc(HWND hwnd,UINT /*uMsg*/,UINT_PTR idEvent, DWORD /*dwTime*/)
{
    CGifSmileyCtrl * me=NULL;//(CGifSmileyCtrl*)idEvent;
    ::KillTimer(hwnd, idEvent);
    MAPTIMER::iterator it=_mapTimers.find(idEvent);
    if (it!=_mapTimers.end())
    {
        me=it->second;
        _mapTimers.erase(it);        
    }
   
    if ( me )  
    {       
        me->OnTimer();
    }
}

VOID CALLBACK CGifSmileyCtrl::PurgeImageListTimerProc( HWND, UINT, UINT_PTR, DWORD )
{
	::KillTimer(NULL, _purgeImageListTimerId);
	LISTIMAGES::iterator it=_listImages.begin();
	while (it!=_listImages.end())
	{
		if ((*it) && (*it)->GetRefCount()<1)
		{			
			delete(*it);
			it = _listImages.erase(it);
			continue;
		}
		it++;
	}
}

VOID CALLBACK CGifSmileyCtrl::ToolTipTimerProc(HWND /*hwnd*/,UINT /*uMsg*/,UINT_PTR idEvent, DWORD /*dwTime*/)
{
   POINT pt;
   GetCursorPos(&pt);              
   HWND focusWnd=WindowFromPoint(pt);
   ::ScreenToClient(focusWnd, &pt);

   static HWND lastHWND=NULL;
   static int lastscroll=0;  
   
   bool recalcPos=FALSE;

   bool skipWindowStuff=false;
   LISTSMILEYS::iterator it;
   for ( it=_listSmileys.begin(); it!=_listSmileys.end(); it++ )
   {
        CGifSmileyCtrl * smiley=*it;   
        if (focusWnd==smiley->m_hwndParent)
        {
            if (!skipWindowStuff)
            {
                if (lastHWND==focusWnd)
                {
                    POINT pt;
                    ::SendMessage(lastHWND, EM_GETSCROLLPOS, 0, (LPARAM) &pt);
                    if (pt.y!=lastscroll) recalcPos=TRUE;
                    lastscroll=pt.y;
                }
                else
                {
                    lastHWND=focusWnd;
                    ::SendMessage(lastHWND, EM_GETSCROLLPOS, 0, (LPARAM) &pt);
                    lastscroll=pt.y;
                }
                skipWindowStuff=true;
            }
            if (recalcPos || (smiley->m_rectPos.bottom!=-100 && !smiley->m_bPaintValid && smiley->m_nFrameCount<2))
                smiley->GetObjectPos(smiley->m_hwndParent);
            if (PtInRect(&smiley->m_rectPos, pt)) 
            {
                smiley->ShowSmileyTooltip();
                break;
            }
            else if (smiley->m_bTipShow) 
            {
                smiley->HideSmileyTip();
            }
        }    
   }
}


void CGifSmileyCtrl::AdvanceFrame()
{
    if ( m_nFrameCount > 1 &&  m_nTimerId == 0)
    {
        m_nTimerId=::SetTimer( NULL /*_hwndTimerWnd*/, (UINT_PTR) this, m_pDelays[m_nCurrentFrame], TimerProc );
        _mapTimers[m_nTimerId]=this;
    }
}

void CGifSmileyCtrl::OnTimer()
{
    m_nTimerId=NULL;
    m_nCurrentFrame=(m_nCurrentFrame+1)%m_nFrameCount;
    FireViewChange();
}

HRESULT CGifSmileyCtrl::GetRECT(RECT *rc)
{
	*rc = m_rectPos;
	return S_OK;
}

HRESULT CGifSmileyCtrl::OnMsg(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, LRESULT *res)
{
	*res = FALSE;
	switch (msg)
	{
	case WM_LBUTTONDOWN:
	case WM_LBUTTONUP:
	case WM_LBUTTONDBLCLK:
	case WM_RBUTTONDOWN:
	case WM_RBUTTONUP:
	case WM_RBUTTONDBLCLK:
	case WM_MBUTTONDOWN: 
	case WM_MBUTTONUP:
	case WM_MBUTTONDBLCLK:
	case WM_MOUSEWHEEL:
		if (m_pFlash == NULL)
			break;
	case WM_MOUSEMOVE:
	{
		if (m_pFlash == NULL)
		{
			ShowHint();
		}
		else
		{
			POINT pt = { LOWORD(lParam), HIWORD(lParam) };
//			LPARAM pos = GetMessagePos();
//			POINT pt = { LOWORD(pos), HIWORD(pos) };

			pt.x -= m_rectPos.left;
			pt.y -= m_rectPos.top;

			*res = m_pFlash->OnMsg(hwnd, msg, wParam, MAKELONG(pt.x, pt.y));
			*res = TRUE; // XXX 
		}

		break;
	}
	case WM_NCHITTEST:
	{
		if (m_pFlash == NULL)
			break;

		*res = HTCLIENT;
		break;
	}

	case WM_KEYDOWN:
	case WM_KEYUP:
	case WM_CHAR:
	case WM_SETCURSOR:
		if (m_pFlash == NULL)
			break;

		*res = m_pFlash->OnMsg(hwnd, msg, wParam, lParam);
		*res = TRUE; // XXX 
	}
	return S_OK;
}

ImageItem::ImageItem() : 
                        m_nRef( 0 ), 
                        m_pBitmap( NULL ),
                        m_nFrameCount( 0 ),
                        m_pFrameDelays( NULL ),
                        m_nHeight( 0 )
{
    
}
ImageItem::~ImageItem()
{
    if ( m_pBitmap ) delete m_pBitmap;
    if ( m_pFrameDelays ) delete [] m_pFrameDelays;
}
BOOL GetBitmapFromFile(Bitmap* &m_pBitmap, ATL::CString& strFilename, 
                       int& m_nFrameCount, Size& m_FrameSize, Size& ImageSize,
                       UINT* &m_pFrameDelays );

Bitmap * ImageItem::LoadImageFromFile(ATL::CString& strFilename, int nHeight)
{
    if ( !m_strFilename.CompareNoCase(strFilename) && nHeight==m_nHeight )
    {
        m_nRef++;
        return m_pBitmap;
    }
    else if ( m_pBitmap==NULL && m_nRef==0)
    {
        Size ImageSize(0, nHeight);
        if (GetBitmapFromFile(m_pBitmap, strFilename, m_nFrameCount, m_FrameSize, ImageSize, m_pFrameDelays ))
        {
            m_nHeight=nHeight;
            m_nRef++;
            m_strFilename=strFilename;
            return m_pBitmap;
        }
        return NULL;

    }
    //NOT REACHABLE
    DebugBreak();
    return NULL;
}

bool ImageItem::UnloadImage()
{
    m_nRef--;
    if (m_nRef < 1)  return true;
    return false;
}


