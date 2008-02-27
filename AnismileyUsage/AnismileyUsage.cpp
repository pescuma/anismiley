// AnismileyUsage.cpp : Defines the entry point for the application.
//

#include "stdafx.h"
#include "resource.h"
#include "..\_GifSmiley.h"
#include <comdef.h>
#include "richedit.h"
#include "richole.h"

#define ASSERT(x) if (!x) return FALSE
#define CHECKRESULT(sub) { HRESULT _hr; if ( FAILED( _hr=sub ) ) _com_issue_error(_hr); }
WINOLEAPI  CoInitializeEx(LPVOID pvReserved, DWORD dwCoInit);
_COM_SMARTPTR_TYPEDEF(IGifSmileyCtrl, __uuidof(IGifSmileyCtrl));
extern "C" const GUID __declspec(selectany)  CLSID_CGifSmileyCtrl = {0xDB35DD77,0x55E2,0x4905,0x80,0x75,0xEB,0x35,0x1B,0xB5,0xCB,0xC1};
extern "C" const IID IID_IGifSmileyCtrl={0xCB64102B,0x8CE4, 0x4A55, 0xB0, 0x50, 0x13,0x1C,0x43,0x5A,0x3A,0x3F};


#define MAX_LOADSTRING 100

// Global Variables:
HINSTANCE hInst;								// current instance
TCHAR szTitle[MAX_LOADSTRING];					// The title bar text
TCHAR szWindowClass[MAX_LOADSTRING];			// the main window class name


// Forward declarations of functions included in this code module:
ATOM				MyRegisterClass(HINSTANCE hInstance);
BOOL				InitInstance(HINSTANCE, int);
LRESULT CALLBACK	WndProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK	About(HWND, UINT, WPARAM, LPARAM);
static BOOL CALLBACK InsertGifImage(HWND hwnd, const TCHAR * filename, COLORREF backColor, int cy);

int APIENTRY _tWinMain(HINSTANCE hInstance,
                     HINSTANCE hPrevInstance,
                     LPTSTR    lpCmdLine,
                     int       nCmdShow)
{
 	// TODO: Place code here.
	MSG msg;
	HACCEL hAccelTable;

	// Initialize global strings
	LoadString(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
	LoadString(hInstance, IDC_ANISMILEYUSAGE, szWindowClass, MAX_LOADSTRING);
	MyRegisterClass(hInstance);

    HMODULE richeditlib=LoadLibraryA("riched20.dll");
	// Perform application initialization:
	if (!InitInstance (hInstance, nCmdShow)) 
	{
		return FALSE;
	}

	hAccelTable = LoadAccelerators(hInstance, (LPCTSTR)IDC_ANISMILEYUSAGE);

	// Main message loop:
	while (GetMessage(&msg, NULL, 0, 0)) 
	{
		if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg)) 
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}
    FreeLibrary(richeditlib);

	return (int) msg.wParam;
}



//
//  FUNCTION: MyRegisterClass()
//
//  PURPOSE: Registers the window class.
//
//  COMMENTS:
//
//    This function and its usage are only necessary if you want this code
//    to be compatible with Win32 systems prior to the 'RegisterClassEx'
//    function that was added to Windows 95. It is important to call this function
//    so that the application will get 'well formed' small icons associated
//    with it.
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
	WNDCLASSEX wcex;

	wcex.cbSize = sizeof(WNDCLASSEX); 

	wcex.style			= CS_HREDRAW | CS_VREDRAW;
	wcex.lpfnWndProc	= (WNDPROC)WndProc;
	wcex.cbClsExtra		= 0;
	wcex.cbWndExtra		= 0;
	wcex.hInstance		= hInstance;
	wcex.hIcon			= LoadIcon(hInstance, (LPCTSTR)IDI_ANISMILEYUSAGE);
	wcex.hCursor		= LoadCursor(NULL, IDC_ARROW);
	wcex.hbrBackground	= (HBRUSH)(COLOR_WINDOW+1);
	wcex.lpszMenuName	= (LPCTSTR)IDC_ANISMILEYUSAGE;
	wcex.lpszClassName	= szWindowClass;
	wcex.hIconSm		= LoadIcon(wcex.hInstance, (LPCTSTR)IDI_SMALL);

	return RegisterClassEx(&wcex);
}

//
//   FUNCTION: InitInstance(HANDLE, int)
//
//   PURPOSE: Saves instance handle and creates main window
//
//   COMMENTS:
//
//        In this function, we save the instance handle in a global variable and
//        create and display the main program window.
//
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
   HWND hWnd;

   hInst = hInstance; // Store instance handle in our global variable

   hWnd = CreateWindow(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
      CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, NULL, NULL, hInstance, NULL);

   if (!hWnd)
   {
      return FALSE;
   }

   //ShowWindow(hWnd, nCmdShow);
   //UpdateWindow(hWnd);

   return TRUE;
}

//
//  FUNCTION: WndProc(HWND, unsigned, WORD, LONG)
//
//  PURPOSE:  Processes messages for the main window.
//
//  WM_COMMAND	- process the application menu
//  WM_PAINT	- Paint the main window
//  WM_DESTROY	- post a quit message and return
//
//
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	int wmId, wmEvent;
	PAINTSTRUCT ps;
	HDC hdc;

	switch (message) 
	{
    case WM_CREATE:
        DialogBox(hInst, (LPCTSTR)IDD_ABOUTBOX, hWnd, (DLGPROC)About);
        PostQuitMessage(0);
        break;

    case WM_COMMAND:
		wmId    = LOWORD(wParam); 
		wmEvent = HIWORD(wParam); 
		// Parse the menu selections:
		switch (wmId)
		{
		case IDM_ABOUT:
			DialogBox(hInst, (LPCTSTR)IDD_ABOUTBOX, hWnd, (DLGPROC)About);
			break;
		case IDM_EXIT:
			DestroyWindow(hWnd);
			break;
		default:
			return DefWindowProc(hWnd, message, wParam, lParam);
		}
		break;
	case WM_PAINT:
		hdc = BeginPaint(hWnd, &ps);
		// TODO: Add any drawing code here...
		EndPaint(hWnd, &ps);
		break;
	case WM_DESTROY:
		PostQuitMessage(0);
		break;
	default:
		return DefWindowProc(hWnd, message, wParam, lParam);
	}
	return 0;
}

// Message handler for about box.
LRESULT CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	switch (message)
	{
	case WM_INITDIALOG:
        InsertGifImage(GetDlgItem(hDlg,IDC_RICHEDIT21), _T("yahoo.gif"), GetSysColor(COLOR_3DFACE),50);
		return TRUE;

	case WM_COMMAND:
		if ( LOWORD(wParam) == IDCANCEL) 
		{
			EndDialog(hDlg, LOWORD(wParam));
			return TRUE;
		}
		else if (LOWORD(wParam) == IDOK)
		{
			IRichEditOle *ole;
			HWND hwnd=GetDlgItem(hDlg,IDC_RICHEDIT21);
			SendMessage(hwnd, EM_GETOLEINTERFACE, 0, (LPARAM)&ole);
			REOBJECT obj={0};
			int i=ole->GetObjectCount();
			i=i;
			obj.cbStruct=sizeof(REOBJECT);
			ole->GetObject(0, &obj, REO_GETOBJ_POLEOBJ);
			IGifSmileyCtrl * iGif;
			obj.poleobj->QueryInterface(IID_IGifSmileyCtrl, (void**)&iGif);
			iGif->LoadFromFileSized(L"drinks.gif",50);
			iGif->Release();
			obj.poleobj->Release();
			ole->Release();
		}

		break;
	}
	return FALSE;
}



static BOOL CALLBACK InsertGifImage(HWND hwnd, const TCHAR * filename, COLORREF backColor, int cy)
{
    LPLOCKBYTES lpLockBytes = NULL;
    SCODE sc;
    //print to RichEdit' s IClientSite
    LPOLECLIENTSITE m_lpClientSite=NULL;
    //A smart point to IAnimator
    IGifSmileyCtrlPtr	m_lpAnimator=NULL;
	//A smart point to IAnimator
	IGifSmileyCtrl	*   m_lpGifSmileyControl=NULL;
    //ptr 2 storage	
    LPSTORAGE m_lpStorage=NULL;
    //the object 2 b insert 2
    LPOLEOBJECT	m_lpObject=NULL;

    BSTR path=NULL;

    //Create lockbytes
    sc = ::CreateILockBytesOnHGlobal(NULL, TRUE, &lpLockBytes);
    if (sc != S_OK)
        _com_issue_error(sc);
    ASSERT(lpLockBytes != NULL);

    //use lockbytes to create storage
    sc = ::StgCreateDocfileOnILockBytes(lpLockBytes,
        STGM_SHARE_EXCLUSIVE|STGM_CREATE|STGM_READWRITE, 0, &m_lpStorage);
    if (sc != S_OK)
    {
        ASSERT(lpLockBytes->Release() == 0);
        lpLockBytes = NULL;
        _com_issue_error(sc);
    }
    ASSERT(m_lpStorage != NULL);

    IRichEditOle *ole;
    SendMessage(hwnd, EM_GETOLEINTERFACE, 0, (LPARAM)&ole);
    //get the ClientSite of the very RichEditCtrl
    ole->GetClientSite(&m_lpClientSite);
    ASSERT(m_lpClientSite != NULL);

    try
    {
        //Initlize COM interface
        CHECKRESULT( ::CoInitializeEx( NULL, COINIT_APARTMENTTHREADED ) );

        //Get GifAnimator object
        //here, I used a smart point, so I do not need to free it
        CHECKRESULT( m_lpAnimator.CreateInstance(CLSID_CGifSmileyCtrl) );	

        //COM operation need BSTR, so get a BSTR
        BSTR path = SysAllocString((const OLECHAR*)filename);

    	CHECKRESULT( m_lpAnimator.QueryInterface(IID_IGifSmileyCtrl, (void**)&m_lpGifSmileyControl) )
        //Load the gif
        CHECKRESULT (m_lpGifSmileyControl->LoadFromFileSized(path, 0))

        OLE_COLOR oleBackColor=(OLE_COLOR)backColor;
        m_lpAnimator->put_BackColor(oleBackColor);

        //get the IOleObject
        CHECKRESULT( m_lpAnimator.QueryInterface(IID_IOleObject, (void**)&m_lpObject) )

        //Set it 2 b inserted
        OleSetContainedObject(m_lpObject, TRUE);

        //2 insert in 2 richedit, you need a struct of REOBJECT
        REOBJECT reobject;
        ZeroMemory(&reobject, sizeof(REOBJECT));

        reobject.cbStruct = sizeof(REOBJECT);	
        CLSID clsid;
        sc = m_lpObject->GetUserClassID(&clsid);
        if (sc != S_OK)
            _com_issue_error(sc);
        //set clsid
        reobject.clsid = clsid;
        //can be selected
        reobject.cp = REO_CP_SELECTION;
        //content, but not static
        reobject.dvaspect = DVASPECT_CONTENT;
        //goes in the same line of text line
        reobject.dwFlags = REO_BELOWBASELINE; //REO_RESIZABLE |
        reobject.dwUser = m_lpAnimator;
        //the very object
        reobject.poleobj = m_lpObject;
        //client site contain the object
        reobject.polesite = m_lpClientSite;
        //the storage 
        reobject.pstg = m_lpStorage;

        SIZEL sizel;
        sizel.cx = sizel.cy = 0;
        reobject.sizel = sizel;

        m_lpObject->SetClientSite(m_lpClientSite);

        m_lpGifSmileyControl->SetHostWindow((long)hwnd, 0);

        ole->InsertObject(&reobject);
        m_lpObject->DoVerb(OLEIVERB_SHOW, NULL, m_lpClientSite, -1, hwnd, NULL);
		

        //redraw the window to show animation
        RedrawWindow(hwnd,NULL, NULL, RDW_INVALIDATE);

		if (m_lpGifSmileyControl)
		{
			m_lpGifSmileyControl->Release();
			m_lpGifSmileyControl = NULL;
		}
        if (m_lpClientSite)
        {
            m_lpClientSite->Release();
            m_lpClientSite = NULL;
        }
        if (m_lpObject)
        {
            m_lpObject->Release();
            m_lpObject = NULL;
        }
        if (m_lpStorage)
        {
            m_lpStorage->Release();
            m_lpStorage = NULL;
        }

        SysFreeString(path);
        ole->Release();
        return TRUE;
    }
    catch( _com_error e )
    {
        SysFreeString(path);
		if (m_lpGifSmileyControl)
		{
			m_lpGifSmileyControl->Release();
			m_lpGifSmileyControl = NULL;
		}
        if (m_lpClientSite)
        {
            m_lpClientSite->Release();
            m_lpClientSite = NULL;
        }
        if (m_lpObject)
        {
            m_lpObject->Release();
            m_lpObject = NULL;
        }
        if (m_lpStorage)
        {
            m_lpStorage->Release();
            m_lpStorage = NULL;
        }
       // ::CoUninitialize();	
        ole->Release();
        return FALSE;
    }
    return 0;
}
