#include "stdafx.h"
#include <comdef.h>
#include <richedit.h>
#include <richole.h>

#include "GifSmileyCtrl.h"
#include "m_anismiley.h"

#include <newpluginapi.h>
#include <m_database.h>

#define THROW(x) _com_issue_error(x)
#define THROWIF(x) if (x) _com_issue_error(x);
#define CHECKRESULT(sub) { HRESULT _hr; if ( FAILED( _hr=sub ) ) _com_issue_error(_hr); }

static PLUGINLINK * pluginLink=NULL;
extern "C" void *pLink;
void *pLink=NULL;

WINOLEAPI  CoInitializeEx(LPVOID pvReserved, DWORD dwCoInit);

static IUnknown * CALLBACK InsertAniSmiley(HWND hwnd, const TCHAR * filename, COLORREF backColor, int cy, const TCHAR * text)
{
	static BOOL hasFault=FALSE;
    static HWND hwndFaultOn=NULL;	
    if (hasFault && hwndFaultOn==hwnd) return FALSE;

    hasFault=TRUE;          //set to true, if all will be ok reset it to False
    hwndFaultOn=hwnd;

	if (pluginLink==NULL) pluginLink=(PLUGINLINK*)pLink;

    ATL::CString strPicPath(filename);
    ATL::CString strExt=strPicPath.Right(4);
	if (strExt.CompareNoCase(_T(".gif")) && strExt.CompareNoCase(_T(".jpg")) && strExt.CompareNoCase(_T(".png")) && strExt.CompareNoCase(_T(".bmp"))) 
		return FALSE;

	IRichEditOle *ole = NULL;
	LPSTORAGE m_lpStorage=NULL;
	LPOLEOBJECT	m_lpObject=NULL;
    LPLOCKBYTES lpLockBytes = NULL;
    LPOLECLIENTSITE m_lpClientSite = NULL;  
	CComObject<::CGifSmileyCtrl> * myObject=NULL;
	BSTR path = NULL;

    //Create lockbytes
    CHECKRESULT (::CreateILockBytesOnHGlobal(NULL, TRUE, &lpLockBytes) )
    THROWIF(lpLockBytes == NULL)

    //use lockbytes to create storage
    SCODE sc = ::StgCreateDocfileOnILockBytes(lpLockBytes, STGM_SHARE_EXCLUSIVE|STGM_CREATE|STGM_READWRITE, 0, &m_lpStorage);
    if (sc != S_OK)   {  lpLockBytes->Release();    THROW( sc );   }  
	THROWIF(m_lpStorage == NULL);

	// retrieve OLE interface for richedit
    ::SendMessage(hwnd, EM_GETOLEINTERFACE, 0, (LPARAM)&ole);

	// Get site
	ole->GetClientSite(&m_lpClientSite);
    THROWIF (m_lpClientSite == NULL);

    try
    {
        //Initialize COM interface
        CHECKRESULT( ::CoInitializeEx( NULL, COINIT_APARTMENTTHREADED ) );
   
		CComObject<::CGifSmileyCtrl>::CreateInstance(&myObject);
        
        THROWIF(myObject==NULL)
		
        myObject->AddRef();

        //COM operation need BSTR, so get a BSTR
        path = strPicPath.AllocSysString();

        //Load the gif
		CHECKRESULT( myObject->LoadFromFileSized(path, cy) )

		//Set back color
        OLE_COLOR oleBackColor=(OLE_COLOR)backColor;
		myObject->put_BackColor(oleBackColor);

        //get the IOleObject
		CHECKRESULT (myObject->QueryInterface(IID_IOleObject, (void**)&m_lpObject) )

        //Set it to be inserted
        OleSetContainedObject(m_lpObject, TRUE);

        //to insert into richedit, you need a struct of REOBJECT
        REOBJECT reobject;
        ZeroMemory(&reobject, sizeof(REOBJECT));

        reobject.cbStruct = sizeof(REOBJECT);	
        
		CLSID clsid;
		CHECKRESULT ( m_lpObject->GetUserClassID(&clsid) )

        //set clsid
        reobject.clsid = clsid;
        //can be selected
        reobject.cp = REO_CP_SELECTION;
        //content, but not static
        reobject.dvaspect = DVASPECT_CONTENT;
        //goes in the same line of text line
        reobject.dwFlags = REO_BELOWBASELINE;
        reobject.dwUser = (DWORD)myObject;
        //the very object
        reobject.poleobj = m_lpObject;
        //client site contain the object
        reobject.polesite = m_lpClientSite;
        //the storage 
        reobject.pstg = m_lpStorage;
        
		SIZEL sizel={0};
        reobject.sizel = sizel;
        
        CHECKRESULT( m_lpObject->SetClientSite(m_lpClientSite) )

		int flag=(DBGetContactSettingByte(NULL, "AnimatorService", "SlowRepaint", 0)==0) ? 0x01 : 0;
		CHECKRESULT (myObject->SetHostWindow((long)hwnd, flag))

        ole->InsertObject(&reobject);
        
        ATL::CString strHint(text);
        if ( strHint.GetLength()>0 )
        {
            BSTR Hint=strHint.AllocSysString();
            myObject->SetTooltip(Hint);
            SysFreeString(Hint);
        }
        
        //redraw the window to show animation
        RedrawWindow(hwnd,NULL, NULL, RDW_INVALIDATE);

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
		myObject->Release();
        hasFault=FALSE;       
    }
    catch( _com_error e )
    {
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

        if (myObject) 
			myObject->Release();
        SysFreeString(path);		
        ::CoUninitialize();	
        ole->Release();  
		return NULL;
    }
	return (IUnknown *)(IOleControl *)myObject;
}
extern "C" int InsertAniSmileyService(WPARAM wParam, LPARAM lParam)
{
    if (wParam==0) return 0; 
    INSERTANISMILEY * ias=(INSERTANISMILEY *)wParam;
    ATL::CString file;
    ATL::CString text;
    if (ias->dwFlags&IASF_UNICODE)
        file=ias->wcFilename;
    else
        file=ias->szFilename;
    if (ias->cbSize==sizeof(INSERTANISMILEY))
    {
        if (ias->dwFlags&IASF_UNICODE)
            text=ias->wcText;
        else
            text=ias->szText;
    }
    else
        text="";

    return (int)InsertAniSmiley(ias->hWnd, (const TCHAR*)file, ias->dwBackColor, ias->nHeight, (const TCHAR*)text );
}
