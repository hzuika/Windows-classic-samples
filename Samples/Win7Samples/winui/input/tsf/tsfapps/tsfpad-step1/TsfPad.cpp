// TsfPad.cpp : Defines the entry point for the application.
//

#include "private.h"
#include "TsfPad.h"
#include "TextInputCtrl.h"
#define MAX_LOADSTRING 100

// Global Variables:
HINSTANCE g_hInst;								// current instance
TCHAR szTitle[MAX_LOADSTRING];					// The title bar text
TCHAR szWindowClass[MAX_LOADSTRING];			// the main window class name

// Forward declarations of functions included in this code module:
ATOM				MyRegisterClass(HINSTANCE hInstance);
BOOL				InitInstance(HINSTANCE, int);
LRESULT CALLBACK	WndProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK	About(HWND, UINT, WPARAM, LPARAM);

HWND g_hwnd;
CTextInputCtrl *g_pTextInputCtrl;

// アプリケーションで一つ持てばいいリソース
// 必ず必要
// CoCreateInstanceで作製
// こいつを使って色々作る
// スレッドマネージャ
// こいつのソース(ITfSource)に登録(AdviseSink)しておくことで，変換中に該当する仮想関数が呼び出される．
// スレッドごとに一つだけ存在して，このスレッドマネージャを経由してあれこれ処理を行う
ITfThreadMgr *g_pThreadMgr = NULL;
TfClientId g_TfClientId = TF_CLIENTID_NULL;

// アプリケーションで一つ持てばいいリソース
// ITfThreadMgrのQueryInterfaceを呼べば入手できる
// TextServiceにキーを処理させる(でないと入力できない)
ITfKeystrokeMgr *g_pKeystrokeMgr = NULL;

//----------------------------------------------------------------
//
//
//
//----------------------------------------------------------------

int APIENTRY MyWinMain(HINSTANCE hInstance,
                     HINSTANCE hPrevInstance,
                     LPTSTR    lpCmdLine,
                     int       nCmdShow)
{
 	// TODO: Place code here.
	MSG msg = {0};

    CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);

    if (FAILED(CoCreateInstance(CLSID_TF_ThreadMgr, NULL, CLSCTX_INPROC_SERVER, 
                                IID_ITfThreadMgr, (void**)&g_pThreadMgr)))
    {
        goto Exit;
    }

    // 適切なタイミングでITfThreadMgr::Activate関数を呼び出す
    // 同じ数だけITfThreadMgr::Deactivate関数を呼び出す必要がある．
    if (FAILED(g_pThreadMgr->Activate(&g_TfClientId)))
    {
        goto Exit;
    }

    if (g_pThreadMgr->QueryInterface(IID_ITfKeystrokeMgr, (void **)&g_pKeystrokeMgr) != S_OK)
    {
        goto Exit;
    }

    
	// Initialize global strings
	LoadString(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
	LoadString(hInstance, IDC_TSFPAD, szWindowClass, MAX_LOADSTRING);
	MyRegisterClass(hInstance);
	CTextInputCtrl::RegisterClass(hInstance);


	// Perform application initialization:
	if (!InitInstance (hInstance, nCmdShow)) 
	{
		goto Exit;
	}

	// Main message loop:
	while (GetMessage(&msg, NULL, 0, 0)) 
	{
 		TranslateMessage(&msg);
	 	DispatchMessage(&msg);
	}

    // ITfThreadMgr::Activate関数と同じ数だけITfThreadMgr::Deactivate関数を呼び出す必要がある．
    g_pThreadMgr->Deactivate();

Exit:
    if (g_pThreadMgr)
    {
        g_pThreadMgr->Release();
    }
    CoUninitialize();

	return (int) msg.wParam;
}



//----------------------------------------------------------------
//
//
//
//----------------------------------------------------------------

ATOM MyRegisterClass(HINSTANCE hInstance)
{
	WNDCLASSEX wcex;

	wcex.cbSize = sizeof(WNDCLASSEX); 

	wcex.style			= CS_HREDRAW | CS_VREDRAW;
	wcex.lpfnWndProc	= (WNDPROC)WndProc;
	wcex.cbClsExtra		= 0;
	wcex.cbWndExtra		= 0;
	wcex.hInstance		= hInstance;
	wcex.hIcon			= LoadIcon(hInstance, (LPCTSTR)IDI_TSFPAD);
	wcex.hCursor		= LoadCursor(NULL, IDC_ARROW);
	wcex.hbrBackground	= (HBRUSH)(COLOR_WINDOW+1);
	wcex.lpszMenuName	= (LPCTSTR)IDC_TSFPAD;
	wcex.lpszClassName	= szWindowClass;
	wcex.hIconSm		= LoadIcon(wcex.hInstance, (LPCTSTR)IDI_SMALL);

	return RegisterClassEx(&wcex);
}

//----------------------------------------------------------------
//
//
//
//----------------------------------------------------------------

BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
   
   g_hInst = hInstance; // Store instance handle in our global variable

   g_hwnd = CreateWindow(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
      CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, NULL, NULL, hInstance, NULL);

   if (!g_hwnd)
   {
      return FALSE;
   }

   g_pTextInputCtrl = new CTextInputCtrl();
   if (!g_pTextInputCtrl)
   {
       return FALSE;
   }
   if (!g_pTextInputCtrl->Create(g_hwnd))
   {
       return FALSE;
   }

   // Keyboard の登録 for WM_INPUT
   RAWINPUTDEVICE rid;

   rid.usUsagePage = 0x01;
   rid.usUsage = 0x06;
   rid.dwFlags = 0;
   rid.hwndTarget = g_hwnd;
   RegisterRawInputDevices(&rid, 1, sizeof(RAWINPUTDEVICE));

   ShowWindow(g_hwnd, nCmdShow);
   UpdateWindow(g_hwnd);
   ShowWindow(g_pTextInputCtrl->GetWnd(), SW_SHOW);
   SetFocus(g_pTextInputCtrl->GetWnd());

   return TRUE;
}

//----------------------------------------------------------------
//
//
//
//----------------------------------------------------------------

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	int wmId, wmEvent;
	PAINTSTRUCT ps;
	HDC hdc;

	switch (message) 
	{
	case WM_COMMAND:
		wmId    = LOWORD(wParam); 
		wmEvent = HIWORD(wParam); 
		// Parse the menu selections:
		switch (wmId)
		{
		    case IDM_ABOUT:
			    DialogBox(g_hInst, (LPCTSTR)IDD_ABOUTBOX, hWnd, (DLGPROC)About);
			    break;

            case IDM_FONT:
                if (g_pTextInputCtrl->GetWnd())
                    g_pTextInputCtrl->SetFont(hWnd);
			    break;

		    case IDM_EXIT:
			    DestroyWindow(hWnd);
			    break;
		}
		break;
	
    case WM_SETFOCUS:
        if (g_pTextInputCtrl->GetWnd())
            SetFocus(g_pTextInputCtrl->GetWnd());
		break;
		
    case WM_SIZE:
		RECT rc;
		GetClientRect(g_hwnd, &rc);
		if (g_pTextInputCtrl)
		    g_pTextInputCtrl->Move(0, 0, rc.right, rc.bottom);
		break;

	case WM_PAINT:
		hdc = BeginPaint(hWnd, &ps);
		EndPaint(hWnd, &ps);
		break;

	case WM_DESTROY:
		PostQuitMessage(0);
		break;
    case WM_INPUT:
        printf("\n");
	default:
		return DefWindowProc(hWnd, message, wParam, lParam);
	}
	return 0;
}



//----------------------------------------------------------------
//
//
//
//----------------------------------------------------------------

LRESULT CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	switch (message)
	{
        case WM_INITDIALOG:
		    return TRUE;

	    case WM_COMMAND:
		    if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL) 
		    {
			    EndDialog(hDlg, LOWORD(wParam));
			    return TRUE;
		    }
		    break;
	}
	return FALSE;
}


//+---------------------------------------------------------------------------
//
// ModuleEntry
//
//----------------------------------------------------------------------------


int
WINAPI 
WinMain(
    HINSTANCE ,
    HINSTANCE ,
    LPSTR ,
    int 
    )
{
    STARTUPINFO si;

    si.dwFlags = 0;
    GetStartupInfo(&si);

    return MyWinMain(GetModuleHandle(NULL), NULL, GetCommandLine(), si.dwFlags & STARTF_USESHOWWINDOW ? si.wShowWindow : SW_SHOWDEFAULT);
}

