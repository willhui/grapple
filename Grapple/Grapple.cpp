/*
** Grapple
** Copyright (C) 2005-2010 Will Hui.
** All rights reserved.
**
** Distributed under the terms of the MIT license.
** See LICENSE file for details.
**
** Grapple.cpp
** Main application entrypoint.
*/

// Allow use of features specific to IE 6.0 or later.
#ifndef _WIN32_IE
#define _WIN32_IE 0x0600
#endif

#include "stdafx.h"

// Windows header files.
#include <Windowsx.h>
#include <commctrl.h>
#include <Shellapi.h>
#include <Shlwapi.h>

// C runtime header files.
#include <stdlib.h>
#include <malloc.h>
#include <memory.h>
#include <tchar.h>
#include <string>

#include "Grapple.h"

#define MY_MSG		(WM_APP+0)
#define MY_ENABLE	(WM_APP+1)
#define MY_DISABLE	(WM_APP+2)
#define MY_ABOUT	(WM_APP+3)
#define MY_QUIT		(WM_APP+4)

typedef bool (WINAPI *InstallHookFn)(void);
typedef void (WINAPI *RemoveHookFn)(void);


const TCHAR *APP_NAME = TEXT("Grapple");
const TCHAR *APP_VERSION = TEXT("3.1");
const TCHAR *DLL_FILE = TEXT("GrappleLib.dll");
const int MAX_LOADSTRING = 100;

static HWND appWnd;
static HINSTANCE hInst;

static NOTIFYICONDATA niData;
static HINSTANCE dllInst;
static bool isHookInstalled = false;
static InstallHookFn InstallHook;
static RemoveHookFn RemoveHook;


// Pesky prototypes.
static ATOM MyRegisterClass(HINSTANCE hInstance);
static bool InitInstance(HINSTANCE, int);
static LRESULT	CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
static LRESULT	CALLBACK About(HWND, UINT, WPARAM, LPARAM);


static void ShowAboutBox(const TCHAR *format, ...)
{
	const int BUF_SIZE = 1024;

	va_list list;
	va_start(list, format);
	TCHAR buf[BUF_SIZE];
	_vsntprintf_s(buf, BUF_SIZE, _TRUNCATE, format, list);
	va_end(list);

	MessageBox(NULL, buf, TEXT("About"), MB_OK);
}

static void EnableGrapple(void)
{
	if (!dllInst) {
		dllInst = LoadLibrary(DLL_FILE);
		if (!dllInst) {
			MessageBox(NULL, TEXT("Unable to load hook DLL."), TEXT("Error"), MB_OK);
			return;
		}
	}
	if (!InstallHook || !RemoveHook) {
		InstallHook = (InstallHookFn) GetProcAddress(dllInst, (LPCSTR) MAKEINTRESOURCE(1));
		RemoveHook = (RemoveHookFn) GetProcAddress(dllInst, (LPCSTR) MAKEINTRESOURCE(2));
		if (!InstallHook || !RemoveHook) {
			MessageBox(
				NULL,
				TEXT("Unable to retrieve install/remove procedures from hook DLL."),
				TEXT("Error"),
				MB_OK);
			return;
		}
	}
	if (!isHookInstalled) {
		if (InstallHook()) {
			isHookInstalled = true;
		}
	}
}

static void DisableGrapple(void)
{
	if (isHookInstalled) {
		RemoveHook();
		isHookInstalled = false;
	}
}

// Set the current working directory to the same one the application is in.
static void ChangeToAppPath(void)
{
	char appPath[MAX_PATH] = "";
	GetModuleFileNameA(0, appPath, sizeof(appPath) - 1);
	std::string appDir = appPath;
	appDir = appDir.substr(0, appDir.rfind("\\"));
	SetCurrentDirectoryA(appDir.c_str());
}

int	APIENTRY _tWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance,
					   LPTSTR lpCmdLine, int nCmdShow)
{
	ChangeToAppPath();
	MyRegisterClass(hInstance);
	if (!InitInstance(hInstance, nCmdShow))
		return 0;

	EnableGrapple();

	// Main	message	loop.
	MSG msg;
	while (GetMessage(&msg,	NULL, 0, 0)) {
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}

	return (int)msg.wParam;
}

static ATOM MyRegisterClass(HINSTANCE hInstance)
{
	WNDCLASSEX wcex;
	wcex.cbSize			= sizeof(WNDCLASSEX); 
	wcex.style			= CS_HREDRAW | CS_VREDRAW;
	wcex.lpfnWndProc	= (WNDPROC)WndProc;
	wcex.cbClsExtra		= 0;
	wcex.cbWndExtra		= 0;
	wcex.hInstance		= hInstance;
	wcex.hIcon			= LoadIcon(hInstance, (LPCTSTR)IDI_GRAPPLE);
	wcex.hCursor		= LoadCursor(NULL, IDC_ARROW);
	wcex.hbrBackground	= (HBRUSH)GetStockObject(WHITE_BRUSH);
	wcex.lpszMenuName	= TEXT("MainMenu");
	wcex.lpszClassName	= APP_NAME;
	wcex.hIconSm		= LoadIcon(wcex.hInstance, (LPCTSTR)IDI_SMALL);
	return RegisterClassEx(&wcex);
}

static ULONGLONG GetDllVersion(LPCTSTR dllname)
{
	ULONGLONG version = 0;
	HINSTANCE dllhandle;
	dllhandle = LoadLibrary(dllname);

	if (dllhandle) {
		DLLGETVERSIONPROC getVersionProc;
		getVersionProc = (DLLGETVERSIONPROC)GetProcAddress(dllhandle, "DllGetVersion");

		if (getVersionProc) {
			DLLVERSIONINFO dvi;
			HRESULT result;
			ZeroMemory(&dvi, sizeof(dvi));
			dvi.cbSize = sizeof(dvi);
			result = (*getVersionProc)(&dvi);

			if (SUCCEEDED(result))
				version = MAKEDLLVERULL(dvi.dwMajorVersion, dvi.dwMinorVersion, 0, 0);
		}
		FreeLibrary(dllhandle);
	}
	
	return version;
}

static void InstallTrayIcon(HWND hWnd, HINSTANCE hInstance)
{
	ZeroMemory(&niData, sizeof(NOTIFYICONDATA));
	niData.hWnd = hWnd;
	ULONGLONG ver = GetDllVersion(TEXT("shell32.dll"));
	if (ver >= MAKEDLLVERULL(5, 0,0,0)) {
		niData.cbSize = sizeof(NOTIFYICONDATA);
	} else {
		niData.cbSize = NOTIFYICONDATA_V2_SIZE;
	}
	niData.uID = 1;
	niData.uFlags = NIF_ICON | NIF_MESSAGE;
	niData.hIcon = LoadIcon(hInstance, (LPCTSTR)IDI_GRAPPLE);
	niData.uCallbackMessage = MY_MSG;
	Shell_NotifyIcon(NIM_ADD, &niData);
}

static bool InitInstance(HINSTANCE hInstance, int nCmdShow)
{
	hInst = hInstance;
	HWND hWnd = CreateWindow(APP_NAME, APP_NAME, WS_OVERLAPPED | WS_THICKFRAME,
		CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, NULL, NULL, hInstance, NULL);
	if (hWnd) {
		InstallTrayIcon(hWnd, hInstance);

		// We don't have much of an interface yet...
		//ShowWindow(hWnd, nCmdShow);
		//UpdateWindow(hWnd);
	}
	return hWnd != NULL;
}

static void SetNormalMenuItem(MENUITEMINFO *item, INT itemID, TCHAR *text)
{
	ZeroMemory(item, sizeof(MENUITEMINFO));
	item->cbSize = sizeof(MENUITEMINFO);
	item->fMask = MIIM_ID | MIIM_FTYPE | MIIM_STRING;
	item->wID = itemID;
	item->fType = MFT_STRING;
	item->dwTypeData = text;
	item->cch = (UINT)_tcslen(text) + 1;
}

static void SetCheckedMenuItem(MENUITEMINFO *item, INT itemID, TCHAR *text, bool isChecked)
{
	SetNormalMenuItem(item, itemID, text);
	item->fMask |= MIIM_CHECKMARKS | MIIM_STATE;
	item->fType |= MFT_RADIOCHECK;
	item->fState = (isChecked ? MFS_CHECKED : 0);
	item->hbmpChecked = NULL;
	item->hbmpUnchecked = NULL;
}

static void ShowContextMenu(HWND hWnd)
{
	POINT pt;
	GetCursorPos(&pt);
	HMENU hMenu = CreatePopupMenu();
	if (hMenu) {
		MENUITEMINFO item;
		SetCheckedMenuItem(&item, MY_ENABLE, TEXT("Enable"), isHookInstalled);
		InsertMenuItem(hMenu, 1, TRUE, &item);
		SetCheckedMenuItem(&item, MY_DISABLE, TEXT("Disable"), !isHookInstalled);
		InsertMenuItem(hMenu, 2, TRUE, &item);
		SetNormalMenuItem(&item, MY_ABOUT, TEXT("About"));
		InsertMenuItem(hMenu, 3, TRUE, &item);
		SetNormalMenuItem(&item, MY_QUIT, TEXT("Quit"));
		InsertMenuItem(hMenu, 4, TRUE, &item);

		// We must set our window to the foreground or the menu won't
		// disappear when it should.
		SetForegroundWindow(hWnd);
		TrackPopupMenu(hMenu, TPM_BOTTOMALIGN, pt.x, pt.y, 0, hWnd, NULL);
		DestroyMenu(hMenu);
	}
}

// Main window procedure.
static LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	int	wmId, wmEvent;
	POINT minsz;
	POINT maxsz;
	minsz.x = minsz.y = 50;
	maxsz.x = maxsz.y = 300;

	switch (message) {
	case MY_MSG:
		// Handle tray icon events.
		switch (lParam) {
		case WM_LBUTTONDBLCLK:
			break;
		case WM_RBUTTONDOWN:
		case WM_CONTEXTMENU:
			ShowContextMenu(hWnd);
			break;
		}
		break;

	case WM_COMMAND:
		wmId = LOWORD(wParam); 
		wmEvent = HIWORD(wParam); 
		switch (wmId) {
		case MY_ENABLE:
			EnableGrapple();
			break;
		case MY_DISABLE:
			DisableGrapple();
			break;
		case MY_ABOUT:
			ShowAboutBox(
				TEXT("%s v%s\nCopyright (C) 2005-2010 Will Hui\nAll rights reserved."),
				APP_NAME, APP_VERSION);
			break;
		case MY_QUIT:
			DestroyWindow(hWnd);
			break;
		default:
			return DefWindowProc(hWnd, message,	wParam,	lParam);
		}
		break;

	case WM_DESTROY:
		DisableGrapple();
		niData.uFlags = 0;
		Shell_NotifyIcon(NIM_DELETE, &niData);
		PostQuitMessage(0);
		break;

	default:
		return DefWindowProc(hWnd, message,	wParam,	lParam);
	}

	return 0;
}
