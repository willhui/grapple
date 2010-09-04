/*
** Grapple
** Copyright (C) 2005-2010 Will Hui.
**
** Distributed under the terms of the MIT license.
** See LICENSE file for details.
**
** GrappleLib.cpp
** DLL entrypoint. This is where all the real work occurs.
**
** Known issues:
** - Send-to-back sometimes fails to give input focus to the next foreground
**   window that is activated. I believe this is a 32/64-bit issue, as I have
**   only been able to repro the bug when switching between a 32-bit app and
**   a 64-bit app.
** - Grapple will not work on administrator-level windows when Grapple itself
**   is running without administrator permissions.
** - We are making the assumption that EnumWindows() enumerates through
**   windows in z-order, from front to back. Nothing in the documentation
**   guarantees that this will always be the case. However, the only
**   other window enumeration method I know of is GetNextWindow(), which
**   isn't too reliable. So if EnumWindows() ever deviates from z-order,
**   I guess we're stuck...
** - We currently allow "Always On Top" windows to get sent to the back of
**   the z-order. They lose their "Always On Top" status when sent to
**   the back. Is this acceptable?
** - There is currently no way to blacklist misbehaving applications from
**   Grapple's influence.
**
** 3.2:
** > Smarter detection of "tangible" windows that should be selected for move
**   and resize operations.
** > Better heuristics for selecting the next foreground window after sending
**   a window to the back of the z-order.
** > Configured the vcproj files to allow building a 64-bit version of Grapple.
**   Technically, this is the correct things to do, but most things seem to work
**   without it. Right now the build defaults to x86.
** > Removed the very first feature this application had. You can no longer
**   middle click (without holding ALT) on a window's title bar in order to send
**   it to the back of the z-order. You must hold ALT. This old behavior was
**   inspired by BeOS, but I've too often activated it by accident when trying
**   to close tabs with middle clicks in Google Chrome.
**
** 3.1:
** > More robust window tracking in the face of mouse drift while dragging.
**   We remember the hwnd used to start the move/resize operation and then reuse
**   it for the duration of the mouse drag gesture.
** > Fixed a bug where move/resize appears to get "stuck down" upon initiating
**   a drag operation inside a Flash application. This behavior could be seen in
**   any web browser on Win 7 (and possibly Win Vista as well).
**
** 3.0:
** > Fixed ALT+MiddleClick disabling on full-screen apps for REAL this time. :)
** > ALT+MiddleClick now sends to back when ALT is held on MOUSEDOWN, not on
**   MOUSEUP. This is the more logical behavior.
** > Now using SetWindowPlacement() instead of SetWindowPos(), to be consistent
**   with the fact that GetWindowPlacement() gives us workspace coordinates.
**   This resolves window "creeping up" issues when the taskbar is put at the
**   top of the screen.
** > Found a better way to deal with attempts to resize below the minimum
**   window size. Windows won't "twitch" anymore when this happens. However,
**   resizing below the minimum from any corner but the bottom-right corner
**   will cause the window position to shift over. Also, AIM still won't
**   play nice with this method.
**
** 2.9.8-beta:
** > Second attempt at fixing the menu bar input focus side-effect.
**
** 2.9.7-beta:
** > Disabled ALT + MiddleClick on full screen applications. Forgot to do this before.
** > Suppressed MOUSEMOVE events when moving and resizing. This fixes a few glitches
**   I've encountered with DC++ and Console2, and potentially other applications.
** > Windows are brought to the foreground when ALT+clicked.
**
** 2.9.6-beta:
** > EnumWindowsProc() was using the original window handle instead of the owned
**   window handle, which is incorrect. I'm surprised I didn't notice this until
**   it failed to work with the e text editor. This bug is now fixed.
** > Attempt to fix the menu bar input focus side-effect.
**
*/

#include "stdafx.h"
#include "GrappleLib.h"
#include <cstdio>
#include <cstdlib>
#include <cstring>

#define QUASIMODE VK_MENU    // ALT key.

static const int ERROR_STRING_SIZE = 1024;

static HANDLE dllHandle;

static bool isMouseHookInstalled = false;
static HHOOK mouseHook;

static bool isKbHookInstalled = false;
static HHOOK kbHook;
static bool quasimodeNeedsKeyUp = false;

static bool inSendBackState = false;

static bool inMoveState = false;
enum ResizeEnum { NONE, TOPLEFT, TOPRIGHT, BOTLEFT, BOTRIGHT };
static ResizeEnum resizeState = NONE;

static POINT wndref, mouseref;
static RECT wndrectref;
static HWND hwndref;

LRESULT WINAPI CALLBACK MouseProc(int nCode, WPARAM wParam, LPARAM lParam);
LRESULT WINAPI CALLBACK KbProc(int nCode, WPARAM wParam, LPARAM lParam);


static void Complain(const TCHAR *s)
{
	const DWORD err = GetLastError();
	TCHAR buf[ERROR_STRING_SIZE];
	_stprintf_s(buf, ERROR_STRING_SIZE, TEXT("%s (error %i)"), s, err);
	MessageBoxW(NULL, buf, TEXT("Error"), MB_OK);
}

static void DebugBox(const TCHAR *format, ...)
{
	va_list list;
	va_start(list, format);
	TCHAR buf[ERROR_STRING_SIZE];
	_vsntprintf_s(buf, ERROR_STRING_SIZE, _TRUNCATE, format, list);
	va_end(list);
	MessageBox(NULL, buf, TEXT("Debug Message"), MB_OK);
}

// Append a message to the log file on disk.
static void Log(const char *format, ...)
{
	va_list list;
	va_start(list, format);
	char buf[ERROR_STRING_SIZE];
	vsnprintf_s(buf, ERROR_STRING_SIZE, _TRUNCATE, format, list);
	va_end(list);
	
	FILE *f;
	errno_t error = fopen_s(&f, "C:/Users/will/Documents/Projects/Grapple/log.txt", "at");
	if (error == 0) {
		fprintf(f, "%s\n", buf);
		fclose(f);
	}
}

static inline bool IsSet(int styles, int mask)
{
	return (styles & mask) != 0;
}

static HWND GetParentWindow(HWND hwnd)
{
	HWND parent = GetAncestor(hwnd, GA_PARENT);

	// GetAncestor() is supposed to return NULL upon reaching the desktop window
	// according to MSDN. But the documentation is wrong, so we explicitly check
	// for the desktop window handle ourselves.
	if (parent == GetDesktopWindow())
		parent = NULL;

	return parent;
}

BOOL APIENTRY DllMain(HMODULE hModule,DWORD ul_reason_for_call, LPVOID lpReserved)
{
	dllHandle = hModule;
	switch (ul_reason_for_call) {
	case DLL_PROCESS_ATTACH:
	case DLL_THREAD_ATTACH:
	case DLL_THREAD_DETACH:
	case DLL_PROCESS_DETACH:
		break;
	}
    return TRUE;
}

GRAPPLELIB_API bool WINAPI InstallHook(void)
{
	if (!isMouseHookInstalled) {
		mouseHook = SetWindowsHookEx(WH_MOUSE, MouseProc, (HINSTANCE)dllHandle, 0);
		if (mouseHook) {
			isMouseHookInstalled = true;
		} else {
			Complain(TEXT("Could not install the global mouse hook."));
		}
	}
	if (!isKbHookInstalled) {
		kbHook = SetWindowsHookEx(WH_KEYBOARD, KbProc, (HINSTANCE)dllHandle, 0);
		if (kbHook) {
			isKbHookInstalled = true;
		} else {
			Complain(TEXT("Could not install the global keyboard hook."));
		}
	}
	return isMouseHookInstalled && isKbHookInstalled;
}

GRAPPLELIB_API void WINAPI RemoveHook(void)
{
	if (isKbHookInstalled) {
		UnhookWindowsHookEx(kbHook);
		isKbHookInstalled = false;
	}
	if (isMouseHookInstalled) {
		UnhookWindowsHookEx(mouseHook);
		isMouseHookInstalled = false;
	}
}

// Returns the highest-level owner the specified window handle can be
// traced to. If the given handle has no owner, returns hwnd.
static HWND GetOwnerWindow(HWND hwnd)
{
	HWND owner = hwnd;
	HWND last;
	
	do {
		last = owner;
		owner = GetWindow(owner, GW_OWNER);
	} while (owner);
	return last;
}

static inline BOOL IsMinimized(HWND hwnd)
{
	return IsIconic(hwnd);
}

// Detect if a given window handle is a full-screen game, movie, etc.
static bool IsFullScreen(HWND hwnd)
{
	const int width = GetSystemMetrics(SM_CXSCREEN);
	const int height = GetSystemMetrics(SM_CYSCREEN);
	RECT r;
	GetWindowRect(hwnd, &r);
	return (width == (r.right - r.left)) && (height == (r.bottom - r.top));
}

// Reference: http://blogs.msdn.com/b/oldnewthing/archive/2007/10/08/5351207.aspx
static bool IsAltTabWindow(HWND hwnd)
{
	// Start at the root owner.
	HWND hwndWalk = GetAncestor(hwnd, GA_ROOTOWNER);

	// See if we are the last active visible popup.
	HWND hwndTry;
	while ((hwndTry = GetLastActivePopup(hwndWalk)) != hwndTry) {
		if (IsWindowVisible(hwndTry))
			break;
		hwndWalk = hwndTry;
	}
	return hwndWalk == hwnd;
}

// Check if a given window is reasonable to activate after a send-to-back operation.
static bool CanBringToTop(HWND hwnd)
{
	const int style = GetWindowLong(hwnd, GWL_STYLE);
	const int exStyle = GetWindowLong(hwnd, GWL_EXSTYLE);

	// If we can't activate it, don't bother.
	
	if (IsSet(exStyle, WS_EX_NOACTIVATE))
		return false;

	if (IsSet(style, WS_DISABLED))
		return false;
	
	if (!IsWindowVisible(hwnd))
		return false;
	
	// These window states would be counter-intuitive to activate.

	if (IsMinimized(hwnd))
		return false;

	if (IsFullScreen(hwnd))
		return false;

	// Tool windows should always be excluded.
	if (IsSet(exStyle, WS_EX_TOOLWINDOW))
		return false;
	
	return IsAltTabWindow(hwnd);
}

// Enumerate over all desktop windows so we can bring the next-highest
// window in the z-order to the foreground and activate it.
static BOOL WINAPI CALLBACK EnumWindowsProc(HWND hwnd, LPARAM lParam)
{
	const HWND sbwnd = (HWND)lParam;
	const HWND owner = GetOwnerWindow(hwnd);
	const HWND sbowner = GetOwnerWindow(sbwnd);

	if (CanBringToTop(hwnd) && (owner != sbowner)) {
		BringWindowToTop(hwnd);
		SetWindowPos(sbwnd, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
		return FALSE;
	}
	return TRUE;
}

// Returns the difference of two POINTs a-b.
static POINT SubtractPoints(const POINT a, const POINT b)
{
	POINT c;
	c.x = a.x - b.x;
	c.y = a.y - b.y;
	return c;
}

// Drags a window based on the new mouse point.
static void DragWindow(const HWND hwnd, const POINT pt)
{
	const POINT change = SubtractPoints(pt, mouseref);
	WINDOWPLACEMENT pl;
	pl.length = sizeof(WINDOWPLACEMENT);
	GetWindowPlacement(hwnd, &pl);
	RECT r = pl.rcNormalPosition;

	const int w = r.right - r.left;
	const int h = r.bottom - r.top;
	r.left = wndref.x + change.x;
	r.top = wndref.y + change.y;
	r.right = r.left + w;
	r.bottom = r.top + h;

	pl.rcNormalPosition = r;
	SetWindowPlacement(hwnd, &pl);
	
	// Don't activate window when moving.
	//SetWindowPos(hwnd, 0, d.x, d.y, 0, 0, SWP_NOSIZE | SWP_NOZORDER /* | SWP_NOACTIVATE */);
}

static bool IsResizable(const HWND hwnd)
{
	// Windows with WS_BORDER or WS_DLGFRAME styles do not have resizing grips.
	int style = GetWindowLong(hwnd, GWL_STYLE);
	if ((style & WS_THICKFRAME) || (style & WS_CAPTION))
		return true;
	if ((style & WS_BORDER) || (style & WS_DLGFRAME))
		return false;

	return true;
}

// Resizes a window based on the new mouse point.
static void ResizeWindow(const HWND hwnd, const POINT pt)
{
	if (!IsResizable(hwnd))
		return;
	POINT change = SubtractPoints(pt, mouseref);
	RECT wndrect = wndrectref;
	
	WINDOWPLACEMENT pl;
	pl.length = sizeof(WINDOWPLACEMENT);
	GetWindowPlacement(hwnd, &pl);
	pl.rcNormalPosition = wndrect;

	switch (resizeState) {
	case TOPLEFT:
		pl.rcNormalPosition.left += change.x;
		pl.rcNormalPosition.top += change.y;
		break;
	case TOPRIGHT:
		pl.rcNormalPosition.right += change.x;
		pl.rcNormalPosition.top += change.y;
		break;
	case BOTLEFT:
		pl.rcNormalPosition.left += change.x;
		pl.rcNormalPosition.bottom += change.y;
		break;
	case BOTRIGHT:
		pl.rcNormalPosition.right += change.x;
		pl.rcNormalPosition.bottom += change.y;
	default:
		break;
	}
	SetWindowPlacement(hwnd, &pl);

	// Don't activate window when resizing.
	//SetWindowPos(hwnd, 0, wndrect.left, wndrect.top, wndrect.right - wndrect.left,
	//	wndrect.bottom - wndrect.top, SWP_NOZORDER /* | SWP_NOACTIVATE */);
}

static LRESULT CALLBACK KbProc(const int code, const WPARAM wParam, const LPARAM lParam)
{
	int ret = 0;
	if (code >= 0 && wParam == VK_MENU) {
		int keyup = int(lParam & 0x80000000);
		if (quasimodeNeedsKeyUp) {
			if (keyup) {
				// Replace Alt SYSKEYUP with KEYUP message -- this prevents
				// input focus from changing to the menu bar. TODO: Spy++
				// tells us that apps don't actually receive this WM_KEYUP
				// message. But everything still seems to be working ok.
				PostMessage(NULL, WM_KEYUP, wParam, lParam);
				quasimodeNeedsKeyUp = false;
				ret = 1;
			}
		}
	}

	// Return non-zero to block the message from being passed on.
	// Otherwise, invoke CallNextHookEx().
	if (ret > 0)
		return ret;
	else
		return CallNextHookEx(kbHook, code, wParam, lParam);
}

// We say a window is "tangible" if it makes sense to move/resize it on-screen
// for the user. This function takes a window handle and searches up the window
// hierarchy for the first tangible window it can find.
static HWND GetTangibleWindow(HWND hwnd, bool debug)
{
	HWND prev = NULL;
	HWND window = hwnd;

	while (window != NULL) {
		const int style = GetWindowLong(window, GWL_STYLE);

		const bool isChild = IsSet(style, WS_CHILD) || IsSet(style, WS_CHILDWINDOW);
		if (!isChild && IsSet(style, WS_POPUP))
			return window;

		prev = window;
		window = GetParentWindow(window);
	}

	return prev;
}

// Global mouse hook procedure.
static LRESULT WINAPI CALLBACK MouseProc(const int nCode, const WPARAM wParam, const LPARAM lParam)
{
	int ret = 0;

	if (nCode >= 0) {
		const MOUSEHOOKSTRUCT *mouseHookStruct = (MOUSEHOOKSTRUCT *)lParam;
		const HWND hwnd = GetTangibleWindow(mouseHookStruct->hwnd, false);
		WINDOWPLACEMENT placement;
		placement.length = sizeof(WINDOWPLACEMENT);
		
		switch (wParam) {
		case WM_NCMBUTTONDOWN:
		case WM_MBUTTONDOWN:
			if (GetKeyState(QUASIMODE) < 0 && !IsFullScreen(hwnd)) {
				inSendBackState = true;
				ret = 1;
			}
			break;

		case WM_NCMBUTTONUP:
		case WM_MBUTTONUP:
			if (inSendBackState) {
				if (!IsFullScreen(hwnd))
					EnumWindows(EnumWindowsProc, (LPARAM)hwnd);

				inSendBackState = false;
				ret = 1;
			}
			break;

		case WM_NCLBUTTONDOWN:
		case WM_LBUTTONDOWN:
			// Drag anywhere.
			GetWindowPlacement(hwnd, &placement);
			
			if (GetKeyState(QUASIMODE) < 0 && placement.showCmd != SW_MAXIMIZE &&
					!inMoveState && resizeState == NONE && !IsFullScreen(hwnd)) {
				BringWindowToTop(hwnd);
				inMoveState = true;
				quasimodeNeedsKeyUp = true;

				// WM_MOUSEMOVE deltas seem to be too inaccurate to track dragging
				// operations. Mouse capture gives us far more accuracy in order
				// to prevent drift.
				//
				// We enable mouse capture for mouseHookStruct->hwnd instead of its
				// root owner window in order to work around odd "sticking" behavior
				// when trying to move/resize a Flash app inside a browser
				// on Win 7. (and possibly Win Vista).
				SetCapture(mouseHookStruct->hwnd);

				// Record starting window and mouse positions.
				wndref.x = placement.rcNormalPosition.left;
				wndref.y = placement.rcNormalPosition.top;
				mouseref = mouseHookStruct->pt;
				hwndref = hwnd;

				ret = 1;
			}
			break;

		case WM_NCRBUTTONDOWN:
		case WM_RBUTTONDOWN:
			// Resize anywhere.
			GetWindowPlacement(hwnd, &placement);
			if ((GetKeyState(QUASIMODE) < 0) && placement.showCmd != SW_MAXIMIZE &&
					!inMoveState && resizeState == NONE && !IsFullScreen(hwnd)) {
				BringWindowToTop(hwnd);
				quasimodeNeedsKeyUp = true;
				const RECT r = placement.rcNormalPosition;
				const int xhalf = (r.left + r.right) / 2;
				const int yhalf = (r.top + r.bottom) / 2;
				POINT pt = mouseHookStruct->pt;
				if (pt.x < xhalf) {
					resizeState = (pt.y < yhalf) ? TOPLEFT : BOTLEFT;
				} else {
					resizeState = (pt.y < yhalf) ? TOPRIGHT : BOTRIGHT;
				}

				// Refer to the comment in WM_LBUTTONDOWN for why we do mouse capture.
				SetCapture(mouseHookStruct->hwnd);

				// Record starting window and mouse positions.
				wndrectref = placement.rcNormalPosition;
				mouseref = mouseHookStruct->pt;
				hwndref = hwnd;

				ret = 1;
			}
			break;

		case WM_NCLBUTTONUP:
		case WM_LBUTTONUP:
			if (inMoveState) {
				ReleaseCapture();
				inMoveState = false;
				ret = 1;
			}
			break;

		case WM_NCRBUTTONUP:
		case WM_RBUTTONUP:
			if (resizeState != NONE) {
				ReleaseCapture();
				resizeState = NONE;
				ret = 1;
			}
			break;

		case WM_NCMOUSEMOVE:
		case WM_MOUSEMOVE:
			if (inMoveState) {
				DragWindow(hwndref, mouseHookStruct->pt);
				ret = 1;
			} else if (resizeState != NONE) {
				ResizeWindow(hwndref, mouseHookStruct->pt);
				ret = 1;
			}
			break;

		default:
			break;
		}
	}

	// Return non-zero to block the message from being passed on.
	// Otherwise, invoke the next hook in the chain.
	if (ret > 0)
		return ret;
	else
		return CallNextHookEx(mouseHook, nCode, wParam, lParam);
}
