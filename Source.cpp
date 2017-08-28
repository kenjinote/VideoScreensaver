#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#pragma comment(lib, "scrnsavw")
#pragma comment(lib, "comctl32")
#pragma comment(lib, "dwmapi")

#include <windows.h>
#include <scrnsave.h>
#include <atlbase.h>
#include <atlwin.h>
#include <wmp.h>
#include <dwmapi.h>
#include <vector>
#include <algorithm>
#include <functional>
#include "resource.h"

CComModule _Module;
WNDPROC DefaultVideoWndProc;

BEGIN_OBJECT_MAP(ObjectMap)
END_OBJECT_MAP()

int GetArea(const LPRECT lpRect)
{
	return (lpRect->right - lpRect->left) * (lpRect->bottom - lpRect->top);
}

bool operator>(const RECT& left, const RECT& right)
{
	return GetArea((LPRECT)&left) > GetArea((LPRECT)&right);
}

BOOL CALLBACK MonitorEnumProc(HMONITOR hMonitor, HDC hdcMonitor, LPRECT lprcMonitor, LPARAM dwData)
{
	MONITORINFOEX MonitorInfoEx;
	MonitorInfoEx.cbSize = sizeof(MonitorInfoEx);
	if (GetMonitorInfo(hMonitor, &MonitorInfoEx) != 0)
	{
		DEVMODE dm = { 0 };
		dm.dmSize = sizeof(DEVMODE);
		if (EnumDisplaySettings(MonitorInfoEx.szDevice, ENUM_CURRENT_SETTINGS, &dm) != 0)
		{
			const int nMonitorWidth = dm.dmPelsWidth;
			const int nMonitorHeight = dm.dmPelsHeight;
			const int nMonitorPosX = dm.dmPosition.x;
			const int nMonitorPosY = dm.dmPosition.y;
			RECT rect = { nMonitorPosX, nMonitorPosY, nMonitorPosX + nMonitorWidth, nMonitorPosY + nMonitorHeight };
			((std::vector<RECT>*)dwData)->push_back(rect);
		}
	}
	return TRUE;
}

LRESULT CALLBACK MyVideoWndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	if (msg == WM_MOUSEMOVE)
	{
		return SendMessage(GetParent(hWnd), msg, wParam, lParam);
	}
	else
	{
		return CallWindowProc(DefaultVideoWndProc, hWnd, msg, wParam, lParam);
	}
}

BOOL PlayVideo(HWND hWnd, LPCTSTR lpszFilePath, BOOL bMute)
{
	BOOL bReturn = FALSE;
	CComPtr<IUnknown> pUnknown;
	if (SUCCEEDED(AtlAxGetControl(hWnd, &pUnknown)))
	{
		CComPtr<IWMPPlayer> pIWMPPlayer;
		if ((SUCCEEDED(pUnknown->QueryInterface(__uuidof(IWMPPlayer), (VOID**)&pIWMPPlayer))))
		{
			BSTR bstrText = SysAllocString(lpszFilePath);
			if (SUCCEEDED(pIWMPPlayer->put_URL(bstrText)))
			{
				bReturn = TRUE;
			}
			SysFreeString(bstrText);
			pIWMPPlayer.Release();
		}
		if (bMute)
		{
			CComPtr<IWMPSettings> pIWMPSettings;
			if ((SUCCEEDED(pUnknown->QueryInterface(__uuidof(IWMPSettings), (VOID**)&pIWMPSettings))))
			{
				pIWMPSettings->put_mute(TRUE);
				pIWMPSettings.Release();
			}
		}
		pUnknown.Release();
	}
	return bReturn;
}

#define REG_KEY TEXT("Software\\VideoScreensaver\\Setting")
class Setting {
	TCHAR m_szFilePath[256];
	DWORD m_dwMute;
public:
	Setting() : m_dwMute(TRUE) {
		m_szFilePath[0] = 0;
	}
	void Load() {
		HKEY hKey;
		DWORD dwPosition;
		if (ERROR_SUCCESS == RegCreateKeyEx(HKEY_CURRENT_USER, REG_KEY, 0, 0, 0, KEY_READ, 0, &hKey, &dwPosition)) {
			DWORD dwType;
			DWORD dwByte;
			dwType = REG_SZ;
			dwByte = sizeof(m_szFilePath);
			RegQueryValueEx(hKey, TEXT("FilePath"), NULL, &dwType, (BYTE *)m_szFilePath, &dwByte);
			dwType = REG_DWORD;
			dwByte = sizeof(DWORD);
			RegQueryValueEx(hKey, TEXT("Mute"), NULL, &dwType, (BYTE *)&m_dwMute, &dwByte);
			RegCloseKey(hKey);
		}
	}
	void Save() {
		HKEY hKey;
		DWORD dwPosition;
		if (ERROR_SUCCESS == RegCreateKeyEx(HKEY_CURRENT_USER, REG_KEY, 0, 0, 0, KEY_WRITE, 0, &hKey, &dwPosition)) {
			RegSetValueEx(hKey, TEXT("FilePath"), 0, REG_SZ, (CONST BYTE *)m_szFilePath, sizeof(TCHAR) * (lstrlen(m_szFilePath) + 1));
			RegSetValueEx(hKey, TEXT("Mute"), 0, REG_DWORD, (CONST BYTE *)&m_dwMute, sizeof(DWORD));
			RegCloseKey(hKey);
		}
	}
	LPTSTR GetFilePath() { return m_szFilePath; }
	void SetFilePath(LPCTSTR lpszText) { lstrcpy(m_szFilePath, lpszText); }
	BOOL GetMute() { return m_dwMute != FALSE; }
	void SetMute(BOOL bMute) { m_dwMute = bMute; }
};

LRESULT WINAPI ScreenSaverProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static Setting setting;
	static HWND hWindowsMediaPlayerControl;
	static std::vector<RECT> MonitorList;
	static std::vector<HTHUMBNAIL> ThumbnailList;
	switch (msg)
	{
	case WM_CREATE:
		EnumDisplayMonitors(NULL, NULL, MonitorEnumProc, (LPARAM)&MonitorList);
		std::sort(MonitorList.begin(), MonitorList.end(), std::greater<RECT>());
		setting.Load();
		AtlAxWinInit();
		_Module.Init(ObjectMap, ((LPCREATESTRUCT)lParam)->hInstance);
		{
			LPOLESTR lpolestr;
			StringFromCLSID(__uuidof(WindowsMediaPlayer), &lpolestr);
			hWindowsMediaPlayerControl = CreateWindow(TEXT(ATLAXWIN_CLASS), lpolestr, WS_POPUP | WS_VISIBLE, 0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
			CoTaskMemFree(lpolestr);
		}
		if (hWindowsMediaPlayerControl)
		{
			DefaultVideoWndProc = (WNDPROC)SetWindowLongPtr(hWindowsMediaPlayerControl, GWL_WNDPROC, (LONG_PTR)MyVideoWndProc);
			CComPtr<IUnknown> pUnknown;
			if (SUCCEEDED(AtlAxGetControl(hWindowsMediaPlayerControl, &pUnknown)))
			{
				CComPtr<IWMPPlayer> pIWMPPlayer;
				if ((SUCCEEDED(pUnknown->QueryInterface(__uuidof(IWMPPlayer), (VOID**)&pIWMPPlayer))))
				{
					BSTR bstrText = SysAllocString(TEXT("none"));
					pIWMPPlayer->put_uiMode(bstrText);
					SysFreeString(bstrText);
					pIWMPPlayer.Release();
				}
				CComPtr<IWMPPlayer2> pIWMPPlayer2;
				if ((SUCCEEDED(pUnknown->QueryInterface(__uuidof(IWMPPlayer2), (VOID**)&pIWMPPlayer2))))
				{
					pIWMPPlayer2->put_stretchToFit(VARIANT_TRUE);
					pIWMPPlayer2.Release();
				}
				CComPtr<IWMPSettings> pIWMPSettings;
				if ((SUCCEEDED(pUnknown->QueryInterface(__uuidof(IWMPSettings), (VOID**)&pIWMPSettings))))
				{
					BSTR bstrText = SysAllocString(TEXT("loop"));
					pIWMPSettings->setMode(bstrText, VARIANT_TRUE);
					SysFreeString(bstrText);
					pIWMPSettings.Release();
				}
				pUnknown.Release();
			}
			for (unsigned int i = 1; i < MonitorList.size(); ++i)
			{
				HTHUMBNAIL thumbnail;
				if (SUCCEEDED(DwmRegisterThumbnail(hWnd, hWindowsMediaPlayerControl, &thumbnail)))
				{
					ThumbnailList.push_back(thumbnail);
				}
			}
			if (PathFileExists(setting.GetFilePath()))
			{
				PlayVideo(hWindowsMediaPlayerControl, setting.GetFilePath(), setting.GetMute());
			}
		}
		break;
	case WM_SIZE:
		MoveWindow(hWindowsMediaPlayerControl, MonitorList[0].left, MonitorList[0].top, MonitorList[0].right - MonitorList[0].left, MonitorList[0].bottom - MonitorList[0].top, TRUE);
		for (unsigned int i = 1; i < MonitorList.size(); ++i)
		{
			RECT dest = MonitorList[i];
			ScreenToClient(hWnd, (LPPOINT)&dest.left);
			ScreenToClient(hWnd, (LPPOINT)&dest.right);
			DWM_THUMBNAIL_PROPERTIES dskThumbProps;
			dskThumbProps.dwFlags = DWM_TNP_RECTDESTINATION | DWM_TNP_VISIBLE | DWM_TNP_SOURCECLIENTAREAONLY;
			dskThumbProps.fSourceClientAreaOnly = FALSE;
			dskThumbProps.fVisible = TRUE;
			dskThumbProps.opacity = 255;
			dskThumbProps.rcDestination = dest;
			DwmUpdateThumbnailProperties(ThumbnailList[i - 1], &dskThumbProps);
		}
		break;
	case WM_DESTROY:
		DestroyWindow(hWindowsMediaPlayerControl);
		AtlAxWinTerm();
		_Module.Term();
		for (auto thumbnail : ThumbnailList)
		{
			DwmUnregisterThumbnail(thumbnail);
		}
		PostQuitMessage(0);
		break;
	default:
		break;
	}
	return DefScreenSaverProc(hWnd, msg, wParam, lParam);
}

BOOL WINAPI ScreenSaverConfigureDialog(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static Setting setting;
	switch (msg)
	{
	case WM_INITDIALOG:
		setting.Load();
		SetDlgItemText(hWnd, IDC_EDIT1, setting.GetFilePath());
		SendDlgItemMessage(hWnd, IDC_CHECK1, BM_SETCHECK, setting.GetMute() ? BST_CHECKED : BST_UNCHECKED, 0);
		return TRUE;
	case WM_COMMAND:
		switch (LOWORD(wParam))
		{
		case IDC_BUTTON1:
		{
			TCHAR szFilePath[MAX_PATH] = { 0 };
			OPENFILENAME of = { 0 };
			of.lStructSize = sizeof(OPENFILENAME);
			of.hwndOwner = hWnd;
			of.lpstrFilter = TEXT("動画ファイル\0*.avi;*.mpg;*.wmv;*.mp4;*.mov;\0すべてのファイル (*.*)\0*.*\0\0");
			of.lpstrFile = szFilePath;
			of.nMaxFile = MAX_PATH;
			of.nMaxFileTitle = MAX_PATH;
			of.Flags = OFN_FILEMUSTEXIST | OFN_HIDEREADONLY;
			of.lpstrTitle = TEXT("動画ファイルの指定");
			if (GetOpenFileName(&of) != 0)
			{
				SetDlgItemText(hWnd, IDC_EDIT1, szFilePath);
			}
		}
		return TRUE;
		case IDOK:
		{
			TCHAR szFilePath[MAX_PATH];
			GetDlgItemText(hWnd, IDC_EDIT1, szFilePath, _countof(szFilePath));
			PathUnquoteSpaces(szFilePath);
			setting.SetFilePath(szFilePath);
			setting.SetMute((BOOL)SendDlgItemMessage(hWnd, IDC_CHECK1, BM_GETCHECK, 0, 0));
			setting.Save();
			EndDialog(hWnd, IDOK);
		}
		return TRUE;
		case IDCANCEL:
			EndDialog(hWnd, IDCANCEL);
			return TRUE;
		}
		return FALSE;
	}
	return FALSE;
}

BOOL WINAPI RegisterDialogClasses(HANDLE hInst)
{
	return TRUE;
}
