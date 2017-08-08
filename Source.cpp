#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#pragma comment(lib, "scrnsavw")
#pragma comment(lib, "comctl32")

#include <windows.h>
#include <scrnsave.h>
#include <atlbase.h>
#include <atlwin.h>
#include <wmp.h>
#include "resource.h"

CComModule _Module;

BEGIN_OBJECT_MAP(ObjectMap)
END_OBJECT_MAP()

BOOL PlayVideo(HWND hWnd, LPCTSTR lpszFilePath)
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
		pUnknown.Release();
	}
	return bReturn;
}

#define REG_KEY TEXT("Software\\VideoScreensaver\\Setting")
class Setting {
	TCHAR m_szFilePath[256];
public:
	Setting() {
		lstrcpy(m_szFilePath, TEXT("Sample.mp4"));
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
			RegCloseKey(hKey);
		}
	}
	void Save() {
		HKEY hKey;
		DWORD dwPosition;
		if (ERROR_SUCCESS == RegCreateKeyEx(HKEY_CURRENT_USER, REG_KEY, 0, 0, 0, KEY_WRITE, 0, &hKey, &dwPosition)) {
			RegSetValueEx(hKey, TEXT("FilePath"), 0, REG_SZ, (CONST BYTE *)m_szFilePath, sizeof(TCHAR) * (lstrlen(m_szFilePath) + 1));
			RegCloseKey(hKey);
		}
	}
	LPTSTR GetFilePath() { return m_szFilePath; }
	void SetFilePath(LPCTSTR lpszText) { lstrcpy(m_szFilePath, lpszText); }
};

LRESULT WINAPI ScreenSaverProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static Setting setting;
	static HWND hWindowsMediaPlayerControl;
	switch (msg)
	{
	case WM_CREATE:
		AtlAxWinInit();
		_Module.Init(ObjectMap, ((LPCREATESTRUCT)lParam)->hInstance);
		{
			LPOLESTR lpolestr;
			StringFromCLSID(__uuidof(WindowsMediaPlayer), &lpolestr);
			hWindowsMediaPlayerControl = CreateWindow(TEXT(ATLAXWIN_CLASS), lpolestr, WS_CHILD | WS_VISIBLE | WS_DISABLED, 0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
			CoTaskMemFree(lpolestr);
		}
		if (hWindowsMediaPlayerControl)
		{
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
			setting.Load();
			if (PathFileExists(setting.GetFilePath()))
			{
				PlayVideo(hWindowsMediaPlayerControl, setting.GetFilePath());
			}
		}
		break;
	case WM_SIZE:
		MoveWindow(hWindowsMediaPlayerControl, 0, 0, LOWORD(lParam), HIWORD(lParam), 1);
		break;
	case WM_DESTROY:
		DestroyWindow(hWindowsMediaPlayerControl);
		AtlAxWinTerm();
		_Module.Term();
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
