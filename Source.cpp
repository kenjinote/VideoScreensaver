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
#include <random>
#include "CWMPEventDispatch.h"
#include "resource.h"

CComModule _Module;
HWND _hMainWindowHandle;
WNDPROC DefaultVideoWndProc;
WNDPROC DefaultListBoxWndProc;

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
			RECT rect = {
				dm.dmPosition.x,
				dm.dmPosition.y,
				dm.dmPosition.x + dm.dmPelsWidth,
				dm.dmPosition.y + dm.dmPelsHeight
			};
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

#define REG_KEY TEXT("Software\\VideoScreensaver\\Setting")
class Setting {
	std::vector<LPTSTR> m_lpszFilePathList;
	DWORD m_dwMute;
	DWORD m_dwRandom;
public:
	Setting() : m_dwMute(TRUE), m_dwRandom(TRUE) {
	}
	~Setting() {
		for (auto item : m_lpszFilePathList) {
			GlobalFree(item);
		}
	}
	void Load() {
		HKEY hKey;
		DWORD dwPosition;
		if (ERROR_SUCCESS == RegCreateKeyEx(HKEY_CURRENT_USER, REG_KEY, 0, 0, 0, KEY_READ, 0, &hKey, &dwPosition)) {
			DWORD dwType;
			DWORD dwByte;
			DWORD nFilePathCount = 0;
			dwType = REG_DWORD;
			dwByte = sizeof(DWORD);
			RegQueryValueEx(hKey, TEXT("FilePathCount"), NULL, &dwType, (BYTE *)&nFilePathCount, &dwByte);
			dwType = REG_SZ;
			for (DWORD i = 0; i < nFilePathCount; ++i) {
				TCHAR szKeyName[16];
				wsprintf(szKeyName, TEXT("FilePath%d"), i);
				if (ERROR_SUCCESS == RegQueryValueEx(hKey, szKeyName, NULL, &dwType, NULL, &dwByte)) {
					LPTSTR lpszFilePath = (LPTSTR)GlobalAlloc(0, dwByte);
					RegQueryValueEx(hKey, szKeyName, NULL, &dwType, (BYTE *)lpszFilePath, &dwByte);
					AddFilePath(lpszFilePath);
				}
			}
			dwType = REG_DWORD;
			dwByte = sizeof(DWORD);
			RegQueryValueEx(hKey, TEXT("Mute"), NULL, &dwType, (BYTE *)&m_dwMute, &dwByte);
			dwType = REG_DWORD;
			dwByte = sizeof(DWORD);
			RegQueryValueEx(hKey, TEXT("Random"), NULL, &dwType, (BYTE *)&m_dwRandom, &dwByte);
			RegCloseKey(hKey);
		}
	}
	void Save() {
		HKEY hKey;
		DWORD dwPosition;
		if (ERROR_SUCCESS == RegCreateKeyEx(HKEY_CURRENT_USER, REG_KEY, 0, 0, 0, KEY_WRITE, 0, &hKey, &dwPosition)) {
			const DWORD nFilePathCount = GetFilePathCount();
			RegSetValueEx(hKey, TEXT("FilePathCount"), 0, REG_DWORD, (CONST BYTE *)&nFilePathCount, sizeof(DWORD));
			for (DWORD i = 0; i < nFilePathCount; ++i)
			{
				TCHAR szKeyName[16];
				wsprintf(szKeyName, TEXT("FilePath%d"), i);
				LPCTSTR lpszFilePath = GetFilePath(i);
				RegSetValueEx(hKey, szKeyName, 0, REG_SZ, (CONST BYTE *)lpszFilePath, sizeof(TCHAR) * (lstrlen(lpszFilePath) + 1));
			}
			RegSetValueEx(hKey, TEXT("Mute"), 0, REG_DWORD, (CONST BYTE *)&m_dwMute, sizeof(DWORD));
			RegSetValueEx(hKey, TEXT("Random"), 0, REG_DWORD, (CONST BYTE *)&m_dwRandom, sizeof(DWORD));
			RegCloseKey(hKey);
		}
	}
	LPTSTR GetFilePath(int nIndex) {
		if (nIndex < 0 || nIndex >= (int)m_lpszFilePathList.size()) return 0;
		return m_lpszFilePathList[nIndex];
	}
	int GetFilePathCount() { return m_lpszFilePathList.size(); }
	void AddFilePath(LPCTSTR lpszText) {
		const int nSize = lstrlen(lpszText);
		LPTSTR lpszFilePath = (LPTSTR)GlobalAlloc(0, (nSize + 1) * sizeof(TCHAR));
		lstrcpy(lpszFilePath, lpszText);
		m_lpszFilePathList.push_back(lpszFilePath);
	}
	BOOL GetMute() { return m_dwMute != FALSE; }
	void SetMute(BOOL bMute) { m_dwMute = bMute; }
	BOOL GetRandom() { return m_dwRandom != FALSE; }
	void SetRandom(BOOL bRandom) { m_dwRandom = bRandom; }
	void ClearFilePath() {
		for (auto item : m_lpszFilePathList) {
			GlobalFree(item);
		}
		m_lpszFilePathList.clear();
	}
	void Shuffle() {
		std::random_device seed_gen;
		std::mt19937 engine(seed_gen());
		std::shuffle(m_lpszFilePathList.begin(), m_lpszFilePathList.end(), engine);
	}
};

LRESULT WINAPI ScreenSaverProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static Setting setting;
	static HWND hWindowsMediaPlayerControl;
	static std::vector<RECT> MonitorList;
	static std::vector<HTHUMBNAIL> ThumbnailList;
	static BOOL bPreviewMode;
	static CComWMPEventDispatch *m_pEventListener;
	static CComPtr<IConnectionPoint>   m_spConnectionPoint;
	static DWORD                       m_dwAdviseCookie;
	static CComPtr<IWMPEvents>         m_spEventListener;
	switch (msg)
	{
	case WM_CREATE:
		_hMainWindowHandle = hWnd;
		{
			int n;
			LPTSTR* argv = CommandLineToArgvW(GetCommandLine(), &n);
			if (argv)
			{
				for (int i = 1; i < n; ++i)
				{
					if (lstrcmpi(argv[i], L"/P") == 0 || lstrcmpi(argv[i], L"/L") == 0)
					{
						bPreviewMode = TRUE;
						break;
					}
				}
				LocalFree(argv);
			}
		}
		if (!bPreviewMode)
		{
			EnumDisplayMonitors(NULL, NULL, MonitorEnumProc, (LPARAM)&MonitorList);
			std::sort(MonitorList.begin(), MonitorList.end(), std::greater<RECT>());
		}
		setting.Load();
		AtlAxWinInit();
		_Module.Init(ObjectMap, ((LPCREATESTRUCT)lParam)->hInstance);
		{
			LPOLESTR lpolestr;
			StringFromCLSID(__uuidof(WindowsMediaPlayer), &lpolestr);
			hWindowsMediaPlayerControl = CreateWindow(TEXT(ATLAXWIN_CLASS), lpolestr, (bPreviewMode ? WS_CHILD | WS_DISABLED : WS_POPUP) | WS_VISIBLE, 0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
			CoTaskMemFree(lpolestr);
		}
		if (hWindowsMediaPlayerControl)
		{
			if (!bPreviewMode)
			{
				DefaultVideoWndProc = (WNDPROC)SetWindowLongPtr(hWindowsMediaPlayerControl, GWL_WNDPROC, (LONG_PTR)MyVideoWndProc);
				for (unsigned int i = 1; i < MonitorList.size(); ++i)
				{
					HTHUMBNAIL thumbnail;
					if (SUCCEEDED(DwmRegisterThumbnail(hWnd, hWindowsMediaPlayerControl, &thumbnail)))
					{
						ThumbnailList.push_back(thumbnail);
					}
				}
			}
			CComPtr<IUnknown> pUnknown;
			if (SUCCEEDED(AtlAxGetControl(hWindowsMediaPlayerControl, &pUnknown)))
			{
				CComPtr<IWMPSettings> pIWMPSettings;
				if ((SUCCEEDED(pUnknown->QueryInterface(__uuidof(IWMPSettings), (VOID**)&pIWMPSettings))))
				{
					pIWMPSettings->put_autoStart(VARIANT_FALSE);
					BSTR bstrText = SysAllocString(TEXT("loop"));
					pIWMPSettings->setMode(bstrText, VARIANT_TRUE);
					SysFreeString(bstrText);
					if (setting.GetRandom())
					{
						setting.Shuffle();
						bstrText = SysAllocString(TEXT("shuffle"));
						pIWMPSettings->setMode(bstrText, VARIANT_TRUE);
						SysFreeString(bstrText);
					}
					pIWMPSettings.Release();
				}
				CComPtr<IWMPPlayer4> pIWMPPlayer;
				if ((SUCCEEDED(pUnknown->QueryInterface(__uuidof(IWMPPlayer4), (VOID**)&pIWMPPlayer))))
				{
					pIWMPPlayer->put_stretchToFit(VARIANT_TRUE);
					BSTR bstrText = SysAllocString(TEXT("none"));
					pIWMPPlayer->put_uiMode(bstrText);
					SysFreeString(bstrText);

					const int nFilePathCount = setting.GetFilePathCount();
					if (nFilePathCount > 0)
					{
						CComPtr<IWMPPlaylist> pIWMPPlaylist;
						if ((SUCCEEDED(pIWMPPlayer->get_currentPlaylist(&pIWMPPlaylist))))
						{
							for (int i = 0; i < nFilePathCount; ++i)
							{
								CComPtr<IWMPMedia> pIWMPMedia;
								if ((SUCCEEDED(pIWMPPlayer->newMedia(setting.GetFilePath(i), &pIWMPMedia))))
								{
									pIWMPPlaylist->appendItem(pIWMPMedia);
									pIWMPMedia.Release();
								}
							}
							pIWMPPlaylist.Release();
						}
					}
					HRESULT hr = CComWMPEventDispatch::CreateInstance(&m_pEventListener);
					m_spEventListener = m_pEventListener;
					if (SUCCEEDED(hr))
					{
						CComPtr<IConnectionPointContainer> spConnectionContainer;
						hr = pIWMPPlayer->QueryInterface(&spConnectionContainer);
						if (SUCCEEDED(hr))
						{
							hr = spConnectionContainer->FindConnectionPoint(__uuidof(IWMPEvents), &m_spConnectionPoint);
							if (FAILED(hr))
							{
								hr = spConnectionContainer->FindConnectionPoint(__uuidof(_WMPOCXEvents), &m_spConnectionPoint);
							}
							spConnectionContainer.Release();
						}
						if (SUCCEEDED(hr))
						{
							hr = m_spConnectionPoint->Advise(m_spEventListener, &m_dwAdviseCookie);
						}
					}
					pIWMPPlayer.Release();
				}
				CComPtr<IWMPControls> pIWMPControls;
				if ((SUCCEEDED(pUnknown->QueryInterface(__uuidof(IWMPControls), (VOID**)&pIWMPControls))))
				{
					pIWMPControls->play();
					pIWMPControls.Release();
				}
				pUnknown.Release();
			}
		}
		break;
	case WM_APP:
		if (setting.GetMute())
		{
			CComPtr<IUnknown> pUnknown;
			if (SUCCEEDED(AtlAxGetControl(hWindowsMediaPlayerControl, &pUnknown)))
			{
				CComPtr<IWMPSettings> pIWMPSettings;
				if ((SUCCEEDED(pUnknown->QueryInterface(__uuidof(IWMPSettings), (VOID**)&pIWMPSettings))))
				{
					pIWMPSettings->put_mute(VARIANT_TRUE);
					pIWMPSettings.Release();
				}
				pUnknown.Release();
			}
		}
		break;
	case WM_SIZE:
		if (bPreviewMode)
		{
			RECT rect;
			GetClientRect(hWnd, &rect);
			MoveWindow(hWindowsMediaPlayerControl, 0, 0, rect.right, rect.bottom, TRUE);
		}
		else
		{
			MoveWindow(hWindowsMediaPlayerControl, MonitorList[0].left, MonitorList[0].top, MonitorList[0].right - MonitorList[0].left, MonitorList[0].bottom - MonitorList[0].top, TRUE);
			for (unsigned int i = 0; i < MonitorList.size() - 1; ++i)
			{
				RECT dest = MonitorList[i + 1];
				ScreenToClient(hWnd, (LPPOINT)&dest.left);
				ScreenToClient(hWnd, (LPPOINT)&dest.right);
				DWM_THUMBNAIL_PROPERTIES dskThumbProps;
				dskThumbProps.dwFlags = DWM_TNP_RECTDESTINATION | DWM_TNP_VISIBLE | DWM_TNP_SOURCECLIENTAREAONLY;
				dskThumbProps.fSourceClientAreaOnly = FALSE;
				dskThumbProps.fVisible = TRUE;
				dskThumbProps.opacity = 255;
				dskThumbProps.rcDestination = dest;
				DwmUpdateThumbnailProperties(ThumbnailList[i], &dskThumbProps);
			}
		}
		break;
	case WM_DESTROY:
		if (m_spConnectionPoint)
		{
			if (0 != m_dwAdviseCookie)
				m_spConnectionPoint->Unadvise(m_dwAdviseCookie);
			//m_spConnectionPoint.Release(); // Release すると終了時に落ちる
		}
		if (m_spEventListener)
		{
			m_spEventListener.Release();
		}
		DestroyWindow(hWindowsMediaPlayerControl);
		AtlAxWinTerm();
		_Module.Term();
		if (!bPreviewMode)
		{
			for (auto thumbnail : ThumbnailList)
			{
				DwmUnregisterThumbnail(thumbnail);
			}
		}
		PostQuitMessage(0);
		break;
	default:
		break;
	}
	return DefScreenSaverProc(hWnd, msg, wParam, lParam);
}

LRESULT CALLBACK MyListBoxProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	switch (msg)
	{
	case WM_KEYDOWN:
		if (wParam == VK_DELETE)
		{
			PostMessage(GetParent(hWnd), WM_COMMAND, IDC_BUTTON3, 0);
		}
		else if (wParam == 'A' && GetAsyncKeyState(VK_CONTROL) < 0)
		{
			SendDlgItemMessage(GetParent(hWnd), IDC_LIST1, LB_SETSEL, 1, -1);
		}
		break;
	}
	return CallWindowProc(DefaultListBoxWndProc, hWnd, msg, wParam, lParam);
}

BOOL WINAPI ScreenSaverConfigureDialog(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static Setting setting;
	switch (msg)
	{
	case WM_INITDIALOG:
		setting.Load();
		{
			const int nFilePathCount = setting.GetFilePathCount();
			for (int i = 0; i < nFilePathCount; ++i)
			{
				LPTSTR lpszFilePath = setting.GetFilePath(i);
				if (lpszFilePath)
				{
					const int nIndex = (int)SendDlgItemMessage(hWnd, IDC_LIST1, LB_ADDSTRING, 0, (LPARAM)lpszFilePath);
					SendDlgItemMessage(hWnd, IDC_LIST1, LB_SETSEL, 1, nIndex);
				}
			}
			PostMessage(hWnd, WM_COMMAND, MAKEWPARAM(IDC_LIST1, LBN_SELCHANGE), 0);
		}
		SendDlgItemMessage(hWnd, IDC_CHECK1, BM_SETCHECK, setting.GetMute() ? BST_CHECKED : BST_UNCHECKED, 0);
		SendDlgItemMessage(hWnd, IDC_CHECK2, BM_SETCHECK, setting.GetRandom() ? BST_CHECKED : BST_UNCHECKED, 0);
		ChangeWindowMessageFilterEx(hWnd, WM_DROPFILES, MSGFLT_ALLOW, 0);
		ChangeWindowMessageFilterEx(hWnd, /*WM_COPYGLOBALDATA*/ 0x0049, MSGFLT_ALLOW, 0);
		DefaultListBoxWndProc = (WNDPROC)SetWindowLongPtr(GetDlgItem(hWnd, IDC_LIST1), GWLP_WNDPROC, (LONG_PTR)MyListBoxProc);
		return TRUE;
	case WM_DROPFILES:
		{
			SendDlgItemMessage(hWnd, IDC_LIST1, LB_SETSEL, 0, -1);
			const UINT nFileCount = DragQueryFile((HDROP)wParam, 0xFFFFFFFF, NULL, 0);
			for (UINT i = 0; i < nFileCount; ++i)
			{
				TCHAR szFilePath[MAX_PATH];
				DragQueryFile((HDROP)wParam, i, szFilePath, _countof(szFilePath));
				if (PathMatchSpec(szFilePath, TEXT("*.avi;*.mpg;*.wmv;*.mp4;*.mov;")))
				{
					const int nIndex = (int)SendDlgItemMessage(hWnd, IDC_LIST1, LB_ADDSTRING, 0, (LPARAM)szFilePath);
					SendDlgItemMessage(hWnd, IDC_LIST1, LB_SETSEL, 1, nIndex);
				}
			}
			DragFinish((HDROP)wParam);
			PostMessage(hWnd, WM_COMMAND, MAKEWPARAM(IDC_LIST1, LBN_SELCHANGE), 0);
		}
		return TRUE;
	case WM_COMMAND:
		switch (LOWORD(wParam))
		{
		case IDC_BUTTON2:
			{
				#define MAX_CFileDialog_FILE_COUNT 99
				#define FILE_LIST_BUFFER_SIZE ((MAX_CFileDialog_FILE_COUNT * (MAX_PATH + 1)) + 1)
				LPTSTR lpszFilePath = (LPTSTR)GlobalAlloc(GMEM_ZEROINIT, sizeof(TCHAR) * FILE_LIST_BUFFER_SIZE);
				OPENFILENAME of = { 0 };
				of.lStructSize = sizeof(OPENFILENAME);
				of.hwndOwner = hWnd;
				of.lpstrFilter = TEXT("動画ファイル\0*.avi;*.mpg;*.wmv;*.mp4;*.mov;\0すべてのファイル (*.*)\0*.*\0\0");
				of.lpstrFile = lpszFilePath;
				of.nMaxFile = FILE_LIST_BUFFER_SIZE;
				of.nMaxFileTitle = MAX_PATH;
				of.Flags = OFN_FILEMUSTEXIST | OFN_HIDEREADONLY | OFN_ALLOWMULTISELECT | OFN_EXPLORER;
				of.lpstrTitle = TEXT("動画ファイルの指定");
				if (GetOpenFileName(&of))
				{
					SendDlgItemMessage(hWnd, IDC_LIST1, LB_SETSEL, 0, -1);
					if (PathIsDirectory(lpszFilePath))
					{
						TCHAR szDirectory[MAX_PATH];
						lstrcpy(szDirectory, lpszFilePath);
						LPTSTR p = lpszFilePath;
						while (*(p += lstrlen(p) + 1) != TEXT('\0'))
						{
							TCHAR szFilePath[MAX_PATH];
							lstrcpy(szFilePath, szDirectory);
							PathAppend(szFilePath, p);
							const int nIndex = (int)SendDlgItemMessage(hWnd, IDC_LIST1, LB_ADDSTRING, 0, (LPARAM)szFilePath);
							SendDlgItemMessage(hWnd, IDC_LIST1, LB_SETSEL, 1, nIndex);
						}
					}
					else if (PathFileExists(lpszFilePath))
					{
						const int nIndex = (int)SendDlgItemMessage(hWnd, IDC_LIST1, LB_ADDSTRING, 0, (LPARAM)lpszFilePath);
						SendDlgItemMessage(hWnd, IDC_LIST1, LB_SETSEL, 1, nIndex);
					}
					PostMessage(hWnd, WM_COMMAND, MAKEWPARAM(IDC_LIST1, LBN_SELCHANGE), 0);
				}
				GlobalFree(lpszFilePath);
			}
			return TRUE;
		case IDC_BUTTON3:
			{
				const int nSelCount = (int)SendDlgItemMessage(hWnd, IDC_LIST1, LB_GETSELCOUNT, 0, 0);
				if (nSelCount > 0)
				{
					int * nSelItems = (int*)GlobalAlloc(0, sizeof(int)*nSelCount);
					SendDlgItemMessage(hWnd, IDC_LIST1, LB_GETSELITEMS, nSelCount, (LPARAM)nSelItems);
					for (int i = nSelCount - 1; i >= 0; --i)
					{
						SendDlgItemMessage(hWnd, IDC_LIST1, LB_DELETESTRING, nSelItems[i], 0);
					}
					const int nGetCount = SendDlgItemMessage(hWnd, IDC_LIST1, LB_GETCOUNT, 0, 0);
					if (nGetCount > 0)
					{
						SendDlgItemMessage(hWnd, IDC_LIST1, LB_SETSEL, TRUE, (nSelCount == 1) ? min(nSelItems[0], nGetCount - 1) : 0);
						SetFocus(GetDlgItem(hWnd, IDC_LIST1));
					}
					GlobalFree(nSelItems);
					PostMessage(hWnd, WM_COMMAND, MAKEWPARAM(IDC_LIST1, LBN_SELCHANGE), 0);
				}
			}
			return TRUE;
		case IDC_LIST1:
			if (HIWORD(wParam) == LBN_SELCHANGE)
			{
				const int nSelCount = (int)SendDlgItemMessage(hWnd, IDC_LIST1, LB_GETSELCOUNT, 0, 0);
				EnableWindow(GetDlgItem(hWnd, IDC_BUTTON3), nSelCount > 0);
			}
			return TRUE;
		case IDOK:
			{
				setting.ClearFilePath();
				{
					const int nItemCount = (int)SendDlgItemMessage(hWnd, IDC_LIST1, LB_GETCOUNT, 0, 0);
					for (int i = 0; i < nItemCount; ++i)
					{
						TCHAR szFilePath[MAX_PATH];
						SendDlgItemMessage(hWnd, IDC_LIST1, LB_GETTEXT, i, (LPARAM)szFilePath);
						setting.AddFilePath(szFilePath);
					}
				}
				setting.SetMute((BOOL)SendDlgItemMessage(hWnd, IDC_CHECK1, BM_GETCHECK, 0, 0));
				setting.SetRandom((BOOL)SendDlgItemMessage(hWnd, IDC_CHECK2, BM_GETCHECK, 0, 0));
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
