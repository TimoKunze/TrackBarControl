// MainDlg.cpp : implementation of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "resource.h"

#include "MainDlg.h"
#include ".\maindlg.h"


BOOL CMainDlg::PreTranslateMessage(MSG* pMsg)
{
	return CWindow::IsDialogMessage(pMsg);
}

LRESULT CMainDlg::OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
	CloseDialog(0);
	return 0;
}

LRESULT CMainDlg::OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled)
{
	if(controls.trackBarU) {
		DispEventUnadvise(controls.trackBarU);
		controls.trackBarU.Release();
	}

	// unregister message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->RemoveMessageFilter(this);

	bHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
	// set icons
	HICON hIcon = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), IMAGE_ICON, GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON), LR_DEFAULTCOLOR));
	SetIcon(hIcon, TRUE);
	HICON hIconSmall = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), LR_DEFAULTCOLOR));
	SetIcon(hIconSmall, FALSE);

	dialogBackgroundTexture = LoadBitmap(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDB_BACKGROUND));
	dialogBackgroundBrush.CreatePatternBrush(dialogBackgroundTexture);

	// register object for message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->AddMessageFilter(this);

	trackBarUWnd.SubclassWindow(GetDlgItem(IDC_TRACKBARU));
	trackBarUWnd.QueryControl(__uuidof(ITrackBar), reinterpret_cast<LPVOID*>(&controls.trackBarU));
	if(controls.trackBarU) {
		DispEventAdvise(controls.trackBarU);
	}

	return TRUE;
}

LRESULT CMainDlg::OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(controls.trackBarU) {
		controls.trackBarU->About();
	}
	return 0;
}

LRESULT CMainDlg::OnCtlColorDlg(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
	return reinterpret_cast<LRESULT>(static_cast<HBRUSH>(dialogBackgroundBrush));
}

LRESULT CMainDlg::OnCtlColorStatic(UINT /*uMsg*/, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	if(LongToHandle(controls.trackBarU->GethWnd()) == reinterpret_cast<HWND>(lParam)){
		if(!trackbarBackgroundBrush.IsNull()) {
			trackbarBackgroundBrush.DeleteObject();
		}

		WTL::CRect windowRectangle;
		::GetWindowRect(reinterpret_cast<HWND>(lParam), &windowRectangle);
		ScreenToClient(&windowRectangle);
		CBitmap trackbarBackgroundTexture;
		CDC trackbarBkDC;
		trackbarBkDC.CreateCompatibleDC(reinterpret_cast<HDC>(wParam));
		trackbarBackgroundTexture.CreateCompatibleBitmap(reinterpret_cast<HDC>(wParam), windowRectangle.Width(), windowRectangle.Height());
		HBITMAP hPreviousBitmap1 = trackbarBkDC.SelectBitmap(trackbarBackgroundTexture);

		CDC dialogBkDC;
		dialogBkDC.CreateCompatibleDC(reinterpret_cast<HDC>(wParam));
		HBITMAP hPreviousBitmap2 = dialogBkDC.SelectBitmap(dialogBackgroundTexture);

		trackbarBkDC.BitBlt(0, 0, windowRectangle.Width(), windowRectangle.Height(), dialogBkDC, windowRectangle.left, windowRectangle.top, SRCCOPY);

		dialogBkDC.SelectBitmap(hPreviousBitmap2);
		trackbarBkDC.SelectBitmap(hPreviousBitmap1);

		trackbarBackgroundBrush.CreatePatternBrush(trackbarBackgroundTexture);
		return reinterpret_cast<LRESULT>(static_cast<HBRUSH>(trackbarBackgroundBrush));
	}
	bHandled = FALSE;
	return 0;
}

void CMainDlg::CloseDialog(int nVal)
{
	DestroyWindow();
	PostQuitMessage(nVal);
}
