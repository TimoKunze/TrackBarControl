// MainDlg.cpp : implementation of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#include <olectl.h>
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

void CMainDlg::CloseDialog(int nVal)
{
	DestroyWindow();
	PostQuitMessage(nVal);
}

void __stdcall CMainDlg::CustomDrawTrackbaru(CustomDrawControlPartConstants controlPart, CustomDrawStageConstants drawStage, CustomDrawControlStateConstants /*controlState*/, long hDC, RECTANGLE* /*drawingRectangle*/, CustomDrawReturnValuesConstants* furtherProcessing)
{
	if(drawStage == cdsPrePaint) {
		*furtherProcessing = cdrvNotifyItemDraw;
	} else if(controlPart == cdcpTics) {
		switch(drawStage) {
			case cdsItemPrePaint:
				*furtherProcessing = cdrvNotifyPostPaint;
				break;
			case cdsItemPostPaint: {
				_variant_t v = controls.trackBarU->GetTickMarks();
				if(v.vt & VT_ARRAY) {
					SetBkMode(static_cast<HDC>(LongToHandle(hDC)), TRANSPARENT);
					HFONT hFont = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
					HFONT hPreviousFont = static_cast<HFONT>(SelectObject(static_cast<HDC>(LongToHandle(hDC)), hFont));

					LONG l = 0;
					SafeArrayGetLBound(v.parray, 1, &l);
					LONG u = 0;
					SafeArrayGetUBound(v.parray, 1, &u);
					for(LONG i = l; i <= u; i++) {
						LONG buffer = 0;
						SafeArrayGetElement(v.parray, &i, &buffer);
						CAtlString s;
						s.Format(TEXT("%i"), buffer);

						CRect textRectangle;
						DrawText(static_cast<HDC>(LongToHandle(hDC)), s, -1, &textRectangle, DT_CALCRECT | DT_SINGLELINE);
						LONG xCenter = controls.trackBarU->GetTickMarkPhysicalPosition(i - l);
						textRectangle.MoveToXY(xCenter - textRectangle.Width() / 2, controls.trackBarU->GetChannelTop() + controls.trackBarU->GetChannelHeight() + 16);

						DrawText(static_cast<HDC>(LongToHandle(hDC)), s, -1, &textRectangle, DT_SINGLELINE);
					}

					SelectObject(static_cast<HDC>(LongToHandle(hDC)), hPreviousFont);
				}
				*furtherProcessing = cdrvDoDefault;
				break;
			}
		}
	}
}
