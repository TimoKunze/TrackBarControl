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
		IDispEventImpl<IDC_TRACKBARU, CMainDlg, &__uuidof(TrackBarCtlLibU::_ITrackBarEvents), &LIBID_TrackBarCtlLibU, 1, 7>::DispEventUnadvise(controls.trackBarU);
		controls.trackBarU.Release();
	}
	if(controls.trackBarA) {
		IDispEventImpl<IDC_TRACKBARA, CMainDlg, &__uuidof(TrackBarCtlLibA::_ITrackBarEvents), &LIBID_TrackBarCtlLibA, 1, 7>::DispEventUnadvise(controls.trackBarA);
		controls.trackBarA.Release();
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
	// Init resizing
	DlgResize_Init(false, false);

	// set icons
	HICON hIcon = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), IMAGE_ICON, GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON), LR_DEFAULTCOLOR));
	SetIcon(hIcon, TRUE);
	HICON hIconSmall = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), LR_DEFAULTCOLOR));
	SetIcon(hIconSmall, FALSE);

	// register object for message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->AddMessageFilter(this);

	controls.logEdit = GetDlgItem(IDC_EDITLOG);
	controls.aboutButton = GetDlgItem(ID_APP_ABOUT);

	trackBarUContainerWnd.SubclassWindow(GetDlgItem(IDC_TRACKBARU));
	trackBarUContainerWnd.QueryControl(__uuidof(TrackBarCtlLibU::ITrackBar), reinterpret_cast<LPVOID*>(&controls.trackBarU));
	if(controls.trackBarU) {
		IDispEventImpl<IDC_TRACKBARU, CMainDlg, &__uuidof(TrackBarCtlLibU::_ITrackBarEvents), &LIBID_TrackBarCtlLibU, 1, 7>::DispEventAdvise(controls.trackBarU);
		trackBarUWnd.SubclassWindow(static_cast<HWND>(LongToHandle(controls.trackBarU->GethWnd())));
	}

	trackBarAContainerWnd.SubclassWindow(GetDlgItem(IDC_TRACKBARA));
	trackBarAContainerWnd.QueryControl(__uuidof(TrackBarCtlLibA::ITrackBar), reinterpret_cast<LPVOID*>(&controls.trackBarA));
	if(controls.trackBarA) {
		IDispEventImpl<IDC_TRACKBARA, CMainDlg, &__uuidof(TrackBarCtlLibA::_ITrackBarEvents), &LIBID_TrackBarCtlLibA, 1, 7>::DispEventAdvise(controls.trackBarA);
		trackBarAWnd.SubclassWindow(static_cast<HWND>(LongToHandle(controls.trackBarA->GethWnd())));
	}

	// force control resize
	WINDOWPOS dummy = {0};
	BOOL b = FALSE;
	OnWindowPosChanged(WM_WINDOWPOSCHANGED, 0, reinterpret_cast<LPARAM>(&dummy), b);

	return TRUE;
}

LRESULT CMainDlg::OnWindowPosChanged(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*bHandled*/)
{
	if(controls.logEdit && controls.aboutButton) {
		LPWINDOWPOS pDetails = reinterpret_cast<LPWINDOWPOS>(lParam);

		if((pDetails->flags & SWP_NOSIZE) == 0) {
			WTL::CRect clientRectangle;
			GetClientRect(&clientRectangle);
			WTL::CRect aboutButtonRectangle;
			controls.aboutButton.GetWindowRect(&aboutButtonRectangle);
			ScreenToClient(&aboutButtonRectangle);
			WTL::CRect controlRectangle;
			trackBarUContainerWnd.GetWindowRect(&controlRectangle);
			ScreenToClient(&controlRectangle);

			int trackBarWidth = static_cast<int>(0.5 * static_cast<double>(clientRectangle.Width()));
			int trackBarHeight = controlRectangle.Height();
			controls.aboutButton.SetWindowPos(NULL, clientRectangle.right - aboutButtonRectangle.Width() - 8, 0, 0, 0, SWP_NOSIZE);
			controls.logEdit.SetWindowPos(NULL, 0, 0, clientRectangle.right - aboutButtonRectangle.Width() - 16, clientRectangle.bottom - trackBarHeight - 8, SWP_NOMOVE);
			trackBarUContainerWnd.SetWindowPos(NULL, 0, clientRectangle.bottom - trackBarHeight, trackBarWidth, trackBarHeight, 0);
			trackBarAContainerWnd.SetWindowPos(NULL, trackBarWidth, clientRectangle.bottom - trackBarHeight, trackBarWidth, trackBarHeight, 0);
		}
	}

	return 0;
}

LRESULT CMainDlg::OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(trackBarAIsFocused) {
		if(controls.trackBarA) {
			controls.trackBarA->About();
		}
	} else {
		if(controls.trackBarU) {
			controls.trackBarU->About();
		}
	}
	return 0;
}

LRESULT CMainDlg::OnSetFocusU(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled)
{
	trackBarAIsFocused = false;
	bHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnSetFocusA(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled)
{
	trackBarAIsFocused = true;
	bHandled = FALSE;
	return 0;
}

void CMainDlg::AddLogEntry(CAtlString text)
{
	static int cLines = 0;
	static CAtlString oldText;

	cLines++;
	if(cLines > 50) {
		// delete the first line
		int pos = oldText.Find(TEXT("\r\n"));
		oldText = oldText.Mid(pos + lstrlen(TEXT("\r\n")), oldText.GetLength());
		cLines--;
	}
	oldText += text;
	oldText += TEXT("\r\n");

	controls.logEdit.SetWindowText(oldText);
	int l = oldText.GetLength();
	controls.logEdit.SetSel(l, l);
}

void CMainDlg::CloseDialog(int nVal)
{
	DestroyWindow();
	PostQuitMessage(nVal);
}

void __stdcall CMainDlg::ClickTrackbaru(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TrackBarU_Click: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuTrackbaru(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TrackBarU_ContextMenu: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CustomDrawTrackbaru(TrackBarCtlLibU::CustomDrawControlPartConstants controlPart, TrackBarCtlLibU::CustomDrawStageConstants drawStage, TrackBarCtlLibU::CustomDrawControlStateConstants controlState, long hDC, TrackBarCtlLibU::RECTANGLE* drawingRectangle, TrackBarCtlLibU::CustomDrawReturnValuesConstants* furtherProcessing)
{
	CAtlString str;
	str.Format(TEXT("TrackBarU_CustomDraw: controlPart=0x%X, drawStage=0x%X, controlState=0x%X, hDC=0x%X, drawingRectangle=(%i, %i)-(%i, %i), furtherProcessing=0x%X"), controlPart, drawStage, controlState, hDC, drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom, *furtherProcessing);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickTrackbaru(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TrackBarU_DblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowTrackbaru(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("TrackBarU_DestroyedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::GetInfoTipTextTrackbaru(long maxInfoTipLength, BSTR* infoTipText, VARIANT_BOOL* abortToolTip)
{
	CAtlString str;
	str.Format(TEXT("TrackBarU_GetInfoTipText: maxInfoTipLength=%i, infoTipText=%s, abortToolTip=%i"), maxInfoTipLength, OLE2W(*infoTipText), *abortToolTip);

	AddLogEntry(str);

	CAtlString tmp;
	tmp.Format(TEXT("Current position: %i"), controls.trackBarU->GetCurrentPosition());

	*infoTipText = _bstr_t(tmp).Detach();
}

void __stdcall CMainDlg::KeyDownTrackbaru(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("TrackBarU_KeyDown: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyPressTrackbaru(short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("TrackBarU_KeyPress: keyAscii=%i"), *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpTrackbaru(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("TrackBarU_KeyUp: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickTrackbaru(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TrackBarU_MClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickTrackbaru(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TrackBarU_MDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownTrackbaru(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TrackBarU_MouseDown: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterTrackbaru(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TrackBarU_MouseEnter: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverTrackbaru(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TrackBarU_MouseHover: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveTrackbaru(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TrackBarU_MouseLeave: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveTrackbaru(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TrackBarU_MouseMove: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpTrackbaru(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TrackBarU_MouseUp: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseWheelTrackbaru(short button, short shift, long x, long y, TrackBarCtlLibU::ScrollAxisConstants scrollAxis, short wheelDelta)
{
	CAtlString str;
	str.Format(TEXT("TrackBarU_MouseWheel: button=%i, shift=%i, x=%i, y=%i, scrollAxis=0x%X, wheelDelta=%i"), button, shift, x, y, scrollAxis, wheelDelta);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropTrackbaru(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<TrackBarCtlLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TrackBarU_OLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TrackBarU_OLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.trackBarU->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterTrackbaru(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<TrackBarCtlLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TrackBarU_OLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TrackBarU_OLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveTrackbaru(LPDISPATCH data, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<TrackBarCtlLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TrackBarU_OLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TrackBarU_OLEDragLeave: data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragMouseMoveTrackbaru(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<TrackBarCtlLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TrackBarU_OLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TrackBarU_OLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::PositionChangedTrackbaru(TrackBarCtlLibU::PositionChangeTypeConstants changeType, LONG newPosition)
{
	CAtlString str;
	str.Format(TEXT("TrackBarU_PositionChanged: changeType=%i, newPosition=%i"), changeType, newPosition);

	AddLogEntry(str);
}

void __stdcall CMainDlg::PositionChangingTrackbaru(TrackBarCtlLibU::PositionChangeTypeConstants changeType, LONG newPosition, VARIANT_BOOL* cancelChange)
{
	CAtlString str;
	str.Format(TEXT("TrackBarU_PositionChanging: changeType=%i, newPosition=%i, cancelChange=%i"), changeType, newPosition, *cancelChange);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RClickTrackbaru(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TrackBarU_RClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickTrackbaru(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TrackBarU_RDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowTrackbaru(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("TrackBarU_RecreatedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowTrackbaru()
{
	AddLogEntry(CAtlString(TEXT("TrackBarU_ResizedControlWindow")));
}

void __stdcall CMainDlg::XClickTrackbaru(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TrackBarU_XClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickTrackbaru(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TrackBarU_XDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}


void __stdcall CMainDlg::ClickTrackbara(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TrackBarA_Click: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuTrackbara(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TrackBarA_ContextMenu: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CustomDrawTrackbara(TrackBarCtlLibA::CustomDrawControlPartConstants controlPart, TrackBarCtlLibA::CustomDrawStageConstants drawStage, TrackBarCtlLibA::CustomDrawControlStateConstants controlState, long hDC, TrackBarCtlLibA::RECTANGLE* drawingRectangle, TrackBarCtlLibA::CustomDrawReturnValuesConstants* furtherProcessing)
{
	CAtlString str;
	str.Format(TEXT("TrackBarA_CustomDraw: controlPart=0x%X, drawStage=0x%X, controlState=0x%X, hDC=0x%X, drawingRectangle=(%i, %i)-(%i, %i), furtherProcessing=0x%X"), controlPart, drawStage, controlState, hDC, drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom, *furtherProcessing);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickTrackbara(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TrackBarA_DblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowTrackbara(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("TrackBarA_DestroyedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::GetInfoTipTextTrackbara(long maxInfoTipLength, BSTR* infoTipText, VARIANT_BOOL* abortToolTip)
{
	CAtlString str;
	str.Format(TEXT("TrackBarA_GetInfoTipText: maxInfoTipLength=%i, infoTipText=%s, abortToolTip=%i"), maxInfoTipLength, OLE2W(*infoTipText), *abortToolTip);

	AddLogEntry(str);

	CAtlString tmp;
	tmp.Format(TEXT("Current position: %i"), controls.trackBarA->GetCurrentPosition());

	*infoTipText = _bstr_t(tmp).Detach();
}

void __stdcall CMainDlg::KeyDownTrackbara(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("TrackBarA_KeyDown: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyPressTrackbara(short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("TrackBarA_KeyPress: keyAscii=%i"), *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpTrackbara(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("TrackBarA_KeyUp: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickTrackbara(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TrackBarA_MClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickTrackbara(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TrackBarA_MDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownTrackbara(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TrackBarA_MouseDown: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterTrackbara(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TrackBarA_MouseEnter: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverTrackbara(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TrackBarA_MouseHover: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveTrackbara(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TrackBarA_MouseLeave: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveTrackbara(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TrackBarA_MouseMove: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpTrackbara(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TrackBarA_MouseUp: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseWheelTrackbara(short button, short shift, long x, long y, TrackBarCtlLibA::ScrollAxisConstants scrollAxis, short wheelDelta)
{
	CAtlString str;
	str.Format(TEXT("TrackBarA_MouseWheel: button=%i, shift=%i, x=%i, y=%i, scrollAxis=0x%X, wheelDelta=%i"), button, shift, x, y, scrollAxis, wheelDelta);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropTrackbara(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<TrackBarCtlLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TrackBarA_OLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TrackBarA_OLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.trackBarA->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterTrackbara(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<TrackBarCtlLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TrackBarA_OLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TrackBarA_OLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveTrackbara(LPDISPATCH data, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<TrackBarCtlLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TrackBarA_OLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TrackBarA_OLEDragLeave: data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragMouseMoveTrackbara(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<TrackBarCtlLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TrackBarA_OLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TrackBarA_OLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::PositionChangedTrackbara(TrackBarCtlLibA::PositionChangeTypeConstants changeType, LONG newPosition)
{
	CAtlString str;
	str.Format(TEXT("TrackBarA_PositionChanged: changeType=%i, newPosition=%i"), changeType, newPosition);

	AddLogEntry(str);
}

void __stdcall CMainDlg::PositionChangingTrackbara(TrackBarCtlLibA::PositionChangeTypeConstants changeType, LONG newPosition, VARIANT_BOOL* cancelChange)
{
	CAtlString str;
	str.Format(TEXT("TrackBarA_PositionChanging: changeType=%i, newPosition=%i, cancelChange=%i"), changeType, newPosition, *cancelChange);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RClickTrackbara(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TrackBarA_RClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickTrackbara(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TrackBarA_RDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowTrackbara(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("TrackBarA_RecreatedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowTrackbara()
{
	AddLogEntry(CAtlString(TEXT("TrackBarA_ResizedControlWindow")));
}

void __stdcall CMainDlg::XClickTrackbara(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TrackBarA_XClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickTrackbara(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TrackBarA_XDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}
