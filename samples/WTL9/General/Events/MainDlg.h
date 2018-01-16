// MainDlg.h : interface of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////
#pragma once
#include <initguid.h>

#import <libid:956B5A46-C53F-45a7-AF0E-EC2E1CC9B567> version("1.7") raw_dispinterfaces
#import <libid:F65D51B5-B8DD-476a-85B7-3E26E69D7B3C> version("1.7") raw_dispinterfaces

DEFINE_GUID(LIBID_TrackBarCtlLibU, 0x956B5A46, 0xC53F, 0x45A7, 0xAF, 0x0E, 0xEC, 0x2E, 0x1C, 0xC9, 0xB5, 0x67);
DEFINE_GUID(LIBID_TrackBarCtlLibA, 0xF65D51B5, 0xB8DD, 0x476a, 0x85, 0xB7, 0x3E, 0x26, 0xE6, 0x9D, 0x7B, 0x3C);

class CMainDlg :
    public CAxDialogImpl<CMainDlg>,
    public CMessageFilter,
    public CDialogResize<CMainDlg>,
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CMainDlg>,
    public IDispEventImpl<IDC_TRACKBARU, CMainDlg, &__uuidof(TrackBarCtlLibU::_ITrackBarEvents), &LIBID_TrackBarCtlLibU, 1, 7>,
    public IDispEventImpl<IDC_TRACKBARA, CMainDlg, &__uuidof(TrackBarCtlLibA::_ITrackBarEvents), &LIBID_TrackBarCtlLibA, 1, 7>
{
public:
	enum { IDD = IDD_MAINDLG };

	CContainedWindowT<CAxWindow> trackBarUContainerWnd;
	CContainedWindowT<CWindow> trackBarUWnd;
	CContainedWindowT<CAxWindow> trackBarAContainerWnd;
	CContainedWindowT<CWindow> trackBarAWnd;

	bool trackBarAIsFocused;

	CMainDlg() :
	    trackBarUContainerWnd(this, 1),
	    trackBarAContainerWnd(this, 2),
	    trackBarUWnd(this, 3),
	    trackBarAWnd(this, 4)
	{
	}

	struct Controls
	{
		CEdit logEdit;
		CButton aboutButton;
		CComPtr<TrackBarCtlLibU::ITrackBar> trackBarU;
		CComPtr<TrackBarCtlLibA::ITrackBar> trackBarA;
	} controls;

	virtual BOOL PreTranslateMessage(MSG* pMsg);

	BEGIN_MSG_MAP(CMainDlg)
		MESSAGE_HANDLER(WM_CLOSE, OnClose)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_WINDOWPOSCHANGED, OnWindowPosChanged)

		COMMAND_ID_HANDLER(ID_APP_ABOUT, OnAppAbout)

		CHAIN_MSG_MAP(CDialogResize<CMainDlg>)

		ALT_MSG_MAP(1)
		ALT_MSG_MAP(2)
		ALT_MSG_MAP(3)
		MESSAGE_HANDLER(WM_SETFOCUS, OnSetFocusU)

		ALT_MSG_MAP(4)
		MESSAGE_HANDLER(WM_SETFOCUS, OnSetFocusA)
	END_MSG_MAP()

	BEGIN_SINK_MAP(CMainDlg)
		SINK_ENTRY_EX(IDC_TRACKBARU, __uuidof(TrackBarCtlLibU::_ITrackBarEvents), 1, ClickTrackbaru)
		SINK_ENTRY_EX(IDC_TRACKBARU, __uuidof(TrackBarCtlLibU::_ITrackBarEvents), 2, ContextMenuTrackbaru)
		//SINK_ENTRY_EX(IDC_TRACKBARU, __uuidof(TrackBarCtlLibU::_ITrackBarEvents), 3, CustomDrawTrackbaru)
		SINK_ENTRY_EX(IDC_TRACKBARU, __uuidof(TrackBarCtlLibU::_ITrackBarEvents), 4, DblClickTrackbaru)
		SINK_ENTRY_EX(IDC_TRACKBARU, __uuidof(TrackBarCtlLibU::_ITrackBarEvents), 5, DestroyedControlWindowTrackbaru)
		SINK_ENTRY_EX(IDC_TRACKBARU, __uuidof(TrackBarCtlLibU::_ITrackBarEvents), 6, GetInfoTipTextTrackbaru)
		SINK_ENTRY_EX(IDC_TRACKBARU, __uuidof(TrackBarCtlLibU::_ITrackBarEvents), 7, KeyDownTrackbaru)
		SINK_ENTRY_EX(IDC_TRACKBARU, __uuidof(TrackBarCtlLibU::_ITrackBarEvents), 8, KeyPressTrackbaru)
		SINK_ENTRY_EX(IDC_TRACKBARU, __uuidof(TrackBarCtlLibU::_ITrackBarEvents), 9, KeyUpTrackbaru)
		SINK_ENTRY_EX(IDC_TRACKBARU, __uuidof(TrackBarCtlLibU::_ITrackBarEvents), 10, MClickTrackbaru)
		SINK_ENTRY_EX(IDC_TRACKBARU, __uuidof(TrackBarCtlLibU::_ITrackBarEvents), 11, MDblClickTrackbaru)
		SINK_ENTRY_EX(IDC_TRACKBARU, __uuidof(TrackBarCtlLibU::_ITrackBarEvents), 12, MouseDownTrackbaru)
		SINK_ENTRY_EX(IDC_TRACKBARU, __uuidof(TrackBarCtlLibU::_ITrackBarEvents), 13, MouseEnterTrackbaru)
		SINK_ENTRY_EX(IDC_TRACKBARU, __uuidof(TrackBarCtlLibU::_ITrackBarEvents), 14, MouseHoverTrackbaru)
		SINK_ENTRY_EX(IDC_TRACKBARU, __uuidof(TrackBarCtlLibU::_ITrackBarEvents), 15, MouseLeaveTrackbaru)
		//SINK_ENTRY_EX(IDC_TRACKBARU, __uuidof(TrackBarCtlLibU::_ITrackBarEvents), 16, MouseMoveTrackbaru)
		SINK_ENTRY_EX(IDC_TRACKBARU, __uuidof(TrackBarCtlLibU::_ITrackBarEvents), 17, MouseUpTrackbaru)
		SINK_ENTRY_EX(IDC_TRACKBARU, __uuidof(TrackBarCtlLibU::_ITrackBarEvents), 18, OLEDragDropTrackbaru)
		SINK_ENTRY_EX(IDC_TRACKBARU, __uuidof(TrackBarCtlLibU::_ITrackBarEvents), 19, OLEDragEnterTrackbaru)
		SINK_ENTRY_EX(IDC_TRACKBARU, __uuidof(TrackBarCtlLibU::_ITrackBarEvents), 20, OLEDragLeaveTrackbaru)
		//SINK_ENTRY_EX(IDC_TRACKBARU, __uuidof(TrackBarCtlLibU::_ITrackBarEvents), 21, OLEDragMouseMoveTrackbaru)
		SINK_ENTRY_EX(IDC_TRACKBARU, __uuidof(TrackBarCtlLibU::_ITrackBarEvents), 0, PositionChangedTrackbaru)
		SINK_ENTRY_EX(IDC_TRACKBARU, __uuidof(TrackBarCtlLibU::_ITrackBarEvents), 22, RClickTrackbaru)
		SINK_ENTRY_EX(IDC_TRACKBARU, __uuidof(TrackBarCtlLibU::_ITrackBarEvents), 23, RDblClickTrackbaru)
		SINK_ENTRY_EX(IDC_TRACKBARU, __uuidof(TrackBarCtlLibU::_ITrackBarEvents), 24, RecreatedControlWindowTrackbaru)
		SINK_ENTRY_EX(IDC_TRACKBARU, __uuidof(TrackBarCtlLibU::_ITrackBarEvents), 25, ResizedControlWindowTrackbaru)
		SINK_ENTRY_EX(IDC_TRACKBARU, __uuidof(TrackBarCtlLibU::_ITrackBarEvents), 26, PositionChangingTrackbaru)
		SINK_ENTRY_EX(IDC_TRACKBARU, __uuidof(TrackBarCtlLibU::_ITrackBarEvents), 27, MouseWheelTrackbaru)
		SINK_ENTRY_EX(IDC_TRACKBARU, __uuidof(TrackBarCtlLibU::_ITrackBarEvents), 28, XClickTrackbaru)
		SINK_ENTRY_EX(IDC_TRACKBARU, __uuidof(TrackBarCtlLibU::_ITrackBarEvents), 29, XDblClickTrackbaru)

		SINK_ENTRY_EX(IDC_TRACKBARA, __uuidof(TrackBarCtlLibA::_ITrackBarEvents), 1, ClickTrackbara)
		SINK_ENTRY_EX(IDC_TRACKBARA, __uuidof(TrackBarCtlLibA::_ITrackBarEvents), 2, ContextMenuTrackbara)
		//SINK_ENTRY_EX(IDC_TRACKBARA, __uuidof(TrackBarCtlLibA::_ITrackBarEvents), 3, CustomDrawTrackbara)
		SINK_ENTRY_EX(IDC_TRACKBARA, __uuidof(TrackBarCtlLibA::_ITrackBarEvents), 4, DblClickTrackbara)
		SINK_ENTRY_EX(IDC_TRACKBARA, __uuidof(TrackBarCtlLibA::_ITrackBarEvents), 5, DestroyedControlWindowTrackbara)
		SINK_ENTRY_EX(IDC_TRACKBARA, __uuidof(TrackBarCtlLibA::_ITrackBarEvents), 6, GetInfoTipTextTrackbara)
		SINK_ENTRY_EX(IDC_TRACKBARA, __uuidof(TrackBarCtlLibA::_ITrackBarEvents), 7, KeyDownTrackbara)
		SINK_ENTRY_EX(IDC_TRACKBARA, __uuidof(TrackBarCtlLibA::_ITrackBarEvents), 8, KeyPressTrackbara)
		SINK_ENTRY_EX(IDC_TRACKBARA, __uuidof(TrackBarCtlLibA::_ITrackBarEvents), 9, KeyUpTrackbara)
		SINK_ENTRY_EX(IDC_TRACKBARA, __uuidof(TrackBarCtlLibA::_ITrackBarEvents), 10, MClickTrackbara)
		SINK_ENTRY_EX(IDC_TRACKBARA, __uuidof(TrackBarCtlLibA::_ITrackBarEvents), 11, MDblClickTrackbara)
		SINK_ENTRY_EX(IDC_TRACKBARA, __uuidof(TrackBarCtlLibA::_ITrackBarEvents), 12, MouseDownTrackbara)
		SINK_ENTRY_EX(IDC_TRACKBARA, __uuidof(TrackBarCtlLibA::_ITrackBarEvents), 13, MouseEnterTrackbara)
		SINK_ENTRY_EX(IDC_TRACKBARA, __uuidof(TrackBarCtlLibA::_ITrackBarEvents), 14, MouseHoverTrackbara)
		SINK_ENTRY_EX(IDC_TRACKBARA, __uuidof(TrackBarCtlLibA::_ITrackBarEvents), 15, MouseLeaveTrackbara)
		//SINK_ENTRY_EX(IDC_TRACKBARA, __uuidof(TrackBarCtlLibA::_ITrackBarEvents), 16, MouseMoveTrackbara)
		SINK_ENTRY_EX(IDC_TRACKBARA, __uuidof(TrackBarCtlLibA::_ITrackBarEvents), 17, MouseUpTrackbara)
		SINK_ENTRY_EX(IDC_TRACKBARA, __uuidof(TrackBarCtlLibA::_ITrackBarEvents), 18, OLEDragDropTrackbara)
		SINK_ENTRY_EX(IDC_TRACKBARA, __uuidof(TrackBarCtlLibA::_ITrackBarEvents), 19, OLEDragEnterTrackbara)
		SINK_ENTRY_EX(IDC_TRACKBARA, __uuidof(TrackBarCtlLibA::_ITrackBarEvents), 20, OLEDragLeaveTrackbara)
		//SINK_ENTRY_EX(IDC_TRACKBARA, __uuidof(TrackBarCtlLibA::_ITrackBarEvents), 21, OLEDragMouseMoveTrackbara)
		SINK_ENTRY_EX(IDC_TRACKBARA, __uuidof(TrackBarCtlLibA::_ITrackBarEvents), 0, PositionChangedTrackbara)
		SINK_ENTRY_EX(IDC_TRACKBARA, __uuidof(TrackBarCtlLibA::_ITrackBarEvents), 22, RClickTrackbara)
		SINK_ENTRY_EX(IDC_TRACKBARA, __uuidof(TrackBarCtlLibA::_ITrackBarEvents), 23, RDblClickTrackbara)
		SINK_ENTRY_EX(IDC_TRACKBARA, __uuidof(TrackBarCtlLibA::_ITrackBarEvents), 24, RecreatedControlWindowTrackbara)
		SINK_ENTRY_EX(IDC_TRACKBARA, __uuidof(TrackBarCtlLibA::_ITrackBarEvents), 25, ResizedControlWindowTrackbara)
		SINK_ENTRY_EX(IDC_TRACKBARA, __uuidof(TrackBarCtlLibA::_ITrackBarEvents), 26, PositionChangingTrackbara)
		SINK_ENTRY_EX(IDC_TRACKBARA, __uuidof(TrackBarCtlLibA::_ITrackBarEvents), 27, MouseWheelTrackbara)
		SINK_ENTRY_EX(IDC_TRACKBARA, __uuidof(TrackBarCtlLibA::_ITrackBarEvents), 28, XClickTrackbara)
		SINK_ENTRY_EX(IDC_TRACKBARA, __uuidof(TrackBarCtlLibA::_ITrackBarEvents), 29, XDblClickTrackbara)
	END_SINK_MAP()

	BEGIN_DLGRESIZE_MAP(CMainDlg)
	END_DLGRESIZE_MAP()

	LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnWindowPosChanged(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*bHandled*/);
	LRESULT OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnSetFocusU(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnSetFocusA(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);

	void AddLogEntry(CAtlString text);
	void CloseDialog(int nVal);

	void __stdcall ClickTrackbaru(short button, short shift, long x, long y);
	void __stdcall ContextMenuTrackbaru(short button, short shift, long x, long y);
	void __stdcall CustomDrawTrackbaru(TrackBarCtlLibU::CustomDrawControlPartConstants controlPart, TrackBarCtlLibU::CustomDrawStageConstants drawStage, TrackBarCtlLibU::CustomDrawControlStateConstants controlState, long hDC, TrackBarCtlLibU::RECTANGLE* drawingRectangle, TrackBarCtlLibU::CustomDrawReturnValuesConstants* furtherProcessing);
	void __stdcall DblClickTrackbaru(short button, short shift, long x, long y);
	void __stdcall DestroyedControlWindowTrackbaru(long hWnd);
	void __stdcall GetInfoTipTextTrackbaru(long maxInfoTipLength, BSTR* infoTipText, VARIANT_BOOL* abortToolTip);
	void __stdcall KeyDownTrackbaru(short* keyCode, short shift);
	void __stdcall KeyPressTrackbaru(short* keyAscii);
	void __stdcall KeyUpTrackbaru(short* keyCode, short shift);
	void __stdcall MClickTrackbaru(short button, short shift, long x, long y);
	void __stdcall MDblClickTrackbaru(short button, short shift, long x, long y);
	void __stdcall MouseDownTrackbaru(short button, short shift, long x, long y);
	void __stdcall MouseEnterTrackbaru(short button, short shift, long x, long y);
	void __stdcall MouseHoverTrackbaru(short button, short shift, long x, long y);
	void __stdcall MouseLeaveTrackbaru(short button, short shift, long x, long y);
	void __stdcall MouseMoveTrackbaru(short button, short shift, long x, long y);
	void __stdcall MouseUpTrackbaru(short button, short shift, long x, long y);
	void __stdcall MouseWheelTrackbaru(short button, short shift, long x, long y, TrackBarCtlLibU::ScrollAxisConstants scrollAxis, short wheelDelta);
	void __stdcall OLEDragDropTrackbaru(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragEnterTrackbaru(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragLeaveTrackbaru(LPDISPATCH data, short button, short shift, long x, long y);
	void __stdcall OLEDragMouseMoveTrackbaru(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall PositionChangedTrackbaru(TrackBarCtlLibU::PositionChangeTypeConstants changeType, LONG newPosition);
	void __stdcall PositionChangingTrackbaru(TrackBarCtlLibU::PositionChangeTypeConstants changeType, LONG newPosition, VARIANT_BOOL* cancelChange);
	void __stdcall RClickTrackbaru(short button, short shift, long x, long y);
	void __stdcall RDblClickTrackbaru(short button, short shift, long x, long y);
	void __stdcall RecreatedControlWindowTrackbaru(long hWnd);
	void __stdcall ResizedControlWindowTrackbaru();
	void __stdcall XClickTrackbaru(short button, short shift, long x, long y);
	void __stdcall XDblClickTrackbaru(short button, short shift, long x, long y);

	void __stdcall ClickTrackbara(short button, short shift, long x, long y);
	void __stdcall ContextMenuTrackbara(short button, short shift, long x, long y);
	void __stdcall CustomDrawTrackbara(TrackBarCtlLibA::CustomDrawControlPartConstants controlPart, TrackBarCtlLibA::CustomDrawStageConstants drawStage, TrackBarCtlLibA::CustomDrawControlStateConstants controlState, long hDC, TrackBarCtlLibA::RECTANGLE* drawingRectangle, TrackBarCtlLibA::CustomDrawReturnValuesConstants* furtherProcessing);
	void __stdcall DblClickTrackbara(short button, short shift, long x, long y);
	void __stdcall DestroyedControlWindowTrackbara(long hWnd);
	void __stdcall GetInfoTipTextTrackbara(long maxInfoTipLength, BSTR* infoTipText, VARIANT_BOOL* abortToolTip);
	void __stdcall KeyDownTrackbara(short* keyCode, short shift);
	void __stdcall KeyPressTrackbara(short* keyAscii);
	void __stdcall KeyUpTrackbara(short* keyCode, short shift);
	void __stdcall MClickTrackbara(short button, short shift, long x, long y);
	void __stdcall MDblClickTrackbara(short button, short shift, long x, long y);
	void __stdcall MouseDownTrackbara(short button, short shift, long x, long y);
	void __stdcall MouseEnterTrackbara(short button, short shift, long x, long y);
	void __stdcall MouseHoverTrackbara(short button, short shift, long x, long y);
	void __stdcall MouseLeaveTrackbara(short button, short shift, long x, long y);
	void __stdcall MouseMoveTrackbara(short button, short shift, long x, long y);
	void __stdcall MouseUpTrackbara(short button, short shift, long x, long y);
	void __stdcall MouseWheelTrackbara(short button, short shift, long x, long y, TrackBarCtlLibA::ScrollAxisConstants scrollAxis, short wheelDelta);
	void __stdcall OLEDragDropTrackbara(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragEnterTrackbara(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragLeaveTrackbara(LPDISPATCH data, short button, short shift, long x, long y);
	void __stdcall OLEDragMouseMoveTrackbara(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall PositionChangedTrackbara(TrackBarCtlLibA::PositionChangeTypeConstants changeType, LONG newPosition);
	void __stdcall PositionChangingTrackbara(TrackBarCtlLibA::PositionChangeTypeConstants changeType, LONG newPosition, VARIANT_BOOL* cancelChange);
	void __stdcall RClickTrackbara(short button, short shift, long x, long y);
	void __stdcall RDblClickTrackbara(short button, short shift, long x, long y);
	void __stdcall RecreatedControlWindowTrackbara(long hWnd);
	void __stdcall ResizedControlWindowTrackbara();
	void __stdcall XClickTrackbara(short button, short shift, long x, long y);
	void __stdcall XDblClickTrackbara(short button, short shift, long x, long y);
};
