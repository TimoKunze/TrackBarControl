// MainDlg.h : interface of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#pragma once
#import <libid:956B5A46-C53F-45a7-AF0E-EC2E1CC9B567> version("1.7") named_guids, no_namespace, raw_dispinterfaces

class CMainDlg :
    public CAxDialogImpl<CMainDlg>,
    public CMessageFilter,
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CMainDlg>,
    public IDispEventImpl<IDC_TRACKBARU, CMainDlg, &__uuidof(_ITrackBarEvents), &LIBID_TrackBarCtlLibU, 1, 7>
{
public:
	enum { IDD = IDD_MAINDLG };

	CContainedWindowT<CAxWindow> trackBarUWnd;

	CMainDlg() : trackBarUWnd(this, 1)
	{
	}

	struct Controls
	{
		CComPtr<ITrackBar> trackBarU;
	} controls;

	virtual BOOL PreTranslateMessage(MSG* pMsg);

	BEGIN_MSG_MAP(CMainDlg)
		MESSAGE_HANDLER(WM_CLOSE, OnClose)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		COMMAND_ID_HANDLER(ID_APP_ABOUT, OnAppAbout)

		ALT_MSG_MAP(1)
	END_MSG_MAP()

	BEGIN_SINK_MAP(CMainDlg)
		SINK_ENTRY_EX(IDC_TRACKBARU, __uuidof(_ITrackBarEvents), 3, CustomDrawTrackbaru)
	END_SINK_MAP()

	LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);

	void CloseDialog(int nVal);

	void __stdcall CustomDrawTrackbaru(CustomDrawControlPartConstants controlPart, CustomDrawStageConstants drawStage, CustomDrawControlStateConstants /*controlState*/, long hDC, RECTANGLE* /*drawingRectangle*/, CustomDrawReturnValuesConstants* furtherProcessing);
};
