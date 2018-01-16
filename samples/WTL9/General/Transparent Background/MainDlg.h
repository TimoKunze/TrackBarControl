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
	CBrush trackbarBackgroundBrush;
	CBitmap dialogBackgroundTexture;
	CBrush dialogBackgroundBrush;

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
		MESSAGE_HANDLER(WM_CTLCOLORDLG, OnCtlColorDlg)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		COMMAND_ID_HANDLER(ID_APP_ABOUT, OnAppAbout)

		ALT_MSG_MAP(1)
		MESSAGE_HANDLER(WM_CTLCOLORSTATIC, OnCtlColorStatic)
	END_MSG_MAP()

	BEGIN_SINK_MAP(CMainDlg)
	END_SINK_MAP()

	LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnCtlColorDlg(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnCtlColorStatic(UINT /*uMsg*/, WPARAM wParam, LPARAM lParam, BOOL& bHandled);

	void CloseDialog(int nVal);
};
