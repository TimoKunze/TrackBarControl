// TrackBar.cpp: Superclasses msctls_trackbar32.

#include "stdafx.h"
#include "TrackBar.h"
#include "ClassFactory.h"

#pragma comment(lib, "comctl32.lib")


TrackBar::TrackBar()
{
	properties.mouseIcon.InitializePropertyWatcher(this, DISPID_TRACKBARCTL_MOUSEICON);

	SIZEL size = {100, 40};
	AtlPixelToHiMetric(&size, &m_sizeExtent);

	// always create a window, even if the container supports windowless controls
	m_bWindowOnly = TRUE;

	pToolTipBuffer = HeapAlloc(GetProcessHeap(), 0, (MAX_PATH + 1) * sizeof(WCHAR));

	// Microsoft couldn't make it more difficult to detect whether we should use themes or not...
	flags.usingThemes = FALSE;
	if(CTheme::IsThemingSupported() && RunTimeHelper::IsCommCtrl6()) {
		HMODULE hThemeDLL = LoadLibrary(TEXT("uxtheme.dll"));
		if(hThemeDLL) {
			typedef BOOL WINAPI IsAppThemedFn();
			IsAppThemedFn* pfnIsAppThemed = static_cast<IsAppThemedFn*>(GetProcAddress(hThemeDLL, "IsAppThemed"));
			if(pfnIsAppThemed()) {
				flags.usingThemes = TRUE;
			}
			FreeLibrary(hThemeDLL);
		}
	}
}

TrackBar::~TrackBar()
{
	if(pToolTipBuffer) {
		HeapFree(GetProcessHeap(), 0, pToolTipBuffer);
	}
}


//////////////////////////////////////////////////////////////////////
// implementation of ISupportsErrorInfo
STDMETHODIMP TrackBar::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_ITrackBar, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportsErrorInfo
//////////////////////////////////////////////////////////////////////


STDMETHODIMP TrackBar::GetSizeMax(ULARGE_INTEGER* pSize)
{
	ATLASSERT_POINTER(pSize, ULARGE_INTEGER);
	if(!pSize) {
		return E_POINTER;
	}

	pSize->LowPart = 0;
	pSize->HighPart = 0;
	pSize->QuadPart = sizeof(LONG/*signature*/) + sizeof(LONG/*version*/) + sizeof(DWORD/*atlVersion*/) + sizeof(m_sizeExtent);

	// we've 19 VT_I4 properties...
	pSize->QuadPart += 19 * (sizeof(VARTYPE) + sizeof(LONG));
	// ...and 1 VT_I2 property...
	pSize->QuadPart += 1 * (sizeof(VARTYPE) + sizeof(SHORT));
	// ...and 11 VT_BOOL properties...
	pSize->QuadPart += 11 * (sizeof(VARTYPE) + sizeof(VARIANT_BOOL));

	// ...and 1 VT_DISPATCH property
	pSize->QuadPart += 1 * (sizeof(VARTYPE) + sizeof(CLSID));

	// we've to query each object for its size
	CComPtr<IPersistStreamInit> pStreamInit = NULL;
	if(properties.mouseIcon.pPictureDisp) {
		if(FAILED(properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
			properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit));
		}
	}
	if(pStreamInit) {
		ULARGE_INTEGER tmp = {0};
		pStreamInit->GetSizeMax(&tmp);
		pSize->QuadPart += tmp.QuadPart;
	}

	return S_OK;
}

STDMETHODIMP TrackBar::Load(LPSTREAM pStream)
{
	ATLASSUME(pStream);
	if(!pStream) {
		return E_POINTER;
	}

	// NOTE: We neither support legacy streams nor streams with a version < 0x0102.

	HRESULT hr = S_OK;
	LONG signature = 0;
	LONG version = 0;
	if(FAILED(hr = pStream->Read(&signature, sizeof(signature), NULL))) {
		return hr;
	}
	if(signature != 0x0B0B0B0B/*4x AppID*/) {
		// might be a legacy stream, that starts with the ATL version
		if(signature == 0x0700 || signature == 0x0710 || signature == 0x0800 || signature == 0x0900) {
			version = 0x0099;
		} else {
			return E_FAIL;
		}
	}
	//LONG version = 0;
	if(version != 0x0099) {
		if(FAILED(hr = pStream->Read(&version, sizeof(version), NULL))) {
			return hr;
		}
		if(version > 0x0103) {
			return E_FAIL;
		}
		// HACK: Old versions of TrackBar used version identifier 0x0000!
		if(version == 0x0000) {
			version = 0x0100;
		}
	}

	DWORD atlVersion;
	if(version == 0x0099) {
		atlVersion = 0x0900;
	} else {
		if(FAILED(hr = pStream->Read(&atlVersion, sizeof(atlVersion), NULL))) {
			return hr;
		}
	}
	if(atlVersion > _ATL_VER) {
		return E_FAIL;
	}

	if(version != 0x0100) {
		if(FAILED(hr = pStream->Read(&m_sizeExtent, sizeof(m_sizeExtent), NULL))) {
			return hr;
		}
	}

	typedef HRESULT ReadVariantFromStreamFn(__in LPSTREAM pStream, VARTYPE expectedVarType, __inout LPVARIANT pVariant);
	ReadVariantFromStreamFn* pfnReadVariantFromStream = NULL;
	if(version == 0x0100) {
		pfnReadVariantFromStream = ReadVariantFromStream_Legacy;
	} else {
		pfnReadVariantFromStream = ReadVariantFromStream;
	}

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	AppearanceConstants appearance = static_cast<AppearanceConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I2, &propertyValue))) {
		return hr;
	}
	SHORT autoTickFrequency = propertyValue.iVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL autoTickMarks = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR backColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	BorderStyleConstants borderStyle = static_cast<BorderStyleConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG currentPosition = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL dontRedraw = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL downIsLeft = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL enabled = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG hoverTime = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG largeStepWidth = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG maximum = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG minimum = propertyValue.lVal;

	VARTYPE vt;
	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_DISPATCH)) {
		return hr;
	}
	CComPtr<IPictureDisp> pMouseIcon = NULL;
	if(FAILED(hr = OleLoadFromStream(pStream, IID_IDispatch, reinterpret_cast<LPVOID*>(&pMouseIcon)))) {
		if(hr != REGDB_E_CLASSNOTREG) {
			return S_OK;
		}
	}

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	MousePointerConstants mousePointer = static_cast<MousePointerConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OrientationConstants orientation = static_cast<OrientationConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL processContextMenuKeys = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG rangeSelectionEnd = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG rangeSelectionStart = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL registerForOLEDragDrop = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL reversed = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL rightToLeftLayout = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	SelectionTypeConstants selectionType = static_cast<SelectionTypeConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL showSlider = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG sliderLength = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG smallStepWidth = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL supportOLEDragImages = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	TickMarksPositionConstants tickMarksPosition = static_cast<TickMarksPositionConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	ToolTipPositionConstants toolTipPosition = static_cast<ToolTipPositionConstants>(propertyValue.lVal);

	BackgroundDrawModeConstants backgroundDrawMode = bdmNormal;
	VARIANT_BOOL detectDoubleClicks = VARIANT_TRUE;
	if(version >= 0x0102) {
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
			return hr;
		}
		backgroundDrawMode = static_cast<BackgroundDrawModeConstants>(propertyValue.lVal);

		if(version >= 0x0103) {
			if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
				return hr;
			}
			detectDoubleClicks = propertyValue.boolVal;
		}
	}


	hr = put_Appearance(appearance);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AutoTickFrequency(autoTickFrequency);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AutoTickMarks(autoTickMarks);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BackColor(backColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BackgroundDrawMode(backgroundDrawMode);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BorderStyle(borderStyle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CurrentPosition(currentPosition);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DetectDoubleClicks(detectDoubleClicks);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisabledEvents(disabledEvents);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DontRedraw(dontRedraw);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DownIsLeft(downIsLeft);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Enabled(enabled);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HoverTime(hoverTime);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_LargeStepWidth(largeStepWidth);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Maximum(maximum);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Minimum(minimum);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MouseIcon(pMouseIcon);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MousePointer(mousePointer);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Orientation(orientation);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ProcessContextMenuKeys(processContextMenuKeys);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RangeSelectionEnd(rangeSelectionEnd);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RangeSelectionStart(rangeSelectionStart);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RegisterForOLEDragDrop(registerForOLEDragDrop);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Reversed(reversed);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RightToLeftLayout(rightToLeftLayout);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SelectionType(selectionType);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ShowSlider(showSlider);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SliderLength(sliderLength);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SmallStepWidth(smallStepWidth);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SupportOLEDragImages(supportOLEDragImages);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_TickMarksPosition(tickMarksPosition);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ToolTipPosition(toolTipPosition);
	if(FAILED(hr)) {
		return hr;
	}

	SetDirty(FALSE);
	return S_OK;
}

STDMETHODIMP TrackBar::Save(LPSTREAM pStream, BOOL clearDirtyFlag)
{
	ATLASSUME(pStream);
	if(!pStream) {
		return E_POINTER;
	}

	HRESULT hr = S_OK;
	LONG signature = 0x0B0B0B0B/*4x AppID*/;
	if(FAILED(hr = pStream->Write(&signature, sizeof(signature), NULL))) {
		return hr;
	}
	LONG version = 0x0103;
	if(FAILED(hr = pStream->Write(&version, sizeof(version), NULL))) {
		return hr;
	}

	DWORD atlVersion = _ATL_VER;
	if(FAILED(hr = pStream->Write(&atlVersion, sizeof(atlVersion), NULL))) {
		return hr;
	}

	if(FAILED(hr = pStream->Write(&m_sizeExtent, sizeof(m_sizeExtent), NULL))) {
		return hr;
	}

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.appearance;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I2;
	propertyValue.iVal = properties.autoTickFrequency;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.autoTickMarks);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.backColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.borderStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.currentPosition;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.disabledEvents;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.dontRedraw);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.downIsLeft);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.enabled);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.hoverTime;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.largeStepWidth;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.maximum;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.minimum;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	CComPtr<IPersistStream> pPersistStream = NULL;
	if(properties.mouseIcon.pPictureDisp) {
		if(FAILED(hr = properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pPersistStream)))) {
			return hr;
		}
	}
	// store some marker
	VARTYPE vt = VT_DISPATCH;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(pPersistStream) {
		if(FAILED(hr = OleSaveToStream(pPersistStream, pStream))) {
			return hr;
		}
	} else {
		if(FAILED(hr = WriteClassStm(pStream, CLSID_NULL))) {
			return hr;
		}
	}

	propertyValue.lVal = properties.mousePointer;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.orientation;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.processContextMenuKeys);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.rangeSelectionEnd;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.rangeSelectionStart;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.registerForOLEDragDrop);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.reversed);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.rightToLeftLayout);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.selectionType;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.showSlider);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.sliderLength;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.smallStepWidth;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.tickMarksPosition;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.toolTipPosition;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	// version 0x0102 starts here
	propertyValue.lVal = properties.backgroundDrawMode;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	// version 0x0103 starts here
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.detectDoubleClicks);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	if(clearDirtyFlag) {
		SetDirty(FALSE);
	}
	return S_OK;
}


HWND TrackBar::Create(HWND hWndParent, ATL::_U_RECT rect/* = NULL*/, LPCTSTR szWindowName/* = NULL*/, DWORD dwStyle/* = 0*/, DWORD dwExStyle/* = 0*/, ATL::_U_MENUorID MenuOrID/* = 0U*/, LPVOID lpCreateParam/* = NULL*/)
{
	INITCOMMONCONTROLSEX data = {0};
	data.dwSize = sizeof(data);
	data.dwICC = ICC_BAR_CLASSES;
	InitCommonControlsEx(&data);

	dwStyle = GetStyleBits();
	dwExStyle = GetExStyleBits();
	return CComControl<TrackBar>::Create(hWndParent, rect, szWindowName, dwStyle, dwExStyle, MenuOrID, lpCreateParam);
}

HRESULT TrackBar::OnDraw(ATL_DRAWINFO& drawInfo)
{
	if(IsInDesignMode()) {
		CAtlString text = TEXT("TrackBar ");
		CComBSTR buffer;
		get_Version(&buffer);
		text += buffer;
		SetTextAlign(drawInfo.hdcDraw, TA_CENTER | TA_BASELINE);
		TextOut(drawInfo.hdcDraw, drawInfo.prcBounds->left + (drawInfo.prcBounds->right - drawInfo.prcBounds->left) / 2, drawInfo.prcBounds->top + (drawInfo.prcBounds->bottom - drawInfo.prcBounds->top) / 2, text, text.GetLength());
	}

	return S_OK;
}

void TrackBar::OnFinalMessage(HWND /*hWnd*/)
{
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	Release();
}

STDMETHODIMP TrackBar::IOleObject_SetClientSite(LPOLECLIENTSITE pClientSite)
{
	return CComControl<TrackBar>::IOleObject_SetClientSite(pClientSite);
}

STDMETHODIMP TrackBar::OnDocWindowActivate(BOOL /*fActivate*/)
{
	return S_OK;
}

BOOL TrackBar::PreTranslateAccelerator(LPMSG pMessage, HRESULT& hReturnValue)
{
	if((pMessage->message >= WM_KEYFIRST) && (pMessage->message <= WM_KEYLAST)) {
		LRESULT dialogCode = SendMessage(pMessage->hwnd, WM_GETDLGCODE, 0, 0);
		//ATLASSERT(dialogCode == DLGC_WANTARROWS);
		if(pMessage->wParam == VK_TAB) {
			if(dialogCode & DLGC_WANTTAB) {
				hReturnValue = S_FALSE;
				return TRUE;
			}
		}
		switch(pMessage->wParam) {
			case VK_LEFT:
			case VK_RIGHT:
			case VK_UP:
			case VK_DOWN:
			case VK_HOME:
			case VK_END:
			case VK_NEXT:
			case VK_PRIOR:
				if(dialogCode & DLGC_WANTARROWS) {
					if(!(GetKeyState(VK_SHIFT) & 0x8000) && !(GetKeyState(VK_CONTROL) & 0x8000) && !(GetKeyState(VK_MENU) & 0x8000)) {
						SendMessage(pMessage->hwnd, pMessage->message, pMessage->wParam, pMessage->lParam);
						hReturnValue = S_OK;
						return TRUE;
					}
				}
				break;
		}
	}
	return CComControl<TrackBar>::PreTranslateAccelerator(pMessage, hReturnValue);
}

//////////////////////////////////////////////////////////////////////
// implementation of IDropTarget
STDMETHODIMP TrackBar::DragEnter(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	if(properties.supportOLEDragImages && !dragDropStatus.pDropTargetHelper) {
		CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_ALL, IID_PPV_ARGS(&dragDropStatus.pDropTargetHelper));
	}

	DROPDESCRIPTION oldDropDescription;
	ZeroMemory(&oldDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, oldDropDescription);

	POINT buffer = {mousePosition.x, mousePosition.y};
	Raise_OLEDragEnter(pDataObject, pEffect, keyState, mousePosition);
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragEnter(*this, pDataObject, &buffer, *pEffect);
	}

	DROPDESCRIPTION newDropDescription;
	ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
	if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION))) {
		InvalidateDragWindow(pDataObject);
	}
	return S_OK;
}

STDMETHODIMP TrackBar::DragLeave(void)
{
	Raise_OLEDragLeave();
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragLeave();
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	return S_OK;
}

STDMETHODIMP TrackBar::DragOver(DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	CComQIPtr<IDataObject> pDataObject = dragDropStatus.pActiveDataObject;
	DROPDESCRIPTION oldDropDescription;
	ZeroMemory(&oldDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, oldDropDescription);

	POINT buffer = {mousePosition.x, mousePosition.y};
	Raise_OLEDragMouseMove(pEffect, keyState, mousePosition);
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragOver(&buffer, *pEffect);
	}

	DROPDESCRIPTION newDropDescription;
	ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
	if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && (newDropDescription.type > DROPIMAGE_NONE || memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION)))) {
		InvalidateDragWindow(pDataObject);
	}
	return S_OK;
}

STDMETHODIMP TrackBar::Drop(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	POINT buffer = {mousePosition.x, mousePosition.y};
	dragDropStatus.drop_pDataObject = pDataObject;
	dragDropStatus.drop_mousePosition = buffer;
	dragDropStatus.drop_effect = *pEffect;

	Raise_OLEDragDrop(pDataObject, pEffect, keyState, mousePosition);
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->Drop(pDataObject, &buffer, *pEffect);
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	dragDropStatus.drop_pDataObject = NULL;
	return S_OK;
}
// implementation of IDropTarget
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of ICategorizeProperties
STDMETHODIMP TrackBar::GetCategoryName(PROPCAT category, LCID /*languageID*/, BSTR* pName)
{
	switch(category) {
		case PROPCAT_Colors:
			*pName = GetResString(IDPC_COLORS).Detach();
			return S_OK;
			break;
	}
	return E_FAIL;
}

STDMETHODIMP TrackBar::MapPropertyToCategory(DISPID property, PROPCAT* pCategory)
{
	if(!pCategory) {
		return E_POINTER;
	}

	switch(property) {
		case DISPID_TRACKBARCTL_APPEARANCE:
		case DISPID_TRACKBARCTL_BACKGROUNDDRAWMODE:
		case DISPID_TRACKBARCTL_BORDERSTYLE:
		case DISPID_TRACKBARCTL_CHANNELHEIGHT:
		case DISPID_TRACKBARCTL_CHANNELLEFT:
		case DISPID_TRACKBARCTL_CHANNELTOP:
		case DISPID_TRACKBARCTL_CHANNELWIDTH:
		case DISPID_TRACKBARCTL_MOUSEICON:
		case DISPID_TRACKBARCTL_MOUSEPOINTER:
		case DISPID_TRACKBARCTL_ORIENTATION:
		case DISPID_TRACKBARCTL_SHOWSLIDER:
		case DISPID_TRACKBARCTL_SLIDERHEIGHT:
		case DISPID_TRACKBARCTL_SLIDERLEFT:
		case DISPID_TRACKBARCTL_SLIDERLENGTH:
		case DISPID_TRACKBARCTL_SLIDERTOP:
		case DISPID_TRACKBARCTL_SLIDERWIDTH:
		case DISPID_TRACKBARCTL_TICKMARKSPOSITION:
			*pCategory = PROPCAT_Appearance;
			return S_OK;
			break;
		case DISPID_TRACKBARCTL_AUTOTICKMARKS:
		case DISPID_TRACKBARCTL_DETECTDOUBLECLICKS:
		case DISPID_TRACKBARCTL_DISABLEDEVENTS:
		case DISPID_TRACKBARCTL_DONTREDRAW:
		case DISPID_TRACKBARCTL_DOWNISLEFT:
		case DISPID_TRACKBARCTL_HOVERTIME:
		case DISPID_TRACKBARCTL_LARGESTEPWIDTH:
		case DISPID_TRACKBARCTL_PROCESSCONTEXTMENUKEYS:
		case DISPID_TRACKBARCTL_RIGHTTOLEFTLAYOUT:
		case DISPID_TRACKBARCTL_SELECTIONTYPE:
		case DISPID_TRACKBARCTL_SMALLSTEPWIDTH:
		case DISPID_TRACKBARCTL_TOOLTIPPOSITION:
			*pCategory = PROPCAT_Behavior;
			return S_OK;
			break;
		case DISPID_TRACKBARCTL_BACKCOLOR:
			*pCategory = PROPCAT_Colors;
			return S_OK;
			break;
		case DISPID_TRACKBARCTL_APPID:
		case DISPID_TRACKBARCTL_APPNAME:
		case DISPID_TRACKBARCTL_APPSHORTNAME:
		case DISPID_TRACKBARCTL_BUILD:
		case DISPID_TRACKBARCTL_CHARSET:
		case DISPID_TRACKBARCTL_HWND:
		case DISPID_TRACKBARCTL_HWNDBOTTOMORRIGHTBUDDY:
		case DISPID_TRACKBARCTL_HWNDTOOLTIP:
		case DISPID_TRACKBARCTL_HWNDTOPORLEFTBUDDY:
		case DISPID_TRACKBARCTL_ISRELEASE:
		case DISPID_TRACKBARCTL_PROGRAMMER:
		case DISPID_TRACKBARCTL_TESTER:
		case DISPID_TRACKBARCTL_VERSION:
			*pCategory = PROPCAT_Data;
			return S_OK;
			break;
		case DISPID_TRACKBARCTL_REGISTERFOROLEDRAGDROP:
		case DISPID_TRACKBARCTL_SUPPORTOLEDRAGIMAGES:
			*pCategory = PROPCAT_DragDrop;
			return S_OK;
			break;
		case DISPID_TRACKBARCTL_ENABLED:
			*pCategory = PROPCAT_Misc;
			return S_OK;
			break;
		case DISPID_TRACKBARCTL_AUTOTICKFREQUENCY:
		case DISPID_TRACKBARCTL_CURRENTPOSITION:
		case DISPID_TRACKBARCTL_MAXIMUM:
		case DISPID_TRACKBARCTL_MINIMUM:
		case DISPID_TRACKBARCTL_REVERSED:
		case DISPID_TRACKBARCTL_TICKMARKS:
			*pCategory = PROPCAT_Scale;
			return S_OK;
			break;
	}
	return E_FAIL;
}
// implementation of ICategorizeProperties
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of ICreditsProvider
CAtlString TrackBar::GetAuthors(void)
{
	CComBSTR buffer;
	get_Programmer(&buffer);
	return CAtlString(buffer);
}

CAtlString TrackBar::GetHomepage(void)
{
	return TEXT("https://www.TimoSoft-Software.de");
}

CAtlString TrackBar::GetPaypalLink(void)
{
	return TEXT("https://www.paypal.com/xclick/business=TKunze71216%40gmx.de&item_name=TrackBar&no_shipping=1&tax=0&currency_code=EUR");
}

CAtlString TrackBar::GetSpecialThanks(void)
{
	return TEXT("Geoff Chappell, Wine Headquarters");
}

CAtlString TrackBar::GetThanks(void)
{
	CAtlString ret = TEXT("Google, various newsgroups and mailing lists, many websites,\n");
	ret += TEXT("Heaven Shall Burn, Arch Enemy, Machine Head, Trivium, Deadlock, Draconian, Soulfly, Delain, Lacuna Coil, Ensiferum, Epica, Nightwish, Guns N' Roses and many other musicians");
	return ret;
}

CAtlString TrackBar::GetVersion(void)
{
	CAtlString ret = TEXT("Version ");
	CComBSTR buffer;
	get_Version(&buffer);
	ret += buffer;
	ret += TEXT(" (");
	get_CharSet(&buffer);
	ret += buffer;
	ret += TEXT(")\nCompilation timestamp: ");
	ret += TEXT(STRTIMESTAMP);
	ret += TEXT("\n");

	VARIANT_BOOL b;
	get_IsRelease(&b);
	if(b == VARIANT_FALSE) {
		ret += TEXT("This version is for debugging only.");
	}

	return ret;
}
// implementation of ICreditsProvider
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IPerPropertyBrowsing
STDMETHODIMP TrackBar::GetDisplayString(DISPID property, BSTR* pDescription)
{
	if(!pDescription) {
		return E_POINTER;
	}

	CComBSTR description;
	HRESULT hr = S_OK;
	switch(property) {
		case DISPID_TRACKBARCTL_APPEARANCE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.appearance), description);
			break;
		case DISPID_TRACKBARCTL_BACKGROUNDDRAWMODE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.backgroundDrawMode), description);
			break;
		case DISPID_TRACKBARCTL_BORDERSTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.borderStyle), description);
			break;
		case DISPID_TRACKBARCTL_MOUSEPOINTER:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.mousePointer), description);
			break;
		case DISPID_TRACKBARCTL_ORIENTATION:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.orientation), description);
			break;
		case DISPID_TRACKBARCTL_SELECTIONTYPE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.selectionType), description);
			break;
		case DISPID_TRACKBARCTL_TICKMARKSPOSITION:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.tickMarksPosition), description);
			break;
		case DISPID_TRACKBARCTL_TOOLTIPPOSITION:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.toolTipPosition), description);
			break;
		default:
			return IPerPropertyBrowsingImpl<TrackBar>::GetDisplayString(property, pDescription);
			break;
	}
	if(SUCCEEDED(hr)) {
		*pDescription = description.Detach();
	}

	return *pDescription ? S_OK : E_OUTOFMEMORY;
}

STDMETHODIMP TrackBar::GetPredefinedStrings(DISPID property, CALPOLESTR* pDescriptions, CADWORD* pCookies)
{
	if(!pDescriptions || !pCookies) {
		return E_POINTER;
	}

	int c = 0;
	switch(property) {
		case DISPID_TRACKBARCTL_BACKGROUNDDRAWMODE:
		case DISPID_TRACKBARCTL_BORDERSTYLE:
		case DISPID_TRACKBARCTL_ORIENTATION:
		case DISPID_TRACKBARCTL_SELECTIONTYPE:
			c = 2;
			break;
		case DISPID_TRACKBARCTL_APPEARANCE:
		case DISPID_TRACKBARCTL_TOOLTIPPOSITION:
			c = 3;
			break;
		case DISPID_TRACKBARCTL_TICKMARKSPOSITION:
			c = 4;
			break;
		case DISPID_TRACKBARCTL_MOUSEPOINTER:
			c = 30;
			break;
		default:
			return IPerPropertyBrowsingImpl<TrackBar>::GetPredefinedStrings(property, pDescriptions, pCookies);
			break;
	}
	pDescriptions->cElems = c;
	pCookies->cElems = c;
	pDescriptions->pElems = static_cast<LPOLESTR*>(CoTaskMemAlloc(pDescriptions->cElems * sizeof(LPOLESTR)));
	pCookies->pElems = static_cast<LPDWORD>(CoTaskMemAlloc(pCookies->cElems * sizeof(DWORD)));

	for(UINT iDescription = 0; iDescription < pDescriptions->cElems; ++iDescription) {
		UINT propertyValue = iDescription;
		if((property == DISPID_TRACKBARCTL_MOUSEPOINTER) && (iDescription == pDescriptions->cElems - 1)) {
			propertyValue = mpCustom;
		}

		CComBSTR description;
		HRESULT hr = GetDisplayStringForSetting(property, propertyValue, description);
		if(SUCCEEDED(hr)) {
			size_t bufferSize = SysStringLen(description) + 1;
			pDescriptions->pElems[iDescription] = static_cast<LPOLESTR>(CoTaskMemAlloc(bufferSize * sizeof(WCHAR)));
			ATLVERIFY(SUCCEEDED(StringCchCopyW(pDescriptions->pElems[iDescription], bufferSize, description)));
			// simply use the property value as cookie
			pCookies->pElems[iDescription] = propertyValue;
		} else {
			return DISP_E_BADINDEX;
		}
	}
	return S_OK;
}

STDMETHODIMP TrackBar::GetPredefinedValue(DISPID property, DWORD cookie, VARIANT* pPropertyValue)
{
	switch(property) {
		case DISPID_TRACKBARCTL_APPEARANCE:
		case DISPID_TRACKBARCTL_BACKGROUNDDRAWMODE:
		case DISPID_TRACKBARCTL_BORDERSTYLE:
		case DISPID_TRACKBARCTL_MOUSEPOINTER:
		case DISPID_TRACKBARCTL_ORIENTATION:
		case DISPID_TRACKBARCTL_SELECTIONTYPE:
		case DISPID_TRACKBARCTL_TICKMARKSPOSITION:
		case DISPID_TRACKBARCTL_TOOLTIPPOSITION:
			VariantInit(pPropertyValue);
			pPropertyValue->vt = VT_I4;
			// we used the property value itself as cookie
			pPropertyValue->lVal = cookie;
			break;
		default:
			return IPerPropertyBrowsingImpl<TrackBar>::GetPredefinedValue(property, cookie, pPropertyValue);
			break;
	}
	return S_OK;
}

STDMETHODIMP TrackBar::MapPropertyToPage(DISPID property, CLSID* pPropertyPage)
{
	return IPerPropertyBrowsingImpl<TrackBar>::MapPropertyToPage(property, pPropertyPage);
}
// implementation of IPerPropertyBrowsing
//////////////////////////////////////////////////////////////////////

HRESULT TrackBar::GetDisplayStringForSetting(DISPID property, DWORD cookie, CComBSTR& description)
{
	switch(property) {
		case DISPID_TRACKBARCTL_APPEARANCE:
			switch(cookie) {
				case a2D:
					description = GetResStringWithNumber(IDP_APPEARANCE2D, a2D);
					break;
				case a3D:
					description = GetResStringWithNumber(IDP_APPEARANCE3D, a3D);
					break;
				case a3DLight:
					description = GetResStringWithNumber(IDP_APPEARANCE3DLIGHT, a3DLight);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_TRACKBARCTL_BACKGROUNDDRAWMODE:
			switch(cookie) {
				case bdmNormal:
					description = GetResStringWithNumber(IDP_BACKGROUNDDRAWMODENORMAL, bdmNormal);
					break;
				case bdmByParent:
					description = GetResStringWithNumber(IDP_BACKGROUNDDRAWMODEBYPARENT, bdmByParent);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_TRACKBARCTL_BORDERSTYLE:
			switch(cookie) {
				case bsNone:
					description = GetResStringWithNumber(IDP_BORDERSTYLENONE, bsNone);
					break;
				case bsFixedSingle:
					description = GetResStringWithNumber(IDP_BORDERSTYLEFIXEDSINGLE, bsFixedSingle);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_TRACKBARCTL_MOUSEPOINTER:
			switch(cookie) {
				case mpDefault:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERDEFAULT, mpDefault);
					break;
				case mpArrow:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROW, mpArrow);
					break;
				case mpCross:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERCROSS, mpCross);
					break;
				case mpIBeam:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERIBEAM, mpIBeam);
					break;
				case mpIcon:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERICON, mpIcon);
					break;
				case mpSize:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZE, mpSize);
					break;
				case mpSizeNESW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENESW, mpSizeNESW);
					break;
				case mpSizeNS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENS, mpSizeNS);
					break;
				case mpSizeNWSE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENWSE, mpSizeNWSE);
					break;
				case mpSizeEW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZEEW, mpSizeEW);
					break;
				case mpUpArrow:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERUPARROW, mpUpArrow);
					break;
				case mpHourglass:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERHOURGLASS, mpHourglass);
					break;
				case mpNoDrop:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERNODROP, mpNoDrop);
					break;
				case mpArrowHourglass:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROWHOURGLASS, mpArrowHourglass);
					break;
				case mpArrowQuestion:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROWQUESTION, mpArrowQuestion);
					break;
				case mpSizeAll:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZEALL, mpSizeAll);
					break;
				case mpHand:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERHAND, mpHand);
					break;
				case mpInsertMedia:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERINSERTMEDIA, mpInsertMedia);
					break;
				case mpScrollAll:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLALL, mpScrollAll);
					break;
				case mpScrollN:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLN, mpScrollN);
					break;
				case mpScrollNE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNE, mpScrollNE);
					break;
				case mpScrollE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLE, mpScrollE);
					break;
				case mpScrollSE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLSE, mpScrollSE);
					break;
				case mpScrollS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLS, mpScrollS);
					break;
				case mpScrollSW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLSW, mpScrollSW);
					break;
				case mpScrollW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLW, mpScrollW);
					break;
				case mpScrollNW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNW, mpScrollNW);
					break;
				case mpScrollNS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNS, mpScrollNS);
					break;
				case mpScrollEW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLEW, mpScrollEW);
					break;
				case mpCustom:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERCUSTOM, mpCustom);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_TRACKBARCTL_ORIENTATION:
			switch(cookie) {
				case oHorizontal:
					description = GetResStringWithNumber(IDP_ORIENTATIONHORIZONTAL, oHorizontal);
					break;
				case oVertical:
					description = GetResStringWithNumber(IDP_ORIENTATIONVERTICAL, oVertical);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_TRACKBARCTL_SELECTIONTYPE:
			switch(cookie) {
				case stDiscreteValue:
					description = GetResStringWithNumber(IDP_SELECTIONTYPEDISCRETEVALUE, stDiscreteValue);
					break;
				case stRangeSelection:
					description = GetResStringWithNumber(IDP_SELECTIONTYPERANGESELECTION, stRangeSelection);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_TRACKBARCTL_TICKMARKSPOSITION:
			switch(cookie) {
				case tmpNone:
					description = GetResStringWithNumber(IDP_TICKMARKSPOSITIONNONE, tmpNone);
					break;
				case tmpBottomOrRight:
					description = GetResStringWithNumber(IDP_TICKMARKSPOSITIONBOTTOMORRIGHT, tmpBottomOrRight);
					break;
				case tmpTopOrLeft:
					description = GetResStringWithNumber(IDP_TICKMARKSPOSITIONTOPORLEFT, tmpTopOrLeft);
					break;
				case tmpBoth:
					description = GetResStringWithNumber(IDP_TICKMARKSPOSITIONBOTH, tmpBoth);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_TRACKBARCTL_TOOLTIPPOSITION:
			switch(cookie) {
				case ttpNone:
					description = GetResStringWithNumber(IDP_TOOLTIPPOSITIONNONE, ttpNone);
					break;
				case ttpBottomOrRight:
					description = GetResStringWithNumber(IDP_TOOLTIPPOSITIONBOTTOMORRIGHT, ttpBottomOrRight);
					break;
				case ttpTopOrLeft:
					description = GetResStringWithNumber(IDP_TOOLTIPPOSITIONTOPORLEFT, ttpTopOrLeft);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		default:
			return DISP_E_BADINDEX;
			break;
	}

	return S_OK;
}

//////////////////////////////////////////////////////////////////////
// implementation of ISpecifyPropertyPages
STDMETHODIMP TrackBar::GetPages(CAUUID* pPropertyPages)
{
	if(!pPropertyPages) {
		return E_POINTER;
	}

	pPropertyPages->cElems = 3;
	pPropertyPages->pElems = static_cast<LPGUID>(CoTaskMemAlloc(sizeof(GUID) * pPropertyPages->cElems));
	if(pPropertyPages->pElems) {
		pPropertyPages->pElems[0] = CLSID_CommonProperties;
		pPropertyPages->pElems[1] = CLSID_StockColorPage;
		pPropertyPages->pElems[2] = CLSID_StockPicturePage;
		return S_OK;
	}
	return E_OUTOFMEMORY;
}
// implementation of ISpecifyPropertyPages
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IOleObject
STDMETHODIMP TrackBar::DoVerb(LONG verbID, LPMSG pMessage, IOleClientSite* pActiveSite, LONG reserved, HWND hWndParent, LPCRECT pBoundingRectangle)
{
	switch(verbID) {
		case 1:     // About...
			return DoVerbAbout(hWndParent);
			break;
		default:
			return IOleObjectImpl<TrackBar>::DoVerb(verbID, pMessage, pActiveSite, reserved, hWndParent, pBoundingRectangle);
			break;
	}
}

STDMETHODIMP TrackBar::EnumVerbs(IEnumOLEVERB** ppEnumerator)
{
	static OLEVERB oleVerbs[3] = {
	    {OLEIVERB_UIACTIVATE, L"&Edit", 0, OLEVERBATTRIB_NEVERDIRTIES | OLEVERBATTRIB_ONCONTAINERMENU},
	    {OLEIVERB_PROPERTIES, L"&Properties...", 0, OLEVERBATTRIB_ONCONTAINERMENU},
	    {1, L"&About...", 0, OLEVERBATTRIB_NEVERDIRTIES | OLEVERBATTRIB_ONCONTAINERMENU},
	};
	return EnumOLEVERB::CreateInstance(oleVerbs, 3, ppEnumerator);
}
// implementation of IOleObject
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IOleControl
STDMETHODIMP TrackBar::GetControlInfo(LPCONTROLINFO pControlInfo)
{
	ATLASSERT_POINTER(pControlInfo, CONTROLINFO);
	if(!pControlInfo) {
		return E_POINTER;
	}

	// our control can have an accelerator
	pControlInfo->cb = sizeof(CONTROLINFO);
	pControlInfo->hAccel = NULL;
	pControlInfo->cAccel = 0;
	pControlInfo->dwFlags = 0;
	return S_OK;
}
// implementation of IOleControl
//////////////////////////////////////////////////////////////////////

HRESULT TrackBar::DoVerbAbout(HWND hWndParent)
{
	HRESULT hr = S_OK;
	//hr = OnPreVerbAbout();
	if(SUCCEEDED(hr))	{
		AboutDlg dlg;
		dlg.SetOwner(this);
		dlg.DoModal(hWndParent);
		hr = S_OK;
		//hr = OnPostVerbAbout();
	}
	return hr;
}

HRESULT TrackBar::OnPropertyObjectChanged(DISPID /*propertyObject*/, DISPID /*objectProperty*/)
{
	return S_OK;
}

HRESULT TrackBar::OnPropertyObjectRequestEdit(DISPID /*propertyObject*/, DISPID /*objectProperty*/)
{
	return S_OK;
}


STDMETHODIMP TrackBar::get_Appearance(AppearanceConstants* pValue)
{
	ATLASSERT_POINTER(pValue, AppearanceConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if(GetExStyle() & WS_EX_CLIENTEDGE) {
			properties.appearance = a3D;
		} else if(GetExStyle() & WS_EX_STATICEDGE) {
			properties.appearance = a3DLight;
		} else {
			properties.appearance = a2D;
		}
	}

	*pValue = properties.appearance;
	return S_OK;
}

STDMETHODIMP TrackBar::put_Appearance(AppearanceConstants newValue)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_APPEARANCE);
	if(properties.appearance != newValue) {
		properties.appearance = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.appearance) {
				case a2D:
					ModifyStyleEx(WS_EX_CLIENTEDGE | WS_EX_STATICEDGE, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case a3D:
					ModifyStyleEx(WS_EX_STATICEDGE, WS_EX_CLIENTEDGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case a3DLight:
					ModifyStyleEx(WS_EX_CLIENTEDGE, WS_EX_STATICEDGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_TRACKBARCTL_APPEARANCE);
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_AppID(SHORT* pValue)
{
	ATLASSERT_POINTER(pValue, SHORT);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = 11;
	return S_OK;
}

STDMETHODIMP TrackBar::get_AppName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"TrackBar");
	return S_OK;
}

STDMETHODIMP TrackBar::get_AppShortName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"TrackBarCtl");
	return S_OK;
}

STDMETHODIMP TrackBar::get_AutoTickFrequency(SHORT* pValue)
{
	ATLASSERT_POINTER(pValue, SHORT);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.autoTickFrequency;
	return S_OK;
}

STDMETHODIMP TrackBar::put_AutoTickFrequency(SHORT newValue)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_AUTOTICKFREQUENCY);
	if(properties.autoTickFrequency != newValue) {
		properties.autoTickFrequency = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(TBM_SETTICFREQ, properties.autoTickFrequency, 0);
		}
		FireOnChanged(DISPID_TRACKBARCTL_AUTOTICKFREQUENCY);
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_AutoTickMarks(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.autoTickMarks = ((GetStyle() & TBS_AUTOTICKS) == TBS_AUTOTICKS);
	}

	*pValue = BOOL2VARIANTBOOL(properties.autoTickMarks);
	return S_OK;
}

STDMETHODIMP TrackBar::put_AutoTickMarks(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_AUTOTICKMARKS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.autoTickMarks != b) {
		properties.autoTickMarks = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_TRACKBARCTL_AUTOTICKMARKS);
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_BackColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.backColor;
	return S_OK;
}

STDMETHODIMP TrackBar::put_BackColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_BACKCOLOR);
	if(properties.backColor != newValue) {
		properties.backColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			WINDOWPOS details = {0};
			details.hwnd = *this;
			details.flags = SWP_DRAWFRAME | SWP_FRAMECHANGED | SWP_NOCOPYBITS | SWP_NOMOVE | SWP_NOSIZE | SWP_NOOWNERZORDER | SWP_NOZORDER;
			SendMessage(WM_WINDOWPOSCHANGED, 0, reinterpret_cast<LPARAM>(&details));
		}
		FireOnChanged(DISPID_TRACKBARCTL_BACKCOLOR);
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_BackgroundDrawMode(BackgroundDrawModeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BackgroundDrawModeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		if(GetStyle() & TBS_TRANSPARENTBKGND) {
			properties.backgroundDrawMode = bdmByParent;
		} else {
			properties.backgroundDrawMode = bdmNormal;
		}
	}

	*pValue = properties.backgroundDrawMode;
	return S_OK;
}

STDMETHODIMP TrackBar::put_BackgroundDrawMode(BackgroundDrawModeConstants newValue)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_BACKGROUNDDRAWMODE);
	if(properties.backgroundDrawMode != newValue) {
		properties.backgroundDrawMode = newValue;
		SetDirty(TRUE);

		if(IsWindow() && IsComctl32Version610OrNewer()) {
			switch(properties.backgroundDrawMode) {
				case bdmNormal:
					ModifyStyle(TBS_TRANSPARENTBKGND, 0);
					break;
				case bdmByParent:
					ModifyStyle(0, TBS_TRANSPARENTBKGND);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_TRACKBARCTL_BACKGROUNDDRAWMODE);
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_BorderStyle(BorderStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BorderStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.borderStyle = ((GetStyle() & WS_BORDER) == WS_BORDER ? bsFixedSingle : bsNone);
	}

	*pValue = properties.borderStyle;
	return S_OK;
}

STDMETHODIMP TrackBar::put_BorderStyle(BorderStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_BORDERSTYLE);
	if(properties.borderStyle != newValue) {
		properties.borderStyle = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.borderStyle) {
				case bsNone:
					ModifyStyle(WS_BORDER, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case bsFixedSingle:
					ModifyStyle(0, WS_BORDER, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_TRACKBARCTL_BORDERSTYLE);
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_Build(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = VERSION_BUILD;
	return S_OK;
}

STDMETHODIMP TrackBar::get_ChannelHeight(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		WTL::CRect channelRectangle;
		SendMessage(TBM_GETCHANNELRECT, 0, reinterpret_cast<LPARAM>(&channelRectangle));
		*pValue = channelRectangle.Height();
	} else {
		*pValue = 0;
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_ChannelLeft(OLE_XPOS_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XPOS_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		RECT channelRectangle = {0};
		SendMessage(TBM_GETCHANNELRECT, 0, reinterpret_cast<LPARAM>(&channelRectangle));
		*pValue = channelRectangle.left;
	} else {
		*pValue = 0;
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_ChannelTop(OLE_YPOS_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		RECT channelRectangle = {0};
		SendMessage(TBM_GETCHANNELRECT, 0, reinterpret_cast<LPARAM>(&channelRectangle));
		*pValue = channelRectangle.top;
	} else {
		*pValue = 0;
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_ChannelWidth(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		WTL::CRect channelRectangle;
		SendMessage(TBM_GETCHANNELRECT, 0, reinterpret_cast<LPARAM>(&channelRectangle));
		*pValue = channelRectangle.Width();
	} else {
		*pValue = 0;
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_CharSet(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	#ifdef UNICODE
		*pValue = SysAllocString(L"Unicode");
	#else
		*pValue = SysAllocString(L"ANSI");
	#endif
	return S_OK;
}

STDMETHODIMP TrackBar::get_CurrentPosition(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.currentPosition = SendMessage(TBM_GETPOS, 0, 0);
	}

	*pValue = properties.currentPosition;
	return S_OK;
}

STDMETHODIMP TrackBar::put_CurrentPosition(LONG newValue)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_CURRENTPOSITION);
	if(properties.currentPosition != newValue) {
		if(IsWindow()) {
			SendMessage(TBM_SETPOS, TRUE, newValue);
		}
		properties.currentPosition = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TRACKBARCTL_CURRENTPOSITION);
		SendOnDataChange();
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_DetectDoubleClicks(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.detectDoubleClicks);
	return S_OK;
}

STDMETHODIMP TrackBar::put_DetectDoubleClicks(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_DETECTDOUBLECLICKS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.detectDoubleClicks != b) {
		properties.detectDoubleClicks = b;
		SetDirty(TRUE);

		FireOnChanged(DISPID_TRACKBARCTL_DETECTDOUBLECLICKS);
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_DisabledEvents(DisabledEventsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, DisabledEventsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.disabledEvents;
	return S_OK;
}

STDMETHODIMP TrackBar::put_DisabledEvents(DisabledEventsConstants newValue)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_DISABLEDEVENTS);
	if(properties.disabledEvents != newValue) {
		if((properties.disabledEvents & deMouseEvents) != (newValue & deMouseEvents)) {
			if(IsWindow()) {
				if(newValue & deMouseEvents) {
					// nothing to do
				} else {
					TRACKMOUSEEVENT trackingOptions = {0};
					trackingOptions.cbSize = sizeof(trackingOptions);
					trackingOptions.hwndTrack = *this;
					trackingOptions.dwFlags = TME_HOVER | TME_LEAVE | TME_CANCEL;
					TrackMouseEvent(&trackingOptions);
				}
			}
		}

		properties.disabledEvents = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TRACKBARCTL_DISABLEDEVENTS);
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_DontRedraw(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.dontRedraw);
	return S_OK;
}

STDMETHODIMP TrackBar::put_DontRedraw(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_DONTREDRAW);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.dontRedraw != b) {
		properties.dontRedraw = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			SetRedraw(!b);
		}
		FireOnChanged(DISPID_TRACKBARCTL_DONTREDRAW);
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_DownIsLeft(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()/* && IsComctl32Version581OrNewer()*/) {
		properties.downIsLeft = ((GetStyle() & TBS_DOWNISLEFT) == TBS_DOWNISLEFT);
	}

	*pValue = BOOL2VARIANTBOOL(properties.downIsLeft);
	return S_OK;
}

STDMETHODIMP TrackBar::put_DownIsLeft(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_DOWNISLEFT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.downIsLeft != b) {
		properties.downIsLeft = b;
		SetDirty(TRUE);

		if(IsWindow()/* && IsComctl32Version581OrNewer()*/) {
			if(properties.downIsLeft) {
				ModifyStyle(0, TBS_DOWNISLEFT);
			} else {
				ModifyStyle(TBS_DOWNISLEFT, 0);
			}
		}
		FireOnChanged(DISPID_TRACKBARCTL_DOWNISLEFT);
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_Enabled(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.enabled = !(GetStyle() & WS_DISABLED);
	}

	*pValue = BOOL2VARIANTBOOL(properties.enabled);
	return S_OK;
}

STDMETHODIMP TrackBar::put_Enabled(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_ENABLED);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.enabled != b) {
		properties.enabled = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			EnableWindow(properties.enabled);
			FireViewChange();
		}

		if(!properties.enabled) {
			IOleInPlaceObject_UIDeactivate();
		}

		FireOnChanged(DISPID_TRACKBARCTL_ENABLED);
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_HoverTime(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.hoverTime;
	return S_OK;
}

STDMETHODIMP TrackBar::put_HoverTime(LONG newValue)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_HOVERTIME);
	if((newValue < 0) && (newValue != -1)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.hoverTime != newValue) {
		properties.hoverTime = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TRACKBARCTL_HOVERTIME);
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_hWnd(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = HandleToLong(static_cast<HWND>(*this));
	return S_OK;
}

STDMETHODIMP TrackBar::get_hWndBottomOrRightBuddy(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = HandleToLong(reinterpret_cast<HWND>(SendMessage(TBM_GETBUDDY, FALSE, 0)));
	}
	return S_OK;
}

STDMETHODIMP TrackBar::put_hWndBottomOrRightBuddy(OLE_HANDLE newValue)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_HWNDBOTTOMORRIGHTBUDDY);
	if(IsWindow()) {
		SendMessage(TBM_SETBUDDY, FALSE, reinterpret_cast<LPARAM>(LongToHandle(newValue)));
	}
	SetDirty(TRUE);
	FireOnChanged(DISPID_TRACKBARCTL_HWNDBOTTOMORRIGHTBUDDY);
	FireViewChange();
	return S_OK;
}

STDMETHODIMP TrackBar::get_hWndToolTip(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = HandleToLong(reinterpret_cast<HWND>(SendMessage(TBM_GETTOOLTIPS, 0, 0)));
	}
	return S_OK;
}

STDMETHODIMP TrackBar::put_hWndToolTip(OLE_HANDLE newValue)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_HWNDTOOLTIP);
	if(IsWindow()) {
		SendMessage(TBM_SETTOOLTIPS, reinterpret_cast<WPARAM>(LongToHandle(newValue)), 0);
	}
	SetDirty(TRUE);
	FireOnChanged(DISPID_TRACKBARCTL_HWNDTOOLTIP);
	return S_OK;
}

STDMETHODIMP TrackBar::get_hWndTopOrLeftBuddy(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = HandleToLong(reinterpret_cast<HWND>(SendMessage(TBM_GETBUDDY, TRUE, 0)));
	}
	return S_OK;
}

STDMETHODIMP TrackBar::put_hWndTopOrLeftBuddy(OLE_HANDLE newValue)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_HWNDTOPORLEFTBUDDY);
	if(IsWindow()) {
		SendMessage(TBM_SETBUDDY, TRUE, reinterpret_cast<LPARAM>(LongToHandle(newValue)));
	}
	SetDirty(TRUE);
	FireOnChanged(DISPID_TRACKBARCTL_HWNDTOPORLEFTBUDDY);
	FireViewChange();
	return S_OK;
}

STDMETHODIMP TrackBar::get_IsRelease(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	#ifdef NDEBUG
		*pValue = VARIANT_TRUE;
	#else
		*pValue = VARIANT_FALSE;
	#endif
	return S_OK;
}

STDMETHODIMP TrackBar::get_LargeStepWidth(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.largeStepWidth = SendMessage(TBM_GETPAGESIZE, 0, 0);
	}

	*pValue = properties.largeStepWidth;
	return S_OK;
}

STDMETHODIMP TrackBar::put_LargeStepWidth(LONG newValue)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_LARGESTEPWIDTH);
	if(properties.largeStepWidth != newValue) {
		properties.largeStepWidth = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(TBM_SETPAGESIZE, 0, properties.largeStepWidth);
		}
		FireOnChanged(DISPID_TRACKBARCTL_LARGESTEPWIDTH);
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_Maximum(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.maximum = SendMessage(TBM_GETRANGEMAX, 0, 0);
	}

	*pValue = properties.maximum;
	return S_OK;
}

STDMETHODIMP TrackBar::put_Maximum(LONG newValue)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_MAXIMUM);
	if(properties.maximum != newValue) {
		properties.maximum = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(TBM_SETRANGEMAX, TRUE, properties.maximum);
		}
		FireOnChanged(DISPID_TRACKBARCTL_MAXIMUM);
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_Minimum(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.minimum = SendMessage(TBM_GETRANGEMIN, 0, 0);
	}

	*pValue = properties.minimum;
	return S_OK;
}

STDMETHODIMP TrackBar::put_Minimum(LONG newValue)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_MINIMUM);
	if(properties.minimum != newValue) {
		properties.minimum = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(TBM_SETRANGEMIN, TRUE, properties.minimum);
		}
		FireOnChanged(DISPID_TRACKBARCTL_MINIMUM);
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_MouseIcon(IPictureDisp** ppMouseIcon)
{
	ATLASSERT_POINTER(ppMouseIcon, IPictureDisp*);
	if(!ppMouseIcon) {
		return E_POINTER;
	}

	if(*ppMouseIcon) {
		(*ppMouseIcon)->Release();
		*ppMouseIcon = NULL;
	}
	if(properties.mouseIcon.pPictureDisp) {
		properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPictureDisp, reinterpret_cast<LPVOID*>(ppMouseIcon));
	}
	return S_OK;
}

STDMETHODIMP TrackBar::put_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_MOUSEICON);
	if(properties.mouseIcon.pPictureDisp != pNewMouseIcon) {
		properties.mouseIcon.StopWatching();
		if(properties.mouseIcon.pPictureDisp) {
			properties.mouseIcon.pPictureDisp->Release();
			properties.mouseIcon.pPictureDisp = NULL;
		}
		if(pNewMouseIcon) {
			// clone the picture by storing it into a stream
			CComQIPtr<IPersistStream, &IID_IPersistStream> pPersistStream(pNewMouseIcon);
			if(pPersistStream) {
				ULARGE_INTEGER pictureSize = {0};
				pPersistStream->GetSizeMax(&pictureSize);
				HGLOBAL hGlobalMem = GlobalAlloc(GHND, pictureSize.LowPart);
				if(hGlobalMem) {
					CComPtr<IStream> pStream = NULL;
					CreateStreamOnHGlobal(hGlobalMem, TRUE, &pStream);
					if(pStream) {
						if(SUCCEEDED(pPersistStream->Save(pStream, FALSE))) {
							LARGE_INTEGER startPosition = {0};
							pStream->Seek(startPosition, STREAM_SEEK_SET, NULL);
							OleLoadPicture(pStream, startPosition.LowPart, FALSE, IID_IPictureDisp, reinterpret_cast<LPVOID*>(&properties.mouseIcon.pPictureDisp));
						}
						pStream.Release();
					}
					GlobalFree(hGlobalMem);
				}
			}
		}
		properties.mouseIcon.StartWatching();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_TRACKBARCTL_MOUSEICON);
	return S_OK;
}

STDMETHODIMP TrackBar::putref_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_MOUSEICON);
	if(properties.mouseIcon.pPictureDisp != pNewMouseIcon) {
		properties.mouseIcon.StopWatching();
		if(properties.mouseIcon.pPictureDisp) {
			properties.mouseIcon.pPictureDisp->Release();
			properties.mouseIcon.pPictureDisp = NULL;
		}
		if(pNewMouseIcon) {
			pNewMouseIcon->QueryInterface(IID_IPictureDisp, reinterpret_cast<LPVOID*>(&properties.mouseIcon.pPictureDisp));
		}
		properties.mouseIcon.StartWatching();
	} else if(pNewMouseIcon) {
		pNewMouseIcon->AddRef();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_TRACKBARCTL_MOUSEICON);
	return S_OK;
}

STDMETHODIMP TrackBar::get_MousePointer(MousePointerConstants* pValue)
{
	ATLASSERT_POINTER(pValue, MousePointerConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.mousePointer;
	return S_OK;
}

STDMETHODIMP TrackBar::put_MousePointer(MousePointerConstants newValue)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_MOUSEPOINTER);
	if(properties.mousePointer != newValue) {
		properties.mousePointer = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TRACKBARCTL_MOUSEPOINTER);
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_Orientation(OrientationConstants* pValue)
{
	ATLASSERT_POINTER(pValue, OrientationConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.orientation = ((GetStyle() & TBS_VERT) == TBS_VERT ? oVertical : oHorizontal);
	}

	*pValue = properties.orientation;
	return S_OK;
}

STDMETHODIMP TrackBar::put_Orientation(OrientationConstants newValue)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_ORIENTATION);
	if(properties.orientation != newValue) {
		properties.orientation = newValue;
		SetDirty(TRUE);

		RECT windowRect = m_rcPos;
		SIZEL himetric = {m_sizeExtent.cx, m_sizeExtent.cy};
		SIZEL pixels = {0};
		AtlHiMetricToPixel(&himetric, &pixels);
		windowRect.right = windowRect.left + pixels.cy;
		windowRect.bottom = windowRect.top + pixels.cx;
		ATLASSUME(m_spInPlaceSite);
		if(m_spInPlaceSite) {
			m_spInPlaceSite->OnPosRectChange(&windowRect);
		}

		if(IsWindow()) {
			switch(properties.orientation) {
				case oHorizontal:
					ModifyStyle(TBS_VERT, TBS_HORZ, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case oVertical:
					ModifyStyle(TBS_HORZ, TBS_VERT, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_TRACKBARCTL_ORIENTATION);
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_ProcessContextMenuKeys(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.processContextMenuKeys);
	return S_OK;
}

STDMETHODIMP TrackBar::put_ProcessContextMenuKeys(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_PROCESSCONTEXTMENUKEYS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.processContextMenuKeys != b) {
		properties.processContextMenuKeys = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TRACKBARCTL_PROCESSCONTEXTMENUKEYS);
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_Programmer(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP TrackBar::get_RangeSelectionEnd(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && ((GetStyle() & TBS_ENABLESELRANGE) == TBS_ENABLESELRANGE)) {
		properties.rangeSelectionEnd = SendMessage(TBM_GETSELEND, 0, 0);
	}

	*pValue = properties.rangeSelectionEnd;
	return S_OK;
}

STDMETHODIMP TrackBar::put_RangeSelectionEnd(LONG newValue)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_RANGESELECTIONEND);
	if(properties.rangeSelectionEnd != newValue) {
		properties.rangeSelectionEnd = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(TBM_SETSELEND, TRUE, properties.rangeSelectionEnd);
		}
		FireOnChanged(DISPID_TRACKBARCTL_RANGESELECTIONEND);
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_RangeSelectionStart(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && ((GetStyle() & TBS_ENABLESELRANGE) == TBS_ENABLESELRANGE)) {
		properties.rangeSelectionStart = SendMessage(TBM_GETSELSTART, 0, 0);
	}

	*pValue = properties.rangeSelectionStart;
	return S_OK;
}

STDMETHODIMP TrackBar::put_RangeSelectionStart(LONG newValue)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_RANGESELECTIONSTART);
	if(properties.rangeSelectionStart != newValue) {
		properties.rangeSelectionStart = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(TBM_SETSELSTART, TRUE, properties.rangeSelectionStart);
		}
		FireOnChanged(DISPID_TRACKBARCTL_RANGESELECTIONSTART);
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_RegisterForOLEDragDrop(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.registerForOLEDragDrop);
	return S_OK;
}

STDMETHODIMP TrackBar::put_RegisterForOLEDragDrop(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_REGISTERFOROLEDRAGDROP);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.registerForOLEDragDrop != b) {
		properties.registerForOLEDragDrop = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.registerForOLEDragDrop) {
				ATLVERIFY(RegisterDragDrop(*this, static_cast<IDropTarget*>(this)) == S_OK);
			} else {
				ATLVERIFY(RevokeDragDrop(*this) == S_OK);
			}
		}
		FireOnChanged(DISPID_TRACKBARCTL_REGISTERFOROLEDRAGDROP);
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_Reversed(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()/* && IsComctl32Version580OrNewer()*/) {
		properties.reversed = ((GetStyle() & TBS_REVERSED) == TBS_REVERSED);
	}

	*pValue = BOOL2VARIANTBOOL(properties.reversed);
	return S_OK;
}

STDMETHODIMP TrackBar::put_Reversed(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_REVERSED);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.reversed != b) {
		properties.reversed = b;
		SetDirty(TRUE);

		if(IsWindow()/* && IsComctl32Version580OrNewer()*/) {
			if(properties.reversed) {
				ModifyStyle(0, TBS_REVERSED);
			} else {
				ModifyStyle(TBS_REVERSED, 0);
			}
		}
		FireOnChanged(DISPID_TRACKBARCTL_REVERSED);
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_RightToLeftLayout(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.rightToLeftLayout = ((GetExStyle() & WS_EX_LAYOUTRTL) == WS_EX_LAYOUTRTL);
	}

	*pValue = BOOL2VARIANTBOOL(properties.rightToLeftLayout);
	return S_OK;
}

STDMETHODIMP TrackBar::put_RightToLeftLayout(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_RIGHTTOLEFTLAYOUT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.rightToLeftLayout != b) {
		properties.rightToLeftLayout = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_TRACKBARCTL_RIGHTTOLEFTLAYOUT);
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_SelectionType(SelectionTypeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, OrientationConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.selectionType = ((GetStyle() & TBS_ENABLESELRANGE) == TBS_ENABLESELRANGE ? stRangeSelection : stDiscreteValue);
	}

	*pValue = properties.selectionType;
	return S_OK;
}

STDMETHODIMP TrackBar::put_SelectionType(SelectionTypeConstants newValue)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_SELECTIONTYPE);
	if(properties.selectionType != newValue) {
		properties.selectionType = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.selectionType) {
				case stDiscreteValue:
					ModifyStyle(TBS_ENABLESELRANGE, 0);
					SendMessage(TBM_CLEARSEL, TRUE, 0);
					break;
				case stRangeSelection:
					ModifyStyle(0, TBS_ENABLESELRANGE);
					SendMessage(TBM_SETSELSTART, FALSE, properties.rangeSelectionStart);
					SendMessage(TBM_SETSELEND, TRUE, properties.rangeSelectionEnd);
					break;
			}
		}
		FireOnChanged(DISPID_TRACKBARCTL_SELECTIONTYPE);
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_ShowSlider(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.showSlider = ((GetStyle() & TBS_NOTHUMB) == 0);
	}

	*pValue = BOOL2VARIANTBOOL(properties.showSlider);
	return S_OK;
}

STDMETHODIMP TrackBar::put_ShowSlider(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_SHOWSLIDER);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.showSlider != b) {
		properties.showSlider = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.showSlider) {
				ModifyStyle(TBS_NOTHUMB, 0);
			} else {
				ModifyStyle(0, TBS_NOTHUMB);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_TRACKBARCTL_SHOWSLIDER);
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_SliderHeight(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		WTL::CRect sliderRectangle;
		SendMessage(TBM_GETTHUMBRECT, 0, reinterpret_cast<LPARAM>(&sliderRectangle));
		*pValue = sliderRectangle.Height();
	} else {
		*pValue = 0;
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_SliderLeft(OLE_XPOS_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XPOS_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		RECT sliderRectangle = {0};
		SendMessage(TBM_GETTHUMBRECT, 0, reinterpret_cast<LPARAM>(&sliderRectangle));
		*pValue = sliderRectangle.left;
	} else {
		*pValue = 0;
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_SliderLength(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		DWORD style = GetStyle();
		if(style & TBS_FIXEDLENGTH) {
			properties.sliderLength = SendMessage(TBM_GETTHUMBLENGTH, 0, 0);
		} else {
			WTL::CRect sliderRectangle;
			SendMessage(TBM_GETTHUMBRECT, 0, reinterpret_cast<LPARAM>(&sliderRectangle));
			if(style & TBS_VERT) {
				properties.sliderLength = sliderRectangle.Width();
			} else {
				properties.sliderLength = sliderRectangle.Height();
			}
		}
	}

	*pValue = properties.sliderLength;
	return S_OK;
}

STDMETHODIMP TrackBar::put_SliderLength(LONG newValue)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_SLIDERLENGTH);
	if(properties.sliderLength != newValue) {
		properties.sliderLength = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.sliderLength == -1) {
				ModifyStyle(TBS_FIXEDLENGTH, 0);
			} else {
				ModifyStyle(0, TBS_FIXEDLENGTH);
				SendMessage(TBM_SETTHUMBLENGTH, properties.sliderLength, 0);
			}
		}
		FireOnChanged(DISPID_TRACKBARCTL_SLIDERLENGTH);
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_SliderTop(OLE_YPOS_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		RECT sliderRectangle = {0};
		SendMessage(TBM_GETTHUMBRECT, 0, reinterpret_cast<LPARAM>(&sliderRectangle));
		*pValue = sliderRectangle.top;
	} else {
		*pValue = 0;
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_SliderWidth(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		WTL::CRect sliderRectangle;
		SendMessage(TBM_GETTHUMBRECT, 0, reinterpret_cast<LPARAM>(&sliderRectangle));
		*pValue = sliderRectangle.Width();
	} else {
		*pValue = 0;
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_SmallStepWidth(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.smallStepWidth = SendMessage(TBM_GETLINESIZE, 0, 0);
	}

	*pValue = properties.smallStepWidth;
	return S_OK;
}

STDMETHODIMP TrackBar::put_SmallStepWidth(LONG newValue)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_SMALLSTEPWIDTH);
	if(properties.smallStepWidth != newValue) {
		properties.smallStepWidth = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(TBM_SETLINESIZE, 0, properties.smallStepWidth);
		}
		FireOnChanged(DISPID_TRACKBARCTL_SMALLSTEPWIDTH);
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_SupportOLEDragImages(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue =  BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	return S_OK;
}

STDMETHODIMP TrackBar::put_SupportOLEDragImages(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_SUPPORTOLEDRAGIMAGES);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.supportOLEDragImages != b) {
		properties.supportOLEDragImages = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TRACKBARCTL_SUPPORTOLEDRAGIMAGES);
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_Tester(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP TrackBar::get_TickMarkPhysicalPosition(LONG tickMark, LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		/* HACK: Comctl32 doesn't tell us the position of the first and last tick marks, so set the current
		         value to the minimum/maximum without redraw, retrieve the slider's rectangle and reset the
		         position. Then use the rectangle to calculate the tick mark's position. */
		if(tickMark == 0) {
			CWindowEx(*this).InternalSetRedraw(FALSE);
			EnterSilentPositionChangesSection();

			LONG currentPosition = SendMessage(TBM_GETPOS, 0, 0);
			SendMessage(TBM_SETPOS, TRUE, SendMessage(TBM_GETRANGEMIN, 0, 0));

			WTL::CRect sliderRectangle;
			SendMessage(TBM_GETTHUMBRECT, 0, reinterpret_cast<LPARAM>(&sliderRectangle));
			if(GetStyle() & TBS_VERT) {
				*pValue = sliderRectangle.CenterPoint().y;
			} else {
				*pValue = sliderRectangle.CenterPoint().x;
			}

			SendMessage(TBM_SETPOS, TRUE, currentPosition);

			LeaveSilentPositionChangesSection();
			CWindowEx(*this).InternalSetRedraw(TRUE);
		} else if(tickMark == (SendMessage(TBM_GETNUMTICS, 0, 0)) - 1) {
			CWindowEx(*this).InternalSetRedraw(FALSE);
			EnterSilentPositionChangesSection();

			LONG currentPosition = SendMessage(TBM_GETPOS, 0, 0);
			SendMessage(TBM_SETPOS, TRUE, SendMessage(TBM_GETRANGEMAX, 0, 0));

			WTL::CRect sliderRectangle;
			SendMessage(TBM_GETTHUMBRECT, 0, reinterpret_cast<LPARAM>(&sliderRectangle));
			if(GetStyle() & TBS_VERT) {
				*pValue = sliderRectangle.CenterPoint().y;
			} else {
				*pValue = sliderRectangle.CenterPoint().x;
			}

			SendMessage(TBM_SETPOS, TRUE, currentPosition);

			LeaveSilentPositionChangesSection();
			CWindowEx(*this).InternalSetRedraw(TRUE);
		} else {
			*pValue = SendMessage(TBM_GETTICPOS, tickMark - 1, 0);
		}
		if(*pValue == -1) {
			// invalid arg - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
		}
		return S_OK;
	}

	return E_FAIL;
}

STDMETHODIMP TrackBar::get_TickMarks(VARIANT* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		VariantClear(pValue);

		int numberOfTicks = SendMessage(TBM_GETNUMTICS, 0, 0);
		if(numberOfTicks > 0) {
			LPDWORD pTickMarks = reinterpret_cast<LPDWORD>(SendMessage(TBM_GETPTICS, 0, 0));

			// create the array
			pValue->vt = VT_ARRAY | VT_I4;
			pValue->parray = SafeArrayCreateVectorEx(VT_I4, 1, numberOfTicks, NULL);

			// set the first and last elements to the control's min and max range
			properties.minimum = SendMessage(TBM_GETRANGEMIN, 0, 0);
			properties.maximum = SendMessage(TBM_GETRANGEMAX, 0, 0);
			LONG i = 1;
			SafeArrayPutElement(pValue->parray, &i, &properties.minimum);
			i = numberOfTicks;
			SafeArrayPutElement(pValue->parray, &i, &properties.maximum);
			// set the remaining elements
			for(i = 2; i < numberOfTicks; ++i) {
				SafeArrayPutElement(pValue->parray, &i, &pTickMarks[i - 2]);
			}
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TrackBar::put_TickMarks(VARIANT newValue)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_TICKMARKS);
	HRESULT hr = E_FAIL;
	if(IsWindow()) {
		hr = S_OK;

		if(newValue.vt == VT_EMPTY) {
			SendMessage(TBM_CLEARTICS, TRUE, 0);
			SetDirty(TRUE);
			FireOnChanged(DISPID_TRACKBARCTL_TICKMARKS);
		} else if(newValue.vt & VT_ARRAY) {
			// an array
			SendMessage(TBM_CLEARTICS, FALSE, 0);

			if(newValue.parray) {
				LONG l = 0;
				SafeArrayGetLBound(newValue.parray, 1, &l);
				LONG u = 0;
				SafeArrayGetUBound(newValue.parray, 1, &u);
				if(u < l) {
					// invalid arg - raise VB runtime error 380
					return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
				}

				properties.minimum = SendMessage(TBM_GETRANGEMIN, 0, 0);
				properties.maximum = SendMessage(TBM_GETRANGEMAX, 0, 0);

				VARTYPE vt = 0;
				SafeArrayGetVartype(newValue.parray, &vt);
				for(LONG i = l; i <= u; ++i) {
					LONG tickMarkPosition = 0;
					if(vt == VT_I1) {
						CHAR buffer = 0;
						SafeArrayGetElement(newValue.parray, &i, &buffer);
						tickMarkPosition = buffer;
					} else if(vt == VT_UI1) {
						BYTE buffer = 0;
						SafeArrayGetElement(newValue.parray, &i, &buffer);
						tickMarkPosition = buffer;
					} else if(vt == VT_I2) {
						SHORT buffer = 0;
						SafeArrayGetElement(newValue.parray, &i, &buffer);
						tickMarkPosition = buffer;
					} else if(vt == VT_UI2) {
						USHORT buffer = 0;
						SafeArrayGetElement(newValue.parray, &i, &buffer);
						tickMarkPosition = buffer;
					} else if(vt == VT_I4) {
						LONG buffer = 0;
						SafeArrayGetElement(newValue.parray, &i, &buffer);
						tickMarkPosition = buffer;
					} else if(vt == VT_UI4) {
						ULONG buffer = 0;
						SafeArrayGetElement(newValue.parray, &i, &buffer);
						tickMarkPosition = buffer;
					} else if(vt == VT_INT) {
						INT buffer = 0;
						SafeArrayGetElement(newValue.parray, &i, &buffer);
						tickMarkPosition = buffer;
					} else if(vt == VT_UINT) {
						UINT buffer = 0;
						SafeArrayGetElement(newValue.parray, &i, &buffer);
						tickMarkPosition = buffer;
					} else if(vt == VT_VARIANT) {
						VARIANT buffer;
						SafeArrayGetElement(newValue.parray, &i, &buffer);
						if(buffer.vt == VT_I1) {
							tickMarkPosition = buffer.cVal;
						} else if(buffer.vt == VT_UI1) {
							tickMarkPosition = buffer.bVal;
						} else if(buffer.vt == VT_I2) {
							tickMarkPosition = buffer.iVal;
						} else if(buffer.vt == VT_UI2) {
							tickMarkPosition = buffer.uiVal;
						} else if(buffer.vt == VT_I4) {
							tickMarkPosition = buffer.lVal;
						} else if(buffer.vt == VT_UI4) {
							tickMarkPosition = buffer.ulVal;
						} else if(buffer.vt == VT_INT) {
							tickMarkPosition = buffer.intVal;
						} else if(buffer.vt == VT_UINT) {
							tickMarkPosition = buffer.uintVal;
						} else {
							// invalid arg - raise VB runtime error 380
							return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
						}
					} else {
						// invalid arg - raise VB runtime error 380
						return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
					}

					if((tickMarkPosition != properties.minimum) && (tickMarkPosition != properties.maximum)) {
						if(!SendMessage(TBM_SETTIC, 0, tickMarkPosition)) {
							hr = E_FAIL;
						}
					}
				}
			}
			SetDirty(TRUE);
			FireViewChange();
			FireOnChanged(DISPID_TRACKBARCTL_TICKMARKS);
		} else {
			// a single value
			LONG tickMarkPosition = 0;
			if(newValue.vt == VT_I1) {
				tickMarkPosition = newValue.cVal;
			} else if(newValue.vt == VT_UI1) {
				tickMarkPosition = newValue.bVal;
			} else if(newValue.vt == VT_I2) {
				tickMarkPosition = newValue.iVal;
			} else if(newValue.vt == VT_UI2) {
				tickMarkPosition = newValue.uiVal;
			} else if(newValue.vt == VT_I4) {
				tickMarkPosition = newValue.lVal;
			} else if(newValue.vt == VT_UI4) {
				tickMarkPosition = newValue.ulVal;
			} else if(newValue.vt == VT_INT) {
				tickMarkPosition = newValue.intVal;
			} else if(newValue.vt == VT_UINT) {
				tickMarkPosition = newValue.uintVal;
			} else {
				// invalid arg - raise VB runtime error 380
				return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			}

			if((tickMarkPosition != properties.minimum) && (tickMarkPosition != properties.maximum)) {
				if(!SendMessage(TBM_SETTIC, 0, tickMarkPosition)) {
					hr = E_FAIL;
				}
			}
			SetDirty(TRUE);
			FireViewChange();
			FireOnChanged(DISPID_TRACKBARCTL_TICKMARKS);
		}
	}
	return hr;
}

STDMETHODIMP TrackBar::get_TickMarksPosition(TickMarksPositionConstants* pValue)
{
	ATLASSERT_POINTER(pValue, TickMarksPositionConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		DWORD style = GetStyle();
		if(style & TBS_NOTICKS) {
			properties.tickMarksPosition = tmpNone;
		} else if(style & TBS_BOTH) {
			properties.tickMarksPosition = tmpBoth;
		} else if(style & (TBS_TOP | TBS_LEFT)) {
			properties.tickMarksPosition = tmpTopOrLeft;
		} else/* if(style & (TBS_BOTTOM | TBS_RIGHT))*/ {
			properties.tickMarksPosition = tmpBottomOrRight;
		}
	}

	*pValue = properties.tickMarksPosition;
	return S_OK;
}

STDMETHODIMP TrackBar::put_TickMarksPosition(TickMarksPositionConstants newValue)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_TICKMARKSPOSITION);
	if(properties.tickMarksPosition != newValue) {
		properties.tickMarksPosition = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.tickMarksPosition) {
				case tmpNone:
					ModifyStyle(TBS_TOP | TBS_BOTTOM | TBS_LEFT | TBS_RIGHT | TBS_BOTH, TBS_NOTICKS);
					break;
				case tmpBottomOrRight:
					ModifyStyle(TBS_NOTICKS | TBS_TOP | TBS_LEFT | TBS_BOTH, TBS_BOTTOM | TBS_RIGHT);
					break;
				case tmpTopOrLeft:
					ModifyStyle(TBS_NOTICKS | TBS_BOTTOM | TBS_RIGHT | TBS_BOTH, TBS_TOP | TBS_LEFT);
					break;
				case tmpBoth:
					ModifyStyle(TBS_NOTICKS | TBS_TOP | TBS_BOTTOM | TBS_LEFT | TBS_RIGHT, TBS_BOTH);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_TRACKBARCTL_TICKMARKSPOSITION);
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_ToolTipPosition(ToolTipPositionConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ToolTipPositionConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if(GetStyle() & TBS_TOOLTIPS) {
			int position = SendMessage(TBM_SETTIPSIDE, TBTS_TOP, 0);
			SendMessage(TBM_SETTIPSIDE, position, 0);
			switch(position) {
				case TBTS_TOP:
				case TBTS_LEFT:
					properties.toolTipPosition = ttpTopOrLeft;
					break;
				case TBTS_BOTTOM:
				case TBTS_RIGHT:
					properties.toolTipPosition = ttpBottomOrRight;
					break;
			}
		} else {
			properties.toolTipPosition = ttpNone;
		}
	}

	*pValue = properties.toolTipPosition;
	return S_OK;
}

STDMETHODIMP TrackBar::put_ToolTipPosition(ToolTipPositionConstants newValue)
{
	PUTPROPPROLOG(DISPID_TRACKBARCTL_TOOLTIPPOSITION);
	if(properties.toolTipPosition != newValue) {
		properties.toolTipPosition = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.toolTipPosition) {
				case ttpNone:
					ModifyStyle(TBS_TOOLTIPS, 0);
					break;
				case ttpBottomOrRight:
					ModifyStyle(0, TBS_TOOLTIPS);
					SendMessage(TBM_SETTIPSIDE, ((GetStyle() & TBS_VERT) == TBS_VERT ? TBTS_RIGHT : TBTS_BOTTOM), 0);
					break;
				case ttpTopOrLeft:
					ModifyStyle(0, TBS_TOOLTIPS);
					SendMessage(TBM_SETTIPSIDE, ((GetStyle() & TBS_VERT) == TBS_VERT ? TBTS_LEFT : TBTS_TOP), 0);
					break;
			}
		}
		FireOnChanged(DISPID_TRACKBARCTL_TOOLTIPPOSITION);
	}
	return S_OK;
}

STDMETHODIMP TrackBar::get_Version(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	TCHAR pBuffer[50];
	ATLVERIFY(SUCCEEDED(StringCbPrintf(pBuffer, 50 * sizeof(TCHAR), TEXT("%i.%i.%i.%i"), VERSION_MAJOR, VERSION_MINOR, VERSION_REVISION1, VERSION_BUILD)));
	*pValue = CComBSTR(pBuffer);
	return S_OK;
}

STDMETHODIMP TrackBar::About(void)
{
	AboutDlg dlg;
	dlg.SetOwner(this);
	dlg.DoModal();
	return S_OK;
}

STDMETHODIMP TrackBar::FinishOLEDragDrop(void)
{
	if(dragDropStatus.pDropTargetHelper && dragDropStatus.drop_pDataObject) {
		dragDropStatus.pDropTargetHelper->Drop(dragDropStatus.drop_pDataObject, &dragDropStatus.drop_mousePosition, dragDropStatus.drop_effect);
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
		return S_OK;
	}
	// Can't perform requested operation - raise VB runtime error 17
	return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 17);
}

STDMETHODIMP TrackBar::LoadSettingsFromFile(BSTR file, VARIANT_BOOL* pSucceeded)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}
	*pSucceeded = VARIANT_FALSE;

	// open the file
	COLE2T converter(file);
	LPTSTR pFilePath = converter;
	CComPtr<IStream> pStream = NULL;
	HRESULT hr = SHCreateStreamOnFile(pFilePath, STGM_READ | STGM_SHARE_DENY_WRITE, &pStream);
	if(SUCCEEDED(hr) && pStream) {
		// read settings
		if(Load(pStream) == S_OK) {
			*pSucceeded = VARIANT_TRUE;
		}
	}
	return S_OK;
}

STDMETHODIMP TrackBar::Refresh(void)
{
	if(IsWindow()) {
		Invalidate();
		UpdateWindow();
	}
	return S_OK;
}

STDMETHODIMP TrackBar::SaveSettingsToFile(BSTR file, VARIANT_BOOL* pSucceeded)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}
	*pSucceeded = VARIANT_FALSE;

	// create the file
	COLE2T converter(file);
	LPTSTR pFilePath = converter;
	CComPtr<IStream> pStream = NULL;
	HRESULT hr = SHCreateStreamOnFile(pFilePath, STGM_CREATE | STGM_WRITE | STGM_SHARE_DENY_WRITE, &pStream);
	if(SUCCEEDED(hr) && pStream) {
		// write settings
		if(Save(pStream, FALSE) == S_OK) {
			if(FAILED(pStream->Commit(STGC_DEFAULT))) {
				return S_OK;
			}
			*pSucceeded = VARIANT_TRUE;
		}
	}
	return S_OK;
}


LRESULT TrackBar::OnChar(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;
	if(!(properties.disabledEvents & deKeyboardEvents)) {
		SHORT keyAscii = static_cast<SHORT>(wParam);
		if(SUCCEEDED(Raise_KeyPress(&keyAscii))) {
			// the client may have changed the key code (actually it can be changed to 0 only)
			wParam = keyAscii;
			if(wParam == 0) {
				wasHandled = TRUE;
			}
		}
	}
	return 0;
}

LRESULT TrackBar::OnContextMenu(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/)
{
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	if((mousePosition.x != -1) && (mousePosition.y != -1)) {
		ScreenToClient(&mousePosition);
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	Raise_ContextMenu(button, shift, mousePosition.x, mousePosition.y);
	return 0;
}

LRESULT TrackBar::OnCreate(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	// we want to receive WM_*BUTTONDBLCLK messages
	DWORD style = GetClassLong(*this, GCL_STYLE);
	style |= CS_DBLCLKS;
	SetClassLong(*this, GCL_STYLE, style);

	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(*this) {
		// this will keep the object alive if the client destroys the control window in an event handler
		AddRef();

		Raise_RecreatedControlWindow(HandleToLong(static_cast<HWND>(*this)));
	}
	return lr;
}

LRESULT TrackBar::OnDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	Raise_DestroyedControlWindow(HandleToLong(static_cast<HWND>(*this)));

	wasHandled = FALSE;
	return 0;
}

LRESULT TrackBar::OnEraseBkGnd(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	// prevents flickering
	return 1;
}

LRESULT TrackBar::OnKeyDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deKeyboardEvents)) {
		SHORT keyCode = static_cast<SHORT>(wParam);
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		if(SUCCEEDED(Raise_KeyDown(&keyCode, shift))) {
			// the client may have changed the key code
			wParam = keyCode;
			if(wParam == 0) {
				return 0;
			}
		}
	}
	return DefWindowProc(message, wParam, lParam);
}

LRESULT TrackBar::OnKillFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	LRESULT lr = CComControl<TrackBar>::OnKillFocus(message, wParam, lParam, wasHandled);
	flags.uiActivationPending = FALSE;
	return lr;
}

LRESULT TrackBar::OnKeyUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deKeyboardEvents)) {
		SHORT keyCode = static_cast<SHORT>(wParam);
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		if(SUCCEEDED(Raise_KeyUp(&keyCode, shift))) {
			// the client may have changed the key code
			wParam = keyCode;
			if(wParam == 0) {
				return 0;
			}
		}
	}
	return DefWindowProc(message, wParam, lParam);
}

LRESULT TrackBar::OnLButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	if(!properties.detectDoubleClicks) {
		return SendMessage(WM_LBUTTONDOWN, wParam, lParam);
	}

	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 1/*MouseButtonConstants.vbLeftButton*/;
		Raise_DblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT TrackBar::OnLButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT TrackBar::OnLButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT TrackBar::OnMButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	if(!properties.detectDoubleClicks) {
		return SendMessage(WM_MBUTTONDOWN, wParam, lParam);
	}

	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 4/*MouseButtonConstants.vbMiddleButton*/;
		Raise_MDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT TrackBar::OnMButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT TrackBar::OnMButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT TrackBar::OnMouseActivate(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	if(m_bInPlaceActive && !m_bUIActive) {
		flags.uiActivationPending = TRUE;
	} else {
		wasHandled = FALSE;
	}
	return MA_ACTIVATE;
}

LRESULT TrackBar::OnMouseHover(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	Raise_MouseHover(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT TrackBar::OnMouseLeave(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	Raise_MouseLeave(button, shift, mouseStatus.lastPosition.x, mouseStatus.lastPosition.y);

	return 0;
}

LRESULT TrackBar::OnMouseMove(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		Raise_MouseMove(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	} else if(!mouseStatus.enteredControl) {
		mouseStatus.EnterControl();
	}

	return 0;
}

LRESULT TrackBar::OnMouseWheel(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		if(message == WM_MOUSEHWHEEL) {
			// wParam and lParam seem to be 0
			WPARAM2BUTTONANDSHIFT(-1, button, shift);
			GetCursorPos(&mousePosition);
		} else {
			WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
		}
		ScreenToClient(&mousePosition);
		Raise_MouseWheel(button, shift, mousePosition.x, mousePosition.y, message == WM_MOUSEHWHEEL ? saHorizontal : saVertical, GET_WHEEL_DELTA_WPARAM(wParam));
	} else if(!mouseStatus.enteredControl) {
		mouseStatus.EnterControl();
	}

	return 0;
}

LRESULT TrackBar::OnNotify(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	switch(reinterpret_cast<LPNMHDR>(lParam)->code) {
		case TTN_GETDISPINFOA:
			return OnToolTipGetDispInfoNotificationA(message, wParam, lParam, wasHandled);
			break;

		case TTN_GETDISPINFOW:
			return OnToolTipGetDispInfoNotificationW(message, wParam, lParam, wasHandled);
			break;

		default:
			wasHandled = FALSE;
			break;
	}
	return 0;
}

LRESULT TrackBar::OnPaint(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	return DefWindowProc(message, wParam, lParam);
}

LRESULT TrackBar::OnRButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	if(!properties.detectDoubleClicks) {
		return SendMessage(WM_RBUTTONDOWN, wParam, lParam);
	}

	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 2/*MouseButtonConstants.vbRightButton*/;
		Raise_RDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT TrackBar::OnRButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT TrackBar::OnRButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT TrackBar::OnSetCursor(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	HCURSOR hCursor = NULL;
	BOOL setCursor = FALSE;

	// Are we really over the control?
	WTL::CRect clientArea;
	GetClientRect(&clientArea);
	ClientToScreen(&clientArea);
	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
	if(clientArea.PtInRect(mousePosition)) {
		// maybe the control is overlapped by a foreign window
		if(WindowFromPoint(mousePosition) == *this) {
			setCursor = TRUE;
		}
	}

	if(setCursor) {
		if(properties.mousePointer == mpCustom) {
			if(properties.mouseIcon.pPictureDisp) {
				CComQIPtr<IPicture, &IID_IPicture> pPicture(properties.mouseIcon.pPictureDisp);
				if(pPicture) {
					OLE_HANDLE h = NULL;
					pPicture->get_Handle(&h);
					hCursor = static_cast<HCURSOR>(LongToHandle(h));
				}
			}
		} else {
			hCursor = MousePointerConst2hCursor(properties.mousePointer);
		}

		if(hCursor) {
			SetCursor(hCursor);
			return TRUE;
		}
	}

	wasHandled = FALSE;
	return FALSE;
}

LRESULT TrackBar::OnSetFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	LRESULT lr = CComControl<TrackBar>::OnSetFocus(message, wParam, lParam, wasHandled);
	if(m_bInPlaceActive && !m_bUIActive && flags.uiActivationPending) {
		flags.uiActivationPending = FALSE;

		// now execute what usually would have been done on WM_MOUSEACTIVATE
		BOOL dummy = TRUE;
		CComControl<TrackBar>::OnMouseActivate(WM_MOUSEACTIVATE, 0, 0, dummy);
	}
	return lr;
}

LRESULT TrackBar::OnSetRedraw(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(lParam == 71216) {
		// the message was sent by ourselves
		lParam = 0;
		if(wParam) {
			// We're gonna activate redrawing - does the client allow this?
			if(properties.dontRedraw) {
				// no, so eat this message
				return 0;
			}
		}
	} else {
		// TODO: Should we really do this?
		properties.dontRedraw = !wParam;
	}

	return DefWindowProc(message, wParam, lParam);
}

LRESULT TrackBar::OnThemeChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	// Microsoft couldn't make it more difficult to detect whether we should use themes or not...
	flags.usingThemes = FALSE;
	if(CTheme::IsThemingSupported() && RunTimeHelper::IsCommCtrl6()) {
		HMODULE hThemeDLL = LoadLibrary(TEXT("uxtheme.dll"));
		if(hThemeDLL) {
			typedef BOOL WINAPI IsAppThemedFn();
			IsAppThemedFn* pfnIsAppThemed = static_cast<IsAppThemedFn*>(GetProcAddress(hThemeDLL, "IsAppThemed"));
			if(pfnIsAppThemed()) {
				flags.usingThemes = TRUE;
			}
			FreeLibrary(hThemeDLL);
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT TrackBar::OnTimer(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	switch(wParam) {
		case timers.ID_REDRAW:
			if(IsWindowVisible()) {
				KillTimer(timers.ID_REDRAW);
				SetRedraw(!properties.dontRedraw);
			} else {
				// wait... (this fixes visibility problems if another control displays a nag screen)
			}
			break;

		default:
			wasHandled = FALSE;
			break;
	}
	return 0;
}

LRESULT TrackBar::OnWindowPosChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	LPWINDOWPOS pDetails = reinterpret_cast<LPWINDOWPOS>(lParam);

	WTL::CRect windowRectangle = m_rcPos;
	/* Ugly hack: We depend on this message being sent without SWP_NOMOVE at least once, but this requirement
	              not always will be fulfilled. Fortunately pDetails seems to contain correct x and y values
	              even if SWP_NOMOVE is set.
	 */
	if(!(pDetails->flags & SWP_NOMOVE) || (windowRectangle.IsRectNull() && pDetails->x != 0 && pDetails->y != 0)) {
		windowRectangle.MoveToXY(pDetails->x, pDetails->y);
	}
	if(!(pDetails->flags & SWP_NOSIZE)) {
		windowRectangle.right = windowRectangle.left + pDetails->cx;
		windowRectangle.bottom = windowRectangle.top + pDetails->cy;
	}

	if(!(pDetails->flags & SWP_NOMOVE) || !(pDetails->flags & SWP_NOSIZE)) {
		Invalidate();
		ATLASSUME(m_spInPlaceSite);
		if(m_spInPlaceSite && !windowRectangle.EqualRect(&m_rcPos)) {
			m_spInPlaceSite->OnPosRectChange(&windowRectangle);
		}
		if(!(pDetails->flags & SWP_NOSIZE)) {
			/* Problem: When the control is resized, m_rcPos already contains the new rectangle, even before the
			 *          message is sent without SWP_NOSIZE. Therefore raise the event even if the rectangles are
			 *          equal. Raising the event too often is better than raising it too few.
			 */
			Raise_ResizedControlWindow();
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT TrackBar::OnXButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	if(!properties.detectDoubleClicks) {
		return SendMessage(WM_XBUTTONDOWN, wParam, lParam);
	}

	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
		switch(GET_XBUTTON_WPARAM(wParam)) {
			case XBUTTON1:
				button = embXButton1;
				break;
			case XBUTTON2:
				button = embXButton2;
				break;
		}
		Raise_XDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT TrackBar::OnXButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
	switch(GET_XBUTTON_WPARAM(wParam)) {
		case XBUTTON1:
			button = embXButton1;
			break;
		case XBUTTON2:
			button = embXButton2;
			break;
	}
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT TrackBar::OnXButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
	switch(GET_XBUTTON_WPARAM(wParam)) {
		case XBUTTON1:
			button = embXButton1;
			break;
		case XBUTTON2:
			button = embXButton2;
			break;
	}
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT TrackBar::OnReflectedCtlColorStatic(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(flags.usingThemes) {
		/* Unfortunately the native control uses double-buffering and sometimes redraws only a part of the control
		 * even if the invalid region contains more than this part. This is done by keeping a bitmap of the drawn
		 * control, updating only those parts in this bitmap that have changed, and copying those parts out of this
		 * bitmap that have to be drawn.
		 * Anything that we draw here, is drawn directly into the double-buffer bitmap. Since we have no chance to
		 * retrieve any information about what actually needs to be drawn, we should not draw anything here. Otherwise
		 * we might draw over parts that won't be redrawn in this paint cycle but that might be copied to the screen.
		 * The brush that we return here will be used to erase the parts that need to be redrawn. So the best thing
		 * we can do is create a pattern brush...
		 */
		/* NOTE: We have to do this on each WM_CTLCOLORSTATIC, because the parent window's appearance might
		         have changed. */
		WTL::CRect clientRectangle;
		::GetClientRect(reinterpret_cast<HWND>(lParam), &clientRectangle);

		CDC memoryDC;
		memoryDC.CreateCompatibleDC(reinterpret_cast<HDC>(wParam));
		CBitmap memoryBitmap;
		memoryBitmap.CreateCompatibleBitmap(reinterpret_cast<HDC>(wParam), clientRectangle.Width(), clientRectangle.Height());
		HBITMAP hPreviousBitmap = memoryDC.SelectBitmap(memoryBitmap);
		memoryDC.SetViewportOrg(-clientRectangle.left, -clientRectangle.top);

		CTheme themingEngine;
		themingEngine.OpenThemeData(*this, VSCLASS_TRACKBAR);
		if(themingEngine.IsThemeNull()) {
			if(!(properties.backColor & 0x80000000)) {
				memoryDC.SetDCBrushColor(OLECOLOR2COLORREF(properties.backColor));
				memoryDC.FillRect(&clientRectangle, static_cast<HBRUSH>(GetStockObject(DC_BRUSH)));
			} else {
				memoryDC.FillRect(&clientRectangle, GetSysColorBrush(properties.backColor & 0x0FFFFFFF));
			}
		} else {
			memoryDC.FillRect(&clientRectangle, GetSysColorBrush(COLOR_3DFACE));
		}
		DrawThemeParentBackground(reinterpret_cast<HWND>(lParam), memoryDC, NULL);

		memoryDC.SelectBitmap(hPreviousBitmap);
		if(!themedBackBrush.IsNull()) {
			themedBackBrush.DeleteObject();
		}
		themedBackBrush.CreatePatternBrush(memoryBitmap);
		return reinterpret_cast<LRESULT>(static_cast<HBRUSH>(themedBackBrush));
	}
	if(!(properties.backColor & 0x80000000)) {
		SetDCBrushColor(reinterpret_cast<HDC>(wParam), OLECOLOR2COLORREF(properties.backColor));
		return reinterpret_cast<LRESULT>(static_cast<HBRUSH>(GetStockObject(DC_BRUSH)));
	} else {
		return reinterpret_cast<LRESULT>(GetSysColorBrush(properties.backColor & 0x0FFFFFFF));
	}
}

LRESULT TrackBar::OnReflectedNotify(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	switch(reinterpret_cast<LPNMHDR>(lParam)->code) {
		case NM_CUSTOMDRAW:
			return OnCustomDrawNotification(message, wParam, lParam, wasHandled);
			break;
		default:
			wasHandled = FALSE;
			break;
	}
	return 0;
}

LRESULT TrackBar::OnReflectedScroll(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	LONG currentPosition = SendMessage(TBM_GETPOS, 0, 0);
	if(currentPosition != properties.currentPosition) {
		PositionChangeTypeConstants changeType = pctJumpToPosition;
		switch(LOWORD(wParam)) {
			case TB_THUMBTRACK:
				// NOTE: HIWORD(wParam) is the new position, but it's only 16 bit wide.
				changeType = pctBeginTrack;
				break;
			case TB_ENDTRACK:
				changeType = pctEndTrack;
				break;
			case TB_THUMBPOSITION:
				changeType = pctTracking;
				break;
			case TB_LINEUP:
				changeType = pctSmallStepUp;
				break;
			case TB_LINEDOWN:
				changeType = pctSmallStepDown;
				break;
			case TB_PAGEUP:
				changeType = pctLargeStepUp;
				break;
			case TB_PAGEDOWN:
				changeType = pctLargeStepDown;
				break;
			case TB_TOP:
			case TB_BOTTOM:
				changeType = pctJumpToPosition;
				break;
			default:
				ATLASSERT(FALSE);
				break;
		}
		properties.currentPosition = currentPosition;
		Raise_PositionChanged(changeType, currentPosition);
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT TrackBar::OnSetPos(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_TRACKBARCTL_CURRENTPOSITION) == S_FALSE) {
		return 0;
	}

	VARIANT_BOOL cancel = VARIANT_FALSE;
	if(lParam != properties.currentPosition) {
		if((flags.silentPositionChanges == 0) && ((GetStyle() & TBS_NOTIFYBEFOREMOVE) == TBS_NOTIFYBEFOREMOVE)) {
			Raise_PositionChanging(pctJumpToPosition, lParam, &cancel);
		}
	}
	if(cancel != VARIANT_FALSE) {
		return 0;
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(flags.silentPositionChanges == 0) {
		LONG buffer = SendMessage(TBM_GETPOS, 0, 0);
		if(buffer != properties.currentPosition) {
			properties.currentPosition = buffer;
			Raise_PositionChanged(pctJumpToPosition, properties.currentPosition);
		}
	}
	SetDirty(TRUE);
	FireOnChanged(DISPID_TRACKBARCTL_CURRENTPOSITION);
	SendOnDataChange();
	return lr;
}

LRESULT TrackBar::OnSetTicFreq(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	properties.autoTickFrequency = static_cast<SHORT>(wParam);
	return DefWindowProc(message, wParam, lParam);
}


LRESULT TrackBar::OnThumbPosChangingNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	NMTRBTHUMBPOSCHANGING* pDetails = reinterpret_cast<NMTRBTHUMBPOSCHANGING*>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMTRBTHUMBPOSCHANGING);
	if(!pDetails) {
		return 0;
	}

	if(static_cast<LONG>(pDetails->dwPos) != properties.currentPosition) {
		PositionChangeTypeConstants changeType = pctJumpToPosition;
		switch(pDetails->nReason) {
			case TB_THUMBTRACK:
				changeType = pctBeginTrack;
				break;
			case TB_ENDTRACK:
				changeType = pctEndTrack;
				break;
			case TB_THUMBPOSITION:
				changeType = pctTracking;
				break;
			case TB_LINEUP:
				changeType = pctSmallStepUp;
				break;
			case TB_LINEDOWN:
				changeType = pctSmallStepDown;
				break;
			case TB_PAGEUP:
				changeType = pctLargeStepUp;
				break;
			case TB_PAGEDOWN:
				changeType = pctLargeStepDown;
				break;
			case TB_TOP:
			case TB_BOTTOM:
				changeType = pctJumpToPosition;
				break;
			default:
				ATLASSERT(FALSE);
				break;
		}

		VARIANT_BOOL cancel = VARIANT_FALSE;
		Raise_PositionChanging(changeType, pDetails->dwPos, &cancel);
		return VARIANTBOOL2BOOL(cancel);
	}
	return 0;
}


LRESULT TrackBar::OnCustomDrawNotification(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LPNMCUSTOMDRAW pDetails = reinterpret_cast<LPNMCUSTOMDRAW>(lParam);
	ATLASSERT_POINTER(pDetails, NMCUSTOMDRAW);
	if(!pDetails) {
		return DefWindowProc(message, wParam, lParam);
	}

	ATLASSERT((pDetails->dwDrawStage & CDDS_PREERASE) != CDDS_PREERASE);
	ATLASSERT((pDetails->dwDrawStage & CDDS_POSTERASE) != CDDS_POSTERASE);
	ATLASSERT((pDetails->dwDrawStage & CDDS_SUBITEM) != CDDS_SUBITEM);
	#ifdef DEBUG
		if((pDetails->dwDrawStage == CDDS_ITEMPREPAINT) || (pDetails->dwDrawStage == CDDS_ITEMPOSTPAINT)) {
			ATLASSERT((pDetails->dwItemSpec >= 0) && (pDetails->dwItemSpec <= TBCD_CHANNEL));
		}
	#endif
	ATLASSERT((pDetails->uItemState & CDIS_GRAYED) == 0);
	ATLASSERT((pDetails->uItemState & CDIS_DISABLED) == 0);
	ATLASSERT((pDetails->uItemState & CDIS_CHECKED) == 0);
	ATLASSERT((pDetails->uItemState & CDIS_DEFAULT) == 0);
	ATLASSERT((pDetails->uItemState & CDIS_HOT) == 0);
	ATLASSERT((pDetails->uItemState & CDIS_MARKED) == 0);
	ATLASSERT((pDetails->uItemState & CDIS_INDETERMINATE) == 0);
	ATLASSERT((pDetails->uItemState & CDIS_SHOWKEYBOARDCUES) == 0);
	ATLASSERT((pDetails->uItemState & CDIS_NEARHOT) == 0);
	ATLASSERT((pDetails->uItemState & CDIS_OTHERSIDEHOT) == 0);
	ATLASSERT((pDetails->uItemState & CDIS_DROPHILITED) == 0);

	CustomDrawReturnValuesConstants returnValue = static_cast<CustomDrawReturnValuesConstants>(DefWindowProc(message, wParam, lParam));

	if(!(properties.disabledEvents & deCustomDraw)) {
		CustomDrawControlPartConstants controlPart = cdcpEntireControl;
		RECTANGLE rectangle = *reinterpret_cast<RECTANGLE*>(&pDetails->rc);
		if(pDetails->dwDrawStage == CDDS_PREPAINT) {
			ZeroMemory(&rectangle, sizeof(RECTANGLE));
		} else {
			if(pDetails->dwItemSpec == TBCD_TICS) {
				ZeroMemory(&rectangle, sizeof(RECTANGLE));
			}
			controlPart = static_cast<CustomDrawControlPartConstants>(pDetails->dwItemSpec);
		}
		Raise_CustomDraw(controlPart, static_cast<CustomDrawStageConstants>(pDetails->dwDrawStage), static_cast<CustomDrawControlStateConstants>(pDetails->uItemState), HandleToLong(pDetails->hdc), &rectangle, &returnValue);
	}

	return static_cast<LRESULT>(returnValue);
}

LRESULT TrackBar::OnToolTipGetDispInfoNotificationA(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	if(properties.disabledEvents & deGetInfoTipText) {
		wasHandled = FALSE;
	} else {
		LPNMTTDISPINFOA pDetails = reinterpret_cast<LPNMTTDISPINFOA>(lParam);
		ZeroMemory(pToolTipBuffer, MAX_PATH * sizeof(CHAR));

		CComBSTR text;
		if(lstrlenA(pDetails->lpszText) == 0) {
			LPTSTR pBuffer = ConvertIntegerToString(SendMessage(TBM_GETPOS, 0, 0));
			text = pBuffer;
			SECUREFREE(pBuffer);
		} else {
			text = pDetails->lpszText;
		}

		VARIANT_BOOL abortToolTip = VARIANT_FALSE;
		Raise_GetInfoTipText(MAX_PATH, &text, &abortToolTip);

		if(abortToolTip == VARIANT_FALSE) {
			if(text.Length() > 0) {
				lstrcpynA(reinterpret_cast<LPSTR>(pToolTipBuffer), CW2A(text), MAX_PATH);
			}
			pDetails->lpszText = reinterpret_cast<LPSTR>(pToolTipBuffer);
		}
	}

	return 0;
}

LRESULT TrackBar::OnToolTipGetDispInfoNotificationW(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	if(properties.disabledEvents & deGetInfoTipText) {
		wasHandled = FALSE;
	} else {
		LPNMTTDISPINFOW pDetails = reinterpret_cast<LPNMTTDISPINFOW>(lParam);
		ZeroMemory(pToolTipBuffer, MAX_PATH * sizeof(WCHAR));

		CComBSTR text;
		if(lstrlenW(pDetails->lpszText) == 0) {
			LPTSTR pBuffer = ConvertIntegerToString(SendMessage(TBM_GETPOS, 0, 0));
			text = pBuffer;
			SECUREFREE(pBuffer);
		} else {
			text = pDetails->lpszText;
		}

		VARIANT_BOOL abortToolTip = VARIANT_FALSE;
		Raise_GetInfoTipText(MAX_PATH, &text, &abortToolTip);

		if(abortToolTip == VARIANT_FALSE) {
			if(text.Length() > 0) {
				#ifdef UNICODE
					lstrcpynW(reinterpret_cast<LPWSTR>(pToolTipBuffer), CW2CW(text), MAX_PATH);
				#else
					//OSVERSIONINFO ovi = {sizeof(OSVERSIONINFO)};
					//if(GetVersionEx(&ovi) && (ovi.dwPlatformId == VER_PLATFORM_WIN32_WINDOWS)) {
					//	/* HACK: On Windows 9x strange things happen. Instead of the ANSI notification, the Unicode
					//	         notification is sent. This isn't really a problem, but lstrcpynW() seems to be
					//	         broken. */
					//	LPCWSTR pToolTip = CW2CW(text);
					//	// as lstrcpynW() seems to be broken, copy the string manually
					//	LPWSTR tmp = reinterpret_cast<LPWSTR>(pToolTipBuffer);
					//	for(int i = 0; (*pToolTip != 0) && (i < MAX_PATH); ++i, ++pToolTip, ++tmp) {
					//		*tmp = *pToolTip;
					//	}
					//	*tmp = 0;
					//} else {
						lstrcpynW(reinterpret_cast<LPWSTR>(pToolTipBuffer), CW2CW(text), MAX_PATH);
					//}
				#endif
			}
			pDetails->lpszText = reinterpret_cast<LPWSTR>(pToolTipBuffer);
		}
	}

	return 0;
}


inline HRESULT TrackBar::Raise_Click(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_Click(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT TrackBar::Raise_ContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		if((x == -1) && (y == -1)) {
			// the event was caused by the keyboard
			if(properties.processContextMenuKeys) {
				// propose the middle of the control's client rectangle as the menu's position
				WTL::CRect clientRectangle;
				GetClientRect(&clientRectangle);
				WTL::CPoint centerPoint = clientRectangle.CenterPoint();
				x = centerPoint.x;
				y = centerPoint.y;
			} else {
				return S_OK;
			}
		}

		return Fire_ContextMenu(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT TrackBar::Raise_CustomDraw(CustomDrawControlPartConstants controlPart, CustomDrawStageConstants drawStage, CustomDrawControlStateConstants controlState, LONG hDC, RECTANGLE* pDrawingRectangle, CustomDrawReturnValuesConstants* pFurtherProcessing)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_CustomDraw(controlPart, drawStage, controlState, hDC, pDrawingRectangle, pFurtherProcessing);
	//}
	//return S_OK;
}

inline HRESULT TrackBar::Raise_DblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_DblClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT TrackBar::Raise_DestroyedControlWindow(LONG hWnd)
{
	KillTimer(timers.ID_REDRAW);
	if(properties.registerForOLEDragDrop) {
		ATLVERIFY(RevokeDragDrop(*this) == S_OK);
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_DestroyedControlWindow(hWnd);
	//}
	//return S_OK;
}

inline HRESULT TrackBar::Raise_GetInfoTipText(LONG maxInfoTipLength, BSTR* pInfoTipText, VARIANT_BOOL* pAbortToolTip)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_GetInfoTipText(maxInfoTipLength, pInfoTipText, pAbortToolTip);
	//}
	//return S_OK;
}

inline HRESULT TrackBar::Raise_KeyDown(SHORT* pKeyCode, SHORT shift)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyDown(pKeyCode, shift);
	//}
	//return S_OK;
}

inline HRESULT TrackBar::Raise_KeyPress(SHORT* pKeyAscii)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyPress(pKeyAscii);
	//}
	//return S_OK;
}

inline HRESULT TrackBar::Raise_KeyUp(SHORT* pKeyCode, SHORT shift)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyUp(pKeyCode, shift);
	//}
	//return S_OK;
}

inline HRESULT TrackBar::Raise_MClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_MClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT TrackBar::Raise_MDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_MDblClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT TrackBar::Raise_MouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		if(!mouseStatus.enteredControl) {
			Raise_MouseEnter(button, shift, x, y);
		}
		if(!mouseStatus.hoveredControl) {
			TRACKMOUSEEVENT trackingOptions = {0};
			trackingOptions.cbSize = sizeof(trackingOptions);
			trackingOptions.hwndTrack = *this;
			trackingOptions.dwFlags = TME_HOVER | TME_CANCEL;
			TrackMouseEvent(&trackingOptions);

			Raise_MouseHover(button, shift, x, y);
		}
		mouseStatus.StoreClickCandidate(button);
		SetCapture();

		return Fire_MouseDown(button, shift, x, y);
	} else {
		if(!mouseStatus.enteredControl) {
			mouseStatus.EnterControl();
		}
		if(!mouseStatus.hoveredControl) {
			mouseStatus.HoverControl();
		}
		mouseStatus.StoreClickCandidate(button);
		SetCapture();
		return S_OK;
	}
}

inline HRESULT TrackBar::Raise_MouseEnter(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	TRACKMOUSEEVENT trackingOptions = {0};
	trackingOptions.cbSize = sizeof(trackingOptions);
	trackingOptions.hwndTrack = *this;
	trackingOptions.dwHoverTime = (properties.hoverTime == -1 ? HOVER_DEFAULT : properties.hoverTime);
	trackingOptions.dwFlags = TME_HOVER | TME_LEAVE;
	TrackMouseEvent(&trackingOptions);

	mouseStatus.EnterControl();

	//if(m_nFreezeEvents == 0) {
		return Fire_MouseEnter(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT TrackBar::Raise_MouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	mouseStatus.HoverControl();

	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		return Fire_MouseHover(button, shift, x, y);
	}
	return S_OK;
}

inline HRESULT TrackBar::Raise_MouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	mouseStatus.LeaveControl();

	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		return Fire_MouseLeave(button, shift, x, y);
	}
	return S_OK;
}

inline HRESULT TrackBar::Raise_MouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(!mouseStatus.enteredControl) {
		Raise_MouseEnter(button, shift, x, y);
	}
	mouseStatus.lastPosition.x = x;
	mouseStatus.lastPosition.y = y;

	//if(m_nFreezeEvents == 0) {
		return Fire_MouseMove(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT TrackBar::Raise_MouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	HRESULT hr = S_OK;
	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		hr = Fire_MouseUp(button, shift, x, y);
	}

	if(mouseStatus.IsClickCandidate(button)) {
		/* Watch for clicks.
		   Are we still within the control's client area? */
		BOOL hasLeftControl = FALSE;
		DWORD position = GetMessagePos();
		POINT cursorPosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		RECT clientArea = {0};
		GetClientRect(&clientArea);
		ClientToScreen(&clientArea);
		if(PtInRect(&clientArea, cursorPosition)) {
			// maybe the control is overlapped by a foreign window
			if(WindowFromPoint(cursorPosition) != *this) {
				hasLeftControl = TRUE;
			}
		} else {
			hasLeftControl = TRUE;
		}
		if(GetCapture() == *this) {
			ReleaseCapture();
		}

		if(!hasLeftControl) {
			// we don't have left the control, so raise the click event
			switch(button) {
				case 1/*MouseButtonConstants.vbLeftButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_Click(button, shift, x, y);
					}
					break;
				case 2/*MouseButtonConstants.vbRightButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_RClick(button, shift, x, y);
					}
					break;
				case 4/*MouseButtonConstants.vbMiddleButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_MClick(button, shift, x, y);
					}
					break;
				case embXButton1:
				case embXButton2:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_XClick(button, shift, x, y);
					}
					break;
			}
		}

		mouseStatus.RemoveClickCandidate(button);
	}

	return hr;
}

inline HRESULT TrackBar::Raise_MouseWheel(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, ScrollAxisConstants scrollAxis, SHORT wheelDelta)
{
	if(!mouseStatus.enteredControl) {
		Raise_MouseEnter(button, shift, x, y);
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_MouseWheel(button, shift, x, y, scrollAxis, wheelDelta);
	//}
	//return S_OK;
}

inline HRESULT TrackBar::Raise_OLEDragDrop(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition)
{
	// NOTE: pData can be NULL

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
		SHORT button = 0;
		SHORT shift = 0;
		OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

		if(pData) {
			/* Actually we wouldn't need the next line, because the data object passed to this method should
				 always be the same as the data object that was passed to Raise_OLEDragEnter. */
			dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
		} else {
			dragDropStatus.pActiveDataObject = NULL;
		}
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragDrop(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), button, shift, mousePosition.x, mousePosition.y);
		}
	//}

	dragDropStatus.pActiveDataObject = NULL;
	dragDropStatus.OLEDragLeaveOrDrop();
	Invalidate();

	return hr;
}

inline HRESULT TrackBar::Raise_OLEDragEnter(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition)
{
	// NOTE: pData can be NULL

	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	dragDropStatus.OLEDragEnter();

	if(pData) {
		dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
	} else {
		dragDropStatus.pActiveDataObject = NULL;
	}
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			return Fire_OLEDragEnter(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), button, shift, mousePosition.x, mousePosition.y);
		}
	//}
	return S_OK;
}

inline HRESULT TrackBar::Raise_OLEDragLeave(void)
{
	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);

		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragLeave(dragDropStatus.pActiveDataObject, button, shift, dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y);
		}
	//}

	dragDropStatus.pActiveDataObject = NULL;
	dragDropStatus.OLEDragLeaveOrDrop();
	Invalidate();

	return hr;
}

inline HRESULT TrackBar::Raise_OLEDragMouseMove(DWORD* pEffect, DWORD keyState, POINTL mousePosition)
{
	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	dragDropStatus.lastMousePosition = mousePosition;
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			return Fire_OLEDragMouseMove(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), button, shift, mousePosition.x, mousePosition.y);
		}
	//}
	return S_OK;
}

inline HRESULT TrackBar::Raise_PositionChanged(PositionChangeTypeConstants changeType, LONG newPosition)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_PositionChanged(changeType, newPosition);
	//}
	//return S_OK;
}

inline HRESULT TrackBar::Raise_PositionChanging(PositionChangeTypeConstants changeType, LONG newPosition, VARIANT_BOOL* pCancelChange)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_PositionChanging(changeType, newPosition, pCancelChange);
	//}
	//return S_OK;
}

inline HRESULT TrackBar::Raise_RClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT TrackBar::Raise_RDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RDblClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT TrackBar::Raise_RecreatedControlWindow(LONG hWnd)
{
	// configure the control
	SendConfigurationMessages();

	if(properties.registerForOLEDragDrop) {
		ATLVERIFY(RegisterDragDrop(*this, static_cast<IDropTarget*>(this)) == S_OK);
	}

	if(properties.dontRedraw) {
		SetTimer(timers.ID_REDRAW, timers.INT_REDRAW);
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_RecreatedControlWindow(hWnd);
	//}
	//return S_OK;
}

inline HRESULT TrackBar::Raise_ResizedControlWindow(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ResizedControlWindow();
	//}
	//return S_OK;
}

inline HRESULT TrackBar::Raise_XClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_XClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT TrackBar::Raise_XDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_XDblClick(button, shift, x, y);
	//}
	//return S_OK;
}


void TrackBar::EnterSilentPositionChangesSection(void)
{
	++flags.silentPositionChanges;
}

void TrackBar::LeaveSilentPositionChangesSection(void)
{
	--flags.silentPositionChanges;
}


void TrackBar::RecreateControlWindow(void)
{
	if(m_bInPlaceActive) {
		BOOL isUIActive = m_bUIActive;
		InPlaceDeactivate();
		ATLASSERT(m_hWnd == NULL);
		InPlaceActivate((isUIActive ? OLEIVERB_UIACTIVATE : OLEIVERB_INPLACEACTIVATE));
	}
}

DWORD TrackBar::GetExStyleBits(void)
{
	DWORD extendedStyle = WS_EX_LEFT | WS_EX_LTRREADING;
	switch(properties.appearance) {
		case a3D:
			extendedStyle |= WS_EX_CLIENTEDGE;
			break;
		case a3DLight:
			extendedStyle |= WS_EX_STATICEDGE;
			break;
	}
	if(properties.rightToLeftLayout) {
		extendedStyle |= WS_EX_LAYOUTRTL;
	}
	return extendedStyle;
}

DWORD TrackBar::GetStyleBits(void)
{
	DWORD style = WS_CHILDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | WS_VISIBLE;
	switch(properties.borderStyle) {
		case bsFixedSingle:
			style |= WS_BORDER;
			break;
	}
	if(!properties.enabled) {
		style |= WS_DISABLED;
	}

	if(IsComctl32Version610OrNewer()) {
		style |= TBS_NOTIFYBEFOREMOVE;
		switch(properties.backgroundDrawMode) {
			case bdmByParent:
				style |= TBS_TRANSPARENTBKGND;
				break;
		}
	}
	if(properties.autoTickMarks) {
		style |= TBS_AUTOTICKS;
	}
	if(properties.downIsLeft/* && IsComctl32Version581OrNewer()*/) {
		style |= TBS_DOWNISLEFT;
	}
	switch(properties.orientation) {
		case oHorizontal:
			style |= TBS_HORZ;
			break;
		case oVertical:
			style |= TBS_VERT;
			break;
	}
	if(properties.reversed/* && IsComctl32Version580OrNewer()*/) {
		style |= TBS_REVERSED;
	}
	switch(properties.selectionType) {
		case stRangeSelection:
			style |= TBS_ENABLESELRANGE;
			break;
	}
	if(!properties.showSlider) {
		style |= TBS_NOTHUMB;
	}
	if(properties.sliderLength != -1) {
		style |= TBS_FIXEDLENGTH;
	}
	switch(properties.tickMarksPosition) {
		case tmpNone:
			style |= TBS_NOTICKS;
			break;
		case tmpBottomOrRight:
			style |= TBS_BOTTOM | TBS_RIGHT;
			break;
		case tmpTopOrLeft:
			style |= TBS_TOP | TBS_LEFT;
			break;
		case tmpBoth:
			style |= TBS_BOTH;
			break;
	}
	switch(properties.toolTipPosition) {
		case ttpBottomOrRight:
		case ttpTopOrLeft:
			style |= TBS_TOOLTIPS;
			break;
	}
	return style;
}

void TrackBar::SendConfigurationMessages(void)
{
	SendMessage(TBM_SETRANGEMIN, FALSE, properties.minimum);
	SendMessage(TBM_SETRANGEMAX, FALSE, properties.maximum);
	SendMessage(TBM_SETSELSTART, FALSE, properties.rangeSelectionStart);
	SendMessage(TBM_SETSELEND, FALSE, properties.rangeSelectionEnd);
	SendMessage(TBM_SETPOS, TRUE, properties.currentPosition);
	SendMessage(TBM_SETTICFREQ, properties.autoTickFrequency, 0);
	SendMessage(TBM_SETPAGESIZE, 0, properties.largeStepWidth);
	SendMessage(TBM_SETLINESIZE, 0, properties.smallStepWidth);
	if(properties.sliderLength != -1) {
		SendMessage(TBM_SETTHUMBLENGTH, properties.sliderLength, 0);
	}

	switch(properties.toolTipPosition) {
		case ttpBottomOrRight:
			SendMessage(TBM_SETTIPSIDE, ((GetStyle() & TBS_VERT) == TBS_VERT ? TBTS_RIGHT : TBTS_BOTTOM), 0);
			break;
		case ttpTopOrLeft:
			SendMessage(TBM_SETTIPSIDE, ((GetStyle() & TBS_VERT) == TBS_VERT ? TBTS_LEFT : TBTS_TOP), 0);
			break;
	}
}

HCURSOR TrackBar::MousePointerConst2hCursor(MousePointerConstants mousePointer)
{
	WORD flag = 0;
	switch(mousePointer) {
		case mpArrow:
			flag = OCR_NORMAL;
			break;
		case mpCross:
			flag = OCR_CROSS;
			break;
		case mpIBeam:
			flag = OCR_IBEAM;
			break;
		case mpIcon:
			flag = OCR_ICOCUR;
			break;
		case mpSize:
			flag = OCR_SIZEALL;     // OCR_SIZE is obsolete
			break;
		case mpSizeNESW:
			flag = OCR_SIZENESW;
			break;
		case mpSizeNS:
			flag = OCR_SIZENS;
			break;
		case mpSizeNWSE:
			flag = OCR_SIZENWSE;
			break;
		case mpSizeEW:
			flag = OCR_SIZEWE;
			break;
		case mpUpArrow:
			flag = OCR_UP;
			break;
		case mpHourglass:
			flag = OCR_WAIT;
			break;
		case mpNoDrop:
			flag = OCR_NO;
			break;
		case mpArrowHourglass:
			flag = OCR_APPSTARTING;
			break;
		case mpArrowQuestion:
			flag = 32651;
			break;
		case mpSizeAll:
			flag = OCR_SIZEALL;
			break;
		case mpHand:
			flag = OCR_HAND;
			break;
		case mpInsertMedia:
			flag = 32663;
			break;
		case mpScrollAll:
			flag = 32654;
			break;
		case mpScrollN:
			flag = 32655;
			break;
		case mpScrollNE:
			flag = 32660;
			break;
		case mpScrollE:
			flag = 32658;
			break;
		case mpScrollSE:
			flag = 32662;
			break;
		case mpScrollS:
			flag = 32656;
			break;
		case mpScrollSW:
			flag = 32661;
			break;
		case mpScrollW:
			flag = 32657;
			break;
		case mpScrollNW:
			flag = 32659;
			break;
		case mpScrollNS:
			flag = 32652;
			break;
		case mpScrollEW:
			flag = 32653;
			break;
		default:
			return NULL;
	}

	return static_cast<HCURSOR>(LoadImage(0, MAKEINTRESOURCE(flag), IMAGE_CURSOR, 0, 0, LR_DEFAULTCOLOR | LR_DEFAULTSIZE | LR_SHARED));
}

BOOL TrackBar::IsInDesignMode(void)
{
	BOOL b = TRUE;
	GetAmbientUserMode(b);
	return !b;
}


BOOL TrackBar::IsComctl32Version610OrNewer(void)
{
	DWORD major = 0;
	DWORD minor = 0;
	HRESULT hr = ATL::AtlGetCommCtrlVersion(&major, &minor);
	if(SUCCEEDED(hr)) {
		return (((major == 6) && (minor >= 10)) || (major > 6));
	}
	return FALSE;
}