VERSION 5.00
Object = "{956B5A46-C53F-45A7-AF0E-EC2E1CC9B567}#1.7#0"; "TrackBarCtlU.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "TrackBar 1.7 - Transparent Background Sample"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6900
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   137
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin TrackBarCtlLibUCtl.TrackBar TrackBarU 
      Height          =   735
      Left            =   1200
      TabIndex        =   0
      Top             =   600
      Width           =   3495
      _cx             =   6165
      _cy             =   1296
      Appearance      =   0
      AutoTickFrequency=   1
      AutoTickMarks   =   -1  'True
      BackColor       =   -2147483633
      BackgroundDrawMode=   0
      BorderStyle     =   0
      CurrentPosition =   0
      DetectDoubleClicks=   -1  'True
      DisabledEvents  =   779
      DontRedraw      =   0   'False
      DownIsLeft      =   -1  'True
      Enabled         =   -1  'True
      HoverTime       =   -1
      LargeStepWidth  =   -1
      Maximum         =   10
      Minimum         =   0
      MousePointer    =   0
      Orientation     =   0
      ProcessContextMenuKeys=   -1  'True
      RangeSelectionEnd=   0
      RangeSelectionStart=   0
      RegisterForOLEDragDrop=   0   'False
      Reversed        =   0   'False
      RightToLeftLayout=   0   'False
      SelectionType   =   0
      ShowSlider      =   -1  'True
      SliderLength    =   -1
      SmallStepWidth  =   1
      SupportOLEDragImages=   -1  'True
      TickMarksPosition=   1
      ToolTipPosition =   2
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Implements ISubclassedWindow


  Private Type POINTAPI
    x As Long
    y As Long
  End Type

  Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
  End Type


  Private hBackBrush As Long


  Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
  Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
  Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
  Private Declare Function CreatePatternBrush Lib "gdi32.dll" (ByVal hBitmap As Long) As Long
  Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
  Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
  Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
  Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long


Private Function ISubclassedWindow_HandleMessage(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal eSubclassID As EnumSubclassID, bCallDefProc As Boolean) As Long
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case eSubclassID
    Case EnumSubclassID.escidFrmMain
      lRet = HandleMessage_Form(hWnd, uMsg, wParam, lParam, bCallDefProc)
    Case Else
      Debug.Print "frmMain.ISubclassedWindow_HandleMessage: Unknown Subclassing ID " & CStr(eSubclassID)
  End Select

StdHandler_Ende:
  ISubclassedWindow_HandleMessage = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in frmMain.ISubclassedWindow_HandleMessage (SubclassID=" & CStr(eSubclassID) & ": ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function

Private Function HandleMessage_Form(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bCallDefProc As Boolean) As Long
  Const SRCCOPY = &HCC0020
  Const WM_CTLCOLORSTATIC = &H138
  Dim hDialogBkDC As Long
  Dim hPreviousBitmap1 As Long
  Dim hPreviousBitmap2 As Long
  Dim hTrackBarBackgroundTexture As Long
  Dim hTrackBarCtlBkDC As Long
  Dim pt As POINTAPI
  Dim windowRectangle As RECT
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case uMsg
    Case WM_CTLCOLORSTATIC
      ' this is required to make the background transparent
      If hBackBrush Then
        DeleteObject hBackBrush
        hBackBrush = 0
      End If
      GetWindowRect lParam, windowRectangle
      pt.x = windowRectangle.Left
      pt.y = windowRectangle.Top
      ScreenToClient Me.hWnd, pt
      windowRectangle.Left = pt.x
      windowRectangle.Top = pt.y
      pt.x = windowRectangle.Right
      pt.y = windowRectangle.Bottom
      ScreenToClient Me.hWnd, pt
      windowRectangle.Right = pt.x
      windowRectangle.Bottom = pt.y

      hTrackBarCtlBkDC = CreateCompatibleDC(wParam)
      hTrackBarBackgroundTexture = CreateCompatibleBitmap(wParam, windowRectangle.Right - windowRectangle.Left, windowRectangle.Bottom - windowRectangle.Top)
      hPreviousBitmap1 = SelectObject(hTrackBarCtlBkDC, hTrackBarBackgroundTexture)

      hDialogBkDC = CreateCompatibleDC(wParam)
      hPreviousBitmap2 = SelectObject(hDialogBkDC, Me.Picture.Handle)

      BitBlt hTrackBarCtlBkDC, 0, 0, windowRectangle.Right - windowRectangle.Left, windowRectangle.Bottom - windowRectangle.Top, hDialogBkDC, windowRectangle.Left, windowRectangle.Top, SRCCOPY

      SelectObject hDialogBkDC, hPreviousBitmap2
      SelectObject hTrackBarCtlBkDC, hPreviousBitmap1

      hBackBrush = CreatePatternBrush(hTrackBarBackgroundTexture)

      DeleteObject hTrackBarBackgroundTexture
      DeleteDC hTrackBarCtlBkDC
      DeleteDC hDialogBkDC

      lRet = hBackBrush
      bCallDefProc = False
  End Select

StdHandler_Ende:
  HandleMessage_Form = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in frmMain.HandleMessage_Form: ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function


Private Sub cmdAbout_Click()
  TrackBarU.About
End Sub

Private Sub Form_Initialize()
  InitCommonControls
End Sub

Private Sub Form_Load()
  #If UseSubClassing Then
    If Not SubclassWindow(Me.hWnd, Me, EnumSubclassID.escidFrmMain) Then
      Debug.Print "Subclassing failed!"
    End If
  #End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not UnSubclassWindow(Me.hWnd, EnumSubclassID.escidFrmMain) Then
    Debug.Print "UnSubclassing failed!"
  End If

  If hBackBrush Then
    DeleteObject hBackBrush
  End If
End Sub
