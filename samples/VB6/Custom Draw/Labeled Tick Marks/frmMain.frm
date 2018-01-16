VERSION 5.00
Object = "{956B5A46-C53F-45A7-AF0E-EC2E1CC9B567}#1.7#0"; "TrackBarCtlU.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "TrackBar 1.7 - Labeled Tick Marks Sample"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5595
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
   ScaleHeight     =   123
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   373
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
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
      Left            =   1950
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
   End
   Begin TrackBarCtlLibUCtl.TrackBar TrackBarU 
      Height          =   735
      Left            =   450
      TabIndex        =   0
      Top             =   240
      Width           =   4695
      _cx             =   8281
      _cy             =   1296
      Appearance      =   0
      AutoTickFrequency=   1
      AutoTickMarks   =   -1  'True
      BackColor       =   -2147483633
      BackgroundDrawMode=   0
      BorderStyle     =   0
      CurrentPosition =   0
      DisabledEvents  =   523
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
      RangeSelectionEnd=   8
      RangeSelectionStart=   2
      RegisterForOLEDragDrop=   -1  'True
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
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Implements IHook


  Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
  End Type


  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)
  Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
  Private Declare Function GetStockObject Lib "gdi32.dll" (ByVal nIndex As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
  Private Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal nBkMode As Long) As Long


Private Function IHook_KeyboardProcAfter(ByVal hookCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  '
End Function

Private Function IHook_KeyboardProcBefore(ByVal hookCode As Long, ByVal wParam As Long, ByVal lParam As Long, eatIt As Boolean) As Long
  Const KF_ALTDOWN = &H2000
  Const UIS_CLEAR = 2
  Const UISF_HIDEACCEL = &H2
  Const UISF_HIDEFOCUS = &H1
  Const VK_DOWN = &H28
  Const VK_LEFT = &H25
  Const VK_RIGHT = &H27
  Const VK_TAB = &H9
  Const VK_UP = &H26
  Const WM_CHANGEUISTATE = &H127

  If HiWord(lParam) And KF_ALTDOWN Then
    ' this will make keyboard and focus cues work
    SendMessageAsLong Me.hWnd, WM_CHANGEUISTATE, MakeDWord(UIS_CLEAR, UISF_HIDEACCEL Or UISF_HIDEFOCUS), 0
  Else
    Select Case wParam
      Case VK_LEFT, VK_RIGHT, VK_UP, VK_DOWN, VK_TAB
        ' this will make focus cues work
        SendMessageAsLong Me.hWnd, WM_CHANGEUISTATE, MakeDWord(UIS_CLEAR, UISF_HIDEFOCUS), 0
    End Select
  End If
End Function


Private Sub cmdAbout_Click()
  TrackBarU.About
End Sub

Private Sub Form_Initialize()
  InitCommonControls
End Sub

Private Sub Form_Load()
  InstallKeyboardHook Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  RemoveKeyboardHook Me
End Sub

Private Sub TrackBarU_CustomDraw(ByVal controlPart As TrackBarCtlLibUCtl.CustomDrawControlPartConstants, ByVal drawStage As TrackBarCtlLibUCtl.CustomDrawStageConstants, ByVal controlState As TrackBarCtlLibUCtl.CustomDrawControlStateConstants, ByVal hDC As Long, drawingRectangle As TrackBarCtlLibUCtl.RECTANGLE, furtherProcessing As TrackBarCtlLibUCtl.CustomDrawReturnValuesConstants)
  Const DEFAULT_GUI_FONT = 17
  Const DT_CALCRECT = &H400
  Const DT_SINGLELINE = &H20
  Const TRANSPARENT = 1
  Dim hFont As Long
  Dim hPreviousFont As Long
  Dim i As Long
  Dim numberOfTickMarks As Long
  Dim s As String
  Dim textRectangle As RECT
  Dim textHeight As Long
  Dim textWidth As Long
  Dim v As Variant
  Dim xCenter As Long

  If drawStage = CustomDrawStageConstants.cdsPrePaint Then
    furtherProcessing = CustomDrawReturnValuesConstants.cdrvNotifyItemDraw
  ElseIf controlPart = CustomDrawControlPartConstants.cdcpTics Then
    Select Case drawStage
      Case CustomDrawStageConstants.cdsItemPrePaint
        furtherProcessing = CustomDrawReturnValuesConstants.cdrvNotifyPostPaint
      Case CustomDrawStageConstants.cdsItemPostPaint
        v = TrackBarU.TickMarks
        If VarType(v) And VbVarType.vbArray Then
          SetBkMode hDC, TRANSPARENT
          hFont = GetStockObject(DEFAULT_GUI_FONT)
          hPreviousFont = SelectObject(hDC, hFont)

          For i = LBound(v) To UBound(v)
            s = CStr(v(i))
            DrawText hDC, StrPtr(s), -1, textRectangle, DT_CALCRECT Or DT_SINGLELINE

            textHeight = textRectangle.Bottom - textRectangle.Top
            textWidth = textRectangle.Right - textRectangle.Left
            xCenter = TrackBarU.TickMarkPhysicalPosition(i - LBound(v))
            textRectangle.Left = xCenter - textWidth / 2
            textRectangle.Right = textRectangle.Left + textWidth
            textRectangle.Top = TrackBarU.ChannelTop + TrackBarU.ChannelHeight + 16
            textRectangle.Bottom = textRectangle.Top + textHeight

            DrawText hDC, StrPtr(s), -1, textRectangle, DT_SINGLELINE
          Next i
          SelectObject hDC, hPreviousFont
        End If

        furtherProcessing = CustomDrawReturnValuesConstants.cdrvDoDefault
    End Select
  End If
End Sub


' returns the higher 16 bits of <value>
Private Function HiWord(ByVal Value As Long) As Integer
  Dim ret As Integer

  CopyMemory VarPtr(ret), VarPtr(Value) + LenB(ret), LenB(ret)

  HiWord = ret
End Function

' makes a 32 bits number out of two 16 bits parts
Private Function MakeDWord(ByVal Lo As Integer, ByVal hi As Integer) As Long
  Dim ret As Long

  CopyMemory VarPtr(ret), VarPtr(Lo), LenB(Lo)
  CopyMemory VarPtr(ret) + LenB(Lo), VarPtr(hi), LenB(hi)

  MakeDWord = ret
End Function
