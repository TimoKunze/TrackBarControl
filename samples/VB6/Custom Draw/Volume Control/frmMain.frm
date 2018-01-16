VERSION 5.00
Object = "{956B5A46-C53F-45A7-AF0E-EC2E1CC9B567}#1.7#0"; "TrackBarCtlU.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "TrackBar 1.7 - Volume Control Sample"
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
      Height          =   615
      Left            =   2047
      TabIndex        =   0
      Top             =   360
      Width           =   1500
      _cx             =   2646
      _cy             =   1085
      Appearance      =   0
      AutoTickFrequency=   1
      AutoTickMarks   =   0   'False
      BackColor       =   -2147483633
      BackgroundDrawMode=   0
      BorderStyle     =   0
      CurrentPosition =   0
      DisabledEvents  =   11
      DontRedraw      =   0   'False
      DownIsLeft      =   -1  'True
      Enabled         =   -1  'True
      HoverTime       =   -1
      LargeStepWidth  =   -1
      Maximum         =   100
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
      TickMarksPosition=   3
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
  Private Declare Function CreatePen Lib "gdi32.dll" (ByVal fnPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
  Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
  Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function LineTo Lib "gdi32.dll" (ByVal hDC As Long, ByVal nXEnd As Long, ByVal nYEnd As Long) As Long
  Private Declare Function MoveToEx Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpPoint As Long) As Long
  Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
  Private Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


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
  Const COLOR_3DSHADOW = 16
  Const COLOR_3DHILIGHT = 20
  Const PS_SOLID = 0
  Dim channelRect As RECT
  Dim hLightPen As Long
  Dim hShadowPen As Long
  Dim hOldPen As Long

  If drawStage = CustomDrawStageConstants.cdsPrePaint Then
    furtherProcessing = CustomDrawReturnValuesConstants.cdrvNotifyItemDraw
  ElseIf controlPart = CustomDrawControlPartConstants.cdcpChannel Then
    Select Case drawStage
      Case CustomDrawStageConstants.cdsItemPrePaint
        With channelRect
          .Left = TrackBarU.ChannelLeft + 3
          .Top = TrackBarU.SliderTop + 3
          .Right = .Left + TrackBarU.ChannelWidth - 6
          .Bottom = .Top + TrackBarU.SliderHeight - 6

          hLightPen = CreatePen(PS_SOLID, 2, GetSysColor(COLOR_3DHILIGHT))
          If hLightPen Then
            hShadowPen = CreatePen(PS_SOLID, 2, GetSysColor(COLOR_3DSHADOW))
            If hShadowPen Then
              hOldPen = SelectObject(hDC, hLightPen)

              ' TODO: antialiasing
              MoveToEx hDC, .Left, .Bottom, 0
              LineTo hDC, .Right, .Bottom
              SelectObject hDC, hShadowPen
              MoveToEx hDC, .Left, .Bottom, 0
              LineTo hDC, .Right, .Top
              SelectObject hDC, hLightPen
              LineTo hDC, .Right, .Bottom + 1

              SelectObject hDC, hOldPen
              DeleteObject hShadowPen
              furtherProcessing = CustomDrawReturnValuesConstants.cdrvSkipDefault
            End If
            DeleteObject hLightPen
          End If
        End With
    End Select
  ElseIf controlPart = CustomDrawControlPartConstants.cdcpTics Then
    furtherProcessing = CustomDrawReturnValuesConstants.cdrvSkipDefault
  End If
End Sub

Private Sub TrackBarU_GetInfoTipText(ByVal maxInfoTipLength As Long, infoTipText As String, abortToolTip As Boolean)
  infoTipText = TrackBarU.CurrentPosition & " %"
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
