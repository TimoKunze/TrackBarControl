VERSION 5.00
Object = "{956B5A46-C53F-45A7-AF0E-EC2E1CC9B567}#1.7#0"; "TrackBarCtlU.ocx"
Object = "{F65D51B5-B8DD-476A-85B7-3E26E69D7B3C}#1.7#0"; "TrackBarCtlA.ocx"
Begin VB.Form frmMain 
   Caption         =   "TrackBar 1.7 - Events Sample"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10545
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
   ScaleHeight     =   385
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   703
   StartUpPosition =   2  'Bildschirmmitte
   Begin TrackBarCtlLibACtl.TrackBar TrackBarA 
      Height          =   615
      Left            =   2640
      TabIndex        =   1
      Top             =   4560
      Width           =   2295
      _cx             =   4048
      _cy             =   1085
      Appearance      =   0
      AutoTickFrequency=   1
      AutoTickMarks   =   -1  'True
      BackColor       =   -2147483633
      BackgroundDrawMode=   0
      BorderStyle     =   0
      CurrentPosition =   0
      DisabledEvents  =   0
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
   Begin VB.CheckBox chkLog 
      Caption         =   "&Log"
      Height          =   255
      Left            =   8760
      TabIndex        =   4
      Top             =   600
      Width           =   975
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
      Left            =   8760
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtLog 
      Height          =   4455
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   2
      Top             =   0
      Width           =   8655
   End
   Begin TrackBarCtlLibUCtl.TrackBar TrackBarU 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   4560
      Width           =   2295
      _cx             =   4048
      _cy             =   1085
      Appearance      =   0
      AutoTickFrequency=   1
      AutoTickMarks   =   -1  'True
      BackColor       =   -2147483633
      BackgroundDrawMode=   0
      BorderStyle     =   0
      CurrentPosition =   0
      DisabledEvents  =   0
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


  Private bLog As Boolean
  Private objActiveCtl As Object


  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
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


Private Sub chkLog_Click()
  bLog = (chkLog.Value = CheckBoxConstants.vbChecked)
End Sub

Private Sub cmdAbout_Click()
  objActiveCtl.About
End Sub

Private Sub Form_Initialize()
  InitCommonControls
End Sub

Private Sub Form_Load()
  chkLog.Value = CheckBoxConstants.vbChecked

  InstallKeyboardHook Me
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    On Error Resume Next
    cmdAbout.Left = Me.ScaleWidth - 5 - cmdAbout.Width
    chkLog.Left = cmdAbout.Left
    txtLog.Width = cmdAbout.Left - 5 - txtLog.Left
    TrackBarU.Top = Me.ScaleHeight - TrackBarU.Height - 10
    TrackBarA.Top = TrackBarU.Top
    txtLog.Height = TrackBarU.Top - 10 - txtLog.Top
    TrackBarU.Width = (Me.ScaleWidth - 16) / 2 - 5
    TrackBarA.Left = TrackBarU.Left + TrackBarU.Width + 10
    TrackBarA.Width = TrackBarU.Width
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  RemoveKeyboardHook Me
End Sub

Private Sub TrackBarA_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TrackBarA_Click: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TrackBarA_ContextMenu(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TrackBarA_ContextMenu: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TrackBarA_CustomDraw(ByVal controlPart As TrackBarCtlLibACtl.CustomDrawControlPartConstants, ByVal drawStage As TrackBarCtlLibACtl.CustomDrawStageConstants, ByVal controlState As TrackBarCtlLibACtl.CustomDrawControlStateConstants, ByVal hDC As Long, drawingRectangle As TrackBarCtlLibACtl.RECTANGLE, furtherProcessing As TrackBarCtlLibACtl.CustomDrawReturnValuesConstants)
  'AddLogEntry "TrackBarA_CustomDraw: controlPart=" & controlPart & ", drawStage=" & drawStage & ", controlState=" & controlState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & "), furtherProcessing=" & furtherProcessing
End Sub

Private Sub TrackBarA_DblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TrackBarA_DblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TrackBarA_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "TrackBarA_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub TrackBarA_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "TrackBarA_DragDrop"
End Sub

Private Sub TrackBarA_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "TrackBarA_DragOver"
End Sub

Private Sub TrackBarA_GetInfoTipText(ByVal maxInfoTipLength As Long, infoTipText As String, abortToolTip As Boolean)
  AddLogEntry "TrackBarA_GetInfoTipText: maxInfoTipLength=" & maxInfoTipLength & ", infoTipText=" & infoTipText & ", abortToolTip=" & abortToolTip

  infoTipText = "Current position: " & TrackBarA.CurrentPosition
End Sub

Private Sub TrackBarA_GotFocus()
  AddLogEntry "TrackBarA_GotFocus"
  Set objActiveCtl = TrackBarA
End Sub

Private Sub TrackBarA_KeyDown(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "TrackBarA_KeyDown: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub TrackBarA_KeyPress(keyAscii As Integer)
  AddLogEntry "TrackBarA_KeyPress: keyAscii=" & keyAscii
End Sub

Private Sub TrackBarA_KeyUp(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "TrackBarA_KeyUp: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub TrackBarA_LostFocus()
  AddLogEntry "TrackBarA_LostFocus"
End Sub

Private Sub TrackBarA_MClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TrackBarA_MClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TrackBarA_MDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TrackBarA_MDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TrackBarA_MouseDown(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TrackBarA_MouseDown: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TrackBarA_MouseEnter(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TrackBarA_MouseEnter: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TrackBarA_MouseHover(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TrackBarA_MouseHover: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TrackBarA_MouseLeave(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TrackBarA_MouseLeave: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TrackBarA_MouseMove(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  'AddLogEntry "TrackBarA_MouseMove: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TrackBarA_MouseUp(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TrackBarA_MouseUp: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TrackBarA_MouseWheel(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal scrollAxis As TrackBarCtlLibACtl.ScrollAxisConstants, ByVal wheelDelta As Integer)
  AddLogEntry "TrackBarA_MouseWheel: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta
End Sub

Private Sub TrackBarA_OLEDragDrop(ByVal Data As TrackBarCtlLibACtl.IOLEDataObject, effect As TrackBarCtlLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "TrackBarA_OLEDragDrop: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str

  If Data.GetFormat(vbCFFiles) Then
    files = Data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    TrackBarA.FinishOLEDragDrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub TrackBarA_OLEDragEnter(ByVal Data As TrackBarCtlLibACtl.IOLEDataObject, effect As TrackBarCtlLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "TrackBarA_OLEDragEnter: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub TrackBarA_OLEDragLeave(ByVal Data As TrackBarCtlLibACtl.IOLEDataObject, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "TrackBarA_OLEDragLeave: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub TrackBarA_OLEDragMouseMove(ByVal Data As TrackBarCtlLibACtl.IOLEDataObject, effect As TrackBarCtlLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "TrackBarA_OLEDragMouseMove: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub TrackBarA_PositionChanged(ByVal changeType As TrackBarCtlLibACtl.PositionChangeTypeConstants, ByVal newPosition As Long)
  AddLogEntry "TrackBarA_PositionChanged: changeType=" & changeType & ", newPosition=" & newPosition
End Sub

Private Sub TrackBarA_PositionChanging(ByVal changeType As TrackBarCtlLibACtl.PositionChangeTypeConstants, ByVal newPosition As Long, cancelChange As Boolean)
  AddLogEntry "TrackBarA_PositionChanging: changeType=" & changeType & ", newPosition=" & newPosition & ", cancelChange=" & cancelChange
End Sub

Private Sub TrackBarA_RClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TrackBarA_RClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TrackBarA_RDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TrackBarA_RDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TrackBarA_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "TrackBarA_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub TrackBarA_ResizedControlWindow()
  AddLogEntry "TrackBarA_ResizedControlWindow"
End Sub

Private Sub TrackBarA_Validate(Cancel As Boolean)
  AddLogEntry "TrackBarA_Validate"
End Sub

Private Sub TrackBarA_XClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TrackBarA_XClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TrackBarA_XDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TrackBarA_XDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TrackBarU_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TrackBarU_Click: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TrackBarU_ContextMenu(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TrackBarU_ContextMenu: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TrackBarU_CustomDraw(ByVal controlPart As TrackBarCtlLibUCtl.CustomDrawControlPartConstants, ByVal drawStage As TrackBarCtlLibUCtl.CustomDrawStageConstants, ByVal controlState As TrackBarCtlLibUCtl.CustomDrawControlStateConstants, ByVal hDC As Long, drawingRectangle As TrackBarCtlLibUCtl.RECTANGLE, furtherProcessing As TrackBarCtlLibUCtl.CustomDrawReturnValuesConstants)
  'AddLogEntry "TrackBarU_CustomDraw: controlPart=" & controlPart & ", drawStage=" & drawStage & ", controlState=" & controlState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & "), furtherProcessing=" & furtherProcessing
End Sub

Private Sub TrackBarU_DblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TrackBarU_DblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TrackBarU_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "TrackBarU_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub TrackBarU_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "TrackBarU_DragDrop"
End Sub

Private Sub TrackBarU_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "TrackBarU_DragOver"
End Sub

Private Sub TrackBarU_GetInfoTipText(ByVal maxInfoTipLength As Long, infoTipText As String, abortToolTip As Boolean)
  AddLogEntry "TrackBarU_GetInfoTipText: maxInfoTipLength=" & maxInfoTipLength & ", infoTipText=" & infoTipText & ", abortToolTip=" & abortToolTip

  infoTipText = "Current position: " & TrackBarU.CurrentPosition
End Sub

Private Sub TrackBarU_GotFocus()
  AddLogEntry "TrackBarU_GotFocus"
  Set objActiveCtl = TrackBarU
End Sub

Private Sub TrackBarU_KeyDown(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "TrackBarU_KeyDown: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub TrackBarU_KeyPress(keyAscii As Integer)
  AddLogEntry "TrackBarU_KeyPress: keyAscii=" & keyAscii
End Sub

Private Sub TrackBarU_KeyUp(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "TrackBarU_KeyUp: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub TrackBarU_LostFocus()
  AddLogEntry "TrackBarU_LostFocus"
End Sub

Private Sub TrackBarU_MClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TrackBarU_MClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TrackBarU_MDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TrackBarU_MDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TrackBarU_MouseDown(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TrackBarU_MouseDown: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TrackBarU_MouseEnter(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TrackBarU_MouseEnter: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TrackBarU_MouseHover(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TrackBarU_MouseHover: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TrackBarU_MouseLeave(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TrackBarU_MouseLeave: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TrackBarU_MouseMove(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  'AddLogEntry "TrackBarU_MouseMove: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TrackBarU_MouseUp(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TrackBarU_MouseUp: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TrackBarU_MouseWheel(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal scrollAxis As TrackBarCtlLibUCtl.ScrollAxisConstants, ByVal wheelDelta As Integer)
  AddLogEntry "TrackBarU_MouseWheel: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta
End Sub

Private Sub TrackBarU_OLEDragDrop(ByVal Data As TrackBarCtlLibUCtl.IOLEDataObject, effect As TrackBarCtlLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "TrackBarU_OLEDragDrop: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str

  If Data.GetFormat(vbCFFiles) Then
    files = Data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    TrackBarU.FinishOLEDragDrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub TrackBarU_OLEDragEnter(ByVal Data As TrackBarCtlLibUCtl.IOLEDataObject, effect As TrackBarCtlLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "TrackBarU_OLEDragEnter: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub TrackBarU_OLEDragLeave(ByVal Data As TrackBarCtlLibUCtl.IOLEDataObject, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "TrackBarU_OLEDragLeave: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub TrackBarU_OLEDragMouseMove(ByVal Data As TrackBarCtlLibUCtl.IOLEDataObject, effect As TrackBarCtlLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "TrackBarU_OLEDragMouseMove: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub TrackBarU_PositionChanged(ByVal changeType As TrackBarCtlLibUCtl.PositionChangeTypeConstants, ByVal newPosition As Long)
  AddLogEntry "TrackBarU_PositionChanged: changeType=" & changeType & ", newPosition=" & newPosition
End Sub

Private Sub TrackBarU_PositionChanging(ByVal changeType As TrackBarCtlLibUCtl.PositionChangeTypeConstants, ByVal newPosition As Long, cancelChange As Boolean)
  AddLogEntry "TrackBarU_PositionChanging: changeType=" & changeType & ", newPosition=" & newPosition & ", cancelChange=" & cancelChange
  'cancelChange = True
End Sub

Private Sub TrackBarU_RClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TrackBarU_RClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TrackBarU_RDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TrackBarU_RDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TrackBarU_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "TrackBarU_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub TrackBarU_ResizedControlWindow()
  AddLogEntry "TrackBarU_ResizedControlWindow"
  txtLog.Height = TrackBarU.Top - 10 - txtLog.Top
End Sub

Private Sub TrackBarU_Validate(Cancel As Boolean)
  AddLogEntry "TrackBarU_Validate"
End Sub

Private Sub TrackBarU_XClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TrackBarU_XClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TrackBarU_XDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TrackBarU_XDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub


Private Sub AddLogEntry(ByVal txt As String)
  Dim pos As Long
  Static cLines As Long
  Static oldtxt As String

  If bLog Then
    cLines = cLines + 1
    If cLines > 50 Then
      ' delete the first line
      pos = InStr(oldtxt, vbNewLine)
      oldtxt = Mid(oldtxt, pos + Len(vbNewLine))
      cLines = cLines - 1
    End If
    oldtxt = oldtxt & (txt & vbNewLine)

    txtLog.Text = oldtxt
    txtLog.SelStart = Len(oldtxt)
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
