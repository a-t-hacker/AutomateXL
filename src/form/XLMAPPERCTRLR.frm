VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} XLMAPPERCTRLR 
   ClientHeight    =   615
   ClientLeft      =   2115
   ClientTop       =   465
   ClientWidth     =   3120
   OleObjectBlob   =   "XLMAPPERCTRLR.frx":0000
End
Attribute VB_Name = "XLMAPPERCTRLR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()

Call getEnvironment(appEnv, appBlk)

'//WinForm #
xWin = 20: Call setWindow(xWin)

MapPointer.ForeColor = vbRed

'//cleanup
ThisWorkbook.Worksheets("Main").Range("xlasBlkAddr175").Value2 = 0
ThisWorkbook.Worksheets("Main").Range("xlasBlkAddr176").Value = 0
ThisWorkbook.Worksheets("Main").Range("ClickType").Value2 = 1
ThisWorkbook.Worksheets("Main").Range("Offset").Value2 = 0
ThisWorkbook.Worksheets("Main").Range("OffsetStart").Value = 0

If ThisWorkbook.Worksheets("Main").Range("MapperX").Value <> vbNullString And ThisWorkbook.Worksheets("Main").Range("MapperX").Value <> 0 Then Me.Left = _
ThisWorkbook.Worksheets("Main").Range("MapperX").Value - ((ThisWorkbook.Worksheets("Main").Range("MapperX").Value / 100) * 33) - 10 Else Me.Left = 0

If ThisWorkbook.Worksheets("Main").Range("MapperY").Value <> vbNullString And ThisWorkbook.Worksheets("Main").Range("MapperY").Value <> 0 Then Me.Top = _
ThisWorkbook.Worksheets("Main").Range("MapperY").Value - ((ThisWorkbook.Worksheets("Main").Range("MapperY").Value / 100) * 36) - 25 Else Me.Top = 0

End Sub
Private Sub AddMapLBtn_Click()

Call AddMapLBtn_Clk
xBtn = 1: Call fxsActive(xBtn)

End Sub
Private Sub AddMapLBtn_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Call AddMapLBtn_DblClk

End Sub
Private Sub AddMapRBtn_Click()

Call AddMapRBtn_Clk
xBtn = 2: Call fxsActive(xBtn)

End Sub
Private Sub AddMapRBtn_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Call AddMapRBtn_DblClk

End Sub
Private Sub AddMapDblBtn_Click()

Call AddMapDblBtn_Clk
xBtn = 3: Call fxsActive(xBtn)

End Sub
Private Sub AddMapDblBtn_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Call AddMapDblBtn_DblClk

End Sub
Private Sub RmvMapBtn_Click()

Call RmvMapBtn_Clk

End Sub
Private Sub RmvAllMapBtn_Click()

Call RmvAllMapBtn_Clk

End Sub
Private Sub PlayBtn_Click()

ThisWorkbook.Worksheets("Main").Range("xlasSilent").Value = 1
Call bldScript
xPath = ThisWorkbook.Worksheets("Main").Range("MapperPath").Value2
Call LoadMappingBtn_Clk(xPath)
XLMAPPERCTRLR.Hide
XLMAPPER.Hide
Call PlayBtn_Clk
XLMAPPER.Show
ThisWorkbook.Worksheets("Main").Range("MapperActive").Value = 1
ThisWorkbook.Worksheets("Main").Range("xlasSilent").Value = 0

End Sub
Private Sub MapPointer_Click()

If ThisWorkbook.Worksheets("Main").Range("ClickType").Value2 = 0 Then Call AddMapLBtn_DblClk: Exit Sub
If ThisWorkbook.Worksheets("Main").Range("ClickType").Value2 = 1 Then Call AddMapLBtn_DblClk: Exit Sub
If ThisWorkbook.Worksheets("Main").Range("ClickType").Value2 = 2 Then Call AddMapRBtn_DblClk: Exit Sub
If ThisWorkbook.Worksheets("Main").Range("ClickType").Value2 = 3 Then Call AddMapRBtn_DblClk: Exit Sub
If ThisWorkbook.Worksheets("Main").Range("ClickType").Value2 = 4 Then Call AddMapDblBtn_DblClk: Exit Sub
If ThisWorkbook.Worksheets("Main").Range("ClickType").Value2 = 99 Then Call setKeyFlow(xKey): _
                                                                      Call AddMapBtn_Clk(cType): Exit Sub

End Sub
Private Sub HideToolBtn_Click()

XLMAPPER.SwitchToolBtn.Caption = vbNullString: XLMAPPER.AppNameLbl.Caption = "Save"
Call HideToolBtn_Clk
    
End Sub
Private Sub KeyFlowBtn_Click()

Call KeyFlowBtn_Clk
xBtn = 4: Call fxsActive(xBtn)

End Sub
Private Sub KeyFlowBtn_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Call KeyFlowBtn_DblClk

End Sub
Private Sub SwitchToolBtn_Click()

ThisWorkbook.Worksheets("Main").Range("xlasBlkAddr174").Value = 1: Me.Hide: _
XLMAPPER.SwitchToolBtn.Caption = "Switch": XLMAPPER.AppNameLbl.Caption = "Save"

End Sub
Private Sub StartPosBtn_Click()

ThisWorkbook.Worksheets("Main").Range("MapperX").Value = 0
ThisWorkbook.Worksheets("Main").Range("MapperY").Value = 0
Me.Left = 0
Me.Top = 0

End Sub
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

'//Special application inputs (can be used outside of key flow activation)
'//
'//esc
If KeyCode.Value = 27 Then
'//check for key flow activation
If ThisWorkbook.Worksheets("Main").Range("ClickType").Value2 = 99 Then ThisWorkbook.Worksheets("Main").Range("xlasBlkAddr176").Value = 1

If ThisWorkbook.Worksheets("Main").Range("xlasBlkAddr176").Value = 0 Then
ThisWorkbook.Worksheets("Main").Range("xlasBlkAddr174").Value = 1: Me.Hide: XLMAPPER.SwitchToolBtn.Caption = "Switch": XLMAPPER.AppNameLbl.Caption = "Save"
ElseIf ThisWorkbook.Worksheets("Main").Range("xlasBlkAddr176").Value = 1 Then
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{ESC}"
ElseIf ThisWorkbook.Worksheets("Main").Range("xlasBlkAddr176").Value = 2 Then
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{ESC}"
    End If
        ThisWorkbook.Worksheets("Main").Range("xlasBlkAddr176").Value = 0: Exit Sub
            End If
'//enter
If KeyCode.Value = 13 Then _
'//check for key flow activation
If ThisWorkbook.Worksheets("Main").Range("ClickType").Value2 = 99 Then
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{ENTER}": Exit Sub
    Else
If ThisWorkbook.Worksheets("Main").Range("ClickType").Value2 = 0 Then Call AddMapLBtn_DblClk: Exit Sub
If ThisWorkbook.Worksheets("Main").Range("ClickType").Value2 = 1 Then Call AddMapLBtn_DblClk: Exit Sub
If ThisWorkbook.Worksheets("Main").Range("ClickType").Value2 = 2 Then Call AddMapRBtn_DblClk: Exit Sub
If ThisWorkbook.Worksheets("Main").Range("ClickType").Value2 = 3 Then Call AddMapRBtn_DblClk: Exit Sub
If ThisWorkbook.Worksheets("Main").Range("ClickType").Value2 = 4 Then Call AddMapDblBtn_DblClk: Exit Sub
If ThisWorkbook.Worksheets("Main").Range("ClickType").Value2 = 99 Then Call setKeyFlow(xKey): _
                                                                      Call AddMapBtn_Clk(cType): Exit Sub
        End If
            End If
            '//#;

'//Check for key flow activation
If ThisWorkbook.Worksheets("Main").Range("ClickType").Value2 <> 99 Then Exit Sub
'//#;

'//Special keyboard inputs
'//
'//backspace
If KeyCode.Value = 8 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{BACKSPACE}": Exit Sub

'//clear
If KeyCode.Value = 12 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{CLEAR}": Exit Sub

'//caps lock
If KeyCode.Value = 20 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{CAPSLOCK}": Exit Sub

'//delete
If KeyCode.Value = 46 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{DELETE}": Exit Sub

'//insert
If KeyCode.Value = 45 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{INSERT}": Exit Sub

'//space
If KeyCode.Value = 32 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & " ": Exit Sub

'//tab
If KeyCode.Value = 9 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{TAB}": Exit Sub

'//up arrow
If KeyCode.Value = 38 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{UP}": Exit Sub

'//down arrow
If KeyCode.Value = 40 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{DOWN}": Exit Sub

'//left arrow
If KeyCode.Value = 37 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{LEFT}": Exit Sub

'//right arrow
If KeyCode.Value = 39 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{RIGHT}": Exit Sub

'//page up
If KeyCode.Value = 33 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{PGUP}": Exit Sub

'//page down
If KeyCode.Value = 34 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{PGDN}": Exit Sub

'//f1
If KeyCode.Value = 112 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{F1}": Exit Sub
'//f2
If KeyCode.Value = 113 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{F2}": Exit Sub
'//f3
If KeyCode.Value = 114 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{F3}": Exit Sub
'//f4
If KeyCode.Value = 115 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{F4}": Exit Sub
'//f5
If KeyCode.Value = 116 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{F5}": Exit Sub
'//f6
If KeyCode.Value = 117 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{F6}": Exit Sub
'//f7
If KeyCode.Value = 118 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{F7}": Exit Sub
'//f8
If KeyCode.Value = 119 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{F8}": Exit Sub
'//f9
If KeyCode.Value = 120 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{F9}": Exit Sub
'//f10
If KeyCode.Value = 121 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{F10}": Exit Sub
'//f11
If KeyCode.Value = 122 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{F11}": Exit Sub
'//f12
If KeyCode.Value = 123 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{F12}": Exit Sub
'//f13
If KeyCode.Value = 124 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{F13}": Exit Sub
'//f14
If KeyCode.Value = 125 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{F14}": Exit Sub
'//f15
If KeyCode.Value = 126 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{F15}": Exit Sub
'//f16
If KeyCode.Value = 127 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{F16}": Exit Sub

'//print screen
If KeyCode.Value = 121 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{PRTSC}": Exit Sub

'//home
If KeyCode.Value = 122 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{HOME}": Exit Sub

'//end
If KeyCode.Value = 123 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{END}": Exit Sub

'//windows key
If KeyCode.Value = 91 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = "^{Esc}": Exit Sub

'//comma
If KeyCode.Value = 188 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & ",": Exit Sub

'//period
If KeyCode.Value = 190 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & ".": Exit Sub

'//forward slash
If KeyCode.Value = 191 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "/": Exit Sub

'//back slash
If KeyCode.Value = 220 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "\": Exit Sub

'//semi-colon
If KeyCode.Value = 186 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & ";": Exit Sub

'//single quote
If KeyCode.Value = 222 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "'": Exit Sub

'//left bracket
If KeyCode.Value = 219 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "[": Exit Sub

'//right bracket
If KeyCode.Value = 221 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "]": Exit Sub

'//ctrl
If KeyCode.Value = 17 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "^": _
ThisWorkbook.Worksheets("Main").Range("xlasBlkAddr176").Value = ThisWorkbook.Worksheets("Main").Range("xlasBlkAddr176").Value + 1: Exit Sub

'//alt
If KeyCode.Value = 18 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "%": _
ThisWorkbook.Worksheets("Main").Range("xlasBlkAddr176").Value = ThisWorkbook.Worksheets("Main").Range("xlasBlkAddr176").Value + 1: Exit Sub

'//shift
If KeyCode.Value = 16 Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "+": _
ThisWorkbook.Worksheets("Main").Range("xlasBlkAddr176").Value = ThisWorkbook.Worksheets("Main").Range("xlasBlkAddr176").Value + 1: Exit Sub
 '//#;
 
'//record key flow (user keyboard input)
'//
'//Below for parsing special character inputs (ctrl, shift)
'//
If ThisWorkbook.Worksheets("Main").Range("xlasBlkAddr176").Value2 = 3 Then _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{" & LCase(Chr(KeyCode.Value)) & "}": _
ThisWorkbook.Worksheets("Main").Range("xlasBlkAddr176").Value2 = 0: Exit Sub

If ThisWorkbook.Worksheets("Main").Range("xlasBlkAddr176").Value2 = 2 Then _
If KeyCode.Value <> vbKeyShift And KeyCode.Value <> vbKeyControl And KeyCode.Value <> 18 _
Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{" & LCase(Chr(KeyCode.Value)) & "}" _
: ThisWorkbook.Worksheets("Main").Range("xlasBlkAddr176").Value2 = 0: Exit Sub

If ThisWorkbook.Worksheets("Main").Range("xlasBlkAddr176").Value2 = 1 Then _
If KeyCode.Value <> vbKeyShift And KeyCode.Value <> vbKeyControl And KeyCode.Value <> 18 _
Then ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = _
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & "{" & LCase(Chr(KeyCode.Value)) & "}" _
: ThisWorkbook.Worksheets("Main").Range("xlasBlkAddr176").Value2 = 0: Exit Sub
'//#;


Call getWindowPos
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value & LCase(Chr(KeyCode.Value))

End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Call getWindowPos

End Sub
Private Sub MapPointer_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

MapPointer.MousePointer = fmMousePointerCross
Call getWindowPos

End Sub
Private Sub AddMapLBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Call getWindowPos

End Sub
Private Sub AddMapRBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Call getWindowPos

End Sub
Private Sub AddMapDblBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Call getWindowPos

End Sub
Private Sub RmvMapBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Call getWindowPos

End Sub
Private Sub RmvAllMapBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Call getWindowPos

End Sub
Private Sub KeyFlowBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Call getWindowPos

End Sub
Private Sub SaveMapBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Call getWindowPos

End Sub
Private Sub UserForm_Terminate()

ThisWorkbook.Worksheets("Main").Range("xlasBlkAddr174").Value = 1
XLMAPPER.SwitchToolBtn.Caption = "Switch"
AUTOMATEXLHOME.Show

End Sub
