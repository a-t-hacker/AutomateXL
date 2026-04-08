VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} XLMAPPER 
   ClientHeight    =   5025
   ClientLeft      =   2115
   ClientTop       =   465
   ClientWidth     =   3120
   OleObjectBlob   =   "XLMAPPER.frx":0000
End
Attribute VB_Name = "XLMAPPER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()

Call getEnvironment(appEnv, appBlk)

'//WinForm #
xWin = 19: Call setWindow(xWin)

Call getData

MapPointer.ForeColor = vbRed

If ThisWorkbook.Worksheets("Main").Range("xlasBlkAddr174").Value <> 1 Then
Art = "<lib>xbas;delayevent(15);$": Call xlas(Art)
XLMAPPERCTRLR.Show
Else
    ThisWorkbook.Worksheets("Main").Range("xlasBlkAddr174").Value = 0
    Exit Sub
        End If

ThisWorkbook.Worksheets("Main").Range("MapperActive").Value = 1

End Sub
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If KeyCode.Value = 27 Then Me.Hide: AUTOMATEXLHOME.Show: Exit Sub

End Sub
Private Sub AppPath_Click()

Dim xPath As String

xPath = InputBox("Enter a new path for your mapping:", "xlMapper", AppPath.Caption)

If xPath <> vbNullString And InStr(1, xPath, "\") Then
Call setProjPath(xPath)
End If

End Sub
Private Sub AppNameLbl_Click()

Call SaveMapBtn_Clk

End Sub
Private Sub HomeBtn_Click()

Me.Hide: AUTOMATEXLHOME.Show

End Sub
Private Sub CtrlrPosBox_Click()

Dim xPos As String

xPos = InputBox("Set controller position:", "xlMapper", CtrlrPosBox.Caption)

If xPos <> vbNullString And InStr(1, xPos, ",") Then
xPos = Replace(xPos, ",", "[,]")
Call setCtrlrPos(xPos)
End If

End Sub

Private Sub KeyFlowBox_Click()

Dim xKey As String

xKey = InputBox("Set key flow:", "xlMapper", KeyFlowBox.Caption)

Call setKeyFlow(xKey)

End Sub

Private Sub LastMapBox_Click()

MsgBox ("Last Map:" & vbNewLine & vbNewLine & LastMapBox.Caption), vbOKOnly, "xlMapper"

End Sub
Private Sub LastMapLbl_Click()

Call AutomateXL_Click.RmvAllMapBtn_Clk

End Sub

Private Sub MapCtLbl_Click()

MsgBox ("Map Count:" & vbNewLine & vbNewLine & MapCtLbl.Caption), vbOKOnly, "xlMapper"

End Sub
Private Sub SwitchToolBtn_Click()

XLMAPPER.SwitchToolBtn.Caption = vbNullString: XLMAPPER.AppNameLbl.Caption = "xlMapper": XLMAPPERCTRLR.Show

End Sub
Private Sub xlFlowStripBar_Click()

Call xlFlowStripBar_Clk

End Sub
Private Sub CtrlrPosLbl_Click()

xPos = "0[,]0"
Call setCtrlrPos(xPos)

End Sub
Private Sub KeyFlowLbl_Click()

Call KeyFlowBtn_DblClk

End Sub
Private Sub xlFlowStrip_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

Call getEnvironment(appEnv, appBlk)

'//Shift (save script)
If KeyCode.Value = vbKeyShift Then
Dim Art As String
Dim lastR As Integer
Art = xlFlowStrip.Value
If InStr(1, Art, "$") Then '//check for run trigger
Art = Replace(Art, "[$]", "*/DOLLAR")
Art = Replace(Art, "$", vbNullString)
lastR = Cells(Rows.Count, "B").End(xlUp).Row
XLMAPPER.LastMapBox.Caption = Art
ThisWorkbook.Worksheets("Main").Range("LastMap").Value2 = Art
ThisWorkbook.Worksheets("Main").Range("MapperXY").Offset(lastR, 0).Value2 = Art
ThisWorkbook.Worksheets("Main").Range("ClickType").Offset(lastR, 0).Value2 = "xlas"
Exit Sub
    End If
        End If

'//Esc
If KeyCode.Value = 27 Then Me.Hide: AUTOMATEXLHOME.Show: Exit Sub

'//Enter key
If KeyCode.Value = 13 Then
xlFlowStrip.EnterKeyBehavior = True
Exit Sub
End If

'//Tab key
If KeyCode.Value = 9 Then
xlFlowStrip.TabKeyBehavior = True
Exit Sub
End If

'//Alt key
If KeyCode.Value = 18 Then
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value = Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value + 18
KeyCode.Value = 0
Exit Sub
End If

'//Ctrl key
If KeyCode.Value = vbKeyControl Then
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value = vbKeyControl
KeyCode.Value = 0
Exit Sub
End If

'//Key Ctrl+D
If KeyCode.Value = vbKeyD Then
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value = 17 Then
'//Clear screen...
xlFlowStrip.Value = vbNullString
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

''//Key Ctrl+F
'If KeyCode.Value = vbKeyF Then
'If Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value = 17 Then
''//Clear screen...
'XLFONTBOX.Show
'Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value = vbNullString
'KeyCode.Value = 0
'Exit Sub
'End If
'End If
'
''//Key Ctrl+H
'If KeyCode.Value = vbKeyH Then
'If Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value = 17 Then
''//Replace...
'XLREPLACE.Show
'Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value = vbNullString
'KeyCode.Value = 0
'Exit Sub
'End If
'End If

'//Key Ctrl+N
If KeyCode.Value = vbKeyN Then
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value = 17 Then
'//Create a new project hotkey...
Call NewMappingBtn_Clk
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If
        
'//Key Ctrl+O
If KeyCode.Value = vbKeyO Then
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value = 17 Then
'//Open new project hotkey...
Call LoadMappingBtn_Clk(xPath)
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If
        
'//Key Ctrl+S
If KeyCode.Value = vbKeyS Then
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value = 17 Then
'//Save project hotkey...
Call SaveMapBtn_Clk
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

'//Key Ctrl+Q
If KeyCode.Value = vbKeyQ Then
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value = 17 Then
'//Close app hotkey w/o saving first...
ThisWorkbook.Close
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If
       
'//Key Ctrl+W
If KeyCode.Value = vbKeyW Then
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value = 17 Then
'//Maximize xlFlowStrip window size
Call xlFlowStripBar_Clk
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

'//Key Ctrl+Alt+Q
If KeyCode.Value = vbKeyQ Then
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value = 35 Then
'//Save and close app hotkey...
ThisWorkbook.Save: ThisWorkbook.Close
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

'//Key Ctrl+Alt+W
If KeyCode.Value = vbKeyW Then
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value = 35 Then
'//Hide application...
Call AutomateXL_Click.HideToolBtn_Clk
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value = vbNullString

End Sub
Private Sub SwitchToolBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

'SwitchToolBtn.Caption = "Switch"

End Sub

Private Sub UserForm_Terminate()

AUTOMATEXLHOME.Show

ThisWorkbook.Worksheets("Main").Range("MapperActive").Value = 0

End Sub

