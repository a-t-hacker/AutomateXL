Attribute VB_Name = "AutomateXL_Click"
Public Function AddMapBtn_Clk(ByVal cType As String)

Dim StartTimer As Double: Dim EndTime As Double: Dim RunTime As Double
Dim lastR As Long

lastR = Cells(Rows.Count, "B").End(xlUp).Row

'//check for click type...
If cType = vbNullString Then
If ThisWorkbook.Worksheets("Main").Range("ClickType").Value2 = 0 Then cType = "-leftdown"
If ThisWorkbook.Worksheets("Main").Range("ClickType").Value2 = 1 Then cType = "-leftup"
If ThisWorkbook.Worksheets("Main").Range("ClickType").Value2 = 2 Then cType = "-rightdown"
If ThisWorkbook.Worksheets("Main").Range("ClickType").Value2 = 3 Then cType = "-rightup"
If ThisWorkbook.Worksheets("Main").Range("ClickType").Value2 = 4 Then cType = "-double"
If ThisWorkbook.Worksheets("Main").Range("ClickType").Value2 = 99 Then cType = "99"
End If

'//end previously running offset timer...
    If ThisWorkbook.Worksheets("Main").Range("Offset").Value2 = 1 Then
    If ThisWorkbook.Worksheets("Main").Range("OffsetStart").Value2 <> 0 Then StartTimer = ThisWorkbook.Worksheets("Main").Range("OffsetStart").Value2
    EndTime = Timer
    RunTime = EndTime - StartTimer
    ThisWorkbook.Worksheets("Main").Range("Offset").Offset(lastR - 1, 0).Value2 = CInt(RunTime)
    ThisWorkbook.Worksheets("Main").Range("Offset").Value = 0
    End If
'//adding click...
If ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value2 = vbNullString Then
ThisWorkbook.Worksheets("Main").Range("MapperXY").Offset(lastR, 0).Value2 = _
CStr(ThisWorkbook.Worksheets("Main").Range("MapperX").Value2) & "[,]" & CStr(ThisWorkbook.Worksheets("Main").Range("MapperY").Value2)
ThisWorkbook.Worksheets("Main").Range("ClickType").Offset(lastR, 0).Value2 = cType
    Else
'//adding key input...
    ThisWorkbook.Worksheets("Main").Range("MapperXY").Offset(lastR, 0).Value2 = ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value2
    ThisWorkbook.Worksheets("Main").Range("ClickType").Offset(lastR, 0).Value2 = "[(]" & ThisWorkbook.Worksheets("Main").Range("xlasBlkAddr175").Value2 & "[)]"
    If ThisWorkbook.Worksheets("Main").Range("xlasBlkAddr175").Value2 < 6 Then ThisWorkbook.Worksheets("Main").Range("xlasBlkAddr175").Value2 = ThisWorkbook.Worksheets("Main").Range("xlasBlkAddr175").Value2 + 1 _
    Else ThisWorkbook.Worksheets("Main").Range("xlasBlkAddr175").Value2 = 0
        End If
        
'//start new offset timer...
If ThisWorkbook.Worksheets("Main").Range("Offset").Value <> 1 Then
ThisWorkbook.Worksheets("Main").Range("Offset").Value = 1
StartTimer = Timer
ThisWorkbook.Worksheets("Main").Range("OffsetStart").Value2 = StartTimer
End If

'//record last map...
ThisWorkbook.Worksheets("Main").Range("LastMap").Value = "(" & ThisWorkbook.Worksheets("Main").Range("MapperXY").Offset(lastR, 0).Value2 & ") " & _
"(" & ThisWorkbook.Worksheets("Main").Range("ClickType").Offset(lastR, 0).Value2 & ")"

'//record map count...
ThisWorkbook.Worksheets("Main").Range("MapCount").Value2 = lastR

'//clear key control...
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value2 = vbNullString

End Function
Public Function AddMapLBtn_Clk()

If ThisWorkbook.Worksheets("Main").Range("ClickType").Value2 = 1 Then Range("ClickType").Value = 0 Else Range("ClickType").Value = 1

End Function
Public Function AddMapLBtn_DblClk()

If ThisWorkbook.Worksheets("Main").Range("ClickType").Value2 = 0 Then Range("ClickType").Value = 1 Else Range("ClickType").Value = 0

Call getWindowPos
Call AddMapBtn_Clk(cType)

End Function
Public Function AddMapRBtn_Clk()

If ThisWorkbook.Worksheets("Main").Range("ClickType").Value2 = 3 Then Range("ClickType").Value = 2 Else Range("ClickType").Value = 3

End Function
Public Function AddMapRBtn_DblClk()

If ThisWorkbook.Worksheets("Main").Range("ClickType").Value2 = 2 Then Range("ClickType").Value = 3 Else Range("ClickType").Value = 2

Call getWindowPos
Call AddMapBtn_Clk(cType)

End Function
Public Function AddMapDblBtn_Clk()

ThisWorkbook.Worksheets("Main").Range("ClickType").Value2 = 4

End Function
Public Function AddMapDblBtn_DblClk()

If ThisWorkbook.Worksheets("Main").Range("ClickType").Value2 = 3 Then Range("ClickType").Value = 4 Else Range("ClickType").Value = 4

Call getWindowPos
Call AddMapBtn_Clk(cType)

End Function
Public Function KeyFlowBtn_Clk()

ThisWorkbook.Worksheets("Main").Range("ClickType").Value2 = 99

End Function
Public Function KeyFlowBtn_DblClk()

XLMAPPER.KeyFlowBox.Caption = vbNullString
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value2 = vbNullString
xMsg = 5: Call AppMsg(xMsg)

End Function
Public Function HideToolBtn_Clk()

On Error Resume Next

Dim X As Byte

XLMAPPERCTRLR.Hide
XLMAPPER.Hide
X = MsgBox("Display mapping tools?", vbOKOnly, AppTag)
    If X = vbOK Then
        Call SH_XLMAPPER
            End If
    
End Function
Public Function RmvMapBtn_Clk()

Dim lastR As Long

lastR = Cells(Rows.Count, "B").End(xlUp).Row

ThisWorkbook.Worksheets("Main").Range("MapperXY").Offset(lastR - 1, 0).Value2 = vbNullString
ThisWorkbook.Worksheets("Main").Range("ClickType").Offset(lastR - 1, 0).Value2 = vbNullString

'//map count
ThisWorkbook.Worksheets("Main").Range("MapCount").Value2 = lastR - 2

'//set mapper to last position
If lastR - 2 >= 0 Then
If ThisWorkbook.Worksheets("Main").Range("MapperXY").Offset(lastR - 2, 0).Value2 <> vbNullString Then
XLMAPPER.LastMapBox.Caption = ThisWorkbook.Worksheets("Main").Range("MapperXY").Offset(lastR - 2, 0).Value2
xPos = ThisWorkbook.Worksheets("Main").Range("MapperXY").Offset(lastR - 2, 0).Value2
ThisWorkbook.Worksheets("Main").Range("LastMap").Value2 = vbNullString
Call setCtrlrPos(xPos)
End If
ElseIf lastR - 2 < 0 Then
ThisWorkbook.Worksheets("Main").Range("LastMap").Value2 = vbNullString
XLMAPPER.LastMapBox.Caption = vbNullString
xPos = "0,0": Call setCtrlrPos(xPos)
    End If

Call getData

End Function
Public Function RmvAllMapBtn_Clk()

Call ClnMainSpace
Call getData
XLMAPPERCTRLR.AddMapLBtn.BackStyle = fmBackStyleTransparent
XLMAPPERCTRLR.AddMapRBtn.BackStyle = fmBackStyleTransparent
XLMAPPERCTRLR.AddMapDblBtn.BackStyle = fmBackStyleTransparent
XLMAPPER.LastMapBox.Caption = vbNullString
ThisWorkbook.Worksheets("Main").Range("LastMap").Value2 = vbNullString
xMsg = 3: Call AppMsg(xMsg)

End Function
Public Function NewMappingBtn_Clk()

On Error Resume Next

Dim I As Byte
Dim xName As String

xName = InputBox("Enter a new name for your mapping:", AppTag)

If xName <> vbNullString Then
ThisWorkbook.Worksheets("Main").Range("MapperPath").Value2 = AppLoc & "\scripts\" & xName & ".xlas"
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value2 = vbNullString
XLMAPPERCTRLR.Caption = "xlMapper - " & xName
Call ClnMainSpace
Call getData
'//show XLMAPPER
XLMAPPER.Show
    Else
    xMsg = 1: Call AppMsg(xMsg)
    AUTOMATEXLHOME.Show
    Exit Function
        End If

End Function
Public Function PlayBtn_Clk()

Dim xPath As String: Dim xStr As String: Dim xStrH As String

xPath = ThisWorkbook.Worksheets("Main").Range("MapperPath").Value2

If xPath <> vbNullString And Dir(xPath) <> vbNullString Then

AUTOMATEXLHOME.Hide

Open xPath For Input As #1
Do Until EOF(1)
Line Input #1, xStr
xStrH = xStrH & xStr
Loop
Close #1

Art = xStrH & "$": Call xlas(Art)

    Else
    xMsg = 4: Call AppMsg(xMsg)
        End If

End Function
Public Function SaveMapBtn_Clk()

ThisWorkbook.Worksheets("Main").Range("Offset").Value2 = 0

Call bldScript

End Function
Public Function LoadMappingBtn_Clk(ByVal xPath As String)

Dim xStr As String: Dim xStrH As String

If xPath = vbNullString Then
ChDir (AppLoc & "\scripts\")
xPath = Application.GetOpenFilename()
If xPath = "False" Then Exit Function
End If

ThisWorkbook.Worksheets("Main").Range("MapperPath").Value2 = xPath: _
xMsg = 6: Call AppMsg(xMsg)

End Function
Public Function xlFlowStripBar_Clk()

If XLMAPPER.Height <> 500 Then _
XLMAPPER.Height = 500: _
XLMAPPER.xlFlowStrip.Height = 280 _
Else XLMAPPER.Height = 280

End Function

