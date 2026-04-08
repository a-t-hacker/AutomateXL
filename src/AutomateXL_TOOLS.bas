Attribute VB_Name = "AutomateXL_TOOLS"
'/##################\
'//Application Tools\\
'///################\\\

Public Function bldScript()

Dim lastR As Long: Dim X As Long
Dim xPath As String: Dim xPos As String: Dim xTime As String: Dim xType As String

xPath = ThisWorkbook.Worksheets("Main").Range("MapperPath").Value2

lastR = Cells(Rows.Count, "B").End(xlUp).Row

Open xPath For Output As #1
Print #1, "<lib> xbas;"
For X = 1 To lastR - 1
xType = ThisWorkbook.Worksheets("Main").Range("ClickType").Offset(X, 0).Value2
xPos = ThisWorkbook.Worksheets("Main").Range("MapperXY").Offset(X, 0).Value2
xTime = ThisWorkbook.Worksheets("Main").Range("Offset").Offset(X, 0).Value2
xPos = Replace(xPos, "[,]", ",")
'//print xlAppScript <---
If InStr(1, xType, "xlas") Then _
Print #1, xPos: GoTo NextLine
'//print key input <---
If InStr(1, xType, "[(]") Or InStr(1, xType, "[)]") Then _
xType = Replace(xType, "[(]", "("): xType = Replace(xType, "[)]", ")"): _
Print #1, "key" & xType & "('" & xPos & "');": GoTo NextLine
'//print click/mouse input <---
If InStr(1, xType, "-") Then _
Print #1, "click(" & xType & " " & xPos & ");": GoTo NextLine
NextLine:
'//print offset <---
If xTime <> vbNullString Then _
Print #1, "wait(" & xTime & "s);" Else Print #1, "wait(2s);"
Next
Close #1

xMsg = 2: Call AppMsg(xMsg)

End Function
Public Function setKeyFlow(ByVal xKey As String)

If xKey = vbNullString Then xKey = ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value
ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value = xKey
XLMAPPER.KeyFlowBox.Caption = xKey

End Function
Public Function setProjPath(xPath)

ThisWorkbook.Worksheets("Main").Range("MapperPath").Value = xPath
XLMAPPER.AppPath.Caption = xPath
xPathArr = Split(xPath, "\")
XLMAPPERCTRLR.Caption = "xlMapper - " & xPathArr(UBound(xPathArr))

End Function
Public Function setCtrlrPos(xPos)

If InStr(1, xPos, "[,]") Then
xPosArr = Split(xPos, "[,]")
ThisWorkbook.Worksheets("Main").Range("MapperX").Value2 = xPosArr(0)
ThisWorkbook.Worksheets("Main").Range("MapperY").Value2 = xPosArr(1)
XLMAPPERCTRLR.Left = xPosArr(0)
XLMAPPERCTRLR.Top = xPosArr(1)
XLMAPPER.CtrlrPosBox.Caption = xPosArr(0) & "," & xPosArr(1)
End If

End Function
Public Function fxsLoading()

Art = "<lib>xbas;delayevent(10);$": Call xlas(Art) '//delay

'//Play Recording
AUTOMATEXLHOME.PlayBtn.Visible = True
AUTOMATEXLHOME.AppIcon1.Visible = True
AUTOMATEXLHOME.PlayBtn.ForeColor = RGB(185, 231, 170) '//set flicker color

Art = "<lib>xbas;delayevent(10);$": Call xlas(Art) '//delay

'//New Recording
AUTOMATEXLHOME.NewMappingBtn.Visible = True
AUTOMATEXLHOME.AppIcon2.Visible = True
AUTOMATEXLHOME.NewMappingBtn.ForeColor = RGB(185, 231, 170) '//set flicker color

AUTOMATEXLHOME.PlayBtn.ForeColor = &HE0E0E0 '//set default color

Art = "<lib>xbas;delayevent(10);$": Call xlas(Art) '//delay

'//Load Recording
AUTOMATEXLHOME.LoadMappingBtn.Visible = True
AUTOMATEXLHOME.AppIcon3.Visible = True
AUTOMATEXLHOME.LoadMappingBtn.ForeColor = RGB(185, 231, 170) '//set flicker color

AUTOMATEXLHOME.NewMappingBtn.ForeColor = &HE0E0E0 '//set default color

Art = "<lib>xbas;delayevent(10);$": Call xlas(Art) '//delay

'//Exit
AUTOMATEXLHOME.ExitBtn.Visible = True
AUTOMATEXLHOME.ExitBtn.ForeColor = RGB(185, 231, 170) '//set flicker color

AUTOMATEXLHOME.LoadMappingBtn.ForeColor = &HE0E0E0 '//set default color

Art = "<lib>xbas;delayevent(10);$": Call xlas(Art) '//delay

'//Application Name
AUTOMATEXLHOME.AppNameLbl.Visible = True
AUTOMATEXLHOME.ExitBtn.ForeColor = &HE0E0E0 '//set default color

Art = "<lib>xbas;delayevent(5);$": Call xlas(Art) '//delay

'//Application Build
AUTOMATEXLHOME.AppBuildLbl.Visible = True
AUTOMATEXLHOME.AppNameLbl.ForeColor = vbBlack

End Function
Public Function fxsHover(ByVal xHov As Byte)

On Error Resume Next

'//Play Recording
If xHov = 1 Then
AUTOMATEXLHOME.PlayBtn.ForeColor = RGB(185, 231, 170)
AUTOMATEXLHOME.PlayBtn.Font.Size = 18
AUTOMATEXLHOME.NewMappingBtn.ForeColor = &HE0E0E0
AUTOMATEXLHOME.LoadMappingBtn.ForeColor = &HE0E0E0
AUTOMATEXLHOME.ExitBtn.ForeColor = &HE0E0E0
AUTOMATEXLHOME.AppIcon1.ForeColor = &H8000000E
AUTOMATEXLHOME.AppIcon2.ForeColor = &H8000000E
AUTOMATEXLHOME.AppIcon3.ForeColor = &H8000000E
Exit Function

'//New Recording
ElseIf xHov = 2 Then
AUTOMATEXLHOME.NewMappingBtn.ForeColor = RGB(185, 231, 170)
AUTOMATEXLHOME.NewMappingBtn.Font.Size = 18
AUTOMATEXLHOME.PlayBtn.ForeColor = &HE0E0E0
AUTOMATEXLHOME.LoadMappingBtn.ForeColor = &HE0E0E0
AUTOMATEXLHOME.ExitBtn.ForeColor = &HE0E0E0
AUTOMATEXLHOME.AppIcon1.ForeColor = &H8000000E
AUTOMATEXLHOME.AppIcon2.ForeColor = &H8000000E
AUTOMATEXLHOME.AppIcon3.ForeColor = &H8000000E
Exit Function

'//Load Recording
ElseIf xHov = 3 Then
AUTOMATEXLHOME.LoadMappingBtn.ForeColor = RGB(185, 231, 170)
AUTOMATEXLHOME.LoadMappingBtn.Font.Size = 18
AUTOMATEXLHOME.NewMappingBtn.ForeColor = &HE0E0E0
AUTOMATEXLHOME.PlayBtn.ForeColor = &HE0E0E0
AUTOMATEXLHOME.ExitBtn.ForeColor = &HE0E0E0
AUTOMATEXLHOME.AppIcon1.ForeColor = &H8000000E
AUTOMATEXLHOME.AppIcon2.ForeColor = &H8000000E
AUTOMATEXLHOME.AppIcon3.ForeColor = &H8000000E
Exit Function

'//Exit
ElseIf xHov = 4 Then
AUTOMATEXLHOME.ExitBtn.ForeColor = RGB(185, 231, 170)
AUTOMATEXLHOME.ExitBtn.Font.Size = 18
AUTOMATEXLHOME.NewMappingBtn.ForeColor = &HE0E0E0
AUTOMATEXLHOME.PlayBtn.ForeColor = &HE0E0E0
AUTOMATEXLHOME.LoadMappingBtn.ForeColor = &HE0E0E0
AUTOMATEXLHOME.AppIcon1.ForeColor = &H8000000E
AUTOMATEXLHOME.AppIcon2.ForeColor = &H8000000E
AUTOMATEXLHOME.AppIcon3.ForeColor = &H8000000E
Exit Function

'//App Icon 1
ElseIf xHov = 5 Then
AUTOMATEXLHOME.AppIcon1.ForeColor = vbRed
AUTOMATEXLHOME.AppIcon2.ForeColor = &H8000000E
AUTOMATEXLHOME.AppIcon3.ForeColor = &H8000000E
Exit Function

'//App Icon 2
ElseIf xHov = 6 Then
AUTOMATEXLHOME.AppIcon2.ForeColor = vbRed
AUTOMATEXLHOME.AppIcon1.ForeColor = &H8000000E
AUTOMATEXLHOME.AppIcon3.ForeColor = &H8000000E
Exit Function


'//App Icon 3
ElseIf xHov = 7 Then
AUTOMATEXLHOME.AppIcon3.ForeColor = vbRed
AUTOMATEXLHOME.AppIcon1.ForeColor = &H8000000E
AUTOMATEXLHOME.AppIcon2.ForeColor = &H8000000E
Exit Function

End If

End Function
Sub dfsHover()

On Error Resume Next

'//Default color...
AUTOMATEXLHOME.NewMappingBtn.ForeColor = &HE0E0E0
AUTOMATEXLHOME.PlayBtn.ForeColor = &HE0E0E0
AUTOMATEXLHOME.LoadMappingBtn.ForeColor = &HE0E0E0
AUTOMATEXLHOME.ExitBtn.ForeColor = &HE0E0E0

'//Default font size...
AUTOMATEXLHOME.NewMappingBtn.Font.Size = 16
AUTOMATEXLHOME.PlayBtn.Font.Size = 16
AUTOMATEXLHOME.LoadMappingBtn.Font.Size = 16
AUTOMATEXLHOME.ExitBtn.Font.Size = 16

'//Remove underline...
AUTOMATEXLHOME.PlayBtn.Font.Underline = False
AUTOMATEXLHOME.NewMappingBtn.Font.Underline = False
AUTOMATEXLHOME.LoadMappingBtn.Font.Underline = False
AUTOMATEXLHOME.ExitBtn.Font.Underline = False

End Sub
Function undHover(ByVal xHov As Byte)

On Error Resume Next

'//Play Recording
If xHov = 1 Then
AUTOMATEXLHOME.PlayBtn.Font.Underline = True
AUTOMATEXLHOME.NewMappingBtn.Font.Underline = False
AUTOMATEXLHOME.LoadMappingBtn.Font.Underline = False
AUTOMATEXLHOME.ExitBtn.Font.Underline = False
Exit Function
End If

'//New Recording
If xHov = 2 Then
AUTOMATEXLHOME.NewMappingBtn.Font.Underline = True
AUTOMATEXLHOME.PlayBtn.Font.Underline = False
AUTOMATEXLHOME.LoadMappingBtn.Font.Underline = False
AUTOMATEXLHOME.ExitBtn.Font.Underline = False
Exit Function
End If

'//Load Recording
If xHov = 3 Then
AUTOMATEXLHOME.LoadMappingBtn.Font.Underline = True
AUTOMATEXLHOME.PlayBtn.Font.Underline = False
AUTOMATEXLHOME.NewMappingBtn.Font.Underline = False
AUTOMATEXLHOME.ExitBtn.Font.Underline = False
Exit Function
End If

'//Exit
If xHov = 4 Then
AUTOMATEXLHOME.ExitBtn.Font.Underline = True
AUTOMATEXLHOME.PlayBtn.Font.Underline = False
AUTOMATEXLHOME.NewMappingBtn.Font.Underline = False
AUTOMATEXLHOME.LoadMappingBtn.Font.Underline = False
Exit Function
End If

End Function
Function fxsActive(ByVal xBtn As Byte)

'//Left click active
If xBtn = 1 Then
XLMAPPERCTRLR.AddMapLBtn.BackStyle = fmBackStyleOpaque
XLMAPPERCTRLR.AddMapLBtn.BackColor = RGB(185, 231, 170)
XLMAPPERCTRLR.AddMapRBtn.BackStyle = fmBackStyleTransparent
XLMAPPERCTRLR.AddMapDblBtn.BackStyle = fmBackStyleTransparent
XLMAPPERCTRLR.KeyFlowBtn.BackStyle = fmBackStyleTransparent
Exit Function

'//Right click active
ElseIf xBtn = 2 Then
XLMAPPERCTRLR.AddMapRBtn.BackStyle = fmBackStyleOpaque
XLMAPPERCTRLR.AddMapRBtn.BackColor = RGB(185, 231, 170)
XLMAPPERCTRLR.AddMapLBtn.BackStyle = fmBackStyleTransparent
XLMAPPERCTRLR.AddMapDblBtn.BackStyle = fmBackStyleTransparent
XLMAPPERCTRLR.KeyFlowBtn.BackStyle = fmBackStyleTransparent
Exit Function

'//Double click active
ElseIf xBtn = 3 Then
XLMAPPERCTRLR.AddMapDblBtn.BackStyle = fmBackStyleOpaque
XLMAPPERCTRLR.AddMapDblBtn.BackColor = RGB(185, 231, 170)
XLMAPPERCTRLR.AddMapLBtn.BackStyle = fmBackStyleTransparent
XLMAPPERCTRLR.AddMapRBtn.BackStyle = fmBackStyleTransparent
XLMAPPERCTRLR.KeyFlowBtn.BackStyle = fmBackStyleTransparent
Exit Function

'//Key flow active
ElseIf xBtn = 4 Then
XLMAPPERCTRLR.KeyFlowBtn.BackStyle = fmBackStyleOpaque
XLMAPPERCTRLR.KeyFlowBtn.BackColor = RGB(185, 231, 170)
XLMAPPERCTRLR.AddMapLBtn.BackStyle = fmBackStyleTransparent
XLMAPPERCTRLR.AddMapRBtn.BackStyle = fmBackStyleTransparent
XLMAPPERCTRLR.AddMapDblBtn.BackStyle = fmBackStyleTransparent
Exit Function

End If

End Function

