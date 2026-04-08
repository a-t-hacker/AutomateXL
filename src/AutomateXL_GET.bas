Attribute VB_Name = "AutomateXL_GET"
Public Function getData()

XLMAPPER.AppPath.Caption = ThisWorkbook.Worksheets("Main").Range("MapperPath").Value2
XLMAPPER.CtrlrPosBox.Caption = _
ThisWorkbook.Worksheets("Main").Range("MapperX").Value2 & ", " & _
ThisWorkbook.Worksheets("Main").Range("MapperY").Value2
XLMAPPER.KeyFlowBox.Caption = ThisWorkbook.Worksheets("Main").Range("xlasKeyCtrl").Value
If ThisWorkbook.Worksheets("Main").Range("LastMap").Value2 <> vbNullString Then
XLMAPPER.LastMapBox.Caption = Replace(ThisWorkbook.Worksheets("Main").Range("LastMap").Value, "[,]", ", ")
XLMAPPER.LastMapBox.Caption = Replace(XLMAPPER.LastMapBox.Caption, "[(]", vbNullString)
XLMAPPER.LastMapBox.Caption = Replace(XLMAPPER.LastMapBox.Caption, "[)]", vbNullString)
    Else
    If ThisWorkbook.Worksheets("Main").Range("MapCount").Value2 >= 1 Then
    XLMAPPER.LastMapBox.Caption = ThisWorkbook.Worksheets("Main").Range("MapperXY").Offset(ThisWorkbook.Worksheets("Main").Range("MapCount").Value2, 0).Value
    XLMAPPER.LastMapBox.Caption = Replace(XLMAPPER.LastMapBox.Caption, "[,]", ", ")
    XLMAPPER.LastMapBox.Caption = Replace(XLMAPPER.LastMapBox.Caption, "[(]", vbNullString)
    XLMAPPER.LastMapBox.Caption = Replace(XLMAPPER.LastMapBox.Caption, "[)]", vbNullString)
    XLMAPPER.LastMapBox.Caption = "(" & XLMAPPER.LastMapBox.Caption & ") (" & ThisWorkbook.Worksheets("Main").Range("ClickType").Offset(ThisWorkbook.Worksheets("Main").Range("MapCount").Value2, 0).Value & ")"
        End If
            End If

XLMAPPER.MapCtLbl.Caption = ThisWorkbook.Worksheets("Main").Range("MapCount").Value2
If XLMAPPER.xlFlowStrip.Text = vbNullString Then XLMAPPER.xlFlowStrip.Text = "Enter xlAppScript code here..."

End Function
Public Function getWindowPos()

Dim X As Integer: Dim Y As Integer
X = (XLMAPPERCTRLR.Left + (XLMAPPERCTRLR.Left / 100) * 33) + 10
Y = (XLMAPPERCTRLR.Top + (XLMAPPERCTRLR.Top / 100) * 36) + 25

ThisWorkbook.Worksheets("Main").Range("MapperX").Value = X
ThisWorkbook.Worksheets("Main").Range("MapperY").Value = Y
Call getData

End Function

