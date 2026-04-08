VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Activate()

If Range("xlasAppLoad").Value <> 1 Then
'//show loading window
XLLOADING_AUTOMATEXL.Show
End If

End Sub
Private Sub Workbook_Open()

On Error GoTo LinkMe

If Range("xlasLink").Value <> 1 Then Call connectWb: Exit Sub

Exit Sub

LinkMe: Call connectWb

End Sub
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

On Error Resume Next

'//reset application load trigger
Range("xlasAppLoad").Value = 0

End Sub


