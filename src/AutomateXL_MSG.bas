Attribute VB_Name = "AutomateXL_MSG"
'/###################################\
'//Application Error & Help Messages \\
'///#################################\\\

Function AppMsg(xMsg) As Integer

On Error Resume Next

Call CloseStrandedFiles

If ThisWorkbook.Worksheets("Main").Range("xlasSilent").Value2 <> 1 Then

'/1/Invalid information
If xMsg = 1 Then
'//msg
MsgBox ("Invalid information entered"), vbExclamation, AppTag
Exit Function

'/2/Mapping saved
ElseIf xMsg = 2 Then
'//msg
MsgBox ("New mapping saved: " & vbNewLine & vbNewLine & ThisWorkbook.Worksheets("Main").Range("MapperPath").Value2), vbInformation, AppTag
Exit Function

'/3/All mappings removed
ElseIf xMsg = 3 Then
'//msg
MsgBox ("All current mappings removed"), vbInformation, AppTag
Exit Function

'/4/Path missing (no mapping found)
ElseIf xMsg = 4 Then
'//msg
MsgBox ("No mapping found"), vbExclamation, AppTag
Exit Function

'/5/Key flow cleared
ElseIf xMsg = 5 Then
'//msg
MsgBox ("Key flow cleared"), vbInformation, AppTag
Exit Function

'/6/Mapping loaded successfully
ElseIf xMsg = 6 Then
'//msg
MsgBox ("Mapping loaded successfully"), vbInformation, AppTag
Exit Function

End If
    End If

End Function
