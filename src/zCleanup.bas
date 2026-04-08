Attribute VB_Name = "zCleanup"
'/####################\
'//Application Cleanup\\
'///##################\\\

Sub CloseStrandedFiles()

'//Close list of potentially opened files

Close #1
Close #2
Close #3
Close #4
Close #5
Close #6
Close #7

End Sub
Sub ClnMainSpace()

Dim lastR As Long

lastR = Cells(Rows.Count, "B").End(xlUp).Row

ThisWorkbook.Worksheets("Main").Range("A2:O" & lastR + 1000).ClearContents

End Sub

