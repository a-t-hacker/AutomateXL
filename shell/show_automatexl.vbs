Option Explicit

Dim objXL, objWb

On Error Resume Next

Set objXL = CreateObject("Excel.Application")

Set objWb = objXL.Workbooks.Open("C:\Users\EDITHERE\.xlas\autokit\automatexl\app\AutomateXL.xlsm")

wscript.Sleep 10

objWb.Save
objWb.Close
Set objWb = Nothing
Set objXL = Nothing

wscript.Quit