Attribute VB_Name = "AutomateXL_INFO"
'/##################################\
'//Important Application Information\\
'///################################\\\

Public Function AppLoc()

AppLoc = ENV & AppPath

End Function
Public Function AppPath()

AppPath = "\.xlas\autokit\automatexl"

End Function
Public Function ENV()

ENV = Environ("USERPROFILE")

End Function
Public Function AppWelcome()

AppWelcome = "Welcome to AutomateXL..."

End Function
Public Function AppTag()

AppTag = "AutomateXL"

End Function
Public Function WbAppName() As String

Dim wbName As String

wbName = ThisWorkbook.name

If InStr(1, wbName, ".xlsm") Then
wbName = Replace(ThisWorkbook.name, ".xlsm", "")
End If

If InStr(1, wbName, ".xlsm") Then
wbName = Replace(ThisWorkbook.name, ".xlsb", "")
End If

WbAppName = wbName

End Function

