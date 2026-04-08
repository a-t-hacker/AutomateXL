VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AUTOMATEXLHOME 
   Caption         =   "AutomateXL"
   ClientHeight    =   9615.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9960.001
   OleObjectBlob   =   "AUTOMATEXLHOME.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AUTOMATEXLHOME"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()

Call getEnvironment(appEnv, appBlk)

'//WinForm #
xWin = 18: Call setWindow(xWin)

AppNameLbl.Visible = False
PlayBtn.Visible = False
NewMappingBtn.Visible = False
ExitBtn.Visible = False
AppBuildLbl.Visible = False

Call fxsLoading

Me.Caption = AppTag
ThisWorkbook.Worksheets("Main").Range("MapperActive").Value = 0

End Sub
Private Sub NewMappingBtn_Click()

ThisWorkbook.Worksheets("Main").Range("MapperActive").Value = 1
Unload Me
Call NewMappingBtn_Clk

End Sub
Private Sub PlayBtn_Click()

Call PlayBtn_Clk
AUTOMATEXLHOME.Show

End Sub
Private Sub LoadMappingBtn_Click()

Call LoadMappingBtn_Clk(xPath)

End Sub
Private Sub ExitBtn_Click()

Me.Hide

End Sub
Private Sub HomeBg1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Call dfsHover

End Sub
Private Sub PlayBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xHov = 1: Call fxsHover(xHov): Call undHover(xHov)

End Sub
Private Sub NewMappingBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xHov = 2: Call fxsHover(xHov): Call undHover(xHov)

End Sub
Private Sub LoadMappingBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xHov = 3: Call fxsHover(xHov): Call undHover(xHov)

End Sub
Private Sub ExitBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xHov = 4: Call fxsHover(xHov): Call undHover(xHov)

End Sub
Private Sub AppIcon1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xHov = 5: Call fxsHover(xHov)

End Sub
Private Sub AppIcon2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xHov = 6: Call fxsHover(xHov)

End Sub
Private Sub AppIcon3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xHov = 7: Call fxsHover(xHov)

End Sub
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If KeyCode.Value = 27 Then Me.Hide

End Sub

Private Sub UserForm_Terminate()

If ThisWorkbook.Worksheets("Main").Range("MapperActive").Value <> 1 Then
ThisWorkbook.Save
ThisWorkbook.Close
End If

End Sub
