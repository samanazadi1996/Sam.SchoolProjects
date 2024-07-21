Attribute VB_Name = "Moduleother"
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpa As String, nSize As Long) As Long
Public Sub usename()
Dim username As String
username = String(255, 0)
GetUserName username, 255
username = Left(username, InStr(username, Chr(0)) - 1)
Form1.lblusername.Caption = username
Form1.Height = 5445 - 375
Form1.Width = 5805
End Sub
Public Sub computername()
Dim s2 As String
s2 = String(255, 0)
GetComputerName s2, 255
s2 = Left(s2, InStr(s2, Chr(0)) - 1)
Form1.textipmsgbox = s2
End Sub
Public Sub tz()
With Form1
.CommonDialogtz.Filter = "AllPhoto|*.jpg;*.gif;*.bmp)"
.CommonDialogtz.ShowOpen
If .CommonDialogtz.FileName <> "" Then
.Textpicfile = .CommonDialogtz.FileName
.Imagetz.Picture = LoadPicture(.CommonDialogtz.FileName)
End If
End With
End Sub

Public Sub allpro()
If Form1.PRONOTEPAD.AutoRedraw = True Then Form1.PRONOTEPAD.Visible = Not Form1.PRONOTEPAD.Visible
If Form1.proruner.AutoRedraw = True Then Form1.proruner.Visible = Not Form1.proruner.Visible
If Form1.pro1.AutoRedraw = True Then Form1.pro1.Visible = Not Form1.pro1.Visible
If Form1.Pro2.AutoRedraw = True Then Form1.Pro2.Visible = Not Form1.Pro2.Visible
If Form1.Pro3.AutoRedraw = True Then Form1.Pro3.Visible = Not Form1.Pro3.Visible
If Form1.pmsgbox.AutoRedraw = True Then Form1.pmsgbox.Visible = Not Form1.pmsgbox.Visible
If Form1.hlp.AutoRedraw = True Then Form1.hlp.Visible = Not Form1.hlp.Visible
End Sub

Public Sub tzs()
On Error Resume Next
With Form1
o = InputBox("Enter Path Files", "Open Folder")
.File1 = o
End With
End Sub
Public Sub MM()
Form1.Label1.BorderStyle = 0
End Sub
