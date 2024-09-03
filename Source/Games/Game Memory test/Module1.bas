Attribute VB_Name = "Module1"
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Sub lo()
With Form1
For w = 0 To 39
If .Picture1(w).Visible = False Then p = p + 1
Next
.Label2 = "20 / " & p / 2
If p = 40 Then o = MsgBox("Time = " & .Label4 + vbNewLine + "Select = " & .Label3 + vbNewLine + "jkjhkjhk = " & .Label2 + vbNewLine + "New Game", vbYesNo + 64, "New Game")
If o = vbYes Then
Call Saman
Else
End If
End With
End Sub
Public Sub Saman()
Randomize Timer
Dim Y(0 To 39) As Integer, X(0 To 39) As Integer
With Form1
.Timer2.Enabled = False
For w = 0 To 39
.Image3.Picture = .Icon
.Image1(w).Move 0, 0, .Picture1(w).ScaleWidth, .Picture1(w).ScaleHeight
.Picture1(w).Visible = True
Y(w) = .Picture1(w).Top
X(w) = .Picture1(w).Left
.Label2 = "20 / 0"
.Label3 = "0"
.Label4 = "0"
.Image1(w).Visible = True
Next
For w = 0 To .Picture1.Count - 1
1:
o = Int(Rnd * .Picture1.Count)
If Y(o) = 0 Then GoTo 1
.Picture1(w).Top = Y(o)
.Picture1(w).Left = X(o)
Y(o) = 0
X(o) = 0
Next
End With
End Sub
Public Sub S(H As Long)
ReleaseCapture
SendMessage H, &HA1, 2, 0&
End Sub
