Attribute VB_Name = "Module1"
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Sub formload()
Form1.Label1.Left = Screen.Height / 3 + 1500
Form1.Label2.Left = Screen.Height / 3 + 1500
Form1.Label3.Left = Screen.Height / 3 + 1500
Form1.ss.Height = Screen.Height
Form1.ss.Width = Screen.Width
Form1.ss.BorderStyle = 0

End Sub


