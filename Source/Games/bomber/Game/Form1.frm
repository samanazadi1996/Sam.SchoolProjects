VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Game Saman"
   ClientHeight    =   6870
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   5460
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer 
      Interval        =   50
      Left            =   3000
      Top             =   840
   End
   Begin VB.PictureBox esc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   720
      ScaleHeight     =   4305
      ScaleWidth      =   3345
      TabIndex        =   13
      Top             =   1080
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFAA00&
         Height          =   495
         Index           =   4
         Left            =   480
         MouseIcon       =   "Form1.frx":0A8A
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Refresh Game"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFAA00&
         Height          =   495
         Index           =   3
         Left            =   480
         MouseIcon       =   "Form1.frx":3CD4
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Save Game"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFAA00&
         Height          =   495
         Index           =   2
         Left            =   480
         MouseIcon       =   "Form1.frx":6F1E
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "New Game"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFAA00&
         Height          =   495
         Index           =   1
         Left            =   480
         MouseIcon       =   "Form1.frx":A168
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Goto Game"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFAA00&
         Height          =   495
         Index           =   0
         Left            =   480
         MouseIcon       =   "Form1.frx":D3B2
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.PictureBox SamanAzadi 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   0
      ScaleHeight     =   6435
      ScaleWidth      =   5235
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   5295
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4950
         Left            =   120
         ScaleHeight     =   4950
         ScaleWidth      =   4950
         TabIndex        =   10
         Top             =   720
         Width           =   4950
         Begin VB.Timer Timerup 
            Enabled         =   0   'False
            Interval        =   200
            Left            =   3240
            Top             =   720
         End
         Begin VB.Timer Timerl 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   2760
            Top             =   1200
         End
         Begin VB.Timer Timerr 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   3720
            Top             =   1200
         End
         Begin VB.Timer Timerd 
            Interval        =   100
            Left            =   3240
            Top             =   1200
         End
         Begin VB.Image q 
            Appearance      =   0  'Flat
            Height          =   495
            Left            =   0
            Picture         =   "Form1.frx":105FC
            Stretch         =   -1  'True
            Top             =   4455
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture4 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5280
         Picture         =   "Form1.frx":10C06
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   9
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   6720
         Picture         =   "Form1.frx":114D0
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   8
         Top             =   840
         Width           =   480
      End
      Begin VB.PictureBox A1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   2160
         Picture         =   "Form1.frx":11912
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   7
         ToolTipText     =   "„Ì Ê«‰Ìœ »— —ÊÌ «Ì‰ ò›ÅÊ‘ Â« —«Â »—ÊÌœ"
         Top             =   120
         Width           =   495
      End
      Begin VB.PictureBox A1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   2760
         Picture         =   "Form1.frx":12638
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   6
         ToolTipText     =   "›ﬁ  „Ì Ê«‰Ìœ —ÊÌ «Ì‰ œÌÊ«— Â« —«Â »—ÊÌœ Ê ÂÌç Êﬁ   Œ—Ì» ‰„Ì‘Ê‰œ"
         Top             =   120
         Width           =   495
      End
      Begin VB.PictureBox A1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   3360
         Picture         =   "Form1.frx":1335E
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   5
         ToolTipText     =   "«Ì‰ œÌÊ«— Â« —« „Ì Ê«‰ »« »„» ‰«»Êœ ò—œ"
         Top             =   120
         Width           =   495
      End
      Begin VB.PictureBox A1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   3960
         Picture         =   "Form1.frx":14084
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   4
         ToolTipText     =   "„Õ«›Ÿ «·„«” Â«"
         Top             =   120
         Width           =   495
      End
      Begin VB.PictureBox A1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   4560
         Picture         =   "Form1.frx":14753
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   3
         ToolTipText     =   "”⁄Ì ò‰Ìœ »Â «Ì‰ «·„«” »—”Ìœ  « »Â „—Õ·Â »⁄œ œ”  ÅÌœ« ò‰Ìœ"
         Top             =   120
         Width           =   495
      End
      Begin VB.PictureBox nn 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   2
         Top             =   2040
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFAA00&
         Caption         =   "Bomb"
         Height          =   495
         Left            =   120
         MouseIcon       =   "Form1.frx":14DEE
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   5760
         Width           =   4935
      End
      Begin VB.Label Text0 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   1440
         TabIndex        =   12
         ToolTipText     =   "My Bombs"
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   1215
      End
      Begin VB.Image Imager 
         Height          =   495
         Left            =   1080
         Picture         =   "Form1.frx":18038
         Stretch         =   -1  'True
         Top             =   7200
         Width           =   495
      End
      Begin VB.Image Imagel 
         Height          =   495
         Left            =   2040
         Picture         =   "Form1.frx":18642
         Stretch         =   -1  'True
         Top             =   7320
         Width           =   495
      End
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2820
      Left            =   2760
      TabIndex        =   19
      Top             =   3960
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox Tex 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   20
      Top             =   2400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image llll 
      Height          =   500
      Left            =   3240
      Picture         =   "Form1.frx":1895D
      Stretch         =   -1  'True
      Top             =   2040
      Visible         =   0   'False
      Width           =   500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s As Integer
Dim l As Integer
Dim a As Integer
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Dim eee As Integer
Private Sub Command1_Click()
On Error Resume Next
Dim u As Boolean
u = False
For e = 0 To 9
For w = 0 To 9
X = w * 495
Y = e * 495
nn.PaintPicture Picture1.Image, 0, 0, 495, 495, X, Y, 495, 495
If nn.Point(200, 200) = Picture2.Point(200, 200) And eee < Picture2.ToolTipText Then
Picture1.PaintPicture A1(0).Image, X, Y
nn.PaintPicture Picture1.Image, 0, 0, 495, 495, X - 495, Y, 495, 495
u = True
eee = eee + 1
If nn.Point(200, 200) = A1(2).Point(200, 200) Or nn.Point(200, 200) = A1(3).Point(200, 200) Then Picture1.PaintPicture A1(0), X - 495, Y
nn.PaintPicture Picture1.Image, 0, 0, 495, 495, X, Y - 495, 495, 495
If nn.Point(200, 200) = A1(2).Point(200, 200) Or nn.Point(200, 200) = A1(3).Point(200, 200) Then Picture1.PaintPicture A1(0), X, Y - 495
nn.PaintPicture Picture1.Image, 0, 0, 495, 495, X + 495, Y, 495, 495
If nn.Point(200, 200) = A1(2).Point(200, 200) Or nn.Point(200, 200) = A1(3).Point(200, 200) Then Picture1.PaintPicture A1(0), X + 495, Y
nn.PaintPicture Picture1.Image, 0, 0, 495, 495, X, Y + 495, 495, 495
If nn.Point(200, 200) = A1(2).Point(200, 200) Or nn.Point(200, 200) = A1(3).Point(200, 200) Then Picture1.PaintPicture A1(0), X, Y + 495
End If
If Picture1.Point(q.Left + 120, q.Top + q.Height) = vbWhite Or Picture1.Point(q.Left + 360, q.Top + q.Height) = vbWhite Then Timerd.Enabled = True
Next
Next
Randomize Timer
uuu = Int(Rnd * 2)
If u = True Then sndPlaySound App.Path & "\Sounds\bomb " & Int(uuu) & ".wav", 1
Text0 = Picture2.ToolTipText - eee
End Sub
Private Sub esc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For w = 0 To 4
Label1(w).FontBold = False
Label1(w).FontUnderline = False
Next
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If esc.Visible = False Then
Select Case KeyCode
Case vbKeyUp
If Timerup.Enabled = False And Timerd.Enabled = False Then
Timerup.Enabled = True
s = q.Top
End If
Case vbKeyLeft
Timerl.Enabled = True
Timerr.Enabled = False
q.Picture = Imagel.Picture
Case vbKeyRight
Timerr.Enabled = True
Timerl.Enabled = False
q.Picture = Imager.Picture
Case 13
Command1_Click
End Select
End If
If KeyCode = vbKeyEscape And Timer.Enabled = False Then ssss
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyLeft
Timerl.Enabled = False
Case vbKeyRight
Timerr.Enabled = False
End Select
End Sub
Private Sub Form_Load()
SamanAzadi.Move Screen.Width / 2 - SamanAzadi.Width / 2, Screen.Height / 2 - SamanAzadi.Height / 2
esc.Move Screen.Width / 2 - esc.Width / 2, Screen.Height / 2 - esc.Height / 2
If App.PrevInstance = True Then
MsgBox " »—‰«„Â œ—Õ«· «Ã—« »ÊœÂ Ê «„ò«‰ «Ã—«Ì Â„“„«‰ ¬‰ ÊÃÊœ ‰œ«—œ ", vbCritical, "Warning !"
End
End If
File1.Path = App.Path & "\Saman_Games"
Tex = "Loading... " & vbNewLine & "Reading Save" & vbNewLine & "Number Levels : " & File1.ListCount & vbNewLine & "Programer : Saman Azadi" & vbNewLine & "Date of made : [2013.10.25] [1392.8.13]" & vbNewLine & "E_Mail : www.saman.com" & vbNewLine & "Goto game...      "
End Sub
Private Sub Form_Resize()
PaintPicture llll.Picture, 0, 0, Screen.Width, Screen.Height
End Sub
Private Sub Label1_Click(Index As Integer)
Select Case Index
Case 1
f = MsgBox("New Game", vbYesNo + 32, "New")
If f = vbYes Then
l = 0
aaa
End If
Case 2
f = MsgBox("Save Game", vbYesNo + 32, "Save")
If f = vbYes Then SaveSetting "Game", "Saman", "Saman", l - 1
Case 3
f = MsgBox("Refresh Game", vbYesNo + 32, "Refresh")
If f = vbYes Then
eee = 0
Timerup.Enabled = False
Timerd.Enabled = False
Timerl.Enabled = False
Timerr.Enabled = False
q.Move 0, Picture1.ScaleHeight - 495
p = App.Path & "\Saman_Games\Game " & l - 1 & ".SamanGame"
r = FreeFile
Open p For Input As #r
Line Input #r, o
t11 = t11 + o
Line Input #r, o1
Picture2.ToolTipText = o1
Text0 = Picture2.ToolTipText
Close #r
For e = 0 To 90 Step 10
For w = 0 To 9
u = Mid(t11, e + w + 1, 1)
Picture1.PaintPicture A1(u).Picture, w * 495, (e / 10) * 495
Next
Next
End If
Case 4
f = MsgBox("Exit Game", vbYesNo + 32, "Exit")
If f = vbYes Then End
End Select
esc.Visible = False
SamanAzadi.Enabled = True
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
For w = 0 To 4
Label1(w).FontBold = False
Label1(w).FontUnderline = False
Next
Label1(Index).FontBold = True
Label1(Index).FontUnderline = True
End Sub

Private Sub nn_Click()
MsgBox nn.Point(100, 100) & vbNewLine & A1(4).Point(100, 100)
End Sub
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
For w = 1 To 495
If X Mod 495 = 0 Then GoTo 1
X = X - 1
Next
1:
For w = 1 To 495
If Y Mod 495 = 0 Then GoTo 2
Y = Y - 1
Next
2:
If Button = 1 Then
If Picture1.Point(X + 200, Y + 200) = vbWhite Then Picture1.PaintPicture Picture2.Picture, X, Y, 495, 495
ElseIf Button = 2 Then
If Picture1.Point(X + 100, Y + 100) = A1(0).Point(100, 100) Then Picture1.PaintPicture A1(0).Picture, X, Y, 495, 495
End If
End Sub
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1_MouseDown Button, Shift, X, Y
End Sub



Private Sub Timer_Timer()
Static yyy As Integer
yyy = yyy + 1
Print Mid(Tex, yyy, 1);
If yyy >= Len(Tex) Then
Timer.Enabled = False
Cls
PaintPicture llll.Picture, 0, 0, Screen.Width, Screen.Height
SamanAzadi.PaintPicture llll.Picture, -(Screen.Width / 2 - SamanAzadi.Width / 2), -(Screen.Height / 2 - SamanAzadi.Height / 2), Screen.Width, Screen.Height

SamanAzadi.Visible = True
l = GetSetting("Game", "Saman", "Saman", 0)
aaa

End If
End Sub

Private Sub Timerd_Timer()
If Picture1.Point(q.Left + 450, q.Top + 495) <> vbWhite Or Picture1.Point(q.Left, q.Top + 495) <> vbWhite Or Picture1.Point(q.Left + 240, q.Top + 495) <> vbWhite Then
Timerd.Enabled = False
Else
q.Top = q.Top + 495
End If
aa
End Sub
Private Sub Timerl_Timer()
If q.Left - 495 >= 0 And Picture1.Point(q.Left - 495, q.Top) = vbWhite And Picture1.Point(q.Left - 495, q.Top + 450) = vbWhite Then
q.Left = q.Left - 495
Timerup.Enabled = False
a = 0
Else
Timerr.Enabled = False
End If
If Picture1.Point(q.Left + 120, q.Top + q.Height) = vbWhite Or Picture1.Point(q.Left + 360, q.Top + q.Height) = vbWhite Then Timerd.Enabled = True
aa
End Sub
Private Sub Timerr_Timer()
If q.Left + 495 <= Picture1.ScaleWidth And Picture1.Point(q.Left + 495, q.Top + 450) = vbWhite And Picture1.Point(q.Left + 495, q.Top) = vbWhite Then
q.Left = q.Left + 495
Timerup.Enabled = False
a = 0
Else
Timerr.Enabled = False
End If
If Picture1.Point(q.Left + 120, q.Top + q.Height) = vbWhite Or Picture1.Point(q.Left + 360, q.Top + q.Height) = vbWhite Then Timerd.Enabled = True
aa
End Sub
Private Sub Timerup_Timer()
If a >= 495 * 2 Or Picture1.Point(q.Left, q.Top - 495) <> vbWhite Or Picture1.Point(q.Left + 450, q.Top - 495) <> vbWhite Then
Timerd.Enabled = True
Timerup.Enabled = False
a = 0
Else
a = a + 495
q.Top = s - a
End If
aa
End Sub
Public Sub aa()
Dim Saman As Boolean
Dim SamanAzadi As Boolean
Saman = False
SamanAzadi = False

On Error Resume Next
nn.PaintPicture Picture1.Image, 0, 0, , , q.Left + 495, q.Top, 495, 495
If nn.Point(200, 200) = A1(4).Point(200, 200) Then Saman = True
If nn.Point(200, 200) = A1(3).Point(200, 200) Then SamanAzadi = True

nn.PaintPicture Picture1.Image, 0, 0, , , q.Left - 495, q.Top, 495, 495
If nn.Point(200, 200) = A1(4).Point(200, 200) Then Saman = True
If nn.Point(200, 200) = A1(3).Point(200, 200) Then SamanAzadi = True

nn.PaintPicture Picture1.Image, 0, 0, , , q.Left, q.Top + 495, 495, 495
If nn.Point(200, 200) = A1(4).Point(200, 200) Then Saman = True
If nn.Point(200, 200) = A1(3).Point(200, 200) Then SamanAzadi = True

nn.PaintPicture Picture1.Image, 0, 0, , , q.Left, q.Top - 495, 495, 495
If nn.Point(200, 200) = A1(4).Point(200, 200) Then Saman = True
If nn.Point(200, 200) = A1(3).Point(200, 200) Then SamanAzadi = True

If Saman = True Then
MsgBox "You win" & vbNewLine & "Next level", , "Win"
aaa
End If
If SamanAzadi = True Then Label1_Click (3)
End Sub
Public Sub aaa()
On Error GoTo 1
SSSSSSS:
If l + 1 < File1.ListCount Then
Text = "„—Õ·Â " & l + 1
eee = 0
Timerup.Enabled = False
Timerd.Enabled = False
Timerl.Enabled = False
Timerr.Enabled = False
q.Move 0, Picture1.ScaleHeight - 495
p = App.Path & "\Saman_Games\Game " & l & ".SamanGame"
r = FreeFile
Open p For Input As #r
Line Input #r, o
t11 = t11 + o
Line Input #r, o1
Picture2.ToolTipText = o1
Text0 = Picture2.ToolTipText
Close #r
For e = 0 To 90 Step 10
For w = 0 To 9
u = Mid(t11, e + w + 1, 1)
Picture1.PaintPicture A1(u).Picture, w * 495, (e / 10) * 495
Next
Next
l = l + 1
Else
l = 0
GoTo SSSSSSS:
End If
1:
End Sub

Public Sub ssss()
esc.Visible = Not esc.Visible
SamanAzadi.Enabled = Not esc.Visible
esc.PaintPicture llll.Picture, -esc.Left, -esc.Top, Screen.Width, Screen.Height
End Sub
