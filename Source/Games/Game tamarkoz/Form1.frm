VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "My Game"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   6015
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":324A
   ScaleHeight     =   8445
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3960
      Top             =   3840
   End
   Begin VB.PictureBox a 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2640
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   4
      Top             =   3600
      Width           =   495
   End
   Begin VB.PictureBox w 
      BackColor       =   &H00FF80FF&
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   2
      Left            =   1320
      ScaleHeight     =   1455
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   4080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox w 
      BackColor       =   &H00FF80FF&
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   3
      Left            =   3840
      ScaleHeight     =   855
      ScaleWidth      =   735
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox w 
      BackColor       =   &H00FF80FF&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   960
      ScaleHeight     =   495
      ScaleWidth      =   1455
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox w 
      BackColor       =   &H00FF80FF&
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   0
      Left            =   2400
      ScaleHeight     =   855
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2400
      Top             =   3240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   405
      Left            =   960
      TabIndex        =   6
      Top             =   120
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
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
      Height          =   405
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   660
   End
   Begin VB.Menu Game 
      Caption         =   "Game"
      Begin VB.Menu ItemNew 
         Caption         =   "NewGame"
         Shortcut        =   ^N
      End
      Begin VB.Menu q3 
         Caption         =   "-"
      End
      Begin VB.Menu ItemExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu ItemAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xx As Integer, yy As Integer
Dim X1(0 To 3) As Boolean, Y1(0 To 3) As Boolean
Private Sub a_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xx = X
yy = Y
Timer1.Enabled = True
Timer2.Enabled = True
End Sub

Private Sub a_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then a.Move a.Left + X - xx, a.Top + Y - yy

End Sub

Private Sub Form_Load()
aaaa
Randomize Timer
xxx
End Sub

Private Sub Itemhlp_Click()
Form2.Show
End Sub

Private Sub ItemAbout_Click()
Form2.Show
End Sub

Private Sub ItemExit_Click()
End
End Sub

Private Sub ItemNew_Click()
p = MsgBox("NewGame", vbYesNo, "New")
If p = vbYes Then theend
End Sub


Private Sub Timer1_Timer()
For e = 0 To 3
If X1(e) = True Then w(e).Left = w(e).Left + 45
If Y1(e) = True Then w(e).Top = w(e).Top + 45
If X1(e) = False Then w(e).Left = w(e).Left - 45
If Y1(e) = False Then w(e).Top = w(e).Top - 45
If Point(w(e).Left, w(e).Top) = vbBlue Then Y1(e) = Not (Y1(e))
If Point(w(e).Left, w(e).Top) = vbBlue + 1 Then X1(e) = Not (X1(e))
If Point(w(e).Left + w(e).Width, w(e).Top + w(e).Height) = vbBlue + 2 Then X1(e) = Not (X1(e))
If Point(w(e).Left + w(e).Width, w(e).Top + w(e).Height) = vbBlue + 3 Then Y1(e) = Not (Y1(e))
Next
aaaa
If Point(a.Left, a.Top) <> BackColor Or Point(a.Left + a.Width, a.Top + a.Height) <> BackColor Or Point(a.Left + a.Width, a.Top) <> BackColor Or Point(a.Left, a.Top + a.Height) <> BackColor Then
MsgBox "##########" + vbNewLine + "#   You   Lost   # " + vbNewLine + "##########" + vbNewLine & "Time is :" & Label2 + vbNewLine + "New Game", vbOKOnly, "Lost"

theend
End If
End Sub

Public Sub aaaa()
Cls
'Line (0, 0)-(ScaleWidth, 700), vbBlue, BF
'Line (0, 0)-(700, ScaleHeight), vbBlue + 1, BF
'Line (ScaleWidth, 0)-(ScaleWidth - 700, ScaleHeight), vbBlue + 2, BF
'Line (0, ScaleHeight)-(ScaleWidth, ScaleHeight - 700), vbBlue + 3, BF
For e = 0 To 3
Line (w(e).Left, w(e).Top)-(w(e).Left + w(e).Width, w(e).Top + w(e).Height), w(0).BackColor, BF
Next
End Sub

Private Sub Timer2_Timer()
Label2 = Label2 + 1
End Sub

Public Sub xxx()
For e1 = 0 To Rnd * 5
For e = 0 To Rnd * (3)
X1(e) = Not X1(e)
Next
Next
For e1 = 0 To Rnd * 5
For e = 0 To Rnd * 3
Y1(e) = Not Y1(e)
Next
Next
End Sub

Public Sub theend()
a.Move ScaleWidth / 2 - a.Height / 2, ScaleHeight / 2 - a.Height / 2
Timer1.Enabled = False
Timer2.Enabled = False
Label2 = 0
w(1).Move 960, 1200
w(3).Move 3840, 2280
w(2).Move 1320, 4080
w(0).Move 2400, 4680
xxx
aaaa
End Sub
