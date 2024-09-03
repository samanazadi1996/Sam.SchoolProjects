VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Saman Azadi"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text 
      Height          =   495
      Left            =   3000
      TabIndex        =   10
      Text            =   "0"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Height          =   495
      Index           =   4
      Left            =   120
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "s"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "o"
      Height          =   495
      Left            =   720
      TabIndex        =   7
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   495
      Left            =   960
      Picture         =   "Form1.frx":069B
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      Picture         =   "Form1.frx":0F65
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   5
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Height          =   495
      Index           =   3
      Left            =   120
      Picture         =   "Form1.frx":3459
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Height          =   495
      Index           =   2
      Left            =   120
      Picture         =   "Form1.frx":3B28
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Height          =   495
      Index           =   1
      Left            =   120
      Picture         =   "Form1.frx":5F30
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Height          =   495
      Index           =   0
      Left            =   120
      Picture         =   "Form1.frx":8424
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   495
   End
   Begin MSComDlg.CommonDialog k 
      Left            =   5160
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4950
      Left            =   840
      ScaleHeight     =   4950
      ScaleWidth      =   4950
      TabIndex        =   0
      Top             =   720
      Width           =   4950
      Begin VB.Image Image1 
         Height          =   495
         Left            =   0
         Picture         =   "Form1.frx":A2E5
         Stretch         =   -1  'True
         Top             =   4455
         Width           =   495
      End
   End
   Begin VB.Label Label1 
      Caption         =   "My Bombs"
      Height          =   495
      Left            =   1320
      TabIndex        =   11
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
On Error Resume Next
k.Filter = "Game|*.SamanGame"
k.ShowOpen
If k.FileName <> "" Then
Text1 = ""
r = FreeFile
Open k.FileName For Input As #r
Line Input #r, o
Text1 = Text1 + o
Line Input #r, o1
Text = o1
Close #r
For e = 0 To 90 Step 10
For w = 0 To 9
u = Mid(Text1, e + w + 1, 1)
Picture1.PaintPicture Command3(u).Picture, w * 495, (e / 10) * 495
Next
Next
End If
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click(Index As Integer)
Picture2.Picture = Command3(Index).Picture
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command5_Click()
Text1 = ""
For w = 0 To 9
For e = 0 To 9
For r = 0 To 4
Picture4.Picture = Command3(r).Picture
If Picture1.Point(100 + e * 495, 100 + w * 495) = Picture4.Point(100, 100) Then Text1 = Text1 + Str(r)
Next
Next
Text1 = Text1
Next
Text1 = Replace(Text1, " ", "")
k.Filter = "Game|*.SamanGame"
k.ShowSave
If k.FileName <> "" Then
r = FreeFile
Open k.FileName For Output As #r
Print #r, Text1
Print #r, Text
Close #r
End If
End Sub

Private Sub Command6_Click()

End Sub




Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
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
Picture1.PaintPicture Picture2.Picture, X, Y, 495, 495
End If
End Sub


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1_MouseDown Button, Shift, X, Y
End Sub

