VERSION 5.00
Begin VB.Form FormA 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6150
   Icon            =   "FormA.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "FormA.frx":324A
   MousePointer    =   99  'Custom
   ScaleHeight     =   4560
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   0
      ScaleHeight     =   360
      ScaleWidth      =   6150
      TabIndex        =   0
      Top             =   0
      Width           =   6150
      Begin VB.CommandButton Command1 
         BackColor       =   &H000000FF&
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5760
         MouseIcon       =   "FormA.frx":3B14
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   30
         Width           =   300
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "&About"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Left            =   0
         MouseIcon       =   "FormA.frx":6D5E
         TabIndex        =   1
         Top             =   0
         Width           =   6135
      End
      Begin VB.Image Image15 
         Height          =   345
         Left            =   0
         Picture         =   "FormA.frx":7628
         Stretch         =   -1  'True
         Top             =   0
         Width           =   6615
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail : WWW.       @        .com"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   450
      Index           =   2
      Left            =   300
      MouseIcon       =   "FormA.frx":7722
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   3720
      Width           =   5550
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date of made: 1392/2/24"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   450
      Index           =   1
      Left            =   960
      TabIndex        =   4
      Top             =   2160
      Width           =   4080
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Programer : Saman Azadi"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   450
      Index           =   0
      Left            =   1080
      TabIndex        =   3
      Top             =   600
      Width           =   3945
   End
   Begin VB.Image Image10 
      Height          =   210
      Left            =   0
      Picture         =   "FormA.frx":A96C
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   6735
   End
End
Attribute VB_Name = "FormA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
Private Sub Form_Load()
Form1.Enabled = False
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For w = 0 To Label2.Count - 1
Label2(w).FontUnderline = False
Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
Form1.Enabled = True
End Sub
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.MousePointer = 99
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Module1.S(Me.hWnd)
End Sub
Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.MousePointer = 0
End Sub
Private Sub Label2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(1, vbCtrlMask, 1000, 1000)
Label2(Index).FontUnderline = True
End Sub
