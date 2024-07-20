VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5190
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form8"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   255
      ItemData        =   "Form8.frx":0442
      Left            =   5040
      List            =   "Form8.frx":056F
      TabIndex        =   1
      Top             =   5400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Text            =   "&nbsp;"
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   98
      Left            =   4635
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   97
      Left            =   4170
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   96
      Left            =   3720
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   95
      Left            =   3270
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   94
      Left            =   2835
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   93
      Left            =   2370
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   92
      Left            =   1920
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   91
      Left            =   1470
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   90
      Left            =   1035
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   89
      Left            =   570
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   88
      Left            =   120
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   87
      Left            =   4635
      Top             =   3285
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   86
      Left            =   4170
      Top             =   3285
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   85
      Left            =   3720
      Top             =   3285
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   84
      Left            =   3270
      Top             =   3285
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   83
      Left            =   2835
      Top             =   3285
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   82
      Left            =   2370
      Top             =   3285
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   81
      Left            =   1920
      Top             =   3285
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   80
      Left            =   1470
      Top             =   3285
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   79
      Left            =   1035
      Top             =   3285
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   78
      Left            =   570
      Top             =   3285
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   77
      Left            =   120
      Top             =   3285
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   76
      Left            =   4635
      Top             =   2820
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   75
      Left            =   4170
      Top             =   2820
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   74
      Left            =   3720
      Top             =   2820
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   73
      Left            =   3270
      Top             =   2820
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   72
      Left            =   2835
      Top             =   2820
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   71
      Left            =   2370
      Top             =   2820
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   70
      Left            =   1920
      Top             =   2820
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   69
      Left            =   1470
      Top             =   2820
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   68
      Left            =   1035
      Top             =   2820
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   67
      Left            =   570
      Top             =   2820
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   66
      Left            =   120
      Top             =   2820
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   65
      Left            =   4635
      Top             =   2400
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   64
      Left            =   4170
      Top             =   2400
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   63
      Left            =   3720
      Top             =   2400
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   62
      Left            =   3270
      Top             =   2400
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   61
      Left            =   2835
      Top             =   2400
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   60
      Left            =   2370
      Top             =   2400
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   59
      Left            =   1920
      Top             =   2400
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   58
      Left            =   1470
      Top             =   2400
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   57
      Left            =   1035
      Top             =   2400
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   56
      Left            =   570
      Top             =   2400
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   55
      Left            =   120
      Top             =   2400
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   54
      Left            =   4635
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   53
      Left            =   4170
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   52
      Left            =   3720
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   51
      Left            =   3270
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   50
      Left            =   2835
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   49
      Left            =   2370
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   48
      Left            =   1920
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   47
      Left            =   1470
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   46
      Left            =   1035
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   45
      Left            =   570
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   44
      Left            =   120
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   43
      Left            =   4635
      Top             =   1485
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   42
      Left            =   4170
      Top             =   1485
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   41
      Left            =   3720
      Top             =   1485
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   40
      Left            =   3270
      Top             =   1485
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   39
      Left            =   2835
      Top             =   1485
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   38
      Left            =   2370
      Top             =   1485
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   37
      Left            =   1920
      Top             =   1485
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   36
      Left            =   1470
      Top             =   1485
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   35
      Left            =   1035
      Top             =   1485
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   34
      Left            =   570
      Top             =   1485
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   33
      Left            =   120
      Top             =   1485
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   32
      Left            =   4635
      Top             =   1020
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   31
      Left            =   4170
      Top             =   1020
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   30
      Left            =   3720
      Top             =   1020
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   29
      Left            =   3270
      Top             =   1020
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   28
      Left            =   2835
      Top             =   1020
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   27
      Left            =   2370
      Top             =   1020
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   26
      Left            =   1920
      Top             =   1020
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   25
      Left            =   1470
      Top             =   1020
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   24
      Left            =   1035
      Top             =   1020
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   23
      Left            =   570
      Top             =   1020
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   22
      Left            =   120
      Top             =   1020
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   21
      Left            =   4635
      Top             =   585
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   20
      Left            =   4170
      Top             =   585
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   19
      Left            =   3720
      Top             =   585
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   18
      Left            =   3270
      Top             =   585
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   17
      Left            =   2835
      Top             =   585
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   16
      Left            =   2370
      Top             =   585
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   15
      Left            =   1920
      Top             =   585
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   14
      Left            =   1470
      Top             =   585
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   13
      Left            =   1035
      Top             =   585
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   12
      Left            =   570
      Top             =   585
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   11
      Left            =   120
      Top             =   585
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   10
      Left            =   4635
      Top             =   120
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   9
      Left            =   4170
      Top             =   120
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   8
      Left            =   3720
      Top             =   120
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   7
      Left            =   3270
      Top             =   120
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   6
      Left            =   2835
      Top             =   120
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   5
      Left            =   2370
      Top             =   120
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   4
      Left            =   1920
      Top             =   120
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   3
      Left            =   1470
      Top             =   120
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   2
      Left            =   1035
      Top             =   120
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   1
      Left            =   570
      Top             =   120
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   3990
      Left            =   120
      Picture         =   "Form8.frx":08F0
      Top             =   120
      Width           =   4890
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MDIForm1.ActiveForm.Text1.SelText = Text1
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
MDIForm1.Enabled = True
End Sub

Private Sub Image2_Click(Index As Integer)
Text1 = List1.List(Index)
End Sub

Private Sub Image2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2(Index).BorderStyle = 1
End Sub

Private Sub Image2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2(Index).BorderStyle = 0
End Sub
