VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10515
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   11520
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timerko 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   3480
      Top             =   240
   End
   Begin VB.Timer Timerhameh 
      Interval        =   1
      Left            =   0
      Top             =   960
   End
   Begin VB.Timer Timeradam 
      Interval        =   20
      Left            =   3000
      Top             =   240
   End
   Begin VB.Timer timerpok 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1920
      Top             =   240
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   1440
      Top             =   240
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   2160
      TabIndex        =   6
      Top             =   4920
      Visible         =   0   'False
      Width           =   1095
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   20
         Height          =   255
         Left            =   480
         Shape           =   2  'Oval
         Top             =   240
         Width           =   255
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   600
         X2              =   120
         Y1              =   960
         Y2              =   1800
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   600
         X2              =   960
         Y1              =   960
         Y2              =   1800
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   600
         X2              =   600
         Y1              =   2040
         Y2              =   480
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   600
         X2              =   600
         Y1              =   2040
         Y2              =   3240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   600
         X2              =   600
         Y1              =   2040
         Y2              =   3240
      End
      Begin VB.Image Imagep 
         Height          =   11520
         Left            =   720
         MouseIcon       =   "Form1.frx":2987E
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":2A148
         Top             =   2880
         Width           =   15360
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1215
      Left            =   -2280
      TabIndex        =   2
      Top             =   4920
      Width           =   3000
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   -2160
      TabIndex        =   1
      Top             =   3480
      Width           =   3000
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   -2400
      TabIndex        =   0
      Top             =   2280
      Width           =   3000
   End
   Begin VB.Label lkh 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "120"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   960
      TabIndex        =   4
      Top             =   10800
      Width           =   1335
   End
   Begin VB.Label ltir 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   2400
      TabIndex        =   3
      Top             =   10800
      Width           =   1215
   End
   Begin VB.Label ss 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   0
      MouseIcon       =   "Form1.frx":539C6
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   0
      Width           =   735
   End
   Begin VB.Image TOF 
      Height          =   4455
      Left            =   4920
      Picture         =   "Form1.frx":54290
      Stretch         =   -1  'True
      Top             =   7080
      Visible         =   0   'False
      Width           =   8535
   End
   Begin VB.Image ba 
      Height          =   975
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   6840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image jtir 
      Height          =   210
      Index           =   0
      Left            =   120
      Picture         =   "Form1.frx":6248C
      Top             =   120
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image pok 
      Height          =   810
      Left            =   120
      Picture         =   "Form1.frx":626CD
      Top             =   360
      Visible         =   0   'False
      Width           =   1050
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer, t As Single, u As Single, X As Single, o As Single



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Label1.Visible = Not Label1.Visible
Label2.Visible = Label1.Visible
Label3.Visible = Label1.Visible
Frame1.Visible = Not Label1.Visible
TOF.Visible = Not TOF.Visible
    If Label1.Visible = True Then
    Form1.MousePointer = 0
    Else
    Form1.MousePointer = 99
    End If
    End If
    If KeyCode = vbKeyR Then Timer1.Enabled = True

End Sub

Private Sub Form_Load()
Module1.formload
For w = 1 To 20
u = u + 3.14159265 / 20
Next
o = 50
End Sub

Private Sub Imagep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ltir.Caption > 0 Then
If Label1.Visible = False Then
If Button = 1 Then
ltir = ltir - 1
owe = sndPlaySound(App.Path & "\1.wav", 1)


timerpok.Enabled = True
pok.Visible = True

End If
End If
End If
If ltir = 0 Then Timer1.Enabled = True
Timeradam.Enabled = False
Timerko.Enabled = True

End Sub

Private Sub Label1_Click()
Label1.Visible = Not Label1.Visible
Label2.Visible = Label1.Visible
Label3.Visible = Label1.Visible
TOF.Visible = Not TOF.Visible
pok.Visible = TOF.Visible
Frame1.Visible = TOF.Visible

    Form1.MousePointer = 99
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = vbCyan
Label2.ForeColor = vbBlue
Label3.ForeColor = vbBlue
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = vbCyan
Label1.ForeColor = vbBlue
Label3.ForeColor = vbBlue

End Sub

Private Sub Label3_Click()
End
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = vbCyan
Label2.ForeColor = vbBlue
Label1.ForeColor = vbBlue

End Sub

Private Sub ss_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ltir.Caption > 0 Then
If Label1.Visible = False Then
If Button = 1 Then
owe = sndPlaySound(App.Path & "\1.wav", 1)


i = i + 1
Load jtir(i)
jtir(i).Visible = True
jtir(i).Move X - 105, Y - 112
ltir = ltir - 1
timerpok.Enabled = True
pok.Visible = True
End If
End If
End If
If ltir = 0 Then Timer1.Enabled = True

End Sub

Private Sub ss_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = vbBlue
Label2.ForeColor = vbBlue
Label3.ForeColor = vbBlue
If Label1.Visible = False Then TOF.Move X + 2000, Y / 3 + Me.Height - 4500
pok.Move TOF.Left + 4440, TOF.Top + 1680
End Sub

Private Sub Timer1_Timer()

lkh = Int(lkh) + Int(ltir)
lkh = lkh - 10
ltir = 10
If lkh < 0 Then
ltir = ltir + lkh
End If
If lkh <= 0 And ltir <= 0 Then
MsgBox ""
End
End If
Timer1.Enabled = False
    Form1.Enabled = True

End Sub

Private Sub Timeradam_Timer()
u = u + 3.14159265 / 20
Line1.X2 = Line1.X1 + Cos(u) * 500
X = X + 3.14159265 / 20
Line2.X2 = Line2.X1 + Cos(X) * 500
Frame1.Left = Frame1.Left + o
Line4.X2 = Line4.X1 + Cos(u) * 900
Line5.X2 = Line5.X1 + Cos(X) * 1000
Imagep.Left = -Frame1.Left
Imagep.Top = -Frame1.Top
End Sub


Private Sub Timerba_Timer()


End Sub

Private Sub Timerhameh_Timer()
If Frame1.Left < -2000 Then
o = 50
ElseIf Frame1.Left > Screen.Width + 2000 Then
o = -50
End If
End Sub

Private Sub Timerko_Timer()
Randomize Timer
p = sndPlaySound(App.Path & "\s" & Int(Rnd * 3) & ".wav", 1)
For w = 1 To 200
PSet (Rnd * 1000 + Frame1.Left, Rnd * 2000 + (Frame1.Top + 500)), vbRed
Next
Dim sss As Integer
sss = Rnd
If sss = 1 Then Frame1.Left = -3000
If sss = 0 Then Frame1.Left = Screen.Width + 3000
Timeradam.Enabled = True
Timerko.Enabled = False
End Sub

Private Sub timerpok_Timer()
t = t - 3.14159265 / 15
Form1.pok.Top = (Form1.TOF.Top + Form1.TOF.Height - 1000) + Cos(t) * 2000
Form1.pok.Left = (Form1.TOF.Left + Form1.TOF.Width - 2000) + Sin(t) * 2000
If t < -5 Then
timerpok.Enabled = False
pok.Visible = False
t = -2.51327412
End If
End Sub

