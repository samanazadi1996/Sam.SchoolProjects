VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "0"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6120
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   20.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":164A
   ScaleHeight     =   6360
   ScaleWidth      =   6120
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2640
      Top             =   4200
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Height          =   6615
      Index           =   3
      Left            =   5760
      TabIndex        =   7
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   6000
      Width           =   6015
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   0
      Width           =   5895
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Height          =   6615
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   -240
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      ForeColor       =   &H00FFFF00&
      Height          =   1200
      Left            =   -5760
      TabIndex        =   3
      Top             =   4560
      Width           =   6240
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Help?"
      ForeColor       =   &H00FFFF00&
      Height          =   1200
      Left            =   -5760
      TabIndex        =   2
      Top             =   3240
      Width           =   6240
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select Map"
      ForeColor       =   &H00FFFF00&
      Height          =   1200
      Left            =   -5760
      TabIndex        =   1
      Top             =   1920
      Width           =   6240
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Play"
      ForeColor       =   &H00FFFF00&
      Height          =   1200
      Left            =   -5760
      TabIndex        =   0
      Top             =   600
      Width           =   6240
   End
   Begin VB.Image j 
      Height          =   255
      Left            =   2760
      Picture         =   "Form1.frx":34DC
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xx As Single, yy As Single, y1(0 To 100000) As Single, x1(0 To 100000) As Single
Dim t As Integer
Dim t1 As Integer
Dim o As Integer
Dim r As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Module1.vlabel
If KeyCode = vbKeyR Then Form1.Cls
Form_Load
End Sub

Private Sub Form_Load()
PSet (2500, 4500), vbWhite
Print "S"
PSet (1600, 360), vbWhite
Print "p"
Module1.load
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X > 2500 And X < 3000 And Y > 4500 And Y < 5000 Then
Form1.Caption = Form1.Caption + 1
xx = X
yy = Y
r = r + 1
Form1.Caption = r
Caption = 0
o = 1
t = 0
End If
If Button = 2 Then o = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If o = 1 Then
If Form1.Point(X, Y) = vbRed Then Form1.Caption = Form1.Caption + 1
t = t + 1
y1(t) = Y
x1(t) = X
Line (xx, yy)-(X, Y)
xx = X
yy = Y
If Button = 2 Then o = 0



End If
End Sub





Private Sub j_Click()
Timer1.Enabled = True
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If o = 1 Then
Form1.Caption = Form1.Caption + 5
o = 0
End If
End Sub

Private Sub Label2_Click()
Module1.vlabel
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Module1.label2mm
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Module1.label3mm
End Sub

Private Sub Label3_Click()
Form2.Show
Unload Me
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Module1.label4mm
End Sub

Private Sub Label5_Click()
Unload Me
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Module1.label5mm
End Sub

Private Sub Timer1_Timer()

If t1 < t Then
t1 = t1 + 1
j.Move x1(t1) - (j.Width / 2), y1(t1) - (j.Height / 2)
Else
Timer1.Enabled = False
t1 = 0
If j.Top > 500 And j.Top < 750 And j.Left > 1400 And j.Left < 1700 Then
If Form1.Caption <= "3" Then
Module1.aks
Else
Unload Me
Form1.Show
End If
End If
End If
End Sub

Public Sub load()
PSet (2500, 4500), vbWhite
Print "S"
PSet (1600, 360), vbWhite
Print "p"
Module1.load
End Sub


Private Sub Timer2_Timer()

End Sub
