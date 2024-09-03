VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox s 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF80FF&
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   2760
      ScaleHeight     =   6465
      ScaleWidth      =   4065
      TabIndex        =   0
      Top             =   600
      Width           =   4095
      Begin VB.PictureBox t 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF80FF&
         BorderStyle     =   0  'None
         Height          =   720
         Left            =   360
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   2
         Top             =   2040
         Width           =   720
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   120
         Top             =   120
      End
      Begin VB.Label a 
         BackColor       =   &H00FF0000&
         Height          =   255
         Left            =   1200
         TabIndex        =   1
         Top             =   6120
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim W As Integer, D As Integer
Private Sub Form_Click()
End
End Sub
Private Sub Form_Load()
W = Screen.Width
h = Screen.Height
s.Move W / 2 - (W / 4), 0, W / 2, h
a.Move s.Width / 2 - a.Width / 2, h - a.Height
W = -100
D = 100
End Sub
Private Sub s_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Timer1.Enabled = False Then t.Move s.Width / 2, s.Height / 2
Timer1.Enabled = True
End Sub
Private Sub s_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
a.Left = X - a.Width / 2
End Sub
Private Sub Timer1_Timer()
If t.Left <= 0 Then D = -1 * D
If t.Left >= s.Width - t.Width Then D = -1 * D
If t.Top <= 0 Then W = -1 * W
If t.Top + t.Height >= a.Top - 120 And t.Left - t.Width > a.Left - 120 And t.Left < a.Left + a.Width Then
W = -(Abs(W))
End If





If t.Top >= a.Top Then Timer1.Enabled = False
t.Left = t.Left + D
t.Top = t.Top + W
End Sub
