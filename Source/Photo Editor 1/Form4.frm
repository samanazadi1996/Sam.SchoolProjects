VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4425
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form4.frx":0A8A
   ScaleHeight     =   3600
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1920
      Top             =   2640
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3240
      Top             =   240
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   240
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1800
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1200
      Top             =   840
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   435
      Left            =   3840
      Picture         =   "Form4.frx":9845
      Stretch         =   -1  'True
      Top             =   120
      Width           =   450
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Photo Editor"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Conveter"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Dim p As Integer

Private Sub Form_Load()
SetWindowLong Me.hWnd, -20, &H80000
SetLayeredWindowAttributes Me.hWnd, 0, 1, &H2
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = vbWhite
Label2.ForeColor = vbWhite
Label3.ForeColor = vbWhite
Image1.BorderStyle = 0
ReleaseCapture
SendMessage Me.hWnd, &HA1, 2, 0


End Sub

Private Sub Image1_Click()
Timer4.Enabled = True
Timer2.Enabled = False
Timer3.Enabled = False
Timer1.Enabled = False
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.BorderStyle = 1
End Sub

Private Sub Label1_Click()
Timer3.Enabled = False
Timer1.Enabled = True
Me.Enabled = False
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = vbRed
Label2.ForeColor = vbWhite
Label3.ForeColor = vbWhite
End Sub

Private Sub Label2_Click()
Timer3.Enabled = False
Timer2.Enabled = True
Me.Enabled = False
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = vbWhite
Label3.ForeColor = vbWhite
Label2.ForeColor = vbRed
End Sub

Private Sub Label3_Click()
Timer3.Enabled = False
Timer5.Enabled = True
Me.Enabled = False
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = vbWhite
Label2.ForeColor = vbWhite
Label3.ForeColor = vbRed
End Sub

Private Sub Timer1_Timer()
p = p - 1
SetWindowLong Me.hWnd, -20, &H80000
SetLayeredWindowAttributes Me.hWnd, 0, p, &H2
If p <= 1 Then
Unload Me
Form3.Show
End If
End Sub

Private Sub Timer2_Timer()
p = p - 1
SetWindowLong Me.hWnd, -20, &H80000
SetLayeredWindowAttributes Me.hWnd, 0, p, &H2
If p <= 1 Then
Unload Me
Form1.Show
End If
End Sub

Private Sub Timer3_Timer()
p = p + 1
SetWindowLong Me.hWnd, -20, &H80000
SetLayeredWindowAttributes Me.hWnd, 0, p, &H2
If p >= 255 Then
Timer3.Enabled = False
End If
End Sub

Private Sub Timer4_Timer()
p = p - 1
SetWindowLong Me.hWnd, -20, &H80000
SetLayeredWindowAttributes Me.hWnd, 0, p, &H2
If p <= 1 Then
End
End If
End Sub

Private Sub Timer5_Timer()
p = p - 1
SetWindowLong Me.hWnd, -20, &H80000
SetLayeredWindowAttributes Me.hWnd, 0, p, &H2
If p <= 1 Then
Form5.Show
Unload Me
End If
End Sub
