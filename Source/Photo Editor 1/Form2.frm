VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6225
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":0A8A
   ScaleHeight     =   3435
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   4800
      Top             =   1080
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Dim p As Integer
Private Sub Form_Load()
                SetWindowLong Me.hWnd, -20, &H80000
    SetLayeredWindowAttributes Me.hWnd, 0, 180, &H2
    p = 256
End Sub



Private Sub Timer1_Timer()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Timer2_Timer()
p = p - 1
                SetWindowLong Me.hWnd, -20, &H80000
    SetLayeredWindowAttributes Me.hWnd, 0, p, &H2
    If p = 1 Then
    Unload Me
    Form4.Show
    End If

End Sub
