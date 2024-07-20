VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saman Azadi"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6765
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   6765
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1800
      Top             =   1320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   201
      Visible         =   0   'False
      X1              =   600
      X2              =   615
      Y1              =   1680
      Y2              =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pp(0 To 200) As Integer
Dim xxx As Integer, yyy As Integer

Private Sub Form_Load()
Randomize Timer
For w = 0 To 200
Load Line1(w)
Line1(w).Visible = True
Line1(w).Y1 = Rnd * 5000
Line1(w).X1 = Rnd * 6820
Line1(w).X2 = Line1(w).X1
Line1(w).Y2 = Line1(w).Y1
pp(w) = 6300
Next
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Enabled = False
xxx = X
yyy = Y
DrawWidth = 2
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Line (X, Y)-(xxx, yyy), vbRed
xxx = X
yyy = Y
ElseIf Button = 2 Then
DrawWidth = 5
Line (X, Y)-(xxx, yyy), vbBlack
xxx = X
yyy = Y
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()
Randomize Timer
For w = 0 To 200
If Line1(w) < 5500 Then
Line1(w).Y1 = Line1(w).Y1 + (Rnd * 30) + 3
Line1(w).X1 = Line1(w).X1 + (Rnd * 10) - 5
Line1(w).X2 = Line1(w).X1
Line1(w).Y2 = Line1(w).Y1
End If
Next
For q = 0 To 200
If Line1(q).Y1 >= pp(q) Then
Line1(q).Y1 = Rnd * 1000
Line1(q).X2 = Line1(q).X1
Line1(q).Y2 = Line1(q).Y1
X = Line1(q).X1
Form1.DrawWidth = Line1(q).BorderWidth
pp(q) = pp(q) - Line1(q).BorderWidth * 10
Form1.PSet (X, pp(q))
ElseIf Point(Line1(q).X1, Line1(q).Y1) = vbRed Or Point(Line1(q).X1, Line1(q).Y1) = vbWhite Then
Form1.DrawWidth = Line1(q).BorderWidth
Form1.PSet (Line1(q).X1, Line1(q).Y1 - 30)
Line1(q).Y1 = Rnd * 1000
Line1(q).X2 = Line1(q).X1
Line1(q).Y2 = Line1(q).Y1




End If

Next
End Sub

