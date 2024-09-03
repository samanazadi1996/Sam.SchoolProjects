VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   8280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9960
   DrawWidth       =   3
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   8280
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1920
      Top             =   3360
   End
   Begin VB.Image i 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   3
      Left            =   120
      Picture         =   "Form1.frx":0948
      Stretch         =   -1  'True
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image i 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   2
      Left            =   120
      Picture         =   "Form1.frx":163C
      Stretch         =   -1  'True
      Top             =   2520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image i 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   1
      Left            =   120
      Picture         =   "Form1.frx":211B
      Stretch         =   -1  'True
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   3000
      Picture         =   "Form1.frx":2B8E
      Stretch         =   -1  'True
      Top             =   240
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   2190
      Picture         =   "Form1.frx":2C12
      Stretch         =   -1  'True
      Top             =   480
      Width           =   405
   End
   Begin VB.Line Line5 
      BorderWidth     =   5
      X1              =   3360
      X2              =   3360
      Y1              =   840
      Y2              =   1560
   End
   Begin VB.Line Line4 
      BorderWidth     =   5
      X1              =   3120
      X2              =   3120
      Y1              =   840
      Y2              =   1560
   End
   Begin VB.Line Line3 
      BorderWidth     =   5
      X1              =   2880
      X2              =   2880
      Y1              =   840
      Y2              =   1560
   End
   Begin VB.Line Line2 
      BorderWidth     =   5
      X1              =   2640
      X2              =   2640
      Y1              =   840
      Y2              =   1560
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   2400
      X2              =   2400
      Y1              =   840
      Y2              =   1560
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   5
      Height          =   495
      Left            =   120
      Shape           =   2  'Oval
      Top             =   120
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   4
      Height          =   495
      Left            =   1320
      Shape           =   3  'Circle
      Top             =   120
      Width           =   495
   End
   Begin VB.Line s3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   3840
      X2              =   4080
      Y1              =   1200
      Y2              =   1440
   End
   Begin VB.Line s2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   3480
      X2              =   3840
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line s1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   3840
      X2              =   3960
      Y1              =   1200
      Y2              =   840
   End
   Begin VB.Line s5 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      X1              =   4320
      X2              =   4560
      Y1              =   1200
      Y2              =   1440
   End
   Begin VB.Line s6 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      X1              =   4320
      X2              =   4560
      Y1              =   1200
      Y2              =   840
   End
   Begin VB.Line s4 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      X1              =   3960
      X2              =   4320
      Y1              =   1200
      Y2              =   1200
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer, e As Integer, c As Integer, p As Single
Private Sub Form_DblClick()
End
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyRight Then Timer1.Enabled = True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyRight Then Timer1.Enabled = False
End Sub
Private Sub Form_Load()
e = 1200
Timer1_Timer
End Sub
Private Sub Timer1_Timer()
e = e + 60
For w = 4000 To Height - 1000 Step 20
If Point(e, w) <> BackColor Then Shape1.Move e - Shape1.Width / 2, w - Shape1.Height
If Point(e - 1000, w) <> BackColor Then Shape2.Move (e - Shape1.Width / 2) - 1000, w - Shape1.Height
Next
If e > Width - 500 Then
e = 1100
c = c + 1
Picture = i(c).Picture
End If
Line1.X1 = Shape1.Left + 240
Line1.Y1 = Shape1.Top + 240
Line1.X2 = Shape2.Left + 240
Line1.Y2 = Shape2.Top + 240

Line2.X1 = Shape2.Left + 1000
Line2.Y1 = Shape1.Top - 400
Line2.X2 = Shape1.Left + 240
Line2.Y2 = Shape1.Top + 240

Line3.X1 = Line2.X1
Line3.Y1 = Line2.Y1










Line4.X1 = Line1.X2 + 500
Line4.Y1 = Line1.Y1 - (Line1.Y1 - Line1.Y2) / 2
Line4.Y2 = (Line1.Y1 - (Line1.Y1 - Line1.Y2) / 2) - 500
Line4.X2 = Shape2.Left + 500
Line3.X2 = Line4.X2
Line3.Y2 = Line4.Y2
Line5.X1 = Line4.X2
Line5.Y1 = Line4.Y2
Line5.X2 = Line1.X2
Line5.Y2 = Line1.Y2
Image1.Move Line5.X1 - 200, Line5.Y1 - 240
Image2.Move Line2.X1 - 300, Line2.Y1 - 240


p = p - 3.14159265 / 15
s1.X2 = Line1.X2
s1.Y2 = Line1.Y2
s1.X1 = s1.X2 + Sin(p) * 200
s1.Y1 = s1.Y2 + Cos(p) * 200
s2.X2 = Line1.X2
s2.Y2 = Line1.Y2
s2.X1 = s2.X2 + Sin(p + 2.30383461) * 200
s2.Y1 = s2.Y2 + Cos(p + 2.30383461) * 200
s3.X2 = Line1.X2
s3.Y2 = Line1.Y2
s3.X1 = s3.X2 + Sin(p + 4.39822971) * 200
s3.Y1 = s3.Y2 + Cos(p + 4.39822971) * 200



s4.X2 = Line1.X1
s4.Y2 = Line1.Y1
s4.X1 = s4.X2 + Sin(p) * 200
s4.Y1 = s4.Y2 + Cos(p) * 200
s5.X2 = Line1.X1
s5.Y2 = Line1.Y1
s5.X1 = s5.X2 + Sin(p + 2.30383461) * 200
s5.Y1 = s5.Y2 + Cos(p + 2.30383461) * 200
s6.X2 = Line1.X1
s6.Y2 = Line1.Y1
s6.X1 = s6.X2 + Sin(p + 4.39822971) * 200
s6.Y1 = s6.Y2 + Cos(p + 4.39822971) * 200

End Sub
