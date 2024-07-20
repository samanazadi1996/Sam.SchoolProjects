VERSION 5.00
Begin VB.Form Form7 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   4455
   ClientLeft      =   2460
   ClientTop       =   945
   ClientWidth     =   5580
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   48
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form7.frx":0442
   ScaleHeight     =   4455
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   375
      Left            =   4620
      MouseIcon       =   "Form7.frx":764F
      MousePointer    =   99  'Custom
      Picture         =   "Form7.frx":A899
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   375
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Sub Form_DblClick()
SavePicture Me.Picture, "C:\Documents and Settings\All Users\Desktop\Saman Azadi.gif"
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub
Private Sub Form_Load()
DoTransparency Me, BackColor
MDIForm1.Enabled = False
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage Me.hWnd, &HA1, 2, 0&
End Sub
Private Sub Form_Unload(Cancel As Integer)
MDIForm1.Enabled = True
End Sub
Public Sub DoTransparency(Frm As Form, transColor)
Dim rgn     As Long
Dim rgn2    As Long
Dim rgn3    As Long
Dim rgn4    As Long
rgn = CreateRectRgn(0, 0, 0, 0)
rgn2 = CreateRectRgn(0, 0, 0, 0)
rgn3 = CreateRectRgn(0, 0, 0, 0)
i = 1
With Frm
     X1 = .Width / Screen.TwipsPerPixelX
     Y1 = .Height / Screen.TwipsPerPixelY
     .AutoRedraw = True
     .ScaleMode = 3
End With

Do While i < X1
    j = 1
    Do While j < Y1
        If GetPixel(Frm.hDC, i, j) <> transColor Then
            tj = j
            Do While GetPixel(Frm.hDC, i, j + 1) <> transColor
                j = j + 1
                If j = Y1 Then Exit Do
            Loop
            rgn4 = CreateRectRgn(i, tj, i + 1, j + 1)
                    CombineRgn rgn3, rgn2, rgn2, 5
            CombineRgn rgn2, rgn4, rgn3, 2
            DeleteObject rgn4
        End If
    j = j + 1
    Loop
    CombineRgn rgn3, rgn, rgn, 5
    CombineRgn rgn, rgn2, rgn3, 2
    i = i + 1
Loop
SetWindowRgn Me.hWnd, rgn, True
End Sub
Private Sub Image1_Click()
Unload Me
End Sub
