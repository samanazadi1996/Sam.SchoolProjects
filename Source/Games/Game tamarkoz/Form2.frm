VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   5280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5775
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":324A
   ScaleHeight     =   5280
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   375
      Left            =   4830
      MouseIcon       =   "Form2.frx":A569
      MousePointer    =   99  'Custom
      Picture         =   "Form2.frx":D7B3
      Stretch         =   -1  'True
      Top             =   4280
      Width           =   375
   End
End
Attribute VB_Name = "Form2"
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
Form1.Enabled = False
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage Me.hWnd, &HA1, 2, 0&
End Sub
Private Sub Form_Unload(Cancel As Integer)
Form1.Enabled = True
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

