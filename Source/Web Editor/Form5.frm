VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Table"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   Icon            =   "Form5.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3840
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   120
      ScaleHeight     =   1905
      ScaleWidth      =   4305
      TabIndex        =   7
      Top             =   1680
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox T 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   2
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "2"
      Top             =   720
      Width           =   3135
   End
   Begin VB.TextBox T 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Index           =   1
      Left            =   1320
      Locked          =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "2"
      Top             =   1200
      Width           =   3135
   End
   Begin VB.TextBox T 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "2"
      Top             =   240
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   960
      Stretch         =   -1  'True
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Seton"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Satr"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Border"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.Menu aa 
      Caption         =   "a"
      Visible         =   0   'False
      Begin VB.Menu ItemBackcolor 
         Caption         =   "BackColor"
      End
      Begin VB.Menu ItemBorderColor 
         Caption         =   "BorderColor"
      End
      Begin VB.Menu ItemBackGround 
         Caption         =   "BackGround"
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BackColor1 As String, BorderColor1 As String
Private Sub Command1_Click()
If BackColor1 = "" Then BackColor1 = vbWhite
If BorderColor1 = "" Then BorderColor1 = vbBlack
MDIForm1.ActiveForm.Text1.SelText = "<Table Border=" & T(0) & " BgColor=" & Hex(BackColor1) & " BorderColor=" & Hex(BorderColor1)
If Image1.ToolTipText <> "" Then MDIForm1.ActiveForm.Text1.SelText = " BackGround=" & Chr(34) & Image1.ToolTipText & Chr(34)
MDIForm1.ActiveForm.Text1.SelText = ">"
For w = 1 To T(2)
MDIForm1.ActiveForm.Text1.SelText = vbNewLine + "<tr>"
For e = 1 To T(1)
MDIForm1.ActiveForm.Text1.SelText = vbNewLine + "   <td>&nbsp;</td>"
Next
MDIForm1.ActiveForm.Text1.SelText = vbNewLine + "</tr>"
Next
MDIForm1.ActiveForm.Text1.SelText = vbNewLine + "</Tabel>"
Unload Me
End Sub
Private Sub Command2_Click()
Unload Me
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Call Command1_Click
If KeyCode = 27 Then Call Command2_Click
End Sub
Private Sub Form_Load()
oo
End Sub
Private Sub Form_Unload(Cancel As Integer)
MDIForm1.Enabled = True
End Sub

Private Sub ItemBackcolor_Click()
CommonDialog1.ShowColor

BackColor1 = CommonDialog1.Color
Image1.ToolTipText = ""
Image1.Picture = LoadPicture()
oo
End Sub

Private Sub ItemBackGround_Click()
CommonDialog1.Filter = "AllPicture|*.jpg;*.gif;*.png;*.bmp;*.jpeg|AllFiles|*.*"
CommonDialog1.ShowOpen
Image1.Picture = LoadPicture(CommonDialog1.FileName)
Image1.ToolTipText = CommonDialog1.FileName
oo

End Sub

Private Sub ItemBorderColor_Click()
CommonDialog1.ShowColor
BorderColor1 = CommonDialog1.Color
oo
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu aa
End Sub

Private Sub T_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then T(Index) = T(Index) + 1
If KeyCode = vbKeyDown And T(Index) > 1 Then T(Index) = T(Index) - 1
oo
End Sub
Private Sub oo()
On Error Resume Next

Picture1.Cls
Picture1.DrawWidth = T(0)
Picture1.Line (60, 60)-((T(1) * 120) + 60, (T(2) * 120) + 60), BackColor1, BF
Picture1.ForeColor = BorderColor1
Picture1.PaintPicture Image1.Picture, 60, 60, T(1) * 120 + 60, T(2) * 120 + 60

s1 = 60
s2 = 60
For A1 = 1 To T(1)
Picture1.Line (s1, 60)-(s1, (T(2) * 120) + 60)
s1 = s1 + 120
Next
For d1 = 1 To T(2)
Picture1.Line (60, s2)-((T(1) * 120) + 60, s2)
s2 = s2 + 120
Next
Picture1.Line (s1, 60)-(s1, (T(2) * 120) + 60)
Picture1.Line (60, s2)-((T(1) * 120) + 60, s2)
1:
End Sub
