VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Image"
   ClientHeight    =   4005
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   ScaleHeight     =   4005
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   405
      Left            =   1200
      TabIndex        =   4
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   405
      Left            =   1200
      TabIndex        =   3
      Top             =   2040
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3120
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Text            =   "Saman"
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   840
      Stretch         =   -1  'True
      ToolTipText     =   "Image"
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Height"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Width"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MDIForm1.ActiveForm.Text1.SelText = "<Img src=" & Chr(34) & CommonDialog1.FileName & Chr(34)
If Text1 <> "" Then MDIForm1.ActiveForm.Text1.SelText = " Name=" & Chr(34) & Text1 & Chr(34)
If Text2 <> "" Then MDIForm1.ActiveForm.Text1.SelText = " Width=" & Chr(34) & Text2 & Chr(34)
If Text3 <> "" Then MDIForm1.ActiveForm.Text1.SelText = " Height=" & Chr(34) & Text3 & Chr(34)
MDIForm1.ActiveForm.Text1.SelText = "/>"
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
Text2 = Screen.Width / 15
Text3 = Screen.Height / 15
End Sub
Private Sub Form_Unload(Cancel As Integer)
MDIForm1.Enabled = True
End Sub
Private Sub Image1_Click()
On Error Resume Next
CommonDialog1.FileName = ""
CommonDialog1.Filter = "AllPicture|*.jpg;*.gif;*.png;*.bmp;*.jpeg|AllFiles|*.*"
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then Image1.Picture = LoadPicture(CommonDialog1.FileName)
End Sub
Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyUp Or KeyCode = vbKeyRight Then Text2 = Text2 + 10
If KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Then Text2 = Text2 - 10
End Sub
Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyUp Or KeyCode = vbKeyRight Then Text3 = Text3 + 10
If KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Then Text3 = Text3 - 10
End Sub
