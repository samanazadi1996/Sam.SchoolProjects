VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "1"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4485
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1680
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   4320
      TabIndex        =   0
      Top             =   3960
      Visible         =   0   'False
      Width           =   3495
      Begin VB.Image Image5 
         Height          =   495
         Index           =   7
         Left            =   720
         Picture         =   "Form2.frx":164A
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   495
      End
      Begin VB.Image Image5 
         Height          =   495
         Index           =   6
         Left            =   120
         Picture         =   "Form2.frx":16DD2
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   495
      End
      Begin VB.Image Image5 
         Height          =   495
         Index           =   5
         Left            =   1320
         Picture         =   "Form2.frx":1B383
         Stretch         =   -1  'True
         Top             =   840
         Width           =   495
      End
      Begin VB.Image Image5 
         Height          =   495
         Index           =   4
         Left            =   720
         Picture         =   "Form2.frx":1DD72
         Stretch         =   -1  'True
         Top             =   840
         Width           =   495
      End
      Begin VB.Image Image5 
         Height          =   495
         Index           =   3
         Left            =   120
         Picture         =   "Form2.frx":1E73F
         Stretch         =   -1  'True
         Top             =   840
         Width           =   495
      End
      Begin VB.Image Image5 
         Height          =   495
         Index           =   2
         Left            =   1320
         Picture         =   "Form2.frx":1FAD2
         Stretch         =   -1  'True
         Top             =   240
         Width           =   495
      End
      Begin VB.Image Image5 
         Height          =   495
         Index           =   1
         Left            =   720
         Picture         =   "Form2.frx":21907
         Stretch         =   -1  'True
         Top             =   240
         Width           =   495
      End
      Begin VB.Image Image5 
         Height          =   495
         Index           =   0
         Left            =   120
         Picture         =   "Form2.frx":22A3A
         Stretch         =   -1  'True
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Image Image10 
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Left            =   3000
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Image Image9 
      Height          =   1215
      Left            =   1560
      Picture         =   "Form2.frx":248CC
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Image Image8 
      Height          =   1215
      Left            =   120
      Picture         =   "Form2.frx":3A054
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Image Image7 
      Height          =   1215
      Left            =   3000
      Picture         =   "Form2.frx":3E605
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Image Image6 
      Height          =   1215
      Left            =   1560
      Picture         =   "Form2.frx":40FF4
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   120
      Picture         =   "Form2.frx":419C1
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1335
   End
   Begin VB.Image Image4 
      Height          =   1215
      Left            =   120
      Picture         =   "Form2.frx":43853
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Image Image3 
      Height          =   1215
      Left            =   3000
      Picture         =   "Form2.frx":44BE6
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   1215
      Left            =   1560
      Picture         =   "Form2.frx":46A1B
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Select Picture"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1215
      Left            =   3000
      TabIndex        =   1
      Top             =   2760
      Width           =   1335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Unload(Cancel As Integer)
Form1.Show
Form1.load

End Sub

Private Sub image1_Click()
Form1.Picture = Image1.Picture
Unload Me
End Sub

Private Sub Image10_DblClick()

p = MsgBox(" ’ÊÌ— „Ê—œ ‰Ÿ— ŒÊœ «‰ —« «‰ ŒÊ«» ò‰Ìœ" + vbNewLine + "»—«Ì „Ê«‰⁄ »«Ìœ «“ —‰ê ﬁ—„“ «” ›«œÂ ò‰Ìœ" + vbNewLine + "· ›« «“  ’«ÊÌ—Ì »«ÿÊ·456 Ê⁄—÷414 «” ›«œÂ ò‰Ìœ", vbOKOnly + 32, "Select Picture")
CommonDialog1.Filter = "(*.jpg)|*.jpg|(*.gif)|*.gif|*.*|*.*"
CommonDialog1.ShowOpen
Image10.Picture = LoadPicture(CommonDialog1.FileName)
Form1.Show
Form1.Picture = Image10.Picture
Unload Me

End Sub

Private Sub image2_Click()
Form1.Picture = Image2.Picture
Unload Me
End Sub

Private Sub image3_Click()
Form1.Picture = Image3.Picture
Unload Me
End Sub

Private Sub image4_Click()
Form1.Picture = Image4.Picture
Unload Me
End Sub

Private Sub Image6_Click()
Form1.Picture = Image6.Picture
Unload Me
End Sub

Private Sub Image7_Click()
Form1.Picture = Image7.Picture
Unload Me
End Sub

Private Sub Image8_Click()
Form1.Picture = Image8.Picture
Unload Me
End Sub

Private Sub Image9_Click()
Form1.Picture = Image9.Picture
Unload Me
End Sub
