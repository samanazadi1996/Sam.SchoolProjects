VERSION 5.00
Begin VB.Form Formsel 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4980
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   2295
   End
   Begin VB.DirListBox p 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF9900&
      Height          =   2610
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4695
   End
   Begin VB.DriveListBox d 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF9900&
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "Formsel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
o = p
If Caption = "Select Folder" Then
Form3.File1 = o
Form3.Command1.ToolTipText = o
Form3.p = o
Form3.d = d
Else
Form3.Command2.ToolTipText = o
End If
Unload Me
End Sub

Private Sub d_Change()
On Error Resume Next
p = d
End Sub

Private Sub Form_Load()
Form3.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form3.Enabled = True
End Sub
