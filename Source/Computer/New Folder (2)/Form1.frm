VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.DirListBox p 
      Height          =   2115
      Left            =   2760
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
      Height          =   315
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   885
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   0
      Left            =   120
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Private Sub Form_Click()
p.Refresh
Form_Load
End Sub
Private Sub Form_Load()
On Error Resume Next
p = "C:\Documents and Settings\All Users\Desktop\"
For w = 0 To p.ListCount - 1
Load Image1(w + 1)
Image1(w).Visible = True
Image1(w).Top = Image1(w - 1).Top + 750
Next
For w = 0 To p.ListCount - 1
Load Label1(w + 1)
Label1(w).Visible = True
Label1(w).Top = Label1(w - 1).Top + 750
Label1(w).Caption = Mid(p.List(w), 45, Len(p.List(w)) - 43)
Next
End Sub

Private Sub Image1_Click(Index As Integer)
WinExec "Explorer.exe " & p.List(Index), 10
End Sub
