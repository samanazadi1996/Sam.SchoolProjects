VERSION 5.00
Begin VB.Form h 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Document 0"
   ClientHeight    =   5445
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4575
   Icon            =   "h.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5445
   ScaleWidth      =   4575
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2535
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "h.frx":0442
      Top             =   120
      Width           =   3375
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1680
      Top             =   2520
   End
End
Attribute VB_Name = "h"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
MDIForm1.LabelDocuments = Me.Caption
MDIForm1.ItemAct.Caption = Caption
MDIForm1.Combo2 = Text1.FontName
MDIForm1.Combo3 = Text1.FontSize
End Sub
Private Sub Form_Resize()
On Error Resume Next

Text1.Move 120, 120, ScaleWidth - 240, ScaleHeight - 240
End Sub
Private Sub Label4_Click()
o = InputBox("Title Is:", "Title")
Label4 = "<Title>" & o & "</Title>"
End Sub

Private Sub Timer1_Timer()
If BackColor <> MDIForm1.Picture1.BackColor Then BackColor = MDIForm1.Picture1.BackColor
If Text1.ForeColor <> MDIForm1.Picture1.BackColor Then Text1.ForeColor = MDIForm1.Picture1.BackColor
End Sub
