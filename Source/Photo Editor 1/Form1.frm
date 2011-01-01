VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paint Saman"
   ClientHeight    =   7575
   ClientLeft      =   180
   ClientTop       =   120
   ClientWidth     =   9120
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   9120
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFC0&
      Height          =   1080
      Left            =   0
      ScaleHeight     =   1020
      ScaleWidth      =   9060
      TabIndex        =   11
      Top             =   0
      Width           =   9120
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   8400
         MouseIcon       =   "Form1.frx":0A8A
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":0BDC
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   1320
         Picture         =   "Form1.frx":153F
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Save"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   720
         Picture         =   "Form1.frx":1AC6
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Open"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         Picture         =   "Form1.frx":1FF7
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "New"
         Top             =   120
         Width           =   495
      End
      Begin VB.Frame i 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   1920
         TabIndex        =   27
         Top             =   0
         Visible         =   0   'False
         Width           =   1575
         Begin VB.Image Image10 
            Height          =   480
            Left            =   120
            Picture         =   "Form1.frx":2576
            Top             =   180
            Width           =   480
         End
         Begin VB.Image Image9 
            Height          =   480
            Left            =   840
            Picture         =   "Form1.frx":2E40
            Top             =   180
            Width           =   480
         End
      End
      Begin VB.Frame yyy 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   1920
         TabIndex        =   17
         Top             =   0
         Visible         =   0   'False
         Width           =   6375
         Begin VB.CheckBox Check5 
            BackColor       =   &H00FFFFC0&
            Caption         =   "TD Text"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF00FF&
            Height          =   255
            Left            =   3960
            TabIndex        =   28
            Top             =   0
            Width           =   1215
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFC0&
            Caption         =   "FontSize"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   615
            Left            =   1560
            TabIndex        =   22
            Top             =   0
            Width           =   855
            Begin VB.ComboBox Combo3 
               ForeColor       =   &H00FF00FF&
               Height          =   315
               Left            =   120
               TabIndex        =   23
               Text            =   "8"
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFC0&
            Caption         =   "FontBold"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   255
            Left            =   2520
            TabIndex        =   21
            Top             =   0
            Width           =   1095
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "FontItalic"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   255
            Left            =   2520
            TabIndex        =   20
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox Check3 
            BackColor       =   &H00FFFFC0&
            Caption         =   "FontUnderline"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   255
            Left            =   2520
            TabIndex        =   19
            Top             =   480
            Width           =   1455
         End
         Begin VB.CheckBox Check4 
            BackColor       =   &H00FFFFC0&
            Caption         =   "FontStrikethru"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   -1  'True
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   255
            Left            =   2520
            TabIndex        =   18
            Top             =   720
            Width           =   1455
         End
         Begin VB.Frame ooo 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            Height          =   1095
            Left            =   3960
            TabIndex        =   29
            Top             =   0
            Visible         =   0   'False
            Width           =   2415
            Begin VB.OptionButton Option7 
               BackColor       =   &H00FFFFFF&
               Height          =   495
               Left            =   1200
               MouseIcon       =   "Form1.frx":370A
               MousePointer    =   99  'Custom
               Picture         =   "Form1.frx":385C
               Style           =   1  'Graphical
               TabIndex        =   36
               Top             =   0
               Value           =   -1  'True
               Width           =   495
            End
            Begin VB.OptionButton Option10 
               BackColor       =   &H00FFFFFF&
               Height          =   495
               Left            =   1680
               MouseIcon       =   "Form1.frx":3C9E
               MousePointer    =   99  'Custom
               Picture         =   "Form1.frx":3DF0
               Style           =   1  'Graphical
               TabIndex        =   35
               Top             =   480
               Width           =   495
            End
            Begin VB.OptionButton Option8 
               BackColor       =   &H00FFFFFF&
               Height          =   495
               Left            =   1680
               MouseIcon       =   "Form1.frx":4232
               MousePointer    =   99  'Custom
               Picture         =   "Form1.frx":4384
               Style           =   1  'Graphical
               TabIndex        =   34
               Top             =   0
               Width           =   495
            End
            Begin VB.OptionButton Option9 
               BackColor       =   &H00FFFFFF&
               Height          =   495
               Left            =   1200
               MouseIcon       =   "Form1.frx":47C6
               MousePointer    =   99  'Custom
               Picture         =   "Form1.frx":4918
               Style           =   1  'Graphical
               TabIndex        =   33
               Top             =   480
               Width           =   495
            End
            Begin VB.ComboBox r1 
               ForeColor       =   &H000000FF&
               Height          =   315
               ItemData        =   "Form1.frx":4D5A
               Left            =   120
               List            =   "Form1.frx":4D5C
               TabIndex        =   32
               Top             =   240
               Width           =   735
            End
            Begin VB.ComboBox g1 
               ForeColor       =   &H0000FF00&
               Height          =   315
               Left            =   120
               TabIndex        =   31
               Top             =   480
               Width           =   735
            End
            Begin VB.ComboBox b1 
               ForeColor       =   &H00FF0000&
               Height          =   315
               Left            =   120
               TabIndex        =   30
               Top             =   720
               Width           =   735
            End
            Begin VB.Image Image1 
               Height          =   255
               Left            =   900
               Picture         =   "Form1.frx":4D5E
               Stretch         =   -1  'True
               Top             =   720
               Width           =   255
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "FontName"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   615
            Left            =   0
            TabIndex        =   24
            Top             =   0
            Width           =   1455
            Begin VB.ComboBox Combo2 
               ForeColor       =   &H00FF00FF&
               Height          =   315
               ItemData        =   "Form1.frx":519C
               Left            =   60
               List            =   "Form1.frx":519E
               Sorted          =   -1  'True
               TabIndex        =   25
               Text            =   "Arial"
               Top             =   240
               Width           =   1335
            End
         End
      End
   End
   Begin MSComDlg.CommonDialog k 
      Left            =   6960
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.HScrollBar sd 
      Height          =   255
      LargeChange     =   500
      Left            =   120
      Max             =   0
      SmallChange     =   100
      TabIndex        =   10
      Top             =   7200
      Width           =   7575
   End
   Begin VB.VScrollBar ss 
      Height          =   6015
      LargeChange     =   500
      Left            =   7800
      Max             =   0
      SmallChange     =   100
      TabIndex        =   9
      Top             =   1080
      Width           =   255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0FF&
      Height          =   6015
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   7575
      Begin VB.Frame n 
         BackColor       =   &H00FFFFFF&
         Caption         =   "##################"
         Height          =   2175
         Left            =   5520
         TabIndex        =   38
         Top             =   0
         Visible         =   0   'False
         Width           =   2055
         Begin VB.Image Image8 
            Height          =   495
            Left            =   120
            Picture         =   "Form1.frx":51A0
            Stretch         =   -1  'True
            Top             =   1440
            Width           =   495
         End
         Begin VB.Image Image7 
            Height          =   495
            Left            =   1320
            Picture         =   "Form1.frx":5613
            Stretch         =   -1  'True
            Top             =   840
            Width           =   495
         End
         Begin VB.Image Image6 
            Height          =   495
            Left            =   720
            Picture         =   "Form1.frx":5A86
            Stretch         =   -1  'True
            Top             =   840
            Width           =   495
         End
         Begin VB.Image Image5 
            Height          =   495
            Left            =   120
            Picture         =   "Form1.frx":5EF9
            Stretch         =   -1  'True
            Top             =   840
            Width           =   495
         End
         Begin VB.Image Image4 
            Height          =   495
            Left            =   1320
            Picture         =   "Form1.frx":636C
            Stretch         =   -1  'True
            Top             =   240
            Width           =   495
         End
         Begin VB.Image Image3 
            Height          =   495
            Left            =   720
            Picture         =   "Form1.frx":67DF
            Stretch         =   -1  'True
            Top             =   240
            Width           =   495
         End
         Begin VB.Image Image2 
            Height          =   495
            Left            =   120
            Picture         =   "Form1.frx":6C52
            Stretch         =   -1  'True
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.PictureBox p 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   4695
         Left            =   0
         MouseIcon       =   "Form1.frx":70C5
         MousePointer    =   99  'Custom
         ScaleHeight     =   4635
         ScaleWidth      =   5235
         TabIndex        =   8
         Top             =   0
         Width           =   5295
         Begin VB.TextBox Text1 
            Height          =   855
            Left            =   2160
            MouseIcon       =   "Form1.frx":798F
            TabIndex        =   37
            Top             =   2160
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Line Line1 
            Visible         =   0   'False
            X1              =   2160
            X2              =   3600
            Y1              =   3120
            Y2              =   3120
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   4  'Align Right
      BackColor       =   &H00C0FFC0&
      Height          =   6495
      Left            =   8145
      ScaleHeight     =   6435
      ScaleWidth      =   915
      TabIndex        =   0
      Top             =   1080
      Width           =   975
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form1.frx":8259
         Left            =   120
         List            =   "Form1.frx":825B
         TabIndex        =   14
         Top             =   5280
         Width           =   735
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   -120
         Top             =   2760
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         MouseIcon       =   "Form1.frx":825D
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":83AF
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4320
         Width           =   735
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         MouseIcon       =   "Form1.frx":845B
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":85AD
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3480
         Width           =   735
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         MouseIcon       =   "Form1.frx":8CD5
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":8E27
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2640
         Width           =   735
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         MouseIcon       =   "Form1.frx":9470
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":95C2
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1800
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         MouseIcon       =   "Form1.frx":9E8C
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":9FDE
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   960
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         MouseIcon       =   "Form1.frx":A8A8
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":A9FA
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   240
         MouseIcon       =   "Form1.frx":AE3C
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   5760
         Width           =   525
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xx As Single, yy As Single, s As String, x6 As Single, y6 As Single
Dim l As String
Private Sub Check1_Click()
Text1.FontBold = Not Text1.FontBold
p.FontBold = Not p.FontBold
End Sub

Private Sub Check2_Click()
Text1.FontItalic = Not Text1.FontItalic
p.FontItalic = Not p.FontItalic
End Sub

Private Sub Check3_Click()
Text1.FontUnderline = Not Text1.FontUnderline
p.FontUnderline = Not p.FontUnderline
End Sub

Private Sub Check4_Click()
Text1.FontStrikethru = Not Text1.FontStrikethru
p.FontStrikethru = Not p.FontStrikethru

End Sub

Private Sub Check5_Click()
ooo.Visible = Not ooo.Visible
End Sub

Private Sub Combo2_Click()
On Error GoTo 111
Text1.FontName = Combo2
p.FontName = Combo2
111:
End Sub

Private Sub Combo3_Click()
Text1.FontSize = Combo3
p.FontSize = Combo3
End Sub

Private Sub Command1_Click()
formnew.Show
End Sub

Private Sub Command2_Click()
k.Filter = "(*.jpg)|*.jpg|(*.gif)|*.gif|(*.bmp)|*.bmp|(*.png)|*.png|(*.*)|*.*"
k.ShowOpen
ss = 0
sd = 0
p.Picture = LoadPicture(k.FileName)
ss.Max = p.Height - Frame1.Height
sd.Max = p.Width - Frame1.Width
End Sub

Private Sub Command3_Click()
k.Filter = "(*.jpg)|*.jpg|(*.gif)|*.gif|(*.bmp)|*.bmp|(*.png)|*.png|(*.*)|*.*"
k.ShowSave
If k.FileName <> "" Then SavePicture p.Image, k.FileName
End Sub

Private Sub Command4_Click()
u = MsgBox("Save Changes", vbYesNoCancel + 48, "Save")
If u = vbYes Then
Call Command3_Click
End
ElseIf u = vbNo Then
Unload Me
Form4.Show
End If
End Sub

Private Sub Command5_Click()

End Sub

Private Sub Command6_Click()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = vbCtrlMask And KeyCode = vbKeyN Then Call Command1_Click
If Shift = vbCtrlMask And KeyCode = vbKeyO Then Call Command2_Click
If Shift = vbCtrlMask And KeyCode = vbKeyS Then Call Command3_Click
If KeyCode = 27 Then Call Command4_Click
End Sub

Private Sub Form_Load()
s = "mostatil"
Combo1.Clear
For w = 1 To 50
Combo1.AddItem (w)
Next
For e = 1 To Screen.FontCount
Combo2.AddItem Screen.Fonts(e)
Next
For t = 6 To 72 Step 3
Combo3.AddItem (t)
Next
r1.AddItem ("")
g1.AddItem ("")
b1.AddItem ("")
For w = 0 To 255
r1.AddItem (w)
g1.AddItem (w)
b1.AddItem (w)
Next

End Sub

Private Sub Itemdayareh_Click()

End Sub

Private Sub ItemMostatil_Click()

End Sub

Private Sub i_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image10.BorderStyle = 0
Image9.BorderStyle = 0
End Sub

Private Sub Image1_Click()
n.Visible = Not n.Visible
End Sub

Private Sub Image10_Click()
s = "mostatil"
End Sub

Private Sub Image10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image10.BorderStyle = 1
Image9.BorderStyle = 0
End Sub

Private Sub Image2_Click()
g1 = ""
b1 = 255
r1 = 70
n.Visible = False
End Sub

Private Sub Image3_Click()
n.Visible = False
r1 = ""
g1 = 0
b1 = 255

End Sub

Private Sub Image4_Click()
n.Visible = False
r1 = 255
g1 = ""
b1 = 0
End Sub

Private Sub Image5_Click()
n.Visible = False
r1 = 0
g1 = 255
b1 = ""
End Sub

Private Sub Image6_Click()
n.Visible = False
r1 = 172
g1 = 185
b1 = ""
End Sub

Private Sub Image7_Click()
n.Visible = False
r1 = 125
g1 = ""
b1 = 198
End Sub

Private Sub Image8_Click()
n.Visible = False
r1 = ""
g1 = 145
b1 = 47
End Sub

Private Sub Image9_Click()
s = "dayareh"
End Sub

Private Sub Image9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image10.BorderStyle = 0
Image9.BorderStyle = 1
End Sub

Private Sub Label1_Click()
k.ShowColor
Label1.BackColor = k.Color
Text1.ForeColor = Label1.BackColor
End Sub

Private Sub Option1_Click()

i.Visible = False
Combo1.Clear
For w = 1 To 50
Combo1.AddItem (w)
Next
Call hideyyy
End Sub

Private Sub Option10_Click()
l = "SD"
End Sub

Private Sub Option11_Click()

End Sub

Private Sub Option2_Click()

i.Visible = False
Combo1.Clear
For w = 1 To 50
Combo1.AddItem (w)
Next
Call hideyyy
End Sub

Private Sub Option3_Click()

i.Visible = True
Combo1.Clear
For w = 1 To 100
Combo1.AddItem (w)
Next
Call hideyyy
End Sub

Private Sub Option3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu r
End Sub

Private Sub Option4_Click()

i.Visible = False
p.DrawWidth = 1
Combo1.Clear
Combo1.AddItem (1)
Combo1.AddItem (250)
Combo1.AddItem (500)
Combo1.AddItem (750)
Combo1.AddItem (1000)
Combo1.AddItem (2000)
Call hideyyy
End Sub

Private Sub Option5_Click()

i.Visible = False
o = MsgBox("Color", vbYesNo + 32, "")
If o = vbYes Then
k.ShowColor
p.BackColor = k.Color
End If
Call hideyyy
End Sub

Private Sub Option6_Click()

i.Visible = False
yyy.Visible = True
End Sub

Private Sub Option7_Click()
l = "AW"
End Sub

Private Sub Option8_Click()
l = "WD"
End Sub

Private Sub Option9_Click()
l = "SA"
End Sub

Private Sub p_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xx = X
yy = Y
If Option2.Value = True Then
Line1.BorderWidth = p.DrawWidth
Line1.BorderColor = Label1.BackColor
Line1.Visible = True
Line1.X1 = X
Line1.X2 = xx
Line1.Y1 = Y
Line1.Y2 = yy
End If
End Sub

Private Sub p_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
n.Visible = False
If Combo1 = "" Then Combo1 = Combo1.List(0)
If Button = 1 Then

If Option2.Value = True And Line1.Visible = True Then
Line1.X1 = X
Line1.X2 = xx
Line1.Y1 = Y
Line1.Y2 = yy
End If
    If Option4.Value = True Then
    Timer1.Enabled = True
    xx = X
    yy = Y
    End If
    
    If Option1.Value = True Then
    p.DrawWidth = Combo1
    
    p.Line (xx, yy)-(X, Y), Label1.BackColor
    xx = X
    yy = Y
    End If
    
    
    
    
End If
End Sub

Private Sub p_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Enabled = False
If Button = 1 Then
       
    If Option2.Value = True Then
    p.DrawWidth = Combo1
    Line1.Visible = False
    p.Line (X, Y)-(xx, yy), Label1.BackColor
    End If
      
    If Option3.Value = True And s = "mostatil" Then
    p.DrawWidth = Combo1
    p.Line (X, Y)-(xx, yy), Label1.BackColor, B
    End If
    
    If Option3.Value = True And s = "dayareh" Then
    p.DrawWidth = Combo1
    p.Circle (xx, yy), Abs(xx - X), Label1.BackColor
    End If
    
    If Option6.Value = True Then
    Text1.Visible = True
    Text1.Move xx, yy, Abs(xx - X), Abs(yy - Y)
    End If
End If
End Sub

Private Sub sd_Change()
p.Left = -sd
End Sub

Private Sub ss_Change()
p.Top = -ss
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
n.Visible = False
If KeyCode = 13 Then
p.DrawWidth = 1
    If ooo.Visible = True Then
    X = Text1.Left
    Y = Text1.Top
    For w = 0 To 255
    Select Case l
    Case "WD"
        p.CurrentX = X + w
    p.CurrentY = Y - w

    Case "SA"
        p.CurrentX = X - w
    p.CurrentY = Y + w

    Case "SD"
    
    p.CurrentX = X - w
    p.CurrentY = Y - w
    Case Else
        p.CurrentX = X + w
    p.CurrentY = Y + w
    End Select
    
    If r1 = "" Then
    r = w
    Else
    r = r1
    End If
    If g1 = "" Then
    g = w
    Else
    g = g1
    End If
    If b1 = "" Then

    b = w
    Else
    b = b1
    End If
    p.ForeColor = RGB(r, g, b)
    p.Print Text1
    Next
    Else
    X = Text1.Left
    Y = Text1.Top
    u = Point(X, Y)
    p.PSet (X, Y), u
    p.ForeColor = Label1.BackColor
    p.Print Text1
    End If
Text1 = ""
Text1.Visible = False
End If
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text1.MousePointer = 99
x6 = X
y6 = Y
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Text1.Move Text1.Left + X - x6, Text1.Top + Y - y6
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text1.MousePointer = 0
End Sub

Private Sub Timer1_Timer()
For w = 1 To 10
o = Rnd * Combo1 / 10
X1 = xx + Cos(o) * Rnd * Combo1
Y1 = yy + Sin(o) * Rnd * Combo1
p.PSet (X1, Y1), Label1.BackColor
Next
End Sub

Private Sub hideyyy()
yyy.Visible = False
Text1.Visible = False
End Sub
