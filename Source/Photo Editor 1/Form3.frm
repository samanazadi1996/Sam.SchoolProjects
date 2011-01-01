VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form3 
   BackColor       =   &H0000FF00&
   Caption         =   "Conveter Saman"
   ClientHeight    =   6345
   ClientLeft      =   285
   ClientTop       =   555
   ClientWidth     =   9750
   ControlBox      =   0   'False
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   6345
   ScaleWidth      =   9750
   StartUpPosition =   1  'CenterOwner
   Begin VB.FileListBox File1 
      Archive         =   0   'False
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
      ForeColor       =   &H00FF00FF&
      Height          =   870
      Left            =   2520
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      Top             =   1440
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFC0&
      DrawWidth       =   5
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1275
      ScaleWidth      =   9690
      TabIndex        =   10
      Top             =   0
      Width           =   9750
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Zakhire tasvir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   2280
         Picture         =   "Form3.frx":0A8A
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Posheye khoroji"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1200
         Picture         =   "Form3.frx":0ECC
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tasavir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         Picture         =   "Form3.frx":DF7B
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   120
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "jpg"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   4920
         Picture         =   "Form3.frx":E3BD
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Value           =   -1  'True
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "gif"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   5520
         Picture         =   "Form3.frx":EA6E
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   6120
         Picture         =   "Form3.frx":F0A3
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "bmp"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   6720
         Picture         =   "Form3.frx":F75B
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "png"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   7320
         Picture         =   "Form3.frx":FE0C
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "tga"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   7920
         Picture         =   "Form3.frx":104BB
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "tiff"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   8520
         Picture         =   "Form3.frx":10B76
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "pcx"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   9120
         Picture         =   "Form3.frx":11236
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   3360
         X2              =   3360
         Y1              =   -120
         Y2              =   1440
      End
      Begin VB.Line Line6 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   4800
         X2              =   4800
         Y1              =   -120
         Y2              =   1440
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1095
         Left            =   3480
         Stretch         =   -1  'True
         ToolTipText     =   "Saman Azadi"
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   3  'Align Left
      BackColor       =   &H00FFC0FF&
      Height          =   4590
      Left            =   0
      ScaleHeight     =   4530
      ScaleWidth      =   2355
      TabIndex        =   7
      Top             =   1335
      Width           =   2415
      Begin VB.DriveListBox d 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   360
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   2175
      End
      Begin VB.DirListBox p 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   1710
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   0
      ScaleHeight     =   360
      ScaleWidth      =   9690
      TabIndex        =   1
      Top             =   5925
      Width           =   9750
      Begin MSComctlLib.ProgressBar sa 
         Height          =   255
         Left            =   2400
         TabIndex        =   22
         Top             =   60
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   60
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0000/00/00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   960
         TabIndex        =   5
         Top             =   60
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Saman azadi.jpg"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4320
         TabIndex        =   4
         Top             =   60
         Width           =   2295
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "File"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   345
         Left            =   6720
         TabIndex        =   3
         Top             =   0
         Width           =   405
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9240
      Top             =   1440
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   5400
      Top             =   5520
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   1485
      Left            =   6750
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form3.frx":118E7
      Top             =   4660
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu ItemV 
         Caption         =   "Tasavir"
         Shortcut        =   ^O
      End
      Begin VB.Menu Itemkh 
         Caption         =   "Posheye Khroji"
         Shortcut        =   ^K
      End
      Begin VB.Menu itemtabdil 
         Caption         =   "Tabdil"
         Shortcut        =   ^T
      End
      Begin VB.Menu itemtabdilall 
         Caption         =   "Tabdil Hame"
         Shortcut        =   ^S
      End
      Begin VB.Menu p4 
         Caption         =   "-"
      End
      Begin VB.Menu iExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpa As String, nSize As Long) As Long
Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Dim MS As String
Private Sub Command1_Click()
On Error Resume Next
Formsel.Show
Formsel.Caption = "Select Folder"
Formsel.p.Path = Command1.ToolTipText
Formsel.d = Mid(Command1.ToolTipText, 1, 3)
End Sub

Private Sub Command2_Click()
On Error Resume Next
Formsel.Show
Formsel.Caption = "Select Folder Output"
Formsel.p.Path = Command2.ToolTipText
Formsel.d = Mid(Command2.ToolTipText, 1, 3)
End Sub

Private Sub Command3_Click()

sa = 0
File1.ListIndex = -1
Timer1.Enabled = Not Timer1.Enabled
File1.Enabled = Not File1.Enabled
d.Enabled = Not d.Enabled
p.Enabled = Not p.Enabled
Command1.Enabled = Not Command1.Enabled
End Sub
Private Sub Command4_Click()
End Sub

Private Sub d_Change()
On Error GoTo 3
p = d
File1 = p
3:
End Sub

Private Sub Exit_Click()

End Sub

Private Sub File1_DblClick()
On Error GoTo 4
Image1.Picture = LoadPicture(p & "\" & File1)
4:
End Sub



Private Sub File1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then

PopupMenu File
End If
End Sub

Private Sub Form_Load()
MS = "JPG"
Command1.ToolTipText = App.Path
Dim ss As String
ss = String(255, 0)
GetUserName ss, 255
ss = Left(ss, InStr(ss, Chr(0)) - 1)
Command2.ToolTipText = "C:\Documents and Settings\" & ss & "\My Documents\My Pictures"
File1 = App.Path
End Sub

Private Sub Form_Resize()
On Error Resume Next
If p.Height < 2000 Then
p.Height = 5000
End If
If Form3.Height > 3000 Then
p.Height = Form3.Height - 2950
File1.Height = Form3.Height - (1700 + 1000)

End If
If Form3.Width > 2800 Then

File1.Width = Form3.Width - 2900
Else
File1.Width = 3000
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub iExit_Click()
Form4.Show
Form3.Hide
End Sub

Private Sub Itemkh_Click()
Call Command2_Click
End Sub

Private Sub ItemTabdil_Click()
On Error GoTo 5
Image1.Picture = LoadPicture(p & "\" & File1)
zs = zs + Mid(File1, 1, Len(File1) - 3)
SavePicture Image1.Picture, Command2.ToolTipText & "\" & zs & MS
5:
End Sub

Private Sub itemtabdilall_Click()
Call Command3_Click
End Sub

Private Sub ItemV_Click()
Call Command1_Click
End Sub

Private Sub Label1_Click()
saman = MsgBox("Time = " & Label1, vbOKOnly, "Time")
End Sub

Private Sub Label2_Click()
saman = MsgBox("Date = " & Label2, vbOKOnly, "Date")
End Sub


Private Sub Label4_Click()
Text1.Visible = Not Text1.Visible
End Sub

Private Sub Option1_Click()
MS = "JPG"
End Sub

Private Sub Option2_Click()
MS = "GIF"
End Sub

Private Sub Option3_Click()
MS = "ICO"
End Sub

Private Sub Option4_Click()
MS = "BMP"
End Sub

Private Sub Option5_Click()
MS = "PNG"
End Sub

Private Sub Option6_Click()
MS = "TGA"
End Sub

Private Sub Option7_Click()
MS = "TIFF"
End Sub

Private Sub Option8_Click()
MS = "PCX"
End Sub

Private Sub p_Change()
Text1.Visible = False
File1 = p
Command1.ToolTipText = p
End Sub





Private Sub Text1_DblClick()
Clipboard.SetText (Text1)
End Sub




Private Sub Timer1_Timer()
On Error Resume Next
If File1.ListIndex + 1 < File1.ListCount Then
File1.Enabled = False
d.Enabled = False
p.Enabled = False
Command1.Enabled = False
sa = sa + 1
File1.ListIndex = File1.ListIndex + 1
On Error GoTo 111
Image1.Picture = LoadPicture(p & "\" & File1)
zs = zs + Mid(File1, 1, Len(File1) - 3)
SavePicture Image1.Picture, Command2.ToolTipText & "\" & zs & MS
111:
Else
sa = 0
Timer1.Enabled = False
File1.Enabled = True
d.Enabled = True
p.Enabled = True
Command1.Enabled = True
 WinExec "Explorer.exe " & Command2.ToolTipText, 10

End If

End Sub

Private Sub Timer2_Timer()
If File1.ListCount <> 0 Then
sa.Max = File1.ListCount
End If
Label1 = Format(Time, "short time")
Label2 = Format(Date, "Medium Date")
Label3 = File1
Text1 = p
End Sub

