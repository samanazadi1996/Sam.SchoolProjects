VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Web Editor Saman"
   ClientHeight    =   6000
   ClientLeft      =   4380
   ClientTop       =   3330
   ClientWidth     =   7740
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.PictureBox Picture2 
      Align           =   4  'Align Right
      BackColor       =   &H00FFC0FF&
      Height          =   5505
      Left            =   6135
      ScaleHeight     =   5445
      ScaleWidth      =   1545
      TabIndex        =   38
      Top             =   495
      Width           =   1605
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   780
         Left            =   120
         TabIndex        =   41
         ToolTipText     =   "Ctrl+q For Clear , Del For DeleteItem , Insert for AddItem"
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Insert"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   120
         Width           =   1335
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   3420
         ItemData        =   "MDIForm1.frx":0442
         Left            =   120
         List            =   "MDIForm1.frx":051E
         Sorted          =   -1  'True
         TabIndex        =   39
         Top             =   540
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1680
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFC0FF&
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   7680
      TabIndex        =   34
      Top             =   0
      Width           =   7740
      Begin VB.TextBox LabelDocuments 
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
         ForeColor       =   &H00FF00FF&
         Height          =   360
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "Document 0"
         Top             =   30
         Width           =   1455
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   6840
         TabIndex        =   36
         ToolTipText     =   "FontSize"
         Top             =   60
         Width           =   735
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   5400
         TabIndex        =   35
         ToolTipText     =   "FontName"
         Top             =   60
         Width           =   1335
      End
      Begin VB.Image Image7 
         Height          =   375
         Left            =   3240
         Picture         =   "MDIForm1.frx":0773
         Stretch         =   -1  'True
         ToolTipText     =   "Preview     F12"
         Top             =   30
         Width           =   375
      End
      Begin VB.Image Image6 
         Height          =   375
         Left            =   2640
         Picture         =   "MDIForm1.frx":0A1C
         Stretch         =   -1  'True
         ToolTipText     =   "Paste     Ctrl+V"
         Top             =   30
         Width           =   375
      End
      Begin VB.Image Image5 
         Height          =   375
         Left            =   2160
         Picture         =   "MDIForm1.frx":13D5
         Stretch         =   -1  'True
         ToolTipText     =   "Copy     Ctrl+C"
         Top             =   30
         Width           =   375
      End
      Begin VB.Image Image4 
         Height          =   375
         Left            =   1680
         Picture         =   "MDIForm1.frx":1ADB
         Stretch         =   -1  'True
         ToolTipText     =   "Cut     Ctrl+X"
         Top             =   30
         Width           =   375
      End
      Begin VB.Image Image3 
         Height          =   375
         Left            =   1080
         Picture         =   "MDIForm1.frx":20C0
         Stretch         =   -1  'True
         ToolTipText     =   "Save     Ctrl+S"
         Top             =   30
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Left            =   600
         Picture         =   "MDIForm1.frx":2647
         Stretch         =   -1  'True
         ToolTipText     =   "Open     Ctrl+O"
         Top             =   30
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   375
         Left            =   120
         Picture         =   "MDIForm1.frx":2B7E
         Stretch         =   -1  'True
         ToolTipText     =   "New     Ctrl+N"
         Top             =   30
         Width           =   375
      End
   End
   Begin VB.PictureBox t 
      Align           =   3  'Align Left
      BackColor       =   &H00FFC0FF&
      Height          =   5505
      Left            =   0
      ScaleHeight     =   5445
      ScaleWidth      =   990
      TabIndex        =   0
      Top             =   495
      Width           =   1050
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   600
         Top             =   2400
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   360
         ItemData        =   "MDIForm1.frx":310E
         Left            =   120
         List            =   "MDIForm1.frx":3110
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "toolbox"
         Top             =   120
         Width           =   735
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         Caption         =   "Object"
         Height          =   5415
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Visible         =   0   'False
         Width           =   735
         Begin VB.CommandButton Command13 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Select"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   4320
            Width           =   495
         End
         Begin VB.CommandButton Command19 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Frame"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   3720
            Width           =   495
         End
         Begin VB.CommandButton Command7 
            BackColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   120
            Picture         =   "MDIForm1.frx":3112
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Image"
            Top             =   3120
            Width           =   495
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Text"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Text"
            Top             =   120
            Width           =   495
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Submit"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Submit"
            Top             =   720
            Width           =   495
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Check box"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "CheckBox"
            Top             =   1320
            Width           =   495
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "File"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "File"
            Top             =   1920
            Width           =   495
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Radio"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Radio"
            Top             =   2520
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         Caption         =   "Text"
         Height          =   4935
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Visible         =   0   'False
         Width           =   735
         Begin VB.CommandButton Command3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Label"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   7
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Label"
            Top             =   120
            Width           =   495
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "h6"
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
            Index           =   6
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "h6"
            Top             =   4320
            Width           =   495
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "h5"
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
            Index           =   5
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "h5"
            Top             =   3720
            Width           =   495
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "h4"
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
            Index           =   4
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "h4"
            Top             =   3120
            Width           =   495
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "h3"
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
            Index           =   3
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "h3"
            Top             =   2520
            Width           =   495
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "h2"
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
            Index           =   2
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "h2"
            Top             =   1920
            Width           =   495
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "h1"
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
            Index           =   1
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "h1"
            Top             =   1320
            Width           =   495
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "P"
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
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "P"
            Top             =   720
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         Caption         =   "Table"
         Height          =   3135
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Visible         =   0   'False
         Width           =   735
         Begin VB.CommandButton CmdRow 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Row"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Rowspan"
            Top             =   1320
            Width           =   495
         End
         Begin VB.CommandButton Cmdcol 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Col"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Colspan"
            Top             =   720
            Width           =   495
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   120
            Picture         =   "MDIForm1.frx":341C
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Table"
            Top             =   120
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         Caption         =   "Other"
         Height          =   5895
         Index           =   4
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Visible         =   0   'False
         Width           =   735
         Begin VB.CommandButton Command15 
            BackColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   120
            Picture         =   "MDIForm1.frx":37F6
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   2520
            Width           =   495
         End
         Begin VB.CommandButton Command8 
            BackColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   120
            Picture         =   "MDIForm1.frx":3D2D
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Link"
            Top             =   1920
            Width           =   495
         End
         Begin VB.CommandButton Command6 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Marquee"
            Top             =   1320
            Width           =   495
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   120
            Picture         =   "MDIForm1.frx":433F
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   720
            Width           =   495
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "C"
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
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Center"
            Top             =   120
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         Caption         =   "Font"
         Height          =   2535
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   735
         Begin VB.CommandButton Command9 
            BackColor       =   &H00FFFFFF&
            Caption         =   "B"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "FontBold"
            Top             =   720
            Width           =   495
         End
         Begin VB.CommandButton Command10 
            BackColor       =   &H00FFFFFF&
            Caption         =   "I"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "FontItalic"
            Top             =   1320
            Width           =   495
         End
         Begin VB.CommandButton Command11 
            BackColor       =   &H00FFFFFF&
            Caption         =   "U"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "FontUnderline"
            Top             =   1920
            Width           =   495
         End
         Begin VB.CommandButton Command12 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Font"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "FontName"
            Top             =   120
            Width           =   495
         End
      End
   End
   Begin VB.Menu ItemFile 
      Caption         =   "&File"
      Begin VB.Menu ItemNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu ItemOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu l4 
         Caption         =   "-"
      End
      Begin VB.Menu ItemSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu ItemSaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu l1 
         Caption         =   "-"
      End
      Begin VB.Menu ItemExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu ItemEdit 
      Caption         =   "&Edit"
      Begin VB.Menu ItemCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu ItemCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu ItemPaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu l2 
         Caption         =   "-"
      End
      Begin VB.Menu ItemPreview 
         Caption         =   "Preview"
         Shortcut        =   {F12}
      End
      Begin VB.Menu l3 
         Caption         =   "-"
      End
      Begin VB.Menu ItemNow 
         Caption         =   "Time And Date"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu ItemView 
      Caption         =   "&View"
      Begin VB.Menu ItemToolbox 
         Caption         =   "Toolbox"
         Checked         =   -1  'True
      End
      Begin VB.Menu ItemItems 
         Caption         =   "Items"
         Checked         =   -1  'True
      End
      Begin VB.Menu ItemHlpsaman 
         Caption         =   "Insert"
         Checked         =   -1  'True
      End
      Begin VB.Menu q6 
         Caption         =   "-"
      End
      Begin VB.Menu ItemGoto 
         Caption         =   "Goto Window"
         WindowList      =   -1  'True
         Begin VB.Menu ItemAct 
            Caption         =   ""
         End
      End
   End
   Begin VB.Menu ItemWin 
      Caption         =   "&Window"
      Begin VB.Menu A1 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu A2 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu A3 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu A4 
         Caption         =   "&Arrange Icons"
      End
      Begin VB.Menu q7 
         Caption         =   "-"
      End
      Begin VB.Menu ItemColor 
         Caption         =   "Color Windows"
         Begin VB.Menu Item1 
            Caption         =   "Automatic"
            Index           =   0
         End
      End
   End
   Begin VB.Menu ItemHelp 
      Caption         =   "&Help"
      Begin VB.Menu ItemAbout 
         Caption         =   "About Programe"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim e1 As Integer

Private Sub A1_Click()
Me.Arrange 0
End Sub
Private Sub A2_Click()
Me.Arrange 1
End Sub
Private Sub A3_Click()
Me.Arrange 2
End Sub
Private Sub A4_Click()
Me.Arrange 3
End Sub
Private Sub Combo2_Click()
On Error Resume Next
ActiveForm.Text1.FontName = Combo2
End Sub
Private Sub Combo3_Click()
On Error Resume Next
ActiveForm.Text1.FontSize = Combo3
End Sub

Private Sub Command14_Click()
MDIForm1.Enabled = False
Form8.Show
End Sub

Private Sub Command15_Click()
CommonDialog1.ShowOpen
MDIForm1.ActiveForm.Text1.SelText = Chr(34) & CommonDialog1.FileName & Chr(34)
End Sub

Private Sub Image1_Click()
ItemNew_Click
End Sub
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.BorderStyle = 1
End Sub
Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.BorderStyle = 0
End Sub
Private Sub Image2_Click()
ItemOpen_Click
End Sub
Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.BorderStyle = 1
End Sub
Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.BorderStyle = 0
End Sub
Private Sub Image3_Click()
ItemSave_Click
End Sub
Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.BorderStyle = 1
End Sub
Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.BorderStyle = 0
End Sub
Private Sub Image4_Click()
ItemCut_Click
End Sub
Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.BorderStyle = 1
End Sub
Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.BorderStyle = 0
End Sub
Private Sub Image5_Click()
ItemCopy_Click
End Sub
Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.BorderStyle = 1
End Sub
Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.BorderStyle = 0
End Sub
Private Sub Image6_Click()
ItemPaste_Click
End Sub
Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image6.BorderStyle = 1
End Sub
Private Sub Image6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image6.BorderStyle = 0
End Sub

Private Sub Image7_Click()
ItemPreview_Click
End Sub

Private Sub Image7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.BorderStyle = 1
End Sub

Private Sub Image7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.BorderStyle = 0
End Sub

Private Sub Item1_Click(Index As Integer)
On Error Resume Next
SaveSetting "WebEditor", "Saman", "Color", Index
For w = 0 To Item1.Count - 1
Item1(w).Checked = False
Next
Item1(Index).Checked = True
Select Case Index
Case 0
Picture1.BackColor = RGB(150, 150, 150)
Case 1
Picture1.BackColor = vbRed
Case 2
Picture1.BackColor = RGB(255, 70, 255)
Case 3
Picture1.BackColor = RGB(0, 162, 255)
Case 4
Picture1.BackColor = RGB(0, 200, 0)
Case 5
Picture1.BackColor = vbYellow
Case 6
Picture1.BackColor = RGB(255, 190, 0)
End Select
ActiveForm.BackColor = Picture1.BackColor
ActiveForm.Text1.ForeColor = Picture1.BackColor
T.BackColor = Picture1.BackColor
Picture2.BackColor = Picture1.BackColor

LabelDocuments.ForeColor = Picture1.BackColor
Combo1.ForeColor = Picture1.BackColor
Combo2.ForeColor = Picture1.BackColor
Combo3.ForeColor = Picture1.BackColor
List1.ForeColor = Picture2.BackColor
List2.ForeColor = Picture2.BackColor
    For w = 0 To Frame1.Count - 1
    Frame1(w).BackColor = T.BackColor
    Next
End Sub

Private Sub ItemAbout_Click()
Form7.Show
End Sub
Private Sub ItemCopy_Click()
Clipboard.SetText (ActiveForm.Text1.SelText)
End Sub
Private Sub ItemCut_Click()
Clipboard.SetText (ActiveForm.Text1.SelText)
ActiveForm.Text1.SelText = ""
End Sub
Private Sub ItemExit_Click()
End
End Sub

Private Sub ItemHlpsaman_Click()
ItemHlpsaman.Checked = Not ItemHlpsaman.Checked
Picture2.Visible = ItemHlpsaman.Checked
End Sub

Private Sub ItemItems_Click()
ItemItems.Checked = Not ItemItems.Checked
Picture1.Visible = ItemItems.Checked
End Sub
Private Sub ItemNew_Click()
Static Namber As Long
Namber = Namber + 1
Set h = New h
h.Show
h.Caption = "Document " & Namber
h.BackColor = Picture1.BackColor
h.Text1.ForeColor = Picture1.BackColor
End Sub
Private Sub ItemNow_Click()
ActiveForm.Text1.SelText = Now
End Sub
Private Sub ItemOpen_Click()
On Error GoTo 1
CommonDialog1.FileName = ""
CommonDialog1.Filter = "(*.html)|*.html|(*.php)|*.php|(*.txt)|*.txt|All File|*.*"
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
Set h = New h
h.Show
h.Caption = Dir(CommonDialog1.FileName)
o = FreeFile
Open CommonDialog1.FileName For Input As #o
Do While Not (EOF(o)) = True
Line Input #o, p
ActiveForm.Text1 = ActiveForm.Text1 + p + vbNewLine
Loop
End If
1:
End Sub
Private Sub ItemPaste_Click()
ActiveForm.Text1.SelText = Clipboard.GetText
End Sub

Private Sub ItemPreview_Click()
On Error GoTo 1
r = FreeFile
xxxx = App.Path & "\WebEditor Saman\Saman.html"
Open xxxx For Output As #r
Print #r, ActiveForm.Text1
Close #r
xxxx = App.Path & "\WebEditor Saman\Saman.html"
Shell "explorer " & xxxx, vbNormalFocus
1:
End Sub

Private Sub ItemSave_Click()
On Error GoTo 1
If CommonDialog1.FileName <> "" Then
r = FreeFile
Open CommonDialog1.FileName For Output As #r
Print #r, ActiveForm.Text1
Close #r
Else
ItemSaveAs_Click
End If
1:
End Sub
Private Sub ItemSaveAs_Click()
On Error GoTo 1
CommonDialog1.Filter = "(*.html)|*.html|(*.php)|*.php|All Files(*.*)|*.*"
CommonDialog1.ShowSave
If CommonDialog1.FileName <> "" Then
r = FreeFile
Open CommonDialog1.FileName For Output As #r
Print #r, ActiveForm.Text1
Close #r
End If
1:
End Sub
Private Sub ItemToolbox_Click()
ItemToolbox.Checked = Not ItemToolbox.Checked
T.Visible = ItemToolbox.Checked
End Sub

Private Sub List1_DblClick()
On Error Resume Next
ActiveForm.Text1.SelText = " " & List1.List(List1.ListIndex)
ActiveForm.Text1.SetFocus
End Sub


Private Sub List2_DblClick()
yyy = False
ttt = Clipboard.GetText
For w = 0 To List2.ListCount - 1
If ttt = List2.List(w) Then yyy = True
Next
If yyy = False Then If MsgBox("Save Clipboard", vbYesNo, "Save") = vbYes Then List2.AddItem ttt
Clipboard.SetText List2.List(List2.ListIndex)
End Sub

Private Sub List2_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyDelete Then
List2.RemoveItem List2.ListIndex
List2.SetFocus
End If
If KeyCode = vbKeyInsert Then List2.AddItem InputBox("Enter Text", "Text")
If KeyCode = vbKeyQ And Shift = vbCtrlMask Then List2.Clear
End Sub
Private Sub MDIForm_Load()
  
  If App.PrevInstance = True Then
     MsgBox " »—‰«„Â œ—Õ«· «Ã—« »ÊœÂ Ê «„ò«‰ «Ã—«Ì Â„“„«‰ ¬‰ ÊÃÊœ ‰œ«—œ ", vbCritical, "Warning !"
     End
  End If
  Hide
  On Error Resume Next
h.Show
For w = 1 To 6
Load Item1(w)
Next
Item1(1).Caption = "Red"
Item1(2).Caption = "Magenta"
Item1(3).Caption = "Blue"
Item1(4).Caption = "Green"
Item1(5).Caption = "Yellow"
Item1(6).Caption = "Orange"
    i = GetSetting("WebEditor", "Saman", "Color", 3)
    Item1_Click (i)
MkDir (App.Path & "\WebEditor Saman")
ItemItems.Checked = GetSetting("WebEditor", "Saman", "Items", ItemItems.Checked)
ItemToolbox.Checked = GetSetting("WebEditor", "Saman", "Tool", ItemToolbox.Checked)
ItemHlpsaman.Checked = GetSetting("WebEditor", "Saman", "Hlp", ItemHlpsaman.Checked)
T.Visible = ItemToolbox.Checked
Picture1.Visible = ItemItems.Checked
Picture2.Visible = ItemHlpsaman.Checked
WindowState = GetSetting("WebEditor", "Saman", "WindowState", WindowState)
If WindowState = 0 Then Move GetSetting("WebEditor", "Saman", "Left", Me.Left), GetSetting("WebEditor", "Saman", "Top", Me.Top), GetSetting("WebEditor", "Saman", "Width", Me.Width), GetSetting("WebEditor", "Saman", "Height", Me.Height)
For w = 0 To Frame1.Count - 1
Combo1.AddItem (Frame1(w).Caption)
Next
For w = 0 To Screen.FontCount
Combo2.AddItem (Screen.Fonts(w))
Next
For w = 6 To 72 Step 3
Combo3.AddItem w
Next
co2 = GetSetting("WebEditor", "Saman", "FontName")
co3 = GetSetting("WebEditor", "Saman", "FontSize")
If co2 <> "" Then Combo2.ListIndex = co2
If co3 <> "" Then Combo3.ListIndex = co3
On Error GoTo 1
r = FreeFile
Open (App.Path & "\WebEditor Saman\Saman.dat") For Input As #r
Do While EOF(r) = False
Line Input #r, ooo
List2.AddItem ooo
Loop
Close #r
1:
Show
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
SaveSetting "WebEditor", "Saman", "WindowState", Me.WindowState
If WindowState = 0 Then
SaveSetting "WebEditor", "Saman", "Left", Me.Left
SaveSetting "WebEditor", "Saman", "Top", Me.Top
SaveSetting "WebEditor", "Saman", "Width", Me.Width
SaveSetting "WebEditor", "Saman", "Height", Me.Height
End If
SaveSetting "WebEditor", "Saman", "Items", ItemItems.Checked
SaveSetting "WebEditor", "Saman", "Tool", ItemToolbox.Checked
SaveSetting "WebEditor", "Saman", "Hlp", ItemHlpsaman.Checked
SaveSetting "WebEditor", "Saman", "FontName", Combo2.ListIndex
SaveSetting "WebEditor", "Saman", "FontSize", Combo3.ListIndex
r = FreeFile
Open (App.Path & "\WebEditor Saman\Saman.dat") For Output As #r
For w = 0 To List2.ListCount - 1
Print #r, List2.List(w)
Next
Close #r
End Sub
Private Sub Combo1_Click()
Call FALS
Frame1(Combo1.ListIndex).Visible = True
End Sub
Private Sub Command1_Click()
MDIForm1.Enabled = False
Form5.Show
End Sub
Private Sub Command10_Click()
ActiveForm.Text1.SelText = vbNewLine + "<I>"
Clipboard.SetText ("</I>")
End Sub
Private Sub Command11_Click()
ActiveForm.Text1.SelText = vbNewLine + "<U>"
Clipboard.SetText ("</U>")
End Sub
Private Sub Command12_Click()
MDIForm1.Enabled = False
Form1.Show
End Sub
Private Sub Command13_Click()
MDIForm1.Enabled = False
Form2.Show
End Sub
Private Sub Cmdcol_Click()
On Error Resume Next
a = InputBox("Enter Colspan")
If a <> 0 Then ActiveForm.Text1.SelText = " Colspan=" & Int(a) & " "
End Sub
Private Sub CmdRow_Click()
On Error Resume Next
a = InputBox("Enter Rowspan")
If a <> 0 Then ActiveForm.Text1.SelText = " Rowspan=" & Int(a) & " "
End Sub
Private Sub Command19_Click()
q1 = InputBox("Enter Src")
q2 = InputBox("Enter Name")
q3 = InputBox("Enter Width")
q4 = InputBox("Enter Height")
ActiveForm.Text1.SelText = "<iframe"
If q1 <> "" Then ActiveForm.Text1.SelText = " Src=" & Chr(34) & q1 & Chr(34)
If q2 <> "" Then ActiveForm.Text1.SelText = " Name=" & Chr(34) & q2 & Chr(34)
If q3 <> "" Then ActiveForm.Text1.SelText = " Width=" & Chr(34) & q3 & Chr(34)
If q4 <> "" Then ActiveForm.Text1.SelText = " Height=" & Chr(34) & q4 & Chr(34)
ActiveForm.Text1.SelText = "/>"
End Sub
Private Sub Command2_Click()
ActiveForm.Text1.SelText = vbNewLine + "<Center>"
Clipboard.SetText ("</Center>")
End Sub
Private Sub Command3_Click(Index As Integer)
o = InputBox("Enter Text")
ActiveForm.Text1.SelText = vbNewLine + "<" & Command3(Index).Caption & ">" & o & "</" & Command3(Index).Caption & ">"
End Sub
Private Sub Command4_Click(Index As Integer)
MDIForm1.Enabled = False
Form3.Show
Form3.Text1.Text = Command4(Index).ToolTipText
End Sub
Private Sub Command5_Click()
ActiveForm.Text1.SelText = vbNewLine + "</Br>"
End Sub
Private Sub Command6_Click()
u = InputBox("direction Is:")
If UCase(u) <> UCase("down") And UCase(u) <> UCase("up") And UCase(u) <> UCase("left") And UCase(u) <> UCase("right") Then u = "Left"
ActiveForm.Text1.SelText = vbNewLine + "<Marquee direction=" & u & ">"
Clipboard.SetText ("</Marquee>")
End Sub
Private Sub Command7_Click()
MDIForm1.Enabled = False
Form4.Show
End Sub
Private Sub Command8_Click()
MDIForm1.Enabled = False
Form6.Show
End Sub
Private Sub Command9_Click()
ActiveForm.Text1.SelText = vbNewLine + "<B>"
Clipboard.SetText ("</B>")
End Sub

Private Sub Picture2_Resize()
On Error Resume Next
List1.Move 120, 480, Picture2.ScaleWidth - 240, Picture2.ScaleHeight - 2000
List2.Move 120, List1.Top + List1.Height + 120, Picture2.ScaleWidth - 240, Picture2.ScaleHeight - (List1.Top + List1.Height) - 240
End Sub

Private Sub Timer1_Timer()
p = "      Marquee     "
e1 = e1 + 1
Command6.Caption = Mid(p, e1, 4)
If e1 = Len(p) Then
e1 = 1
End If
End Sub
Public Sub FALS()
For w = 0 To Frame1.Count - 1
Frame1(w).Visible = False
Next
End Sub
