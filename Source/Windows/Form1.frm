VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  "
   ClientHeight    =   8475
   ClientLeft      =   2295
   ClientTop       =   885
   ClientWidth     =   10395
   ControlBox      =   0   'False
   ForeColor       =   &H00FF9900&
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   Picture         =   "Form1.frx":324A
   ScaleHeight     =   8475
   ScaleWidth      =   10395
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox pro1 
      BackColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   1320
      MouseIcon       =   "Form1.frx":C682
      MousePointer    =   99  'Custom
      ScaleHeight     =   2595
      ScaleWidth      =   2955
      TabIndex        =   21
      Top             =   240
      Visible         =   0   'False
      Width           =   3015
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Photo"
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
         Height          =   255
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   480
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Photos"
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
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   480
         Width           =   855
      End
      Begin VB.Frame FramePhotos 
         BackColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   120
         TabIndex        =   41
         Top             =   720
         Visible         =   0   'False
         Width           =   2775
         Begin VB.CommandButton Command7 
            BackColor       =   &H000080FF&
            Caption         =   "Path Files"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   120
            Width           =   1215
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00FF80FF&
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   44
            ToolTipText     =   "Speed By Secound"
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton Command8 
            BackColor       =   &H0000FF00&
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1320
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   1440
            Width           =   615
         End
         Begin VB.CommandButton Command9 
            BackColor       =   &H0000FFFF&
            Caption         =   "Apply"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2040
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   1440
            Width           =   615
         End
         Begin VB.Timer Timer 
            Enabled         =   0   'False
            Interval        =   5000
            Left            =   240
            Top             =   1320
         End
         Begin VB.FileListBox File1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1230
            Left            =   120
            Pattern         =   "*.jpg;*.gif;*.bmp"
            TabIndex        =   45
            Top             =   480
            Width           =   1215
         End
         Begin VB.Image Imagemo 
            BorderStyle     =   1  'Fixed Single
            Height          =   900
            Left            =   1440
            Stretch         =   -1  'True
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.Frame Framephoto 
         BackColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   120
         TabIndex        =   35
         Top             =   720
         Width           =   2775
         Begin VB.CommandButton Command10 
            BackColor       =   &H00FF00FF&
            Caption         =   "...Reset..."
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1320
            TabIndex        =   47
            Top             =   720
            Width           =   1335
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H0000FFFF&
            Caption         =   "Apply"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2040
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   1440
            Width           =   615
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H0000FF00&
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1320
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   1440
            Width           =   615
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H000000FF&
            Caption         =   "Select"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2040
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox Textpicfile 
            BackColor       =   &H000000FF&
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   240
            Width           =   1815
         End
         Begin VB.Image Imagetz 
            BorderStyle     =   1  'Fixed Single
            Height          =   975
            Left            =   120
            Picture         =   "Form1.frx":CF4C
            Stretch         =   -1  'True
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.Image cmdend 
         Height          =   375
         Left            =   2520
         MouseIcon       =   "Form1.frx":626FE
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":65948
         Stretch         =   -1  'True
         ToolTipText     =   "Close"
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         Caption         =   "Properties"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   3015
      End
   End
   Begin VB.ListBox ListFolders 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1410
      Left            =   4200
      TabIndex        =   89
      Top             =   5520
      Width           =   1095
   End
   Begin VB.DirListBox Folders 
      Height          =   765
      Left            =   5760
      TabIndex        =   87
      Top             =   3600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox pturnoff 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   5760
      ScaleHeight     =   1695
      ScaleWidth      =   2895
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   2895
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2040
         TabIndex        =   11
         Top             =   1380
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   2100
         MouseIcon       =   "Form1.frx":65AA0
         MousePointer    =   99  'Custom
         TabIndex        =   10
         ToolTipText     =   "Refresh"
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   520
         MouseIcon       =   "Form1.frx":68CEA
         MousePointer    =   99  'Custom
         TabIndex        =   9
         ToolTipText     =   "Minisize"
         Top             =   690
         Width           =   255
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1320
         MouseIcon       =   "Form1.frx":6BF34
         MousePointer    =   99  'Custom
         TabIndex        =   8
         ToolTipText     =   "End"
         Top             =   720
         Width           =   255
      End
      Begin VB.Image Image4 
         Height          =   1695
         Left            =   0
         Picture         =   "Form1.frx":6F17E
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2895
      End
   End
   Begin VB.ListBox ppro 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   3165
      ItemData        =   "Form1.frx":73D66
      Left            =   8400
      List            =   "Form1.frx":73D79
      MouseIcon       =   "Form1.frx":73DB7
      MousePointer    =   99  'Custom
      Sorted          =   -1  'True
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox Pro3 
      BackColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   960
      MouseIcon       =   "Form1.frx":74681
      MousePointer    =   99  'Custom
      ScaleHeight     =   2595
      ScaleWidth      =   2955
      TabIndex        =   25
      Top             =   4920
      Visible         =   0   'False
      Width           =   3015
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   120
         Max             =   0
         TabIndex        =   57
         Top             =   2280
         Width           =   2535
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1335
         Left            =   2640
         Max             =   0
         TabIndex        =   56
         Top             =   960
         Width           =   255
      End
      Begin VB.Frame Frame 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   120
         TabIndex        =   54
         Top             =   960
         Width           =   2535
         Begin VB.PictureBox Picturepaint 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Height          =   975
            Left            =   0
            ScaleHeight     =   915
            ScaleWidth      =   1680
            TabIndex        =   55
            Top             =   0
            Width           =   1740
            Begin VB.Label LblPaint 
               BackColor       =   &H000000FF&
               Height          =   135
               Left            =   1560
               MouseIcon       =   "Form1.frx":74F4B
               MousePointer    =   99  'Custom
               TabIndex        =   58
               Top             =   840
               Width           =   255
            End
         End
      End
      Begin VB.OptionButton Optionda 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1680
         Picture         =   "Form1.frx":75815
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Circle"
         Top             =   480
         Width           =   375
      End
      Begin MSComDlg.CommonDialog CommonDialogp 
         Left            =   1080
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.ComboBox border 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form1.frx":76657
         Left            =   2160
         List            =   "Form1.frx":76659
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   480
         Width           =   615
      End
      Begin VB.OptionButton Optionmo 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1320
         Picture         =   "Form1.frx":7665B
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Rectangle"
         Top             =   480
         Width           =   375
      End
      Begin VB.OptionButton Optionline 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   960
         Picture         =   "Form1.frx":7749D
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Line"
         Top             =   480
         Width           =   375
      End
      Begin VB.OptionButton Optionpen 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   600
         Picture         =   "Form1.frx":782DF
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Pen"
         Top             =   480
         Value           =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdcolor 
         BackColor       =   &H00000000&
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Color"
         Top             =   480
         Width           =   375
      End
      Begin VB.Image Image12 
         Height          =   255
         Left            =   840
         Picture         =   "Form1.frx":79121
         Stretch         =   -1  'True
         Top             =   60
         Width           =   255
      End
      Begin VB.Image Image10 
         Height          =   255
         Left            =   480
         Picture         =   "Form1.frx":796A8
         Stretch         =   -1  'True
         Top             =   60
         Width           =   255
      End
      Begin VB.Image Itemnew 
         Height          =   255
         Left            =   120
         Picture         =   "Form1.frx":79BDF
         Stretch         =   -1  'True
         Top             =   60
         Width           =   255
      End
      Begin VB.Image Image7 
         Height          =   375
         Left            =   2520
         MouseIcon       =   "Form1.frx":7A16F
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":7D3B9
         Stretch         =   -1  'True
         ToolTipText     =   "Close"
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00FF00FF&
         Caption         =   "Paint"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   3015
      End
   End
   Begin VB.PictureBox hlp 
      BackColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   5640
      MouseIcon       =   "Form1.frx":7D511
      MousePointer    =   99  'Custom
      ScaleHeight     =   2595
      ScaleWidth      =   2955
      TabIndex        =   84
      Top             =   4920
      Width           =   3015
      Begin VB.Image Imageicon 
         Height          =   300
         Left            =   30
         Picture         =   "Form1.frx":7DDDB
         Stretch         =   -1  'True
         ToolTipText     =   "Icon"
         Top             =   30
         Width           =   300
      End
      Begin VB.Image Imageabout 
         BorderStyle     =   1  'Fixed Single
         Height          =   1995
         Left            =   100
         MouseIcon       =   "Form1.frx":81025
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":8426F
         Stretch         =   -1  'True
         ToolTipText     =   "About"
         Top             =   495
         Width           =   2745
      End
      Begin VB.Image Image14 
         Height          =   375
         Left            =   2520
         MouseIcon       =   "Form1.frx":8B58E
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":8E7D8
         Stretch         =   -1  'True
         ToolTipText     =   "Close"
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "About Program"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   85
         Top             =   0
         Width           =   3000
      End
   End
   Begin VB.PictureBox pstart 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2480
      Left            =   8760
      MouseIcon       =   "Form1.frx":8E930
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":8F1FA
      ScaleHeight     =   2475
      ScaleWidth      =   1500
      TabIndex        =   13
      Top             =   3360
      Visible         =   0   'False
      Width           =   1500
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF80FF&
         Height          =   255
         Left            =   120
         TabIndex        =   86
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Computer"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF80FF&
         Height          =   255
         Left            =   120
         TabIndex        =   83
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Run"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF80FF&
         Height          =   260
         Left            =   120
         TabIndex        =   17
         Top             =   1650
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "All Programs"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lblusername 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "User Name"
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Shutdown"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2220
         Width           =   1095
      End
      Begin VB.Image Image3 
         Height          =   255
         Left            =   180
         Picture         =   "Form1.frx":8F461
         Stretch         =   -1  'True
         Top             =   2220
         Width           =   255
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   5040
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   10395
      TabIndex        =   4
      Top             =   8100
      Width           =   10395
      Begin VB.Label TPro 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   600
         TabIndex        =   6
         Top             =   0
         Width           =   165
      End
      Begin VB.Image Image9 
         Height          =   375
         Left            =   5520
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Lbltime 
         Alignment       =   2  'Center
         BackColor       =   &H00FF80FF&
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   4440
         TabIndex        =   5
         Top             =   0
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   375
         Left            =   0
         Picture         =   "Form1.frx":8F9EA
         Stretch         =   -1  'True
         ToolTipText     =   "Start"
         Top             =   0
         Width           =   375
      End
   End
   Begin MSComDlg.CommonDialog CommonDialogtz 
      Left            =   4920
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox pmsgbox 
      BackColor       =   &H00FFC0FF&
      Height          =   1455
      Left            =   5880
      MouseIcon       =   "Form1.frx":8FEA9
      MousePointer    =   99  'Custom
      ScaleHeight     =   1395
      ScaleWidth      =   1755
      TabIndex        =   0
      Top             =   2040
      Visible         =   0   'False
      Width           =   1815
      Begin VB.TextBox textipmsgbox 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF9900&
         Height          =   615
         Left            =   80
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   2
         Top             =   420
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image Image6 
         Height          =   375
         Left            =   1320
         MouseIcon       =   "Form1.frx":90773
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":939BD
         Stretch         =   -1  'True
         ToolTipText     =   "Close"
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Lblmb 
         BackColor       =   &H00FF9900&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1815
      End
   End
   Begin VB.PictureBox proruner 
      BackColor       =   &H0080C0FF&
      Height          =   1575
      Left            =   5760
      MouseIcon       =   "Form1.frx":93B15
      MousePointer    =   99  'Custom
      ScaleHeight     =   1515
      ScaleWidth      =   1995
      TabIndex        =   27
      Top             =   1920
      Width           =   2055
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sel Pro"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Run"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox textrun 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   1215
      End
      Begin MSComDlg.CommonDialog CommonDialogrun 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Image Image8 
         Height          =   375
         Left            =   1560
         MouseIcon       =   "Form1.frx":943DF
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":97629
         Stretch         =   -1  'True
         ToolTipText     =   "Close"
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "Runer"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.PictureBox Pro2 
      BackColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   840
      MouseIcon       =   "Form1.frx":97781
      MousePointer    =   99  'Custom
      ScaleHeight     =   2595
      ScaleWidth      =   2955
      TabIndex        =   23
      Top             =   4800
      Visible         =   0   'False
      Width           =   3015
      Begin VB.TextBox Textc 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   81
         Top             =   480
         Width           =   2775
      End
      Begin VB.Frame FrameCOL 
         BackColor       =   &H00FFC0FF&
         Height          =   1455
         Left            =   370
         TabIndex        =   59
         Top             =   840
         Width           =   2175
         Begin VB.CommandButton Cmdc 
            BackColor       =   &H00FFC0FF&
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   720
            Style           =   1  'Graphical
            TabIndex        =   80
            Top             =   360
            Width           =   375
         End
         Begin VB.CommandButton CmdC3 
            BackColor       =   &H00FFFF00&
            Caption         =   "tan"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   79
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton CmdC3 
            BackColor       =   &H00FFFF00&
            Caption         =   "cos"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   78
            Top             =   720
            Width           =   375
         End
         Begin VB.CommandButton CmdC3 
            BackColor       =   &H00FFFF00&
            Caption         =   "sin"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   77
            Top             =   360
            Width           =   375
         End
         Begin VB.CommandButton Cmdc 
            BackColor       =   &H00FFC0FF&
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   720
            Style           =   1  'Graphical
            TabIndex        =   76
            Top             =   0
            Width           =   375
         End
         Begin VB.CommandButton Cmdc 
            BackColor       =   &H00FFC0FF&
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1080
            Style           =   1  'Graphical
            TabIndex        =   75
            Top             =   0
            Width           =   375
         End
         Begin VB.CommandButton Cmdc 
            BackColor       =   &H00FFC0FF&
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   74
            Top             =   0
            Width           =   375
         End
         Begin VB.CommandButton Cmdc 
            BackColor       =   &H00FFC0FF&
            Caption         =   "5"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   1080
            Style           =   1  'Graphical
            TabIndex        =   73
            Top             =   360
            Width           =   375
         End
         Begin VB.CommandButton Cmdc 
            BackColor       =   &H00FFC0FF&
            Caption         =   "6"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   72
            Top             =   360
            Width           =   375
         End
         Begin VB.CommandButton Cmdc 
            BackColor       =   &H00FFC0FF&
            Caption         =   "7"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   6
            Left            =   720
            Style           =   1  'Graphical
            TabIndex        =   71
            Top             =   720
            Width           =   375
         End
         Begin VB.CommandButton Cmdc 
            BackColor       =   &H00FFC0FF&
            Caption         =   "8"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   7
            Left            =   1080
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   720
            Width           =   375
         End
         Begin VB.CommandButton Cmdc 
            BackColor       =   &H00FFC0FF&
            Caption         =   "9"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   8
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   69
            Top             =   720
            Width           =   375
         End
         Begin VB.CommandButton Cmdc 
            BackColor       =   &H00FFC0FF&
            Caption         =   "."
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   9
            Left            =   720
            Style           =   1  'Graphical
            TabIndex        =   68
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton Cmdc 
            BackColor       =   &H00FFC0FF&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   10
            Left            =   1080
            Style           =   1  'Graphical
            TabIndex        =   67
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton Cmdc2 
            BackColor       =   &H0000FFFF&
            Caption         =   "/"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   360
            Width           =   375
         End
         Begin VB.CommandButton Cmdc2 
            BackColor       =   &H0000FFFF&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   65
            Top             =   720
            Width           =   375
         End
         Begin VB.CommandButton Command13 
            BackColor       =   &H00FF9900&
            Caption         =   "="
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   64
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton Cmdc2 
            BackColor       =   &H0000FFFF&
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   63
            Top             =   0
            Width           =   375
         End
         Begin VB.CommandButton Cmdc2 
            BackColor       =   &H0000FFFF&
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   62
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton Command12 
            BackColor       =   &H000000FF&
            Caption         =   "C"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   0
            Width           =   375
         End
         Begin VB.CommandButton Command11 
            BackColor       =   &H0000FF00&
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.Image Image5 
         Height          =   375
         Left            =   2520
         MouseIcon       =   "Form1.frx":9804B
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":9B295
         Stretch         =   -1  'True
         ToolTipText     =   "Close"
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00FF00FF&
         Caption         =   "Calculator"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   3015
      End
   End
   Begin VB.PictureBox PRONOTEPAD 
      BackColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   720
      MouseIcon       =   "Form1.frx":9B3ED
      MousePointer    =   99  'Custom
      ScaleHeight     =   2595
      ScaleWidth      =   2955
      TabIndex        =   18
      Top             =   4680
      Visible         =   0   'False
      Width           =   3015
      Begin VB.TextBox TextNOTAPAD 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   90
         MousePointer    =   3  'I-Beam
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   19
         Top             =   480
         Width           =   2775
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   1680
         Top             =   1320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Image Image11 
         Height          =   375
         Left            =   2520
         MouseIcon       =   "Form1.frx":9BCB7
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":9EF01
         Stretch         =   -1  'True
         ToolTipText     =   "Close"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image CMDOPEN 
         Height          =   255
         Left            =   420
         Picture         =   "Form1.frx":9F059
         Stretch         =   -1  'True
         ToolTipText     =   "Open"
         Top             =   60
         Width           =   255
      End
      Begin VB.Image CMDSAVE 
         Height          =   255
         Left            =   720
         Picture         =   "Form1.frx":9F590
         Stretch         =   -1  'True
         ToolTipText     =   "Save"
         Top             =   60
         Width           =   255
      End
      Begin VB.Image CMDNEW 
         Height          =   255
         Left            =   120
         Picture         =   "Form1.frx":9FB17
         Stretch         =   -1  'True
         ToolTipText     =   "New"
         Top             =   60
         Width           =   255
      End
      Begin VB.Label f 
         Alignment       =   2  'Center
         BackColor       =   &H00FF00FF&
         Caption         =   "Note Pad"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   3015
      End
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Folders"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   88
      Top             =   3960
      Width           =   615
   End
   Begin VB.Image ImageFolder 
      Height          =   375
      Left            =   120
      Picture         =   "Form1.frx":A00A7
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "My Computer"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   0
      TabIndex        =   82
      Top             =   600
      Width           =   780
   End
   Begin VB.Image Image13 
      Height          =   495
      Left            =   120
      Picture         =   "Form1.frx":A3469
      Stretch         =   -1  'True
      ToolTipText     =   "My Computer"
      Top             =   120
      Width           =   495
   End
   Begin VB.Label q2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Paint"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   90
      TabIndex        =   33
      Top             =   2880
      Width           =   405
   End
   Begin VB.Label q1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note Pad"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   0
      TabIndex        =   32
      Top             =   1800
      Width           =   690
   End
   Begin VB.Image iFarhangLogat 
      Height          =   375
      Left            =   120
      Picture         =   "Form1.frx":A6C80
      Stretch         =   -1  'True
      ToolTipText     =   "Paint"
      Top             =   2520
      Width           =   375
   End
   Begin VB.Image INOTEPAD 
      Height          =   375
      Left            =   120
      Picture         =   "Form1.frx":AC9DA
      Stretch         =   -1  'True
      ToolTipText     =   "Note Pad"
      Top             =   1440
      Width           =   375
   End
   Begin VB.Menu a 
      Caption         =   "a"
      Visible         =   0   'False
      Begin VB.Menu ItemRefreshComputer 
         Caption         =   "Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu w3 
         Caption         =   "-"
      End
      Begin VB.Menu Itemnew1 
         Caption         =   "New"
         WindowList      =   -1  'True
         Begin VB.Menu ItemNewFolder 
            Caption         =   "Folder"
         End
         Begin VB.Menu w2 
            Caption         =   "-"
         End
         Begin VB.Menu ItemNewTextDocument 
            Caption         =   "New Text Document"
         End
         Begin VB.Menu ItemNewBitmapImage 
            Caption         =   "New Bitmap Image"
         End
      End
      Begin VB.Menu w1 
         Caption         =   "-"
      End
      Begin VB.Menu IetmProperties 
         Caption         =   "Properties"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim T00 As Integer, xx As Single, yy As Single
Dim numbernewFolder As Integer
Private Sub border_Click()
Picturepaint.DrawWidth = border
End Sub

Private Sub Cmdc_Click(Index As Integer)
Call Module.Cmdc_Click(Index)
End Sub
Private Sub Cmdc2_Click(Index As Integer)
Call Module.Cmdc2_Click(Index)
End Sub

Private Sub CmdC3_Click(Index As Integer)
Call Module.CmdC3_Click(Index)
End Sub

Private Sub cmdcolor_Click()
CommonDialogp.ShowColor
cmdcolor.BackColor = CommonDialogp.Color
Picturepaint.ForeColor = CommonDialogp.Color
End Sub
Private Sub cmdend_Click()
TPro = TPro - 1
pro1.Visible = False
pro1.AutoRedraw = False
End Sub
Private Sub CMDNEW_Click()
If pmsgbox.Visible = True Then Exit Sub
pmsgbox.Visible = True
pmsgbox.Move 1500, 1500
textipmsgbox = "New Text"
TPro = TPro + 1
End Sub
Private Sub CMDOPEN_Click()
TextNOTAPAD.Text = ""
CommonDialog1.Filter = "(*.txt)|*.txt|(*.html)|*.html(*.*)|*.*"
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
T = FreeFile
Open CommonDialog1.FileName For Input As #T
Do While Not EOF(T)
Line Input #T, r
TextNOTAPAD.Text = TextNOTAPAD.Text + r + vbNewLine
Loop
Close #T
End If
End Sub

Private Sub CMDSAVE_Click()
CommonDialog1.Filter = "(*.txt)|*.txt|(*.html)|*.html|(*.php)|*.php|(*.*)|*.*"
CommonDialog1.ShowSave
If CommonDialog1.FileName <> "" Then
r = FreeFile
Open CommonDialog1.FileName For Output As #r
Print #r, TextNOTAPAD
Close #r
End If
End Sub
Private Sub Combo1_Click()
On Error GoTo 1
Timer.Interval = Combo1.List(Combo1.ListIndex) * 1000
1:
End Sub

Private Sub Command1_Click()
Moduleother.tz
End Sub

Private Sub Command10_Click()

Imagetz.Picture = Picture
Textpicfile = ""
End Sub

Private Sub Command11_Click()
Call Module.C3_Click
End Sub

Private Sub Command12_Click()
Call Module.Command12_Click
End Sub

Private Sub Command13_Click()
Module.C_Click
End Sub

Private Sub Command2_Click()
On Error GoTo 1
Me.PaintPicture Imagetz.Picture, 0, 0, ScaleWidth, ScaleHeight
cmdend_Click
1:
End Sub

Private Sub Command3_Click()
On Error GoTo 1
Me.PaintPicture Imagetz.Picture, 0, 0, ScaleWidth, ScaleHeight
1:
End Sub

Private Sub Command4_Click()
If textipmsgbox = "New Text" Then TextNOTAPAD.Text = ""
If textipmsgbox = "New Picture" Then Picturepaint.Picture = LoadPicture()
Image6_Click
End Sub

Private Sub Command5_Click()
On Error Resume Next
If UCase(textrun) = "NOTEPAD" Or UCase(textrun) = "PAINT" Or UCase(textrun) = UCase("Calculator") Then
If UCase(textrun) = "NOTEPAD" Then INOTEPAD_DblClick
If UCase(textrun) = "PAINT" Then iFarhangLogat_DblClick
If UCase(textrun) = UCase("Calculator") Then
ppro.ListIndex = 0
ppro_DblClick
End If
textrun = ""
Exit Sub
End If
Shell textrun, vbNormalFocus
textrun = ""
End Sub

Private Sub Command6_Click()
CommonDialogrun.Filter = "Programs|*.exe|AllFiles|*.*"
CommonDialogrun.ShowOpen
If CommonDialogrun.FileName <> "" Then textrun.Text = CommonDialogrun.FileName
End Sub

Private Sub Command7_Click()
Moduleother.tzs
End Sub

Private Sub Command8_Click()
Timer.Enabled = True
cmdend_Click
T00 = -1
End Sub

Private Sub Command9_Click()
Timer.Enabled = True
T00 = -1
End Sub

Private Sub f_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage PRONOTEPAD.hWnd, &HA1, 2, 0&
If Button = 1 Then
If PRONOTEPAD.Top < 0 Then PRONOTEPAD.Move 0, 0, ScaleWidth, ScaleHeight
If PRONOTEPAD.Top > 0 Then PRONOTEPAD.Move PRONOTEPAD.Left, PRONOTEPAD.Top, 3015, 2655
If PRONOTEPAD.Left + X >= Me.ScaleWidth - 120 Then PRONOTEPAD.Move Me.ScaleWidth / 2, 0, PRONOTEPAD.ScaleWidth, Me.ScaleHeight
If PRONOTEPAD.Left + X <= 120 Then PRONOTEPAD.Move 0, 0, Me.ScaleWidth / 2, Me.ScaleHeight
End If
f.Move 0, 0, PRONOTEPAD.Width
Image11.Move PRONOTEPAD.Width - 500
TextNOTAPAD.Move TextNOTAPAD.Left, TextNOTAPAD.Top, PRONOTEPAD.ScaleWidth - 180, PRONOTEPAD.ScaleHeight - 540
End Sub

Private Sub File1_Click()
On Error GoTo 1
Imagemo.Picture = LoadPicture(File1.Path & "\" & File1.List(File1.ListIndex))
1:
End Sub

Private Sub Form_Click()

pstart.Visible = False
ppro.Visible = False
If pturnoff.Visible = True Or pmsgbox.Visible = True Then Beep
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = vbAltMask And KeyCode = vbKeyF4 Then Label1_Click
If KeyCode = vbKeyF1 Then Label15_Click
End Sub

Private Sub Form_Load()
Caption = App.CompanyName
Module.fl
Timer1_Timer
Moduleother.usename
Module.AllFolder
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ppro.Visible = True Then ppro.Visible = False
If ListFolders.Visible = True Then ListFolders.Visible = False
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu a

End Sub

Private Sub HScroll1_Change()
Picturepaint.Left = -(HScroll1 * 200)
End Sub

Private Sub IetmProperties_Click()
Moduleppro.Proper
Timer.Enabled = False
End Sub
Private Sub iFarhangLogat_DblClick()
ppro = "Paint"
Call Moduleppro.LISTPRO
End Sub

Private Sub Image1_Click()
If ListFolders.Visible = True Then ListFolders.Visible = False
pstart.Visible = Not pstart.Visible
pstart.Move 0, 1750

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.BorderStyle = 1
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moduleother.MM
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.BorderStyle = 0
End Sub

Private Sub Image10_Click()
CommonDialogp.FileName = ""
CommonDialogp.Filter = "(*.jpg)|*.jpg|(*.gif)|*.gif|(*.ico)|*.ico|(*.bmp)|*.bmp|All Files|*.*"
CommonDialogp.ShowOpen
If CommonDialogp.FileName <> "" Then
Picturepaint.Picture = LoadPicture(CommonDialogp.FileName)
If Picturepaint.Height > Frame.Height Then VScroll1.Max = Int((Picturepaint.Height - Frame.Height) / 200)
If Picturepaint.Width > Frame.Width Then HScroll1.Max = Int((Picturepaint.Width - Frame.Width) / 200)
LblPaint.Move Picturepaint.Width - 120, Picturepaint.Height - 120
End If
End Sub

Private Sub Image11_Click()
TPro = TPro - 1
PRONOTEPAD.Visible = False
TextNOTAPAD.Text = ""
PRONOTEPAD.AutoRedraw = False
End Sub

Private Sub Image12_Click()
CommonDialogp.FileName = ""
CommonDialogp.Filter = "(*.jpg)|*.jpg|(*.gif)|*.gif|(*.ico)|*.ico|(*.bmp)|*.bmp|All Files|*.*"
CommonDialogp.ShowSave
If CommonDialogp.FileName <> "" Then SavePicture Picturepaint.Image, CommonDialogp.FileName
End Sub

Private Sub Image13_DblClick()
Shell "Explorer.exe", vbNormalFocus
End Sub

Private Sub Image14_Click()
TPro = TPro - 1
hlp.AutoRedraw = False
hlp.Visible = False
End Sub



Private Sub Image5_Click()
TPro = TPro - 1
Pro2.Visible = False
Pro2.AutoRedraw = False
Textc = ""
End Sub

Private Sub Image6_Click()
pmsgbox.Visible = False
pmsgbox.AutoRedraw = False
TPro = TPro - 1
End Sub

Private Sub Image7_Click()
TPro = TPro - 1
Pro3.AutoRedraw = False
Pro3.Visible = False
Picturepaint.Picture = LoadPicture()
End Sub

Private Sub Image9_Click()
Moduleother.allpro
End Sub

Private Sub Image8_Click()
TPro = TPro - 1
textrun = ""
proruner.Visible = False
proruner.AutoRedraw = False
End Sub

Private Sub Image9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image9.BorderStyle = 1
End Sub

Private Sub Image9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image9.BorderStyle = 0
End Sub
Private Sub Imageabout_Click()
SavePicture Imageabout.Picture, "C:\Documents and Settings\All Users\Desktop\WindowsSaman2013.gif"

End Sub
Private Sub ImageFolder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ListFolders.ListCount <> 0 Then
ListFolders.Visible = True
ListFolders.Move X + ImageFolder.Left, Y + ImageFolder.Top - ListFolders.Height
End If
End Sub
Private Sub INOTEPAD_DblClick()
ppro = "Note Pad"
Call Moduleppro.LISTPRO
End Sub

Private Sub Itemnew_Click()
If pmsgbox.Visible = True Then Exit Sub
pmsgbox.Visible = True
pmsgbox.Move 1500, 1500
textipmsgbox = "New Picture"
TPro = TPro + 1
End Sub

Private Sub ItemNewBitmapImage_Click()
Randomize Timer
r = FreeFile
Open "C:\Documents and Settings\All Users\Desktop\New Bitmap Image" & Rnd * 10000 & ".bmp" For Output As #r
Close #r
End Sub

Private Sub ItemNewFolder_Click()
numbernewFolder = numbernewFolder + 1
newfolder = "New Folder " & numbernewFolder
GoTo 1
2:
Call ItemNewFolder_Click
Exit Sub
1:
On Error GoTo 2
MkDir "C:\Documents and Settings\" & lblusername & "\Desktop\" & newfolder
Module.AllFolder
End Sub

Private Sub ItemNewTextDocument_Click()
Module.NewTextDocument

End Sub

Private Sub ItemRefreshComputer_Click()
Module.AllFolder
End Sub

Private Sub Label1_Click()
pturnoff.Visible = Not pturnoff.Visible
pturnoff.Move 1320, 1080
Call Form_Click
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.BorderStyle = 1
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage proruner.hWnd, &HA1, 2, 0&
End Sub



Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage Pro3.hWnd, &HA1, 2, 0&
If Button = 1 Then
If Pro3.Top < 0 Then Pro3.Move 0, 0, ScaleWidth, ScaleHeight
If Pro3.Top > 0 Then Pro3.Move Pro3.Left, Pro3.Top, 3015, 2655
If Pro3.Left + X >= Me.ScaleWidth - 120 Then Pro3.Move Me.ScaleWidth / 2, 0, Pro3.ScaleWidth, Me.ScaleHeight
If Pro3.Left + X <= 120 Then Pro3.Move 0, 0, Me.ScaleWidth / 2, Me.ScaleHeight
End If
Label12.Width = Pro3.Width
Image7.Left = Pro3.Width - 500
Frame.Move 120, Frame.Top, Pro3.Width - 540, Pro3.Height - 340 - Frame.Top
Module.Ppaint
End Sub

Private Sub Label13_Click()
Image13_DblClick
Form_Click
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label13.FontUnderline = True
End Sub

Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moduleother.MM
End Sub

Private Sub Label13_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label13.FontUnderline = False
End Sub

Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage hlp.hWnd, &HA1, 2, 0&
If Button = 1 Then
If hlp.Top < 0 Then hlp.Move 0, 0, ScaleWidth, ScaleHeight
If hlp.Top > 0 Then hlp.Move hlp.Left, hlp.Top, 3015, 2655
If hlp.Left + X >= Me.ScaleWidth - 120 Then hlp.Move Me.ScaleWidth / 2, 0, hlp.ScaleWidth, Me.ScaleHeight
If hlp.Left + X <= 120 Then hlp.Move 0, 0, Me.ScaleWidth / 2, Me.ScaleHeight
End If
Label14.Width = hlp.ScaleWidth
Image14.Left = hlp.Width - 500
Imageabout.Move Imageabout.Left, Imageabout.Top, hlp.Width - 260, hlp.Height - 660
End Sub

Private Sub Label15_Click()
Moduleppro.AboutPro
Form_Click
End Sub

Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label15.FontUnderline = True
End Sub

Private Sub Label15_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label15.FontUnderline = False
End Sub

Private Sub Label16_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ListFolders.Visible = True Then ListFolders.Visible = False

End Sub

Private Sub Label2_Click()
End
End Sub
Private Sub Label3_Click()
Form1.WindowState = 1
pturnoff.Visible = False
End Sub

Private Sub Label4_Click()
Unload Me
Show
End Sub

Private Sub Picture2_Click()

End Sub

Private Sub Label5_Click()
pturnoff.Visible = False
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage pro1.hWnd, &HA1, 2, 0&

End Sub

Private Sub Label7_Click()
Call Form_Click
If proruner.AutoRedraw = True Then Exit Sub
proruner.AutoRedraw = True
proruner.Visible = True
proruner.Move 1700, 1500
TPro = TPro + 1
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.FontUnderline = True
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moduleother.MM
End Sub

Private Sub Label7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.FontUnderline = False
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage Pro2.hWnd, &HA1, 2, 0&
If Button = 1 Then
If Pro2.Top < 0 Then Pro2.Move 0, 0, ScaleWidth, ScaleHeight
If Pro2.Top > 0 Then Pro2.Move Pro2.Left, Pro2.Top, 3015, 2655
If Pro2.Left + X >= Me.ScaleWidth - 120 Then Pro2.Move Me.ScaleWidth / 2, 0, Pro2.ScaleWidth, Me.ScaleHeight
If Pro2.Left + X <= 120 Then Pro2.Move 0, 0, Me.ScaleWidth / 2, Me.ScaleHeight
End If
Label8.Width = Pro2.Width
Image5.Left = Pro2.Width - 500
Call Module.Col
End Sub

Private Sub Label9_Click()
ppro.Visible = Not ppro.Visible
ppro.Move 1440, 1080

End Sub

Private Sub Text1_Change()

End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.BorderStyle = 1
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moduleother.MM
End Sub

Private Sub Label9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.BorderStyle = 0
End Sub

Private Sub Lblmb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage pmsgbox.hWnd, &HA1, 2, 0&
End Sub




Private Sub LblPaint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then LblPaint.Move LblPaint.Left + X, LblPaint.Top + Y
Picturepaint.Move Picturepaint.Left, Picturepaint.Top, LblPaint.Left + 120, LblPaint.Top + 120
Module.Ppaint
End Sub

Private Sub lblusername_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
If pmsgbox.AutoRedraw = True Then Exit Sub
TPro = TPro + 1

pmsgbox.AutoRedraw = True
pmsgbox.Visible = True
pmsgbox.Move 1500, 1500
Lblmb.Caption = "User Name"
textipmsgbox = "User Name is : " & lblusername
End If
End Sub

Private Sub ListFolders_DblClick()
WinExec "Explorer.exe " & Folders.List(ListFolders.ListIndex), 10
ListFolders.Visible = False
End Sub
Private Sub Option1_Click()
Framephoto.Visible = True
FramePhotos.Visible = False
End Sub
Private Sub Option2_Click()
Framephoto.Visible = False
FramePhotos.Visible = True
Timer.Enabled = False
End Sub

Private Sub Picturepaint_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xx = X
yy = Y
End Sub

Private Sub Picturepaint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Module.PPaint_m_m Button, X, Y, xx, yy
End Sub

Private Sub Picturepaint_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Module.PPaint_m_u Button, X, Y, xx, yy
End Sub

Private Sub ppro_DblClick()
Call Moduleppro.LISTPRO
Call Form_Click
End Sub

Private Sub pstart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moduleother.MM
End Sub


Private Sub textrun_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Command5_Click
End Sub

Private Sub Timer_Timer()
T00 = T00 + 1
On Error GoTo 1
Imagemo.Picture = LoadPicture(File1.Path & "\" & File1.List(T00))
Me.PaintPicture Imagemo.Picture, 0, 0, ScaleWidth, ScaleHeight
If T00 >= File1.ListCount - 1 Then T00 = 0
1:
End Sub

Private Sub Timer1_Timer()
Lbltime.Caption = Format(Time, "short time")
Lbltime.ToolTipText = "Date is : " & Date
If pturnoff.Visible = True Then pturnoff.ZOrder
End Sub

Private Sub VScroll1_Change()
Picturepaint.Top = -(VScroll1 * 200)
End Sub
