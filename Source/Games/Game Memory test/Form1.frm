VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Game For Windows"
   ClientHeight    =   8610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12480
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form1.frx":324A
   MousePointer    =   99  'Custom
   ScaleHeight     =   8610
   ScaleWidth      =   12480
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   10800
      Top             =   720
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FF00FF&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   10080
      MouseIcon       =   "Form1.frx":3B14
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "About"
      Top             =   30
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000C000&
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   10680
      MouseIcon       =   "Form1.frx":6D5E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "New Game"
      Top             =   30
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF0000&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   11280
      MouseIcon       =   "Form1.frx":9FA8
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "MiniSize Window"
      Top             =   30
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   11880
      MouseIcon       =   "Form1.frx":D1F2
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Exit Game"
      Top             =   30
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   11280
      Top             =   720
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   39
      Left            =   9480
      Picture         =   "Form1.frx":1043C
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   39
      Top             =   6960
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   39
         Left            =   0
         Picture         =   "Form1.frx":14ED1
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   38
      Left            =   8160
      Picture         =   "Form1.frx":1D5BB
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   38
      Top             =   6960
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   38
         Left            =   0
         Picture         =   "Form1.frx":22050
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   37
      Left            =   6840
      Picture         =   "Form1.frx":2A73A
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   37
      Top             =   6960
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   37
         Left            =   0
         Picture         =   "Form1.frx":2E19F
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   36
      Left            =   5520
      Picture         =   "Form1.frx":36889
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   36
      Top             =   6960
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   36
         Left            =   0
         Picture         =   "Form1.frx":3A2EE
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   35
      Left            =   4200
      Picture         =   "Form1.frx":429D8
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   35
      Top             =   6960
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   35
         Left            =   0
         Picture         =   "Form1.frx":467CC
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   34
      Left            =   2880
      Picture         =   "Form1.frx":4EEB6
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   34
      Top             =   6960
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   34
         Left            =   0
         Picture         =   "Form1.frx":52CAA
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   33
      Left            =   1560
      Picture         =   "Form1.frx":5B394
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   33
      Top             =   6960
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   33
         Left            =   0
         Picture         =   "Form1.frx":5C51F
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   32
      Left            =   240
      Picture         =   "Form1.frx":64C09
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   32
      Top             =   6960
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   32
         Left            =   0
         Picture         =   "Form1.frx":65D94
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   31
      Left            =   9480
      Picture         =   "Form1.frx":6E47E
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   31
      Top             =   5400
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   31
         Left            =   0
         Picture         =   "Form1.frx":6F739
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   30
      Left            =   8160
      Picture         =   "Form1.frx":77E23
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   30
      Top             =   5400
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   30
         Left            =   0
         Picture         =   "Form1.frx":790DE
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   29
      Left            =   6840
      Picture         =   "Form1.frx":817C8
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   29
      Top             =   5400
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   29
         Left            =   0
         Picture         =   "Form1.frx":8583D
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   28
      Left            =   5520
      Picture         =   "Form1.frx":8DF27
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   28
      Top             =   5400
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   28
         Left            =   0
         Picture         =   "Form1.frx":91F9C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   27
      Left            =   4200
      Picture         =   "Form1.frx":9A686
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   27
      Top             =   5400
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   27
         Left            =   0
         Picture         =   "Form1.frx":A05CB
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   26
      Left            =   2880
      Picture         =   "Form1.frx":A8CB5
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   26
      Top             =   5400
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   26
         Left            =   0
         Picture         =   "Form1.frx":AEBFA
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   25
      Left            =   1560
      Picture         =   "Form1.frx":B72E4
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   25
      Top             =   5400
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   25
         Left            =   0
         Picture         =   "Form1.frx":BAF20
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   24
      Left            =   240
      Picture         =   "Form1.frx":C360A
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   24
      Top             =   5400
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   24
         Left            =   0
         Picture         =   "Form1.frx":C7246
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   23
      Left            =   9480
      Picture         =   "Form1.frx":CF930
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   23
      Top             =   3840
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   23
         Left            =   0
         Picture         =   "Form1.frx":D44AE
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   22
      Left            =   8160
      Picture         =   "Form1.frx":DCB98
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   22
      Top             =   3840
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   22
         Left            =   0
         Picture         =   "Form1.frx":E1716
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   21
      Left            =   6840
      Picture         =   "Form1.frx":E9E00
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   21
      Top             =   3840
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   21
         Left            =   0
         Picture         =   "Form1.frx":EE895
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   20
      Left            =   5520
      Picture         =   "Form1.frx":F6F7F
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   20
      Top             =   3840
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   20
         Left            =   0
         Picture         =   "Form1.frx":FBA14
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   19
      Left            =   4200
      Picture         =   "Form1.frx":1040FE
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   19
      Top             =   3840
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   19
         Left            =   0
         Picture         =   "Form1.frx":108966
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   18
      Left            =   2880
      Picture         =   "Form1.frx":111050
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   18
      Top             =   3840
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   18
         Left            =   0
         Picture         =   "Form1.frx":1158B8
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   17
      Left            =   1560
      Picture         =   "Form1.frx":11DFA2
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   17
      Top             =   3840
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   17
         Left            =   0
         Picture         =   "Form1.frx":11EFD8
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   16
      Left            =   240
      Picture         =   "Form1.frx":1276C2
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   16
      Top             =   3840
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   16
         Left            =   0
         Picture         =   "Form1.frx":1286F8
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   15
      Left            =   9480
      Picture         =   "Form1.frx":130DE2
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   15
      Top             =   2280
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   15
         Left            =   0
         Picture         =   "Form1.frx":132C0D
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   14
      Left            =   8160
      Picture         =   "Form1.frx":13B2F7
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   14
      Top             =   2280
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   14
         Left            =   0
         Picture         =   "Form1.frx":13D122
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   13
      Left            =   6840
      Picture         =   "Form1.frx":14580C
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   13
      Top             =   2280
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   13
         Left            =   0
         Picture         =   "Form1.frx":146DE0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   12
      Left            =   5520
      Picture         =   "Form1.frx":14F4CA
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   12
      Top             =   2280
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   12
         Left            =   0
         Picture         =   "Form1.frx":150A9E
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   11
      Left            =   4200
      Picture         =   "Form1.frx":159188
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   11
      Top             =   2280
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   11
         Left            =   0
         Picture         =   "Form1.frx":159D15
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   10
      Left            =   2880
      Picture         =   "Form1.frx":1623FF
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   10
      Top             =   2280
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   10
         Left            =   0
         Picture         =   "Form1.frx":162F8C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   9
      Left            =   1560
      Picture         =   "Form1.frx":16B676
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   9
      Top             =   2280
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   9
         Left            =   0
         Picture         =   "Form1.frx":16D39B
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   8
      Left            =   240
      Picture         =   "Form1.frx":175A85
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   8
      Top             =   2280
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   8
         Left            =   0
         Picture         =   "Form1.frx":1777AA
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   7
      Left            =   9480
      Picture         =   "Form1.frx":17FE94
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   7
      Top             =   720
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   7
         Left            =   0
         Picture         =   "Form1.frx":18101F
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   6
      Left            =   8160
      Picture         =   "Form1.frx":189709
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   6
      Top             =   720
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   6
         Left            =   0
         Picture         =   "Form1.frx":18A894
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   5
      Left            =   6840
      Picture         =   "Form1.frx":192F7E
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   5
      Top             =   720
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   5
         Left            =   0
         Picture         =   "Form1.frx":196D72
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   4
      Left            =   5520
      Picture         =   "Form1.frx":19F45C
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   4
      Top             =   720
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   4
         Left            =   0
         Picture         =   "Form1.frx":1A3250
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   3
      Left            =   4200
      Picture         =   "Form1.frx":1AB93A
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   3
      Top             =   720
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   3
         Left            =   0
         Picture         =   "Form1.frx":1ACBCA
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   2
      Left            =   2880
      Picture         =   "Form1.frx":1B52B4
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   2
      Top             =   720
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   2
         Left            =   0
         Picture         =   "Form1.frx":1B6544
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   1
      Left            =   1560
      Picture         =   "Form1.frx":1BEC2E
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   1
      Top             =   720
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   1
         Left            =   0
         Picture         =   "Form1.frx":1BFEE9
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   0
      Left            =   240
      Picture         =   "Form1.frx":1C85D3
      ScaleHeight     =   1455
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   720
      Width           =   1215
      Begin VB.Image Image1 
         Height          =   615
         Index           =   0
         Left            =   0
         Picture         =   "Form1.frx":1C988E
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.Image Image4 
      Height          =   3855
      Left            =   10920
      Picture         =   "Form1.frx":1D1F78
      Stretch         =   -1  'True
      ToolTipText     =   "Programer : Saman Azadi"
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   10920
      TabIndex        =   47
      ToolTipText     =   "Timer"
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   10920
      TabIndex        =   46
      ToolTipText     =   "Select"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "20 / 0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   10920
      TabIndex        =   45
      Top             =   720
      Width           =   1335
   End
   Begin VB.Image Image3 
      Height          =   400
      Left            =   120
      Stretch         =   -1  'True
      Top             =   30
      Width           =   400
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Game For Windows"
      DragIcon        =   "Form1.frx":1F726A
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   0
      MouseIcon       =   "Form1.frx":1F7B34
      TabIndex        =   44
      Top             =   0
      Width           =   12495
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   0
      MouseIcon       =   "Form1.frx":1F83FE
      Picture         =   "Form1.frx":1F8CC8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim text1 As String, text2 As String
Private Sub Command1_Click()
o = MsgBox("Exit Game", vbYesNo + 64, "End")
If o = vbYes Then End
End Sub
Private Sub Command2_Click()
WindowState = vbMinimized
End Sub
Private Sub Command3_Click()
o = MsgBox("New Game", vbYesNo + 64, "New")
If o = vbYes Then Call Module1.Saman
End Sub
Private Sub Command4_Click()
FormA.Show
End Sub

Private Sub Form_Load()
  If App.PrevInstance = True Then
     MsgBox " »—‰«„Â œ—Õ«· «Ã—« »ÊœÂ Ê «„ò«‰ «Ã—«Ì Â„“„«‰ ¬‰ ÊÃÊœ ‰œ«—œ ", vbCritical, "Warning !"
     End
  End If
  Module1.Saman
End Sub

Private Sub Image1_Click(Index As Integer)
Timer2.Enabled = True
If text1 = "" Then
text1 = Index
Image1(text1).Visible = False
Exit Sub
ElseIf text2 = "" Then
text2 = Index
Image1(text2).Visible = False
Timer1.Enabled = True
Label3 = Label3 + 1
End If
End Sub
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.MousePointer = 99
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Module1.S(Me.hWnd)
End Sub
Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.MousePointer = 0
End Sub

Private Sub Timer1_Timer()
For w = 0 To Image1.Count - 1
Image1(w).Visible = True
Next
On Error Resume Next
If Picture1(text1).Point(240, 240) = Picture1(text2).Point(240, 240) Then
Picture1(text1).Visible = False
Picture1(text2).Visible = False


X1 = sndPlaySound(App.Path & "\Saman.wav", 1)
End If
If text1 <> "" And text2 <> "" Then
text1 = ""
text2 = ""
End If
Module1.lo
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
Label4 = Label4 + 1
End Sub
