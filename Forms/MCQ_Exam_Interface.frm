VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamalButton.ocx"
Object = "{08654D78-6636-11D3-87BF-B4980CC10374}#2.0#0"; "MyEllipticButton.ocx"
Begin VB.Form MCQ 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "SPEEDUP TEST SOLUTIONS MCQ TEST "
   ClientHeight    =   10515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20430
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form14"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10515
   ScaleWidth      =   20430
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4Five 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   535
      Left            =   1440
      TabIndex        =   142
      Top             =   5160
      Width           =   14505
      Begin ChamaleonButton.ChameleonBtn sub51 
         Height          =   465
         Left            =   0
         TabIndex        =   143
         ToolTipText     =   "Move Next Question"
         Top             =   15
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   820
         BTYPE           =   1
         TX              =   "English"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   104
         BCOLO           =   104
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "MCQ_Exam_Interface.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn sub52 
         Height          =   465
         Left            =   2880
         TabIndex        =   144
         ToolTipText     =   "Move Next Question"
         Top             =   15
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   820
         BTYPE           =   1
         TX              =   "English"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   7552000
         FCOLO           =   7552000
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "MCQ_Exam_Interface.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn sub53 
         Height          =   465
         Left            =   5760
         TabIndex        =   145
         ToolTipText     =   "Move Next Question"
         Top             =   15
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   820
         BTYPE           =   1
         TX              =   "English"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   7552000
         FCOLO           =   7552000
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "MCQ_Exam_Interface.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn Sub54 
         Height          =   465
         Left            =   8640
         TabIndex        =   146
         ToolTipText     =   "Move Next Question"
         Top             =   15
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   820
         BTYPE           =   1
         TX              =   "English"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   7552000
         FCOLO           =   7552000
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "MCQ_Exam_Interface.frx":0054
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn sub55 
         Height          =   465
         Left            =   11520
         TabIndex        =   132
         ToolTipText     =   "Move Next Question"
         Top             =   15
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   820
         BTYPE           =   1
         TX              =   "English"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   7552000
         FCOLO           =   7552000
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "MCQ_Exam_Interface.frx":0070
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame Frame4Four 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   1440
      TabIndex        =   137
      Top             =   4200
      Width           =   14505
      Begin ChamaleonButton.ChameleonBtn sub44 
         Height          =   465
         Left            =   10800
         TabIndex        =   138
         ToolTipText     =   "Move Next Question"
         Top             =   15
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   820
         BTYPE           =   1
         TX              =   "English"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   7552000
         FCOLO           =   7552000
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "MCQ_Exam_Interface.frx":008C
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn sub43 
         Height          =   465
         Left            =   7200
         TabIndex        =   139
         ToolTipText     =   "Move Next Question"
         Top             =   15
         Width           =   3585
         _ExtentX        =   6324
         _ExtentY        =   820
         BTYPE           =   1
         TX              =   "English"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   7552000
         FCOLO           =   7552000
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "MCQ_Exam_Interface.frx":00A8
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn sub42 
         Height          =   465
         Left            =   3600
         TabIndex        =   140
         ToolTipText     =   "Move Next Question"
         Top             =   15
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   820
         BTYPE           =   1
         TX              =   "English"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   7552000
         FCOLO           =   7552000
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "MCQ_Exam_Interface.frx":00C4
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn sub41 
         Height          =   465
         Left            =   0
         TabIndex        =   141
         ToolTipText     =   "Move Next Question"
         Top             =   15
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   820
         BTYPE           =   1
         TX              =   "English"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   104
         BCOLO           =   104
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "MCQ_Exam_Interface.frx":00E0
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame Frame4Three 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   535
      Left            =   1440
      TabIndex        =   133
      Top             =   2640
      Width           =   14505
      Begin ChamaleonButton.ChameleonBtn sub33 
         Height          =   465
         Left            =   9600
         TabIndex        =   134
         ToolTipText     =   "Move Next Question"
         Top             =   15
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   820
         BTYPE           =   1
         TX              =   "English"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   7552000
         FCOLO           =   7552000
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "MCQ_Exam_Interface.frx":00FC
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn sub32 
         Height          =   465
         Left            =   4800
         TabIndex        =   135
         ToolTipText     =   "Move Next Question"
         Top             =   15
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   820
         BTYPE           =   1
         TX              =   "English"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   7552000
         FCOLO           =   7552000
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "MCQ_Exam_Interface.frx":0118
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn sub31 
         Height          =   465
         Left            =   0
         TabIndex        =   136
         ToolTipText     =   "Move Next Question"
         Top             =   15
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   820
         BTYPE           =   1
         TX              =   "English"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   104
         BCOLO           =   104
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "MCQ_Exam_Interface.frx":0134
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame Frame4Two 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   1440
      TabIndex        =   129
      Top             =   1800
      Width           =   14505
      Begin ChamaleonButton.ChameleonBtn sub21 
         Height          =   465
         Left            =   0
         TabIndex        =   130
         ToolTipText     =   "Move Next Question"
         Top             =   15
         Width           =   7200
         _ExtentX        =   12700
         _ExtentY        =   820
         BTYPE           =   1
         TX              =   "English"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   104
         BCOLO           =   104
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "MCQ_Exam_Interface.frx":0150
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn sub22 
         Height          =   465
         Left            =   7200
         TabIndex        =   131
         ToolTipText     =   "Move Next Question"
         Top             =   15
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   820
         BTYPE           =   1
         TX              =   "English"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   7552000
         FCOLO           =   7552000
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "MCQ_Exam_Interface.frx":016C
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   9165
      Left            =   16000
      Picture         =   "MCQ_Exam_Interface.frx":0188
      ScaleHeight     =   9105
      ScaleWidth      =   4395
      TabIndex        =   55
      Top             =   0
      Width           =   4455
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   1
         Left            =   120
         TabIndex        =   66
         Top             =   2300
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":1A3A
         DisabledPicture =   "MCQ_Exam_Interface.frx":1A56
         DownPicture     =   "MCQ_Exam_Interface.frx":1A72
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":1A8E
         Caption         =   "1"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   2
         Left            =   840
         TabIndex        =   68
         Top             =   2300
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":1BF0
         DisabledPicture =   "MCQ_Exam_Interface.frx":1C0C
         DownPicture     =   "MCQ_Exam_Interface.frx":1C28
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":1C44
         Caption         =   "2"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   3
         Left            =   1560
         TabIndex        =   69
         Top             =   2300
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":1DA6
         DisabledPicture =   "MCQ_Exam_Interface.frx":1DC2
         DownPicture     =   "MCQ_Exam_Interface.frx":1DDE
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":1DFA
         Caption         =   "3"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   4
         Left            =   2280
         TabIndex        =   70
         Top             =   2300
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":1F5C
         DisabledPicture =   "MCQ_Exam_Interface.frx":1F78
         DownPicture     =   "MCQ_Exam_Interface.frx":1F94
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":1FB0
         Caption         =   "4"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   5
         Left            =   3000
         TabIndex        =   71
         Top             =   2300
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":2112
         DisabledPicture =   "MCQ_Exam_Interface.frx":212E
         DownPicture     =   "MCQ_Exam_Interface.frx":214A
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":2166
         Caption         =   "5"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   6
         Left            =   3720
         TabIndex        =   72
         Top             =   2300
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":22C8
         DisabledPicture =   "MCQ_Exam_Interface.frx":22E4
         DownPicture     =   "MCQ_Exam_Interface.frx":2300
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":231C
         Caption         =   "6"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   7
         Left            =   120
         TabIndex        =   73
         Top             =   2970
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":247E
         DisabledPicture =   "MCQ_Exam_Interface.frx":249A
         DownPicture     =   "MCQ_Exam_Interface.frx":24B6
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":24D2
         Caption         =   "7"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   8
         Left            =   840
         TabIndex        =   74
         Top             =   2970
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":2634
         DisabledPicture =   "MCQ_Exam_Interface.frx":2650
         DownPicture     =   "MCQ_Exam_Interface.frx":266C
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":2688
         Caption         =   "8"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   9
         Left            =   1560
         TabIndex        =   75
         Top             =   2970
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":27EA
         DisabledPicture =   "MCQ_Exam_Interface.frx":2806
         DownPicture     =   "MCQ_Exam_Interface.frx":2822
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":283E
         Caption         =   "9"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   10
         Left            =   2280
         TabIndex        =   76
         Top             =   2970
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":29A0
         DisabledPicture =   "MCQ_Exam_Interface.frx":29BC
         DownPicture     =   "MCQ_Exam_Interface.frx":29D8
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":29F4
         Caption         =   "10"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   11
         Left            =   3000
         TabIndex        =   77
         Top             =   2970
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":2B56
         DisabledPicture =   "MCQ_Exam_Interface.frx":2B72
         DownPicture     =   "MCQ_Exam_Interface.frx":2B8E
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":2BAA
         Caption         =   "11"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   12
         Left            =   3720
         TabIndex        =   78
         Top             =   2970
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":2D0C
         DisabledPicture =   "MCQ_Exam_Interface.frx":2D28
         DownPicture     =   "MCQ_Exam_Interface.frx":2D44
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":2D60
         Caption         =   "12"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   13
         Left            =   120
         TabIndex        =   79
         Top             =   3640
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":2EC2
         DisabledPicture =   "MCQ_Exam_Interface.frx":2EDE
         DownPicture     =   "MCQ_Exam_Interface.frx":2EFA
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":2F16
         Caption         =   "13"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   14
         Left            =   840
         TabIndex        =   80
         Top             =   3640
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":3078
         DisabledPicture =   "MCQ_Exam_Interface.frx":3094
         DownPicture     =   "MCQ_Exam_Interface.frx":30B0
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":30CC
         Caption         =   "14"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   15
         Left            =   1560
         TabIndex        =   81
         Top             =   3640
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":322E
         DisabledPicture =   "MCQ_Exam_Interface.frx":324A
         DownPicture     =   "MCQ_Exam_Interface.frx":3266
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":3282
         Caption         =   "15"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   16
         Left            =   2280
         TabIndex        =   82
         Top             =   3640
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":33E4
         DisabledPicture =   "MCQ_Exam_Interface.frx":3400
         DownPicture     =   "MCQ_Exam_Interface.frx":341C
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":3438
         Caption         =   "16"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   17
         Left            =   3000
         TabIndex        =   83
         Top             =   3640
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":359A
         DisabledPicture =   "MCQ_Exam_Interface.frx":35B6
         DownPicture     =   "MCQ_Exam_Interface.frx":35D2
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":35EE
         Caption         =   "17"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   18
         Left            =   3720
         TabIndex        =   84
         Top             =   3640
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":3750
         DisabledPicture =   "MCQ_Exam_Interface.frx":376C
         DownPicture     =   "MCQ_Exam_Interface.frx":3788
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":37A4
         Caption         =   "18"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   19
         Left            =   120
         TabIndex        =   85
         Top             =   4285
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":3906
         DisabledPicture =   "MCQ_Exam_Interface.frx":3922
         DownPicture     =   "MCQ_Exam_Interface.frx":393E
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":395A
         Caption         =   "19"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   20
         Left            =   840
         TabIndex        =   86
         Top             =   4285
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":3ABC
         DisabledPicture =   "MCQ_Exam_Interface.frx":3AD8
         DownPicture     =   "MCQ_Exam_Interface.frx":3AF4
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":3B10
         Caption         =   "20"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   21
         Left            =   1560
         TabIndex        =   87
         Top             =   4285
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":3C72
         DisabledPicture =   "MCQ_Exam_Interface.frx":3C8E
         DownPicture     =   "MCQ_Exam_Interface.frx":3CAA
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":3CC6
         Caption         =   "21"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   22
         Left            =   2280
         TabIndex        =   88
         Top             =   4285
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":3E28
         DisabledPicture =   "MCQ_Exam_Interface.frx":3E44
         DownPicture     =   "MCQ_Exam_Interface.frx":3E60
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":3E7C
         Caption         =   "22"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   23
         Left            =   3000
         TabIndex        =   89
         Top             =   4285
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":3FDE
         DisabledPicture =   "MCQ_Exam_Interface.frx":3FFA
         DownPicture     =   "MCQ_Exam_Interface.frx":4016
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":4032
         Caption         =   "23"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   24
         Left            =   3720
         TabIndex        =   90
         Top             =   4285
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":4194
         DisabledPicture =   "MCQ_Exam_Interface.frx":41B0
         DownPicture     =   "MCQ_Exam_Interface.frx":41CC
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":41E8
         Caption         =   "24"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   25
         Left            =   120
         TabIndex        =   91
         Top             =   4960
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":434A
         DisabledPicture =   "MCQ_Exam_Interface.frx":4366
         DownPicture     =   "MCQ_Exam_Interface.frx":4382
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":439E
         Caption         =   "25"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   26
         Left            =   840
         TabIndex        =   92
         Top             =   4960
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":4500
         DisabledPicture =   "MCQ_Exam_Interface.frx":451C
         DownPicture     =   "MCQ_Exam_Interface.frx":4538
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":4554
         Caption         =   "26"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   27
         Left            =   1560
         TabIndex        =   93
         Top             =   4960
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":46B6
         DisabledPicture =   "MCQ_Exam_Interface.frx":46D2
         DownPicture     =   "MCQ_Exam_Interface.frx":46EE
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":470A
         Caption         =   "27"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   28
         Left            =   2280
         TabIndex        =   94
         Top             =   4960
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":486C
         DisabledPicture =   "MCQ_Exam_Interface.frx":4888
         DownPicture     =   "MCQ_Exam_Interface.frx":48A4
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":48C0
         Caption         =   "28"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   29
         Left            =   3000
         TabIndex        =   95
         Top             =   4960
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":4A22
         DisabledPicture =   "MCQ_Exam_Interface.frx":4A3E
         DownPicture     =   "MCQ_Exam_Interface.frx":4A5A
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":4A76
         Caption         =   "29"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   30
         Left            =   3720
         TabIndex        =   96
         Top             =   4960
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":4BD8
         DisabledPicture =   "MCQ_Exam_Interface.frx":4BF4
         DownPicture     =   "MCQ_Exam_Interface.frx":4C10
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":4C2C
         Caption         =   "30"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   31
         Left            =   120
         TabIndex        =   97
         Top             =   5670
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":4D8E
         DisabledPicture =   "MCQ_Exam_Interface.frx":4DAA
         DownPicture     =   "MCQ_Exam_Interface.frx":4DC6
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":4DE2
         Caption         =   "31"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   32
         Left            =   840
         TabIndex        =   98
         Top             =   5670
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":4F44
         DisabledPicture =   "MCQ_Exam_Interface.frx":4F60
         DownPicture     =   "MCQ_Exam_Interface.frx":4F7C
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":4F98
         Caption         =   "32"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   33
         Left            =   1560
         TabIndex        =   99
         Top             =   5670
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":50FA
         DisabledPicture =   "MCQ_Exam_Interface.frx":5116
         DownPicture     =   "MCQ_Exam_Interface.frx":5132
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":514E
         Caption         =   "33"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   34
         Left            =   2280
         TabIndex        =   100
         Top             =   5670
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":52B0
         DisabledPicture =   "MCQ_Exam_Interface.frx":52CC
         DownPicture     =   "MCQ_Exam_Interface.frx":52E8
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":5304
         Caption         =   "34"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   35
         Left            =   3000
         TabIndex        =   101
         Top             =   5670
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":5466
         DisabledPicture =   "MCQ_Exam_Interface.frx":5482
         DownPicture     =   "MCQ_Exam_Interface.frx":549E
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":54BA
         Caption         =   "35"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   36
         Left            =   3720
         TabIndex        =   102
         Top             =   5670
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":561C
         DisabledPicture =   "MCQ_Exam_Interface.frx":5638
         DownPicture     =   "MCQ_Exam_Interface.frx":5654
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":5670
         Caption         =   "36"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   37
         Left            =   120
         TabIndex        =   103
         Top             =   6350
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":57D2
         DisabledPicture =   "MCQ_Exam_Interface.frx":57EE
         DownPicture     =   "MCQ_Exam_Interface.frx":580A
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":5826
         Caption         =   "37"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   38
         Left            =   840
         TabIndex        =   104
         Top             =   6350
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":5988
         DisabledPicture =   "MCQ_Exam_Interface.frx":59A4
         DownPicture     =   "MCQ_Exam_Interface.frx":59C0
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":59DC
         Caption         =   "38"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   39
         Left            =   1560
         TabIndex        =   105
         Top             =   6350
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":5B3E
         DisabledPicture =   "MCQ_Exam_Interface.frx":5B5A
         DownPicture     =   "MCQ_Exam_Interface.frx":5B76
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":5B92
         Caption         =   "39"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   40
         Left            =   2280
         TabIndex        =   106
         Top             =   6350
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":5CF4
         DisabledPicture =   "MCQ_Exam_Interface.frx":5D10
         DownPicture     =   "MCQ_Exam_Interface.frx":5D2C
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":5D48
         Caption         =   "40"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   41
         Left            =   3000
         TabIndex        =   107
         Top             =   6350
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":5EAA
         DisabledPicture =   "MCQ_Exam_Interface.frx":5EC6
         DownPicture     =   "MCQ_Exam_Interface.frx":5EE2
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":5EFE
         Caption         =   "41"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   42
         Left            =   3720
         TabIndex        =   108
         Top             =   6350
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":6060
         DisabledPicture =   "MCQ_Exam_Interface.frx":607C
         DownPicture     =   "MCQ_Exam_Interface.frx":6098
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":60B4
         Caption         =   "42"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   43
         Left            =   120
         TabIndex        =   109
         Top             =   7000
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":6216
         DisabledPicture =   "MCQ_Exam_Interface.frx":6232
         DownPicture     =   "MCQ_Exam_Interface.frx":624E
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":626A
         Caption         =   "43"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   44
         Left            =   840
         TabIndex        =   110
         Top             =   7000
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":63CC
         DisabledPicture =   "MCQ_Exam_Interface.frx":63E8
         DownPicture     =   "MCQ_Exam_Interface.frx":6404
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":6420
         Caption         =   "44"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   45
         Left            =   1560
         TabIndex        =   111
         Top             =   7000
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":6582
         DisabledPicture =   "MCQ_Exam_Interface.frx":659E
         DownPicture     =   "MCQ_Exam_Interface.frx":65BA
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":65D6
         Caption         =   "45"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   46
         Left            =   2280
         TabIndex        =   112
         Top             =   7000
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":6738
         DisabledPicture =   "MCQ_Exam_Interface.frx":6754
         DownPicture     =   "MCQ_Exam_Interface.frx":6770
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":678C
         Caption         =   "46"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   47
         Left            =   3000
         TabIndex        =   113
         Top             =   7000
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":68EE
         DisabledPicture =   "MCQ_Exam_Interface.frx":690A
         DownPicture     =   "MCQ_Exam_Interface.frx":6926
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":6942
         Caption         =   "47"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   48
         Left            =   3720
         TabIndex        =   114
         Top             =   7000
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":6AA4
         DisabledPicture =   "MCQ_Exam_Interface.frx":6AC0
         DownPicture     =   "MCQ_Exam_Interface.frx":6ADC
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":6AF8
         Caption         =   "48"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   49
         Left            =   120
         TabIndex        =   115
         Top             =   7700
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":6C5A
         DisabledPicture =   "MCQ_Exam_Interface.frx":6C76
         DownPicture     =   "MCQ_Exam_Interface.frx":6C92
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":6CAE
         Caption         =   "49"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   50
         Left            =   840
         TabIndex        =   116
         Top             =   7700
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":6E10
         DisabledPicture =   "MCQ_Exam_Interface.frx":6E2C
         DownPicture     =   "MCQ_Exam_Interface.frx":6E48
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":6E64
         Caption         =   "50"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   51
         Left            =   1560
         TabIndex        =   117
         Top             =   7700
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":6FC6
         DisabledPicture =   "MCQ_Exam_Interface.frx":6FE2
         DownPicture     =   "MCQ_Exam_Interface.frx":6FFE
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":701A
         Caption         =   "51"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   52
         Left            =   2280
         TabIndex        =   118
         Top             =   7700
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":717C
         DisabledPicture =   "MCQ_Exam_Interface.frx":7198
         DownPicture     =   "MCQ_Exam_Interface.frx":71B4
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":71D0
         Caption         =   "52"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   53
         Left            =   3000
         TabIndex        =   119
         Top             =   7700
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":7332
         DisabledPicture =   "MCQ_Exam_Interface.frx":734E
         DownPicture     =   "MCQ_Exam_Interface.frx":736A
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":7386
         Caption         =   "53"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   54
         Left            =   3720
         TabIndex        =   120
         Top             =   7700
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":74E8
         DisabledPicture =   "MCQ_Exam_Interface.frx":7504
         DownPicture     =   "MCQ_Exam_Interface.frx":7520
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":753C
         Caption         =   "54"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   55
         Left            =   120
         TabIndex        =   121
         Top             =   8400
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":769E
         DisabledPicture =   "MCQ_Exam_Interface.frx":76BA
         DownPicture     =   "MCQ_Exam_Interface.frx":76D6
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":76F2
         Caption         =   "55"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   56
         Left            =   840
         TabIndex        =   122
         Top             =   8400
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":7854
         DisabledPicture =   "MCQ_Exam_Interface.frx":7870
         DownPicture     =   "MCQ_Exam_Interface.frx":788C
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":78A8
         Caption         =   "56"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   57
         Left            =   1560
         TabIndex        =   123
         Top             =   8400
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":7A0A
         DisabledPicture =   "MCQ_Exam_Interface.frx":7A26
         DownPicture     =   "MCQ_Exam_Interface.frx":7A42
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":7A5E
         Caption         =   "57"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   58
         Left            =   2280
         TabIndex        =   124
         Top             =   8400
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":7BC0
         DisabledPicture =   "MCQ_Exam_Interface.frx":7BDC
         DownPicture     =   "MCQ_Exam_Interface.frx":7BF8
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":7C14
         Caption         =   "58"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   59
         Left            =   3000
         TabIndex        =   125
         Top             =   8400
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":7D76
         DisabledPicture =   "MCQ_Exam_Interface.frx":7D92
         DownPicture     =   "MCQ_Exam_Interface.frx":7DAE
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":7DCA
         Caption         =   "59"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   60
         Left            =   3720
         TabIndex        =   126
         Top             =   8400
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MCQ_Exam_Interface.frx":7F2C
         DisabledPicture =   "MCQ_Exam_Interface.frx":7F48
         DownPicture     =   "MCQ_Exam_Interface.frx":7F64
         MousePointer    =   99
         MouseIcon       =   "MCQ_Exam_Interface.frx":7F80
         Caption         =   "60"
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   60
         Left            =   3800
         Shape           =   3  'Circle
         Top             =   8400
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   59
         Left            =   3080
         Shape           =   3  'Circle
         Top             =   8400
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   58
         Left            =   2350
         Shape           =   3  'Circle
         Top             =   8400
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   57
         Left            =   1635
         Shape           =   3  'Circle
         Top             =   8400
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   56
         Left            =   920
         Shape           =   3  'Circle
         Top             =   8400
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   55
         Left            =   200
         Shape           =   3  'Circle
         Top             =   8400
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   54
         Left            =   3800
         Shape           =   3  'Circle
         Top             =   7695
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   53
         Left            =   3070
         Shape           =   3  'Circle
         Top             =   7695
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   52
         Left            =   2350
         Shape           =   3  'Circle
         Top             =   7695
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   51
         Left            =   1635
         Shape           =   3  'Circle
         Top             =   7695
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   50
         Left            =   920
         Shape           =   3  'Circle
         Top             =   7695
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   49
         Left            =   200
         Shape           =   3  'Circle
         Top             =   7695
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   48
         Left            =   3800
         Shape           =   3  'Circle
         Top             =   7005
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   47
         Left            =   3070
         Shape           =   3  'Circle
         Top             =   7005
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   46
         Left            =   2350
         Shape           =   3  'Circle
         Top             =   7005
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   45
         Left            =   1630
         Shape           =   3  'Circle
         Top             =   7005
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   44
         Left            =   920
         Shape           =   3  'Circle
         Top             =   7005
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   43
         Left            =   200
         Shape           =   3  'Circle
         Top             =   7005
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   42
         Left            =   3800
         Shape           =   3  'Circle
         Top             =   6345
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   41
         Left            =   3070
         Shape           =   3  'Circle
         Top             =   6345
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   40
         Left            =   2350
         Shape           =   3  'Circle
         Top             =   6345
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   39
         Left            =   1630
         Shape           =   3  'Circle
         Top             =   6345
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   38
         Left            =   920
         Shape           =   3  'Circle
         Top             =   6345
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   37
         Left            =   200
         Shape           =   3  'Circle
         Top             =   6345
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   36
         Left            =   3800
         Shape           =   3  'Circle
         Top             =   5670
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   35
         Left            =   3070
         Shape           =   3  'Circle
         Top             =   5670
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   34
         Left            =   2350
         Shape           =   3  'Circle
         Top             =   5670
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   33
         Left            =   1630
         Shape           =   3  'Circle
         Top             =   5670
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   32
         Left            =   920
         Shape           =   3  'Circle
         Top             =   5670
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   31
         Left            =   200
         Shape           =   3  'Circle
         Top             =   5670
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   30
         Left            =   3800
         Shape           =   3  'Circle
         Top             =   4960
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   29
         Left            =   3070
         Shape           =   3  'Circle
         Top             =   4960
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   28
         Left            =   2350
         Shape           =   3  'Circle
         Top             =   4960
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   27
         Left            =   1630
         Shape           =   3  'Circle
         Top             =   4960
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   26
         Left            =   920
         Shape           =   3  'Circle
         Top             =   4960
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   25
         Left            =   200
         Shape           =   3  'Circle
         Top             =   4960
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   24
         Left            =   3800
         Shape           =   3  'Circle
         Top             =   4285
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   23
         Left            =   3070
         Shape           =   3  'Circle
         Top             =   4285
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   22
         Left            =   2350
         Shape           =   3  'Circle
         Top             =   4285
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   21
         Left            =   1630
         Shape           =   3  'Circle
         Top             =   4285
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   20
         Left            =   920
         Shape           =   3  'Circle
         Top             =   4285
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   19
         Left            =   200
         Shape           =   3  'Circle
         Top             =   4285
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   18
         Left            =   3800
         Shape           =   3  'Circle
         Top             =   3650
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   17
         Left            =   3070
         Shape           =   3  'Circle
         Top             =   3650
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   16
         Left            =   2350
         Shape           =   3  'Circle
         Top             =   3650
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   15
         Left            =   1630
         Shape           =   3  'Circle
         Top             =   3650
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   14
         Left            =   920
         Shape           =   3  'Circle
         Top             =   3650
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   13
         Left            =   200
         Shape           =   3  'Circle
         Top             =   3650
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   12
         Left            =   3800
         Shape           =   3  'Circle
         Top             =   2970
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   11
         Left            =   3070
         Shape           =   3  'Circle
         Top             =   2970
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   10
         Left            =   2350
         Shape           =   3  'Circle
         Top             =   2970
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   9
         Left            =   1630
         Shape           =   3  'Circle
         Top             =   2970
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   8
         Left            =   920
         Shape           =   3  'Circle
         Top             =   2970
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   7
         Left            =   200
         Shape           =   3  'Circle
         Top             =   2970
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   1
         Left            =   200
         Shape           =   3  'Circle
         Top             =   2300
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   6
         Left            =   3800
         Shape           =   3  'Circle
         Top             =   2300
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   5
         Left            =   3080
         Shape           =   3  'Circle
         Top             =   2300
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   4
         Left            =   2350
         Shape           =   3  'Circle
         Top             =   2300
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   3
         Left            =   1635
         Shape           =   3  'Circle
         Top             =   2300
         Width           =   495
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         Height          =   615
         Index           =   2
         Left            =   920
         Shape           =   3  'Circle
         Top             =   2300
         Width           =   495
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000C000&
         FillColor       =   &H00004000&
         Height          =   345
         Left            =   60
         Shape           =   3  'Circle
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Answered"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   520
         TabIndex        =   64
         Top             =   765
         Width           =   945
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Marked"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   2010
         TabIndex        =   63
         Top             =   765
         Width           =   690
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Not visited"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   3210
         TabIndex        =   62
         Top             =   765
         Width           =   1005
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Marked && answered"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   520
         TabIndex        =   61
         Top             =   1275
         Width           =   1890
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Not Answered"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   2970
         TabIndex        =   60
         Top             =   1275
         Width           =   1335
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H000100FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000100FF&
         FillColor       =   &H00004000&
         Height          =   345
         Left            =   2565
         Shape           =   3  'Circle
         Top             =   1230
         Width           =   375
      End
      Begin VB.Shape Shape8 
         BackColor       =   &H00CA30DF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00CA30DF&
         FillColor       =   &H00004000&
         Height          =   345
         Left            =   1575
         Shape           =   3  'Circle
         Top             =   720
         Width           =   375
      End
      Begin VB.Shape Shape9 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0FFFF&
         FillColor       =   &H00004000&
         Height          =   345
         Left            =   2745
         Shape           =   3  'Circle
         Top             =   720
         Width           =   375
      End
      Begin VB.Shape Shape11 
         BackColor       =   &H00800080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800080&
         FillColor       =   &H00004000&
         Height          =   345
         Left            =   60
         Shape           =   3  'Circle
         Top             =   1230
         Width           =   375
      End
      Begin VB.Label oName 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Purushottam Kumar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   705
         TabIndex        =   59
         Top             =   135
         Width           =   2250
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   2
         X1              =   0
         X2              =   4450
         Y1              =   635
         Y2              =   635
      End
      Begin VB.Image stuPIC 
         Height          =   600
         Left            =   15
         Picture         =   "MCQ_Exam_Interface.frx":80E2
         Stretch         =   -1  'True
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "(PUR003)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   3065
         TabIndex        =   58
         Top             =   140
         Width           =   1260
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "MCQ  Questions"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1260
         TabIndex        =   57
         Top             =   1800
         Width           =   1650
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Section :"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   190
         TabIndex        =   56
         Top             =   1800
         Width           =   900
      End
      Begin VB.Shape Shape10 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFC0C0&
         Height          =   495
         Left            =   -15
         Top             =   1650
         Width           =   4575
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   1350
      Left            =   16000
      Picture         =   "MCQ_Exam_Interface.frx":8924
      ScaleHeight     =   1290
      ScaleWidth      =   4395
      TabIndex        =   45
      Top             =   9200
      Width           =   4455
      Begin VB.CommandButton btnSubmit 
         BackColor       =   &H00C0C000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   200
         MouseIcon       =   "MCQ_Exam_Interface.frx":A1D6
         MousePointer    =   99  'Custom
         Picture         =   "MCQ_Exam_Interface.frx":A328
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Submit The Test"
         Top             =   700
         Width           =   4020
      End
      Begin VB.CommandButton vkCommand3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2380
         MouseIcon       =   "MCQ_Exam_Interface.frx":AF4C
         MousePointer    =   99  'Custom
         Picture         =   "MCQ_Exam_Interface.frx":B09E
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Show Instruction Page"
         Top             =   120
         Width           =   1845
      End
      Begin VB.CommandButton vkCommand2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   200
         MouseIcon       =   "MCQ_Exam_Interface.frx":B9CD
         MousePointer    =   99  'Custom
         Picture         =   "MCQ_Exam_Interface.frx":BB1F
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "See All Questions at a glance"
         Top             =   120
         Width           =   2000
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   8445
      Left            =   0
      TabIndex        =   9
      Top             =   2160
      Width           =   16050
      Begin VB.PictureBox Picture2 
         Height          =   855
         Left            =   0
         Picture         =   "MCQ_Exam_Interface.frx":C67E
         ScaleHeight     =   795
         ScaleWidth      =   15915
         TabIndex        =   49
         Top             =   7560
         Width           =   15975
         Begin VB.TextBox Text1 
            Height          =   330
            Left            =   6120
            TabIndex        =   149
            Text            =   "Text1"
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CommandButton svNext 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   13440
            MouseIcon       =   "MCQ_Exam_Interface.frx":E9E4
            MousePointer    =   99  'Custom
            Picture         =   "MCQ_Exam_Interface.frx":EB36
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "Save and Move to Next Question"
            Top             =   150
            Width           =   1965
         End
         Begin VB.CommandButton nextq 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   10440
            MouseIcon       =   "MCQ_Exam_Interface.frx":F597
            MousePointer    =   99  'Custom
            Picture         =   "MCQ_Exam_Interface.frx":F6E9
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Move to Next Question"
            Top             =   150
            Width           =   1845
         End
         Begin VB.CommandButton prevQ 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   8360
            MouseIcon       =   "MCQ_Exam_Interface.frx":FF84
            MousePointer    =   99  'Custom
            Picture         =   "MCQ_Exam_Interface.frx":100D6
            Style           =   1  'Graphical
            TabIndex        =   52
            ToolTipText     =   "Move to Previous Question"
            Top             =   150
            Width           =   1845
         End
         Begin VB.CommandButton clear 
            BackColor       =   &H80000005&
            Caption         =   "Clear Response"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   3230
            MouseIcon       =   "MCQ_Exam_Interface.frx":10974
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "Clear Answer"
            Top             =   150
            Width           =   1845
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000005&
            Caption         =   "Bookmark This Question"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   325
            Left            =   225
            MouseIcon       =   "MCQ_Exam_Interface.frx":10AC6
            MousePointer    =   99  'Custom
            TabIndex        =   50
            ToolTipText     =   "Mark this Question"
            Top             =   215
            Width           =   2655
         End
         Begin MyEllipticButton.EllipticButton btn 
            Height          =   645
            Index           =   0
            Left            =   6840
            TabIndex        =   67
            Top             =   0
            Visible         =   0   'False
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   1138
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "MCQ_Exam_Interface.frx":10C18
            DisabledPicture =   "MCQ_Exam_Interface.frx":10C34
            DownPicture     =   "MCQ_Exam_Interface.frx":10C50
            MousePointer    =   99
            MouseIcon       =   "MCQ_Exam_Interface.frx":10C6C
            Caption         =   "1"
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   6
            Height          =   615
            Index           =   0
            Left            =   6915
            Shape           =   3  'Circle
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Shape Shape2 
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   465
            Left            =   150
            Top             =   150
            Width           =   2790
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   3015
         Left            =   6720
         TabIndex        =   38
         Top             =   1140
         Width           =   4695
         Begin VB.Image Image1 
            Height          =   2925
            Left            =   360
            Picture         =   "MCQ_Exam_Interface.frx":10DCE
            Stretch         =   -1  'True
            Top             =   0
            Width           =   3450
         End
      End
      Begin VB.Frame MainFrame 
         BackColor       =   &H00FFFFFF&
         Height          =   2535
         Left            =   6960
         TabIndex        =   24
         Top             =   1560
         Width           =   4455
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   495
            Left            =   120
            Top             =   1920
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   873
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   "Provider=MSDAORA.1;Password=sts;User ID=sts;Persist Security Info=True"
            OLEDBString     =   "Provider=MSDAORA.1;Password=sts;User ID=sts;Persist Security Info=True"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "select * from mcqtest order by q_no"
            Caption         =   "Adodc1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin VB.Frame Frame2 
            Caption         =   "Timing"
            Height          =   975
            Left            =   1200
            TabIndex        =   37
            Top             =   120
            Visible         =   0   'False
            Width           =   1095
            Begin VB.Timer Timer2 
               Enabled         =   0   'False
               Interval        =   1000
               Left            =   600
               Top             =   360
            End
            Begin VB.Timer Timer1 
               Enabled         =   0   'False
               Interval        =   1000
               Left            =   120
               Top             =   360
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Prev_Next"
            Height          =   975
            Left            =   120
            TabIndex        =   25
            Top             =   120
            Visible         =   0   'False
            Width           =   975
            Begin VB.Timer btnNXT_PREV 
               Enabled         =   0   'False
               Interval        =   100
               Left            =   240
               Top             =   360
            End
         End
         Begin VB.Label tmp 
            BackColor       =   &H008080FF&
            Height          =   375
            Left            =   120
            TabIndex        =   36
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label Optkey 
            BackColor       =   &H00FF80FF&
            Height          =   375
            Left            =   2280
            TabIndex        =   35
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFC0&
            Caption         =   "OptKey"
            Height          =   255
            Left            =   2280
            TabIndex        =   34
            Top             =   120
            Width           =   2055
         End
         Begin VB.Label indx 
            BackColor       =   &H00FF8080&
            Height          =   255
            Left            =   3600
            TabIndex        =   33
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            BackColor       =   &H0080C0FF&
            Caption         =   "Correct_Ans"
            Height          =   255
            Left            =   2280
            TabIndex        =   32
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            BackColor       =   &H0080C0FF&
            Caption         =   "User Ans"
            Height          =   255
            Left            =   2280
            TabIndex        =   31
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label lbl_corr_ans 
            BackColor       =   &H0000FF00&
            Height          =   255
            Left            =   3600
            TabIndex        =   30
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label lbl_usr_ans 
            BackColor       =   &H0000FF00&
            Height          =   255
            Left            =   3600
            TabIndex        =   29
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label25 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Caption         =   "Indx"
            Height          =   255
            Left            =   2280
            TabIndex        =   28
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label bookmark 
            BackColor       =   &H00FF8080&
            DataField       =   "BOOKMRK"
            DataSource      =   "Adodc2"
            Height          =   255
            Left            =   3600
            TabIndex        =   27
            Top             =   2160
            Width           =   735
         End
         Begin VB.Label Label26 
            Alignment       =   2  'Center
            BackColor       =   &H0080C0FF&
            Caption         =   "Bookmark"
            Height          =   255
            Left            =   2280
            TabIndex        =   26
            Top             =   2160
            Width           =   1095
         End
      End
      Begin VB.OptionButton options 
         BackColor       =   &H00FFFFFF&
         Caption         =   "180 Runs ieuwie ieei fe ebeibeiee he biebebiev bvivbbvv vjhhfjhjj vvfj jijjijjJBJKSKJNN j NJ NJSKNSNKJSKSKS"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Index           =   3
         Left            =   960
         MouseIcon       =   "MCQ_Exam_Interface.frx":15100
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   3000
         Width           =   5535
      End
      Begin VB.OptionButton options 
         BackColor       =   &H00FFFFFF&
         Caption         =   "180 Runs ieuwie ieei fe ebeibeiee he biebebiev bvivbbvv vjhhfjhjj vvfj jijjijjJBJKSKJNN j NJ NJSKNSNKJSKSKS"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Index           =   2
         Left            =   960
         MouseIcon       =   "MCQ_Exam_Interface.frx":15252
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   2400
         Width           =   5535
      End
      Begin VB.OptionButton options 
         BackColor       =   &H00FFFFFF&
         Caption         =   "180 Runs ieuwie ieei fe ebeibeiee he biebebiev bvivbbvv vjhhfjhjj vvfj jijjijjJBJKSKJNN j NJ NJSKNSNKJSKSKS"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Index           =   1
         Left            =   960
         MouseIcon       =   "MCQ_Exam_Interface.frx":153A4
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   1800
         Width           =   5535
      End
      Begin VB.OptionButton options 
         BackColor       =   &H00FFFFFF&
         Caption         =   "180 Runs ieuwie ieei fe ebeibeiee he biebebiev bvivbbvv vjhhfjhjj vvfj jijjijjJBJKSKJNN j NJ NJSKNSNKJSKSKS"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Index           =   0
         Left            =   960
         MouseIcon       =   "MCQ_Exam_Interface.frx":154F6
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   1200
         Width           =   5415
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         X1              =   0
         X2              =   16100
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFFFF&
         Caption         =   "D."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   14
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFFFF&
         Caption         =   "C."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   13
         Top             =   2505
         Width           =   375
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFFFF&
         Caption         =   "B."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   12
         Top             =   1905
         Width           =   375
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "A."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   11
         Top             =   1310
         Width           =   375
      End
      Begin VB.Label qtext 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   $"MCQ_Exam_Interface.frx":15648
         DataField       =   "Q_TXT"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   630
         Left            =   360
         TabIndex        =   10
         Top             =   360
         Width           =   15015
      End
   End
   Begin VB.PictureBox vkFrame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   16035
      TabIndex        =   0
      Top             =   -120
      Width           =   16095
      Begin VB.CommandButton vkCommand1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   14325
         MouseIcon       =   "MCQ_Exam_Interface.frx":15726
         MousePointer    =   99  'Custom
         Picture         =   "MCQ_Exam_Interface.frx":15878
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Terminate Exam "
         Top             =   260
         Width           =   1485
      End
      Begin VB.Label Label31 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Course :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   240
         TabIndex        =   147
         Top             =   210
         Width           =   1515
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time Left -"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9960
         TabIndex        =   1
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Timer_remain 
         Alignment       =   2  'Center
         BackColor       =   &H00D1E0D0&
         Caption         =   "04 : 36"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10965
         TabIndex        =   23
         Top             =   330
         Width           =   1215
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   11.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   11640
         TabIndex        =   22
         Top             =   360
         Width           =   75
      End
      Begin VB.Label t_left_hour 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "05"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   11.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   11265
         TabIndex        =   21
         Top             =   360
         Width           =   300
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   11.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   12120
         TabIndex        =   20
         Top             =   360
         Width           =   75
      End
      Begin VB.Label t_left_min 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "05"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   11.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   11760
         TabIndex        =   3
         Top             =   360
         Width           =   300
      End
      Begin VB.Label Heading 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "[ SSC - CGL ]"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   1800
         TabIndex        =   4
         Top             =   210
         Width           =   5115
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time Left"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   0
         TabIndex        =   2
         Top             =   1440
         Width           =   855
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00D1E0D0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   495
         Left            =   9600
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame4One 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   535
      Left            =   1440
      TabIndex        =   127
      Top             =   840
      Width           =   14505
      Begin ChamaleonButton.ChameleonBtn Sub11 
         Height          =   465
         Left            =   0
         TabIndex        =   128
         ToolTipText     =   "Move Next Question"
         Top             =   15
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   820
         BTYPE           =   1
         TX              =   "English"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   104
         BCOLO           =   104
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "MCQ_Exam_Interface.frx":160F7
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.CommandButton TestTypeInfo 
      Caption         =   "Subject Wise Test (Mathematics)"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1440
      MouseIcon       =   "MCQ_Exam_Interface.frx":16113
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   148
      ToolTipText     =   "Test Type Information"
      Top             =   840
      Width           =   14445
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Listen"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   15315
      MouseIcon       =   "MCQ_Exam_Interface.frx":16265
      MousePointer    =   99  'Custom
      TabIndex        =   44
      Top             =   1950
      Width           =   435
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rough"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   14025
      MouseIcon       =   "MCQ_Exam_Interface.frx":163B7
      MousePointer    =   99  'Custom
      TabIndex        =   43
      Top             =   1935
      Width           =   435
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   14670
      MouseIcon       =   "MCQ_Exam_Interface.frx":16509
      MousePointer    =   99  'Custom
      TabIndex        =   42
      Top             =   1950
      Width           =   450
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   7200
      TabIndex        =   41
      Top             =   1640
      Width           =   240
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   330
      Left            =   7080
      Shape           =   4  'Rounded Rectangle
      Top             =   1605
      Width           =   525
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   330
      Left            =   3720
      Shape           =   4  'Rounded Rectangle
      Top             =   6360
      Width           =   525
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+4"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   6540
      TabIndex        =   40
      Top             =   1640
      Width           =   300
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Marks :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5560
      TabIndex        =   39
      Top             =   1590
      Width           =   795
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   325
      Left            =   6480
      Shape           =   4  'Rounded Rectangle
      Top             =   1610
      Width           =   560
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   14010
      MouseIcon       =   "MCQ_Exam_Interface.frx":1665B
      MousePointer    =   99  'Custom
      Picture         =   "MCQ_Exam_Interface.frx":167AD
      ToolTipText     =   "Use Rough page here"
      Top             =   1455
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   15285
      MouseIcon       =   "MCQ_Exam_Interface.frx":17077
      MousePointer    =   99  'Custom
      Picture         =   "MCQ_Exam_Interface.frx":171C9
      Stretch         =   -1  'True
      ToolTipText     =   "Click To Listen Question"
      Top             =   1485
      Width           =   525
   End
   Begin VB.Image Command1 
      Height          =   450
      Left            =   14640
      MouseIcon       =   "MCQ_Exam_Interface.frx":189F2
      MousePointer    =   99  'Custom
      Picture         =   "MCQ_Exam_Interface.frx":18B44
      Stretch         =   -1  'True
      ToolTipText     =   "Show Current Status"
      Top             =   1500
      Width           =   480
   End
   Begin VB.Label t_time_min 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Elapsed Time - 00 : 00 : 44"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9645
      TabIndex        =   8
      Top             =   1620
      Width           =   3420
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Question No. "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   14280
      TabIndex        =   7
      Top             =   -240
      Width           =   1470
   End
   Begin VB.Label q_id 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      DataField       =   "Q_NO"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1860
      TabIndex        =   6
      Top             =   1605
      Width           =   630
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sections :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   195
      TabIndex        =   5
      Top             =   900
      Width           =   1140
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Question No. "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   19
      Top             =   1620
      Width           =   1530
   End
   Begin VB.Shape Shape12 
      BackColor       =   &H00FFEFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   0  'Transparent
      Height          =   810
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   16050
   End
   Begin VB.Shape Shape18 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFC0C0&
      Height          =   690
      Left            =   -120
      Shape           =   4  'Rounded Rectangle
      Top             =   735
      Width           =   16200
   End
End
Attribute VB_Name = "MCQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim totQues As Integer
Dim tmpp As Byte
Dim minute As Integer, second As Integer 'Timer Control
Dim currenttime As Variant
Dim Quespic As String

Private Sub btn_Click(Index As Integer)
On Error Resume Next
For i = 1 To 60
 If i = Index Then
  Shape(i).Visible = True
 Else
  Shape(i).Visible = False
 End If
Next i
Set r = New ADODB.Recordset
Optkey.Caption = Index
 q_id.Caption = btn(Index).Caption     'Question No
sql = "select * from mcqtest where Q_no =" & btn(Index).Caption & " "
Set r = c.Execute(sql)
 qtext.Caption = r.Fields(1)  'Question
For i = 0 To 3
 options(i).Caption = r.Fields(i + 2) 'Options
Next i
lbl_corr_ans.Caption = r.Fields(7)  'Answer
indx.Caption = lbl_corr_ans.Caption - 1
If IsNull(r.Fields(8)) = False Then
 Image1.Visible = True
 Quespic = r.Fields(8)
 Image1.Picture = LoadPicture(Quespic)
 Else
 Quespic = ""
 Image1.Visible = False
End If
ChkOption 'Matching option
ChkUserANS 'matching answer
CHKbookmark 'checking bookmark
bookmrkColor 'setting color
End Sub

Public Function ChkOption() 'Checking User Answer
Set r = New ADODB.Recordset
Dim tempANS As Integer
Set r = c.Execute("select user_ans from ANSWERHOLD where id =" & Val(q_id.Caption) & " ")
If IsNull(r.Fields(0)) = False Then
 tempANS = r.Fields(0)
 If tempANS = 0 Then
  For i = 0 To 3
   options(i).Value = False
  Next i
 Else
   options(tempANS - 1).Value = True
 End If
Else
 For i = 0 To 3
   options(i).Value = False
 Next i
End If
End Function

Public Function ChkUserANS() 'Check Which option is checked
Set r = New ADODB.Recordset
Dim userAns As Integer
Set r = c.Execute("select user_ans from ANSWERHOLD where id =" & Val(q_id.Caption) & "")
If IsNull(r.Fields(0)) = False Then
 userAns = r.Fields(0)
 If userAns >= 1 And userAns <= 4 Then
  lbl_usr_ans.Caption = userAns
  indx.Caption = Val(lbl_usr_ans.Caption) - 1
 Else
  lbl_usr_ans.Caption = ""
  indx.Caption = ""
 End If
Else
  lbl_usr_ans.Caption = ""
  indx.Caption = ""
End If
End Function

Public Function CHKbookmark()
 Set r = New ADODB.Recordset
 Set r = c.Execute("select * from answerhold where id='" & Val(q_id.Caption) & "' ")
 If r.Fields(2) <> 0 And r.Fields(3) = 1 Then 'Ansered and bookmarked
     Check1.Value = vbChecked
     options(r.Fields(2) - 1).Value = True
 ElseIf r.Fields(2) <> 0 And r.Fields(3) <> 1 Then 'Ansered but not bookmarked
     Check1.Value = vbUnchecked
     options(r.Fields(2) - 1).Value = True
 ElseIf r.Fields(2) = 0 And r.Fields(3) = 1 Then 'Not answer but bookmarked
     Check1.Value = vbChecked
     For i = 0 To 3
      options(i).Value = False
     Next i
 ElseIf r.Fields(2) = 0 And r.Fields(3) <> 1 Then 'Not answered and not bookmarked
     Check1.Value = vbUnchecked
     For i = 0 To 3
      options(i).Value = False
     Next i
 End If
End Function

Public Function bookmrkColor() 'Providing Equivalent Color
If IsFullLengthSelected = 1 Then
 If FTOTSUB = 2 Then
  If Val(q_id.Caption) >= Question22 Then
    sub21.BackColor = &HE0E0E0
    sub21.ForeColor = &H733C00
    sub22.BackColor = &H68&
    sub22.ForeColor = &HFFFFFF
   ElseIf Val(q_id.Caption) < Question22 Then
    sub21.BackColor = &H68&
    sub21.ForeColor = &HFFFFFF
    sub22.BackColor = &HE0E0E0
    sub22.ForeColor = &H733C00
  End If
 ElseIf FTOTSUB = 3 Then
  If Val(q_id.Caption) >= Question33 Then
   sub32.BackColor = &HE0E0E0
   sub32.ForeColor = &H733C00
   sub31.BackColor = &HE0E0E0
   sub31.ForeColor = &H733C00
   sub33.BackColor = &H68&
   sub33.ForeColor = &HFFFFFF
  ElseIf Val(q_id.Caption) < Question33 And Val(q_id.Caption) >= Question32 Then
   sub33.BackColor = &HE0E0E0
   sub33.ForeColor = &H733C00
   sub31.BackColor = &HE0E0E0
   sub31.ForeColor = &H733C00
   sub32.BackColor = &H68&
   sub32.ForeColor = &HFFFFFF
  ElseIf Val(q_id.Caption) < Question32 Then
   sub32.BackColor = &HE0E0E0
   sub32.ForeColor = &H733C00
   sub33.BackColor = &HE0E0E0
   sub33.ForeColor = &H733C00
   sub31.BackColor = &H68&
   sub31.ForeColor = &HFFFFFF
  End If
 ElseIf FTOTSUB = 4 Then
  If Val(q_id.Caption) >= Question44 Then
   sub42.BackColor = &HE0E0E0
   sub42.ForeColor = &H733C00
   sub43.BackColor = &HE0E0E0
   sub43.ForeColor = &H733C00
   sub41.BackColor = &HE0E0E0
   sub41.ForeColor = &H733C00
   sub44.BackColor = &H68&
   sub44.ForeColor = &HFFFFFF
  ElseIf Val(q_id.Caption) < Question44 And Val(q_id.Caption) >= Question43 Then
   sub42.BackColor = &HE0E0E0
   sub42.ForeColor = &H733C00
   sub44.BackColor = &HE0E0E0
   sub44.ForeColor = &H733C00
   sub41.BackColor = &HE0E0E0
   sub41.ForeColor = &H733C00
   sub43.BackColor = &H68&
   sub43.ForeColor = &HFFFFFF
  ElseIf Val(q_id.Caption) < Question43 And Val(q_id.Caption) >= Question42 Then
    sub44.BackColor = &HE0E0E0
    sub44.ForeColor = &H733C00
    sub43.BackColor = &HE0E0E0
    sub43.ForeColor = &H733C00
    sub41.BackColor = &HE0E0E0
    sub41.ForeColor = &H733C00
    sub42.BackColor = &H68&
    sub42.ForeColor = &HFFFFFF
  ElseIf Val(q_id.Caption) < Question42 Then
    sub42.BackColor = &HE0E0E0
    sub42.ForeColor = &H733C00
    sub43.BackColor = &HE0E0E0
    sub43.ForeColor = &H733C00
    sub44.BackColor = &HE0E0E0
    sub44.ForeColor = &H733C00
    sub41.BackColor = &H68&
    sub41.ForeColor = &HFFFFFF
  End If
 ElseIf FTOTSUB = 5 Then
    If Val(q_id.Caption) >= Question55 Then
     sub52.BackColor = &HE0E0E0
     sub52.ForeColor = &H733C00
     sub53.BackColor = &HE0E0E0
     sub53.ForeColor = &H733C00
     Sub54.BackColor = &HE0E0E0
     Sub54.ForeColor = &H733C00
     sub51.BackColor = &HE0E0E0
     sub51.ForeColor = &H733C00
     sub55.BackColor = &H68&
     sub55.ForeColor = &HFFFFFF
  ElseIf Val(q_id.Caption) < Question55 And Val(q_id.Caption) >= Question54 Then
     sub52.BackColor = &HE0E0E0
     sub52.ForeColor = &H733C00
     sub53.BackColor = &HE0E0E0
     sub53.ForeColor = &H733C00
     sub55.BackColor = &HE0E0E0
     sub55.ForeColor = &H733C00
     sub51.BackColor = &HE0E0E0
     sub51.ForeColor = &H733C00
     Sub54.BackColor = &H68&
     Sub54.ForeColor = &HFFFFFF
  ElseIf Val(q_id.Caption) < Question54 And Val(q_id.Caption) >= Question53 Then
     sub52.BackColor = &HE0E0E0
     sub52.ForeColor = &H733C00
     Sub54.BackColor = &HE0E0E0
     Sub54.ForeColor = &H733C00
     sub55.BackColor = &HE0E0E0
     sub55.ForeColor = &H733C00
     sub51.BackColor = &HE0E0E0
     sub51.ForeColor = &H733C00
     sub53.BackColor = &H68&
     sub53.ForeColor = &HFFFFFF
  ElseIf Val(q_id.Caption) < Question53 And Val(q_id.Caption) >= Question52 Then
     Sub54.BackColor = &HE0E0E0
     Sub54.ForeColor = &H733C00
     sub53.BackColor = &HE0E0E0
     sub53.ForeColor = &H733C00
     sub55.BackColor = &HE0E0E0
     sub55.ForeColor = &H733C00
     sub51.BackColor = &HE0E0E0
     sub51.ForeColor = &H733C00
     sub52.BackColor = &H68&
     sub52.ForeColor = &HFFFFFF
  ElseIf Val(q_id.Caption) < Question52 Then
     sub52.BackColor = &HE0E0E0
     sub52.ForeColor = &H733C00
     sub53.BackColor = &HE0E0E0
     sub53.ForeColor = &H733C00
     sub55.BackColor = &HE0E0E0
     sub55.ForeColor = &H733C00
     Sub54.BackColor = &HE0E0E0
     Sub54.ForeColor = &H733C00
     sub51.BackColor = &H68&
     sub51.ForeColor = &HFFFFFF
  End If
 End If
End If
'----------+++++++++++++++------------+++++++++++++++++--------
If Check1.Value = vbChecked And lbl_usr_ans.Caption = "" Then 'Not Answered but bookmark
   btn(Val(q_id.Caption)).BackColor = &HCA30DF
   c.Execute ("update answerhold set BOOKMRK=1 where id=" & q_id.Caption & "")
 ElseIf Check1.Value = vbChecked And lbl_usr_ans.Caption <> "" Then 'Answered and bookmark
   btn(Val(q_id.Caption)).BackColor = &H800080
   c.Execute ("update answerhold set BOOKMRK=1 where id=" & q_id.Caption & "")
ElseIf Check1.Value = vbUnchecked And lbl_usr_ans.Caption = "" Then 'NOT ANSWERED and not bookmark
     btn(Val(q_id.Caption)).BackColor = &H100FF     '&H8130FF
     c.Execute ("update answerhold set BOOKMRK=2 where id=" & q_id.Caption & "")
ElseIf Check1.Value = vbUnchecked And lbl_usr_ans.Caption <> "" Then 'ANswered and not bookmark
     btn(Val(q_id.Caption)).BackColor = &HC000&
     c.Execute ("update answerhold set BOOKMRK=2 where id=" & q_id.Caption & "")
End If
End Function

Private Sub btn_GotFocus(Index As Integer) '''''' checking
For i = 1 To totQues
If i = Index Then
 btn(Index).Font.Bold = True
Else
 btn(Index).Font.Bold = False
End If
Next i
bookmrkColor
End Sub

Private Sub btn_LostFocus(Index As Integer)   '''''' checking
btn(Index).Font.Bold = False
bookmrkColor
End Sub

Private Sub Check1_Click() 'BookMark Facilities
bookmrkColor
End Sub

Private Sub btnNXT_PREV_Timer() 'Timer For Button to be enabled
If Val(q_id.Caption) = 1 And totQues = 1 Then
 nextq.Enabled = False  'Next
 svNext.Enabled = False
 prevQ.Enabled = False 'Previous
 btnSubmit.Enabled = True
ElseIf Val(q_id.Caption) = 1 And totQues > 1 Then
 nextq.Enabled = True  'Next
 svNext.Enabled = True
  prevQ.Enabled = False 'Previous
 btnSubmit.Enabled = False
ElseIf Val(q_id.Caption) = totQues Then 'reach at last
 nextq.Enabled = False 'Next
 svNext.Enabled = False
 btnSubmit.Enabled = True
 prevQ.Enabled = True  'Previous
Else
 nextq.Enabled = True 'Next
 svNext.Enabled = True
 prevQ.Enabled = True 'Previous
 btnSubmit.Enabled = False
 End If
End Sub

Private Sub clear_Click() 'Clearing the answer
bookmark.Caption = ""
indx.Caption = ""
For i = 0 To 3
 options(i).Value = False
Next i
 lbl_usr_ans.Caption = ""
c.Execute ("update answerhold set BOOKMRK=2,user_ans=0 where id=" & Val(q_id.Caption) & "")
Check1_Click
End Sub

'++++++++++++++No  More Required But Important ++++++++++++++++++'
'Public Function bookmrk_module()
'Set r = New ADODB.Recordset
'Set r = c.Execute("select USER_ANS,BOOKMRK from answerhold where ID='" & Val(q_id.Caption) & "' ")
' If r.Fields(1) = 1 And r.Fields(0) <> 0 Then
'  check1.Value = Checked
'  options(r.Fields(0) - 1).Value = True
' ElseIf r.Fields(1) = 1 And r.Fields(0) = 0 Then
'  check1.Value = Checked
'  For i = 0 To 3
'   options(i).Value = False
'  Next i
'  ElseIf r.Fields(1) = 2 And r.Fields(0) <> 0 Then
'   check1.Value = Unchecked
'   options(r.Fields(0) - 1).Value = True
'ElseIf r.Fields(1) = 2 And r.Fields(0) = 0 Then
'  check1.Value = Unchecked
'  For i = 0 To 3
'   options(i).Value = False
'  Next i
' End If
'  Check1_Click
'End Function

Private Sub btnsubmit_Click()
Set r = New ADODB.Recordset
Set r = c.Execute("select BOOKMRK from answerhold")
While r.EOF = False
 If r.Fields(0) = 1 Then
 MsgBox "Some Questions are still BookMarked." & vbCrLf & "Clear Them to Submit Test ", vbInformation + vbOKOnly, "Submit Test"
 Exit Sub
 End If
r.MoveNext
Wend

If Timer1.Enabled = True Then
 If (MsgBox("Are You Sure To Submit .", vbYesNo + vbInformation, "SUBMIT ") = vbYes) Then
  Timer1.Enabled = False
  Timer2.Enabled = False
  remainTIM = Format(minute, "00") & ":" & Format(second, "00")
  btnNXT_PREV.Enabled = False
  Unload Me
  Summary_Test.Show
 Exit Sub
 End If
End If
End Sub

Private Sub Command1_Click() 'Pause Test
ResumeTEST.Show 1, MDI
End Sub

Private Sub Form_Load()  '++++++++++++FORM LOADING ++++++++++++++++
On Error Resume Next
Dim i As Integer
Dim img_pth As String
Dim nell As String
conn
Label8.Caption = "+" & FMRKPERCOR
If FMRKPERWRONG = 0 Then
 Label27.Caption = 0
Else
 Label27.Caption = "-" & FMRKPERWRONG
End If
'Setting Section Button Here'
If IsFullLengthSelected = 1 Then
 TestTypeInfo.Visible = False
  Frame4One.Top = 840
  Frame4Two.Top = 840
  Frame4Three.Top = 840
  Frame4Four.Top = 840
  Frame4Five.Top = 840
  Frame4One.Left = 1440
  Frame4Two.Left = 1440
  Frame4Three.Left = 1440
  Frame4Four.Left = 1440
  Frame4Five.Left = 1440
 If FTOTSUB = 1 Then
 Set r = New ADODB.Recordset
 Set r = c.Execute("select initcap(sub_nm) from sub where sub_id='" & subName11 & "' ")
  Sub11.Caption = r.Fields(0)
  Frame4One.Visible = True
  Frame4Two.Visible = False
  Frame4Three.Visible = False
  Frame4Four.Visible = False
  Frame4Five.Visible = False
 ElseIf FTOTSUB = 2 Then
 Set r1 = New ADODB.Recordset
 Set r1 = c.Execute("select initcap(sub_nm) from sub where sub_id='" & subName21 & "' ")
 Set r2 = New ADODB.Recordset
 Set r2 = c.Execute("select initcap(sub_nm) from sub where sub_id='" & subName22 & "' ")
  sub21.Caption = r1.Fields(0)
  sub22.Caption = r2.Fields(0)
  Frame4One.Visible = False
  Frame4Two.Visible = True
  Frame4Three.Visible = False
  Frame4Four.Visible = False
  Frame4Five.Visible = False
 ElseIf FTOTSUB = 3 Then
 Set r1 = New ADODB.Recordset
 Set r1 = c.Execute("select initcap(sub_nm) from sub where sub_id='" & subName31 & "' ")
 Set r2 = New ADODB.Recordset
 Set r2 = c.Execute("select initcap(sub_nm) from sub where sub_id='" & subName32 & "' ")
 Set r3 = New ADODB.Recordset
 Set r3 = c.Execute("select initcap(sub_nm) from sub where sub_id='" & subName33 & "' ")
  sub31.Caption = r1.Fields(0)
  sub32.Caption = r2.Fields(0)
  sub33.Caption = r3.Fields(0)
  Frame4One.Visible = False
  Frame4Two.Visible = False
  Frame4Three.Visible = True
  Frame4Four.Visible = False
  Frame4Five.Visible = False
 ElseIf FTOTSUB = 4 Then
 Set r1 = New ADODB.Recordset
 Set r1 = c.Execute("select initcap(sub_nm) from sub where sub_id='" & subName41 & "' ")
 Set r2 = New ADODB.Recordset
 Set r2 = c.Execute("select initcap(sub_nm) from sub where sub_id='" & subName42 & "' ")
 Set r3 = New ADODB.Recordset
 Set r3 = c.Execute("select initcap(sub_nm) from sub where sub_id='" & subName43 & "' ")
 Set r4 = New ADODB.Recordset
 Set r4 = c.Execute("select initcap(sub_nm) from sub where sub_id='" & subName44 & "' ")
  sub41.Caption = r1.Fields(0)
  sub42.Caption = r2.Fields(0)
  sub43.Caption = r3.Fields(0)
  sub44.Caption = r4.Fields(0)
  Frame4One.Visible = False
  Frame4Two.Visible = False
  Frame4Three.Visible = False
  Frame4Four.Visible = True
  Frame4Five.Visible = False
 ElseIf FTOTSUB = 5 Then
 Set r1 = New ADODB.Recordset
 Set r1 = c.Execute("select initcap(sub_nm) from sub where sub_id='" & subName51 & "' ")
 Set r2 = New ADODB.Recordset
 Set r2 = c.Execute("select initcap(sub_nm) from sub where sub_id='" & subName52 & "' ")
 Set r3 = New ADODB.Recordset
 Set r3 = c.Execute("select initcap(sub_nm) from sub where sub_id='" & subName53 & "' ")
 Set r4 = New ADODB.Recordset
 Set r4 = c.Execute("select initcap(sub_nm) from sub where sub_id='" & subName54 & "' ")
 Set r = New ADODB.Recordset
 Set r = c.Execute("select initcap(sub_nm) from sub where sub_id='" & subName55 & "' ")
  sub51.Caption = r1.Fields(0)
  sub52.Caption = r2.Fields(0)
  sub53.Caption = r3.Fields(0)
  Sub54.Caption = r4.Fields(0)
  sub55.Caption = r.Fields(0)
  Frame4One.Visible = False
  Frame4Two.Visible = False
  Frame4Three.Visible = False
  Frame4Four.Visible = False
  Frame4Five.Visible = True
 End If
 Else 'If Full Length Test Is Not Selected
  Frame4One.Visible = False
  Frame4Two.Visible = False
  Frame4Three.Visible = False
  Frame4Four.Visible = False
  Frame4Five.Visible = False
  TestTypeInfo.Visible = True
  If UCase(selectedType) = UCase("Subject Wise Test") Then
   TestTypeInfo.Caption = " Subject Wise Test : [Subject]:- (" & CurrentSub & ") "
  Else
   TestTypeInfo.Caption = " Topic Wise Test : [Subject]:- " & CurrentSub & " [Topic] :- " & CurrentTopic & ") "
  End If
End If
'++++++++++++++++++++++++++++++
MCQ.Top = 25
MCQ.Left = 50
Me.Width = MDI.Width
Me.Height = MDI.Height
Set r = New ADODB.Recordset
Set r = c.Execute("select RSTUD_NM,C_ID,RSTUD_PIC from rstud where RSTUD_REG_NO='" & Current_Logged_ID & "' ")
 oName.Caption = r.Fields(0)
 StuNam = r.Fields(0) 'Display On Summary
 Label3.Caption = "(" & Current_Logged_ID & ")"
If IsNull(r.Fields(2)) = False Then
  nell = r.Fields(1)
  img_pth = r.Fields(2)
  stuPIC.Picture = LoadPicture(img_pth)
End If
  Set r1 = New ADODB.Recordset
  Set r1 = c.Execute("select initcap(c_nm) from course where c_id='" & nell & "' ")
  Heading.Caption = r1.Fields(0)
  GivenTESTCourse = Heading.Caption 'Dislay while seeing summary
'Showing Only Equivalent Button
For i = 1 To 60
   btn(i).Visible = False
   Shape(i).Visible = False 'InVisible All Yellow Circle
Next i
Set r = New ADODB.Recordset
Set r = c1.Execute("select count(*) from MCQTEST")
 totQues = r.Fields(0)
 For i = 1 To totQues
  btn(i).Visible = True
 Next i
'--Loading Question Into Panel--
Set r1 = New ADODB.Recordset
sql = "select * from MCQTEST "
Set r1 = c.Execute(sql)
q_id.Caption = r1.Fields(0)
qtext.Caption = r1.Fields(1)
For i = 0 To 3
 options(i).Caption = r1.Fields(i + 2)
Next i
lbl_corr_ans.Caption = r1.Fields(7)
indx.Caption = lbl_corr_ans.Caption - 1
Shape(1).Visible = True
'++++++++++++++++++++++++++++++++++++++
If IsNull(r1.Fields(8)) = False Then
 Image1.Visible = True
 Image1.Stretch = True
 Quespic = r1.Fields(8)
 Image1.Picture = LoadPicture(Quespic)
 Else
 Quespic = ""
 Image1.Stretch = True
 Image1.Visible = False
 Image1.Picture = Nothing
End If
''------------To maintain Answer Sheet-----------------
c1.Execute ("delete from answerhold")
Dim I1 As Integer
For I1 = 1 To totQues
 Set r = New ADODB.Recordset
 Set r = c.Execute("select ANS_NO from MCQTEST where Q_NO=" & I1 & " order by Q_NO")
  c.Execute ("Insert Into answerHold values(" & I1 & "," & r.Fields(0) & ",DEFAULT,DEFAULT)")
 Next I1

Timer1.Enabled = False
Timer2.Enabled = False
btnNXT_PREV.Enabled = False
'Set Time Properties
minute = FTOTTIMEMINUTE 'Set Here Time In Minute
second = FTOTTIMESECOND 'Set Here Time in Second
ToTaTiMe = Format(minute, "00") & " : " & Format(second, "00")
Total4InstructionPage = totQues
min4InstructionPage = minute
Timer1.Interval = 1000 'For Remaining Time
Timer2.Interval = 1000 'For Elapsed Time

currenttime = Now
t_time_min = Space(2) & Format(Now - currenttime, "hh : MM : ss")
Timer1.Enabled = True
Timer2.Enabled = True
''-----------------------------------------------------------------
'Option Panel Button Status
For i = 1 To totQues
 btn(i).Font.Bold = True
 btn(i).Font.Name = "Microsoft Sans Serif"
 btn(i).Font.Size = 12
 btn(i).BackColor = &HC0FFFF
Next i

Optkey.Caption = "" 'initialise starting to blank
'+++++++++++++++Used Timers++++++++++++++++++
btnNXT_PREV.Enabled = True
Check1.Value = Unchecked
bookmrkColor
End Sub

Private Sub Image2_Click() 'For Listening the Question
Set objSpeech = CreateObject("SAPI.SpVoice")
objSpeech.speak "Question " & qtext.Caption
objSpeech.speak "Option A" & options(0).Caption
objSpeech.speak "Option B" & options(1).Caption
objSpeech.speak "Option C" & options(2).Caption
objSpeech.speak "Option D" & options(3).Caption
End Sub

Private Sub Image3_Click() 'Rough Page
Rstud_Rough.Show 1, MDI
End Sub

Private Sub Label28_Click()
Command1_Click
End Sub

Private Sub Label29_Click()
Image3_Click
End Sub

Private Sub Label30_Click()
Image2_Click
End Sub

Private Sub nextQ_Click() 'Next Button
On Error Resume Next
Set r = New ADODB.Recordset
Set r = c.Execute("select * from MCQTEST where q_no=" & Val(q_id.Caption) + 1 & "")
 q_id.Caption = r.Fields(0)
 btn(Val(q_id.Caption)).SetFocus
 qtext.Caption = r.Fields(1)
For i = 0 To 3
 options(i).Caption = r.Fields(i + 2)
Next i
lbl_corr_ans.Caption = r.Fields(7)  'Answer
indx.Caption = lbl_corr_ans.Caption - 1
If IsNull(r.Fields(8)) = False Then
 Image1.Visible = True
 Quespic = r.Fields(8)
 Image1.Picture = LoadPicture(Quespic)
 Else
 Quespic = ""
 Image1.Visible = False
End If
For i = 1 To 60
 If i = Val(q_id.Caption) Then
  Shape(i).Visible = True
 Else
  Shape(i).Visible = False
 End If
Next i
ChkOption
ChkUserANS
CHKbookmark
bookmrkColor
End Sub

Private Sub prevQ_Click() 'Previous Button
Text1.SetFocus
Set r = New ADODB.Recordset
Set r = c.Execute("select * from MCQTEST where Q_no=" & Val(q_id.Caption) - 1 & "")
q_id.Caption = r.Fields(0)
btn(Val(q_id.Caption)).SetFocus
qtext.Caption = r.Fields(1)
For i = 0 To 3
 options(i).Caption = r.Fields(i + 2)
Next i
lbl_corr_ans.Caption = r.Fields(7)
indx.Caption = lbl_corr_ans.Caption - 1
If IsNull(r.Fields(8)) = False Then
 Image1.Visible = True
 Quespic = r.Fields(8)
 Image1.Picture = LoadPicture(Quespic)
 Else
 Quespic = ""
 Image1.Visible = False
End If
For i = 1 To 60
 If i = Val(q_id.Caption) Then
  Shape(i).Visible = True
 Else
  Shape(i).Visible = False
 End If
Next i
ChkOption
ChkUserANS
CHKbookmark
bookmrkColor
End Sub

Private Sub options_Click(Index As Integer) 'click on options
indx.Caption = Index
lbl_usr_ans.Caption = Index + 1
c.Execute ("update answerhold set user_ans=" & Val(lbl_usr_ans.Caption) & " where ID=" & Val(q_id.Caption) & "")
CHKbookmark
bookmrkColor
End Sub

Private Sub Sub11_Click() 'If One Section Exist
btn_Click (Question11)
End Sub

Private Sub Sub21_Click()
sub22.BackColor = &HE0E0E0
sub22.BackOver = &HE0E0E0
sub22.ForeColor = &H733C00
sub22.ForeOver = &H733C00
sub21.BackColor = &H68&
sub21.BackOver = &H68&
sub21.ForeColor = &HFFFFFF
sub21.ForeOver = &HFFFFFF
btn_Click (Question21)
End Sub

Private Sub Sub22_Click()
sub21.BackColor = &HE0E0E0
sub21.BackOver = &HE0E0E0
sub21.ForeColor = &H733C00
sub21.ForeOver = &H733C00
sub22.BackColor = &H68&
sub22.BackOver = &H68&
sub22.ForeColor = &HFFFFFF
sub22.ForeOver = &HFFFFFF
btn_Click (Question22)
End Sub

Private Sub Sub31_Click()
sub32.BackColor = &HE0E0E0
sub32.BackOver = &HE0E0E0
sub32.ForeColor = &H733C00
sub32.ForeOver = &H733C00
sub33.BackColor = &HE0E0E0
sub33.ForeColor = &H733C00
sub32.ForeOver = &H733C00
sub33.BackOver = &HE0E0E0
sub31.BackColor = &H68&
sub31.BackOver = &H68&
sub31.ForeColor = &HFFFFFF
sub31.ForeOver = &HFFFFFF
btn_Click (Question31)
End Sub

Private Sub sub32_Click()
sub33.BackColor = &HE0E0E0
sub33.BackOver = &HE0E0E0
sub33.ForeColor = &H733C00
sub33.ForeOver = &H733C00
sub31.BackColor = &HE0E0E0
sub31.ForeColor = &H733C00
sub31.ForeOver = &H733C00
sub31.BackOver = &HE0E0E0
sub32.BackColor = &H68&
sub32.BackOver = &H68&
sub32.ForeColor = &HFFFFFF
sub32.ForeOver = &HFFFFFF
btn_Click (Question32)
End Sub

Private Sub sub33_Click()
sub32.BackColor = &HE0E0E0
sub32.BackOver = &HE0E0E0
sub32.ForeColor = &H733C00
sub32.ForeOver = &H733C00
sub31.BackColor = &HE0E0E0
sub31.ForeColor = &H733C00
sub31.ForeOver = &H733C00
sub31.BackOver = &HE0E0E0
sub33.BackColor = &H68&
sub33.BackOver = &H68&
sub33.ForeColor = &HFFFFFF
sub33.ForeOver = &HFFFFFF
btn_Click (Question33)
End Sub

Private Sub sub41_Click()
sub42.BackColor = &HE0E0E0
sub42.BackOver = &HE0E0E0
sub42.ForeColor = &H733C00
sub42.ForeOver = &H733C00
sub44.BackColor = &HE0E0E0
sub44.ForeColor = &H733C00
sub44.ForeOver = &H733C00
sub44.BackOver = &HE0E0E0
sub43.BackColor = &HE0E0E0
sub43.ForeColor = &H733C00
sub43.ForeOver = &H733C00
sub43.BackOver = &HE0E0E0
sub41.BackColor = &H68&
sub41.BackOver = &H68&
sub41.ForeColor = &HFFFFFF
sub41.ForeOver = &HFFFFFF
btn_Click (Question41)
End Sub

Private Sub sub42_Click()
sub43.BackColor = &HE0E0E0
sub43.BackOver = &HE0E0E0
sub43.ForeColor = &H733C00
sub43.ForeOver = &H733C00
sub44.BackColor = &HE0E0E0
sub44.ForeColor = &H733C00
sub44.ForeOver = &H733C00
sub44.BackOver = &HE0E0E0
sub41.BackColor = &HE0E0E0
sub41.ForeColor = &H733C00
sub41.ForeOver = &H733C00
sub41.BackOver = &HE0E0E0
sub42.BackColor = &H68&
sub42.BackOver = &H68&
sub42.ForeColor = &HFFFFFF
sub42.ForeOver = &HFFFFFF
btn_Click (Question42)
End Sub

Private Sub sub43_Click()
sub42.BackColor = &HE0E0E0
sub42.ForeColor = &H733C00
sub42.ForeOver = &H733C00
sub42.BackOver = &HE0E0E0
sub44.BackColor = &HE0E0E0
sub44.ForeColor = &H733C00
sub44.ForeOver = &H733C00
sub44.BackOver = &HE0E0E0
sub41.BackColor = &HE0E0E0
sub41.ForeColor = &H733C00
sub41.ForeOver = &H733C00
sub41.BackOver = &HE0E0E0
sub43.BackColor = &H68&
sub43.BackOver = &H68&
sub43.ForeColor = &HFFFFFF
sub43.ForeOver = &HFFFFFF
btn_Click (Question43)
End Sub

Private Sub sub44_Click()
sub42.BackColor = &HE0E0E0
sub42.ForeColor = &H733C00
sub42.ForeOver = &H733C00
sub43.BackColor = &HE0E0E0
sub43.ForeColor = &H733C00
sub43.ForeOver = &H733C00
sub43.BackOver = &HE0E0E0
sub41.BackColor = &HE0E0E0
sub41.ForeColor = &H733C00
sub41.ForeOver = &H733C00
sub41.BackOver = &HE0E0E0
sub44.BackColor = &H68&
sub44.BackOver = &H68&
sub44.ForeColor = &HFFFFFF
sub44.ForeOver = &HFFFFFF
btn_Click (Question44)
End Sub

Private Sub sub51_Click()
sub52.BackColor = &HE0E0E0
sub52.ForeColor = &H733C00
sub52.ForeOver = &H733C00
sub52.BackOver = &HE0E0E0
sub53.BackColor = &HE0E0E0
sub53.ForeColor = &H733C00
sub53.ForeOver = &H733C00
sub53.BackOver = &HE0E0E0
Sub54.BackColor = &HE0E0E0
Sub54.ForeColor = &H733C00
Sub54.ForeOver = &H733C00
Sub54.BackOver = &HE0E0E0
sub55.BackColor = &HE0E0E0
sub55.ForeColor = &H733C00
sub55.ForeOver = &H733C00
sub55.BackOver = &HE0E0E0
sub51.BackColor = &H68&
sub51.BackOver = &H68&
sub51.ForeColor = &HFFFFFF
sub51.ForeOver = &HFFFFFF
btn_Click (Question51)
End Sub

Private Sub sub52_Click()
sub51.BackColor = &HE0E0E0
sub51.BackOver = &HE0E0E0
sub51.ForeColor = &H733C00
sub51.ForeOver = &H733C00
sub51.BackColor = &HE0E0E0
sub53.BackColor = &HE0E0E0
sub53.ForeColor = &H733C00
sub53.ForeOver = &H733C00
sub53.BackColor = &HE0E0E0
sub53.BackOver = &HE0E0E0
Sub54.BackColor = &HE0E0E0
Sub54.ForeColor = &H733C00
Sub54.ForeOver = &H733C00
Sub54.BackColor = &HE0E0E0
Sub54.BackOver = &HE0E0E0
sub55.BackColor = &HE0E0E0
sub55.ForeColor = &H733C00
sub55.ForeOver = &H733C00
sub55.BackColor = &HE0E0E0
sub55.BackOver = &HE0E0E0
sub52.BackColor = &H68&
sub52.BackOver = &HE0E0E0
sub52.ForeColor = &HFFFFFF
sub52.ForeOver = &HFFFFFF
btn_Click (Question52)
End Sub

Private Sub sub53_Click()
sub52.BackColor = &HE0E0E0
sub52.BackOver = &HE0E0E0
sub52.ForeColor = &H733C00
sub52.ForeOver = &H733C00
sub51.BackColor = &HE0E0E0
sub51.ForeColor = &H733C00
sub51.ForeOver = &H733C00
sub51.BackOver = &HE0E0E0
Sub54.BackColor = &HE0E0E0
Sub54.ForeColor = &H733C00
Sub54.ForeOver = &H733C00
Sub54.BackOver = &HE0E0E0
sub55.BackColor = &HE0E0E0
sub55.ForeColor = &H733C00
sub55.ForeOver = &H733C00
sub55.BackOver = &HE0E0E0
sub53.BackColor = &H68&
sub53.BackOver = &H68&
sub53.ForeColor = &HFFFFFF
sub53.ForeOver = &HFFFFFF
btn_Click (Question53)
End Sub

Private Sub sub54_Click()
sub52.BackColor = &HE0E0E0
sub52.BackOver = &HE0E0E0
sub52.ForeColor = &H733C00
sub52.ForeOver = &H733C00
sub53.BackColor = &HE0E0E0
sub53.ForeColor = &H733C00
sub53.ForeOver = &H733C00
sub53.BackOver = &HE0E0E0
sub51.BackColor = &HE0E0E0
sub51.ForeColor = &H733C00
sub51.ForeOver = &H733C00
sub51.BackOver = &HE0E0E0
sub55.BackColor = &HE0E0E0
sub55.ForeColor = &H733C00
sub55.ForeOver = &H733C00
sub55.BackOver = &HE0E0E0
Sub54.BackColor = &H68&
Sub54.BackOver = &H68&
Sub54.ForeColor = &HFFFFFF
Sub54.ForeOver = &HFFFFFF
btn_Click (Question54)
End Sub

Private Sub sub55_Click()
sub52.BackColor = &HE0E0E0
sub52.BackOver = &HE0E0E0
sub52.ForeColor = &H733C00
sub52.ForeOver = &H733C00
sub53.BackColor = &HE0E0E0
sub53.ForeColor = &H733C00
sub53.ForeOver = &H733C00
sub53.BackOver = &HE0E0E0
Sub54.BackColor = &HE0E0E0
Sub54.ForeColor = &H733C00
Sub54.ForeOver = &H733C00
Sub54.BackOver = &HE0E0E0
sub51.BackColor = &HE0E0E0
sub51.ForeColor = &H733C00
sub51.ForeOver = &H733C00
sub51.BackOver = &HE0E0E0
sub55.BackColor = &H68&
sub55.BackOver = &H68&
sub55.ForeColor = &HFFFFFF
sub55.ForeOver = &HFFFFFF
btn_Click (Question55)
End Sub

Private Sub svNext_Click()
nextQ_Click
End Sub

Private Sub Timer1_Timer() 'For time remaining
Timer_remain.Caption = Format(minute, "00") & " : " & Format(second, "00")
If second = 0 And minute <> 0 Then
 minute = minute - 1
 second = 59
ElseIf second = 0 And minute = 0 Then
 Timer1.Enabled = False
 Timer2.Enabled = False
 btnNXT_PREV.Enabled = False
 opt = MsgBox("Time Over ! Go To Test Summary" & vbCrLf & "Bookmarked Questions have been cleared automatically ", vbInformation + vbOKOnly, "EXAM TIME OVER")
 Set r = New ADODB.Recordset
 Set r = c.Execute("select ID,BOOKMRK from answerhold")
 While r.EOF = False
 tmpp = r.Fields(0)
  If r.Fields(1) = 1 Then
   c.Execute ("update answerhold set bookmrk=2 where ID=" & tmpp & " ")
  End If
r.MoveNext
Wend

 remainTIM = Format(minute, "00") & ":" & Format(second, "00")
 Unload Me
 Summary_Test.Show
End If
second = second - 1
End Sub

Private Sub Timer2_Timer() 'total Passed Time
t_time_min.Caption = "Elapsed Time - " & Format$(Now - currenttime, "hh : mm : ss")  'should be stop when minute =5
End Sub

Private Sub vkCommand1_Click() 'Exit Button
If MsgBox("        Are You Sure To Terminate Exam Abnormally " & vbCrLf & "Remember this test will not be present in Your Previous record section", vbCritical + vbYesNo, "Warning") = vbYes Then
 c.Execute ("delete from answerhold")
 c.Execute ("delete from mcqtest")
 Timer1.Enabled = False
 Timer2.Enabled = False
 btnNXT_PREV.Enabled = False
 Unload Me
stu_dash.Show
End If
End Sub

Private Sub vkCommand2_Click() 'Showing Paper
'mcqTestRunPPr.Orientation = rptOrientPortrait
mcqTestRunPPr.Show vbModal, MDI
End Sub

Private Sub vkCommand3_Click() 'Show Instructions
Instruction_test.Show vbModal, MDI
End Sub
