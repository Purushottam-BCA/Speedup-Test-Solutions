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
   ClientWidth     =   20355
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "mS_Exam_Interface.frx":0000
   LinkTopic       =   "Form14"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10515
   ScaleWidth      =   20355
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton TestTypeInfo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Topic Wise Test"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1440
      MouseIcon       =   "mS_Exam_Interface.frx":6062
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   137
      Top             =   855
      Width           =   14445
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H80000013&
      Height          =   1380
      Left            =   15980
      ScaleHeight     =   1320
      ScaleWidth      =   4335
      TabIndex        =   133
      Top             =   9180
      Width           =   4400
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
         MouseIcon       =   "mS_Exam_Interface.frx":61B4
         MousePointer    =   99  'Custom
         Picture         =   "mS_Exam_Interface.frx":6306
         Style           =   1  'Graphical
         TabIndex        =   136
         ToolTipText     =   "See All Questions at a glance"
         Top             =   120
         Width           =   1960
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
         Left            =   2330
         MouseIcon       =   "mS_Exam_Interface.frx":6E65
         MousePointer    =   99  'Custom
         Picture         =   "mS_Exam_Interface.frx":6FB7
         Style           =   1  'Graphical
         TabIndex        =   135
         ToolTipText     =   "Show Instruction Page"
         Top             =   120
         Width           =   1845
      End
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
         Height          =   480
         Left            =   200
         MouseIcon       =   "mS_Exam_Interface.frx":78E6
         MousePointer    =   99  'Custom
         Picture         =   "mS_Exam_Interface.frx":7A38
         Style           =   1  'Graphical
         TabIndex        =   134
         ToolTipText     =   "Submit The Test"
         Top             =   700
         Width           =   3965
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000013&
      Height          =   9165
      Left            =   15980
      ScaleHeight     =   9105
      ScaleWidth      =   4335
      TabIndex        =   56
      Top             =   0
      Width           =   4395
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   1
         Left            =   120
         TabIndex        =   66
         Top             =   2295
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
         Picture         =   "mS_Exam_Interface.frx":865C
         DisabledPicture =   "mS_Exam_Interface.frx":8678
         DownPicture     =   "mS_Exam_Interface.frx":8694
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":86B0
         Caption         =   "1"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   2
         Left            =   840
         TabIndex        =   67
         Top             =   2295
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1138
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "mS_Exam_Interface.frx":8812
         DisabledPicture =   "mS_Exam_Interface.frx":882E
         DownPicture     =   "mS_Exam_Interface.frx":884A
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":8866
         Caption         =   "2"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   3
         Left            =   1560
         TabIndex        =   68
         Top             =   2295
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
         Picture         =   "mS_Exam_Interface.frx":89C8
         DisabledPicture =   "mS_Exam_Interface.frx":89E4
         DownPicture     =   "mS_Exam_Interface.frx":8A00
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":8A1C
         Caption         =   "3"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   4
         Left            =   2280
         TabIndex        =   69
         Top             =   2295
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
         Picture         =   "mS_Exam_Interface.frx":8B7E
         DisabledPicture =   "mS_Exam_Interface.frx":8B9A
         DownPicture     =   "mS_Exam_Interface.frx":8BB6
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":8BD2
         Caption         =   "4"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   6
         Left            =   3720
         TabIndex        =   70
         Top             =   2295
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
         Picture         =   "mS_Exam_Interface.frx":8D34
         DisabledPicture =   "mS_Exam_Interface.frx":8D50
         DownPicture     =   "mS_Exam_Interface.frx":8D6C
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":8D88
         Caption         =   "6"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   5
         Left            =   3000
         TabIndex        =   71
         Top             =   2295
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
         Picture         =   "mS_Exam_Interface.frx":8EEA
         DisabledPicture =   "mS_Exam_Interface.frx":8F06
         DownPicture     =   "mS_Exam_Interface.frx":8F22
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":8F3E
         Caption         =   "5"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   55
         Left            =   120
         TabIndex        =   72
         Top             =   8400
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
         Picture         =   "mS_Exam_Interface.frx":90A0
         DisabledPicture =   "mS_Exam_Interface.frx":90BC
         DownPicture     =   "mS_Exam_Interface.frx":90D8
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":90F4
         Caption         =   "55"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   56
         Left            =   840
         TabIndex        =   73
         Top             =   8400
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
         Picture         =   "mS_Exam_Interface.frx":9256
         DisabledPicture =   "mS_Exam_Interface.frx":9272
         DownPicture     =   "mS_Exam_Interface.frx":928E
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":92AA
         Caption         =   "56"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   57
         Left            =   1560
         TabIndex        =   74
         Top             =   8400
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
         Picture         =   "mS_Exam_Interface.frx":940C
         DisabledPicture =   "mS_Exam_Interface.frx":9428
         DownPicture     =   "mS_Exam_Interface.frx":9444
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":9460
         Caption         =   "57"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   58
         Left            =   2280
         TabIndex        =   75
         Top             =   8400
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
         Picture         =   "mS_Exam_Interface.frx":95C2
         DisabledPicture =   "mS_Exam_Interface.frx":95DE
         DownPicture     =   "mS_Exam_Interface.frx":95FA
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":9616
         Caption         =   "58"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   60
         Left            =   3720
         TabIndex        =   76
         Top             =   8400
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
         Picture         =   "mS_Exam_Interface.frx":9778
         DisabledPicture =   "mS_Exam_Interface.frx":9794
         DownPicture     =   "mS_Exam_Interface.frx":97B0
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":97CC
         Caption         =   "60"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   59
         Left            =   3000
         TabIndex        =   77
         Top             =   8400
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
         Picture         =   "mS_Exam_Interface.frx":992E
         DisabledPicture =   "mS_Exam_Interface.frx":994A
         DownPicture     =   "mS_Exam_Interface.frx":9966
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":9982
         Caption         =   "59"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   53
         Left            =   3000
         TabIndex        =   78
         Top             =   7700
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
         Picture         =   "mS_Exam_Interface.frx":9AE4
         DisabledPicture =   "mS_Exam_Interface.frx":9B00
         DownPicture     =   "mS_Exam_Interface.frx":9B1C
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":9B38
         Caption         =   "53"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   54
         Left            =   3720
         TabIndex        =   79
         Top             =   7700
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
         Picture         =   "mS_Exam_Interface.frx":9C9A
         DisabledPicture =   "mS_Exam_Interface.frx":9CB6
         DownPicture     =   "mS_Exam_Interface.frx":9CD2
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":9CEE
         Caption         =   "54"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   52
         Left            =   2280
         TabIndex        =   80
         Top             =   7700
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
         Picture         =   "mS_Exam_Interface.frx":9E50
         DisabledPicture =   "mS_Exam_Interface.frx":9E6C
         DownPicture     =   "mS_Exam_Interface.frx":9E88
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":9EA4
         Caption         =   "52"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   51
         Left            =   1560
         TabIndex        =   81
         Top             =   7700
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
         Picture         =   "mS_Exam_Interface.frx":A006
         DisabledPicture =   "mS_Exam_Interface.frx":A022
         DownPicture     =   "mS_Exam_Interface.frx":A03E
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":A05A
         Caption         =   "51"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   50
         Left            =   840
         TabIndex        =   82
         Top             =   7700
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
         Picture         =   "mS_Exam_Interface.frx":A1BC
         DisabledPicture =   "mS_Exam_Interface.frx":A1D8
         DownPicture     =   "mS_Exam_Interface.frx":A1F4
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":A210
         Caption         =   "50"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   49
         Left            =   120
         TabIndex        =   83
         Top             =   7700
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
         Picture         =   "mS_Exam_Interface.frx":A372
         DisabledPicture =   "mS_Exam_Interface.frx":A38E
         DownPicture     =   "mS_Exam_Interface.frx":A3AA
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":A3C6
         Caption         =   "49"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   47
         Left            =   3000
         TabIndex        =   84
         Top             =   7000
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
         Picture         =   "mS_Exam_Interface.frx":A528
         DisabledPicture =   "mS_Exam_Interface.frx":A544
         DownPicture     =   "mS_Exam_Interface.frx":A560
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":A57C
         Caption         =   "47"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   48
         Left            =   3720
         TabIndex        =   85
         Top             =   7000
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
         Picture         =   "mS_Exam_Interface.frx":A6DE
         DisabledPicture =   "mS_Exam_Interface.frx":A6FA
         DownPicture     =   "mS_Exam_Interface.frx":A716
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":A732
         Caption         =   "48"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   46
         Left            =   2280
         TabIndex        =   86
         Top             =   7000
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
         Picture         =   "mS_Exam_Interface.frx":A894
         DisabledPicture =   "mS_Exam_Interface.frx":A8B0
         DownPicture     =   "mS_Exam_Interface.frx":A8CC
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":A8E8
         Caption         =   "46"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   45
         Left            =   1560
         TabIndex        =   87
         Top             =   7000
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
         Picture         =   "mS_Exam_Interface.frx":AA4A
         DisabledPicture =   "mS_Exam_Interface.frx":AA66
         DownPicture     =   "mS_Exam_Interface.frx":AA82
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":AA9E
         Caption         =   "45"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   44
         Left            =   840
         TabIndex        =   88
         Top             =   7000
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
         Picture         =   "mS_Exam_Interface.frx":AC00
         DisabledPicture =   "mS_Exam_Interface.frx":AC1C
         DownPicture     =   "mS_Exam_Interface.frx":AC38
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":AC54
         Caption         =   "44"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   43
         Left            =   120
         TabIndex        =   89
         Top             =   7000
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
         Picture         =   "mS_Exam_Interface.frx":ADB6
         DisabledPicture =   "mS_Exam_Interface.frx":ADD2
         DownPicture     =   "mS_Exam_Interface.frx":ADEE
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":AE0A
         Caption         =   "43"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   37
         Left            =   120
         TabIndex        =   90
         Top             =   6350
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
         Picture         =   "mS_Exam_Interface.frx":AF6C
         DisabledPicture =   "mS_Exam_Interface.frx":AF88
         DownPicture     =   "mS_Exam_Interface.frx":AFA4
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":AFC0
         Caption         =   "37"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   38
         Left            =   840
         TabIndex        =   91
         Top             =   6350
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
         Picture         =   "mS_Exam_Interface.frx":B122
         DisabledPicture =   "mS_Exam_Interface.frx":B13E
         DownPicture     =   "mS_Exam_Interface.frx":B15A
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":B176
         Caption         =   "38"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   39
         Left            =   1560
         TabIndex        =   92
         Top             =   6350
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
         Picture         =   "mS_Exam_Interface.frx":B2D8
         DisabledPicture =   "mS_Exam_Interface.frx":B2F4
         DownPicture     =   "mS_Exam_Interface.frx":B310
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":B32C
         Caption         =   "39"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   40
         Left            =   2280
         TabIndex        =   93
         Top             =   6350
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
         Picture         =   "mS_Exam_Interface.frx":B48E
         DisabledPicture =   "mS_Exam_Interface.frx":B4AA
         DownPicture     =   "mS_Exam_Interface.frx":B4C6
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":B4E2
         Caption         =   "40"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   42
         Left            =   3720
         TabIndex        =   94
         Top             =   6350
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
         Picture         =   "mS_Exam_Interface.frx":B644
         DisabledPicture =   "mS_Exam_Interface.frx":B660
         DownPicture     =   "mS_Exam_Interface.frx":B67C
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":B698
         Caption         =   "42"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   41
         Left            =   3000
         TabIndex        =   95
         Top             =   6350
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
         Picture         =   "mS_Exam_Interface.frx":B7FA
         DisabledPicture =   "mS_Exam_Interface.frx":B816
         DownPicture     =   "mS_Exam_Interface.frx":B832
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":B84E
         Caption         =   "41"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   31
         Left            =   120
         TabIndex        =   96
         Top             =   5670
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
         Picture         =   "mS_Exam_Interface.frx":B9B0
         DisabledPicture =   "mS_Exam_Interface.frx":B9CC
         DownPicture     =   "mS_Exam_Interface.frx":B9E8
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":BA04
         Caption         =   "31"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   32
         Left            =   840
         TabIndex        =   97
         Top             =   5670
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
         Picture         =   "mS_Exam_Interface.frx":BB66
         DisabledPicture =   "mS_Exam_Interface.frx":BB82
         DownPicture     =   "mS_Exam_Interface.frx":BB9E
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":BBBA
         Caption         =   "32"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   33
         Left            =   1560
         TabIndex        =   98
         Top             =   5670
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
         Picture         =   "mS_Exam_Interface.frx":BD1C
         DisabledPicture =   "mS_Exam_Interface.frx":BD38
         DownPicture     =   "mS_Exam_Interface.frx":BD54
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":BD70
         Caption         =   "33"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   34
         Left            =   2280
         TabIndex        =   99
         Top             =   5670
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
         Picture         =   "mS_Exam_Interface.frx":BED2
         DisabledPicture =   "mS_Exam_Interface.frx":BEEE
         DownPicture     =   "mS_Exam_Interface.frx":BF0A
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":BF26
         Caption         =   "34"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   36
         Left            =   3720
         TabIndex        =   100
         Top             =   5670
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
         Picture         =   "mS_Exam_Interface.frx":C088
         DisabledPicture =   "mS_Exam_Interface.frx":C0A4
         DownPicture     =   "mS_Exam_Interface.frx":C0C0
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":C0DC
         Caption         =   "36"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   35
         Left            =   3000
         TabIndex        =   101
         Top             =   5670
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
         Picture         =   "mS_Exam_Interface.frx":C23E
         DisabledPicture =   "mS_Exam_Interface.frx":C25A
         DownPicture     =   "mS_Exam_Interface.frx":C276
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":C292
         Caption         =   "35"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   25
         Left            =   120
         TabIndex        =   102
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
         Picture         =   "mS_Exam_Interface.frx":C3F4
         DisabledPicture =   "mS_Exam_Interface.frx":C410
         DownPicture     =   "mS_Exam_Interface.frx":C42C
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":C448
         Caption         =   "25"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   26
         Left            =   840
         TabIndex        =   103
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
         Picture         =   "mS_Exam_Interface.frx":C5AA
         DisabledPicture =   "mS_Exam_Interface.frx":C5C6
         DownPicture     =   "mS_Exam_Interface.frx":C5E2
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":C5FE
         Caption         =   "26"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   27
         Left            =   1560
         TabIndex        =   104
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
         Picture         =   "mS_Exam_Interface.frx":C760
         DisabledPicture =   "mS_Exam_Interface.frx":C77C
         DownPicture     =   "mS_Exam_Interface.frx":C798
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":C7B4
         Caption         =   "27"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   28
         Left            =   2280
         TabIndex        =   105
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
         Picture         =   "mS_Exam_Interface.frx":C916
         DisabledPicture =   "mS_Exam_Interface.frx":C932
         DownPicture     =   "mS_Exam_Interface.frx":C94E
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":C96A
         Caption         =   "28"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   30
         Left            =   3720
         TabIndex        =   106
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
         Picture         =   "mS_Exam_Interface.frx":CACC
         DisabledPicture =   "mS_Exam_Interface.frx":CAE8
         DownPicture     =   "mS_Exam_Interface.frx":CB04
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":CB20
         Caption         =   "30"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   29
         Left            =   3000
         TabIndex        =   107
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
         Picture         =   "mS_Exam_Interface.frx":CC82
         DisabledPicture =   "mS_Exam_Interface.frx":CC9E
         DownPicture     =   "mS_Exam_Interface.frx":CCBA
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":CCD6
         Caption         =   "29"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   19
         Left            =   120
         TabIndex        =   108
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
         Picture         =   "mS_Exam_Interface.frx":CE38
         DisabledPicture =   "mS_Exam_Interface.frx":CE54
         DownPicture     =   "mS_Exam_Interface.frx":CE70
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":CE8C
         Caption         =   "19"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   20
         Left            =   840
         TabIndex        =   109
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
         Picture         =   "mS_Exam_Interface.frx":CFEE
         DisabledPicture =   "mS_Exam_Interface.frx":D00A
         DownPicture     =   "mS_Exam_Interface.frx":D026
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":D042
         Caption         =   "20"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   21
         Left            =   1560
         TabIndex        =   110
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
         Picture         =   "mS_Exam_Interface.frx":D1A4
         DisabledPicture =   "mS_Exam_Interface.frx":D1C0
         DownPicture     =   "mS_Exam_Interface.frx":D1DC
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":D1F8
         Caption         =   "21"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   22
         Left            =   2280
         TabIndex        =   111
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
         Picture         =   "mS_Exam_Interface.frx":D35A
         DisabledPicture =   "mS_Exam_Interface.frx":D376
         DownPicture     =   "mS_Exam_Interface.frx":D392
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":D3AE
         Caption         =   "22"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   24
         Left            =   3720
         TabIndex        =   112
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
         Picture         =   "mS_Exam_Interface.frx":D510
         DisabledPicture =   "mS_Exam_Interface.frx":D52C
         DownPicture     =   "mS_Exam_Interface.frx":D548
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":D564
         Caption         =   "24"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   23
         Left            =   3000
         TabIndex        =   113
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
         Picture         =   "mS_Exam_Interface.frx":D6C6
         DisabledPicture =   "mS_Exam_Interface.frx":D6E2
         DownPicture     =   "mS_Exam_Interface.frx":D6FE
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":D71A
         Caption         =   "23"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   13
         Left            =   120
         TabIndex        =   114
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
         Picture         =   "mS_Exam_Interface.frx":D87C
         DisabledPicture =   "mS_Exam_Interface.frx":D898
         DownPicture     =   "mS_Exam_Interface.frx":D8B4
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":D8D0
         Caption         =   "13"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   14
         Left            =   840
         TabIndex        =   115
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
         Picture         =   "mS_Exam_Interface.frx":DA32
         DisabledPicture =   "mS_Exam_Interface.frx":DA4E
         DownPicture     =   "mS_Exam_Interface.frx":DA6A
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":DA86
         Caption         =   "14"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   15
         Left            =   1560
         TabIndex        =   116
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
         Picture         =   "mS_Exam_Interface.frx":DBE8
         DisabledPicture =   "mS_Exam_Interface.frx":DC04
         DownPicture     =   "mS_Exam_Interface.frx":DC20
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":DC3C
         Caption         =   "15"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   16
         Left            =   2280
         TabIndex        =   117
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
         Picture         =   "mS_Exam_Interface.frx":DD9E
         DisabledPicture =   "mS_Exam_Interface.frx":DDBA
         DownPicture     =   "mS_Exam_Interface.frx":DDD6
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":DDF2
         Caption         =   "16"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   18
         Left            =   3720
         TabIndex        =   118
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
         Picture         =   "mS_Exam_Interface.frx":DF54
         DisabledPicture =   "mS_Exam_Interface.frx":DF70
         DownPicture     =   "mS_Exam_Interface.frx":DF8C
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":DFA8
         Caption         =   "18"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   17
         Left            =   3000
         TabIndex        =   119
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
         Picture         =   "mS_Exam_Interface.frx":E10A
         DisabledPicture =   "mS_Exam_Interface.frx":E126
         DownPicture     =   "mS_Exam_Interface.frx":E142
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":E15E
         Caption         =   "17"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   7
         Left            =   120
         TabIndex        =   120
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
         Picture         =   "mS_Exam_Interface.frx":E2C0
         DisabledPicture =   "mS_Exam_Interface.frx":E2DC
         DownPicture     =   "mS_Exam_Interface.frx":E2F8
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":E314
         Caption         =   "7"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   8
         Left            =   840
         TabIndex        =   121
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
         Picture         =   "mS_Exam_Interface.frx":E476
         DisabledPicture =   "mS_Exam_Interface.frx":E492
         DownPicture     =   "mS_Exam_Interface.frx":E4AE
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":E4CA
         Caption         =   "8"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   9
         Left            =   1560
         TabIndex        =   122
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
         Picture         =   "mS_Exam_Interface.frx":E62C
         DisabledPicture =   "mS_Exam_Interface.frx":E648
         DownPicture     =   "mS_Exam_Interface.frx":E664
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":E680
         Caption         =   "9"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   10
         Left            =   2280
         TabIndex        =   123
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
         Picture         =   "mS_Exam_Interface.frx":E7E2
         DisabledPicture =   "mS_Exam_Interface.frx":E7FE
         DownPicture     =   "mS_Exam_Interface.frx":E81A
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":E836
         Caption         =   "10"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   12
         Left            =   3720
         TabIndex        =   124
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
         Picture         =   "mS_Exam_Interface.frx":E998
         DisabledPicture =   "mS_Exam_Interface.frx":E9B4
         DownPicture     =   "mS_Exam_Interface.frx":E9D0
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":E9EC
         Caption         =   "12"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   11
         Left            =   3000
         TabIndex        =   125
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
         Picture         =   "mS_Exam_Interface.frx":EB4E
         DisabledPicture =   "mS_Exam_Interface.frx":EB6A
         DownPicture     =   "mS_Exam_Interface.frx":EB86
         MousePointer    =   99
         MouseIcon       =   "mS_Exam_Interface.frx":EBA2
         Caption         =   "11"
      End
      Begin VB.Label Label39 
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
         TabIndex        =   65
         Top             =   1800
         Width           =   900
      End
      Begin VB.Label Label38 
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
         TabIndex        =   64
         Top             =   1800
         Width           =   1650
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "(PUR003)"
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
         Height          =   260
         Left            =   3200
         TabIndex        =   63
         Top             =   150
         Width           =   1035
      End
      Begin VB.Image StuPic 
         Height          =   600
         Left            =   15
         Picture         =   "mS_Exam_Interface.frx":ED04
         Stretch         =   -1  'True
         Top             =   0
         Width           =   615
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   2
         X1              =   0
         X2              =   4450
         Y1              =   635
         Y2              =   635
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
         TabIndex        =   62
         Top             =   135
         Width           =   2250
      End
      Begin VB.Shape Shape16 
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
      Begin VB.Shape Shape15 
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
      Begin VB.Shape Shape14 
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
      Begin VB.Shape Shape13 
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
      Begin VB.Label Label35 
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
         TabIndex        =   61
         Top             =   1275
         Width           =   1335
      End
      Begin VB.Label Label34 
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
         TabIndex        =   60
         Top             =   1275
         Width           =   1890
      End
      Begin VB.Label Label33 
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
         TabIndex        =   59
         Top             =   765
         Width           =   1005
      End
      Begin VB.Label Label32 
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
         TabIndex        =   58
         Top             =   765
         Width           =   690
      End
      Begin VB.Label Label31 
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
         TabIndex        =   57
         Top             =   765
         Width           =   945
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000C000&
         FillColor       =   &H00004000&
         Height          =   345
         Index           =   1
         Left            =   60
         Shape           =   3  'Circle
         Top             =   720
         Width           =   375
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
         Index           =   60
         Left            =   3800
         Shape           =   3  'Circle
         Top             =   8400
         Width           =   495
      End
      Begin VB.Shape Shape17 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFC0C0&
         Height          =   495
         Left            =   -15
         Top             =   1650
         Width           =   4575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   8445
      Left            =   0
      TabIndex        =   4
      Top             =   2280
      Width           =   16050
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000013&
         Height          =   855
         Left            =   0
         ScaleHeight     =   795
         ScaleWidth      =   15915
         TabIndex        =   127
         Top             =   7520
         Width           =   15975
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
            Left            =   13800
            MouseIcon       =   "mS_Exam_Interface.frx":F081
            MousePointer    =   99  'Custom
            Picture         =   "mS_Exam_Interface.frx":F1D3
            Style           =   1  'Graphical
            TabIndex        =   126
            ToolTipText     =   "Save and Move to Next Question"
            Top             =   120
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
            Left            =   10320
            MouseIcon       =   "mS_Exam_Interface.frx":FC34
            MousePointer    =   99  'Custom
            Picture         =   "mS_Exam_Interface.frx":FD86
            Style           =   1  'Graphical
            TabIndex        =   131
            ToolTipText     =   "Move to Next Question"
            Top             =   120
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
            MouseIcon       =   "mS_Exam_Interface.frx":10621
            MousePointer    =   99  'Custom
            Picture         =   "mS_Exam_Interface.frx":10773
            Style           =   1  'Graphical
            TabIndex        =   130
            ToolTipText     =   "Move to Previous Question"
            Top             =   120
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
            MouseIcon       =   "mS_Exam_Interface.frx":11011
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   129
            ToolTipText     =   "Clear Answer"
            Top             =   120
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
            MouseIcon       =   "mS_Exam_Interface.frx":11163
            MousePointer    =   99  'Custom
            TabIndex        =   128
            ToolTipText     =   "Mark this Question"
            Top             =   200
            Width           =   2655
         End
         Begin MyEllipticButton.EllipticButton btn 
            Height          =   405
            Index           =   0
            Left            =   11760
            TabIndex        =   132
            Top             =   120
            Visible         =   0   'False
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   714
            BackColor       =   65535
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "mS_Exam_Interface.frx":112B5
            DisabledPicture =   "mS_Exam_Interface.frx":112D1
            DownPicture     =   "mS_Exam_Interface.frx":112ED
            MousePointer    =   99
            MouseIcon       =   "mS_Exam_Interface.frx":11309
            Caption         =   ""
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   6
            Height          =   615
            Index           =   0
            Left            =   11640
            Shape           =   3  'Circle
            Top             =   0
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape Shape7 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000013&
            Height          =   465
            Left            =   150
            Top             =   120
            Width           =   2790
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   4455
         Left            =   6840
         TabIndex        =   29
         Top             =   960
         Width           =   5535
         Begin VB.Image Image1 
            Height          =   2685
            Left            =   0
            Picture         =   "mS_Exam_Interface.frx":11BE3
            Stretch         =   -1  'True
            Top             =   120
            Width           =   3210
         End
      End
      Begin VB.Frame MainFrame 
         Height          =   2535
         Left            =   7560
         TabIndex        =   15
         Top             =   1920
         Width           =   4455
         Begin VB.Frame Frame2 
            Caption         =   "Timing"
            Height          =   975
            Left            =   1200
            TabIndex        =   28
            Top             =   120
            Visible         =   0   'False
            Width           =   1095
            Begin VB.Timer Timer2 
               Enabled         =   0   'False
               Interval        =   1005
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
            TabIndex        =   16
            Top             =   120
            Visible         =   0   'False
            Width           =   975
            Begin VB.Timer btnNXT_PREV 
               Enabled         =   0   'False
               Interval        =   20
               Left            =   240
               Top             =   360
            End
         End
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   375
            Left            =   120
            Top             =   1680
            Visible         =   0   'False
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
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
            Connect         =   "Provider=MSDAORA.1;User ID=sts/sts;Persist Security Info=True"
            OLEDBString     =   "Provider=MSDAORA.1;User ID=sts/sts;Persist Security Info=True"
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
         Begin VB.Label tmp 
            BackColor       =   &H008080FF&
            Height          =   375
            Left            =   120
            TabIndex        =   27
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label Optkey 
            BackColor       =   &H00FF80FF&
            Height          =   375
            Left            =   2280
            TabIndex        =   26
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFC0&
            Caption         =   "OptKey"
            Height          =   255
            Left            =   2280
            TabIndex        =   25
            Top             =   120
            Width           =   2055
         End
         Begin VB.Label indx 
            BackColor       =   &H00FF8080&
            Height          =   255
            Left            =   3600
            TabIndex        =   24
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            BackColor       =   &H0080C0FF&
            Caption         =   "Correct_Ans"
            Height          =   255
            Left            =   2280
            TabIndex        =   23
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            BackColor       =   &H0080C0FF&
            Caption         =   "User Ans"
            Height          =   255
            Left            =   2280
            TabIndex        =   22
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label lbl_corr_ans 
            BackColor       =   &H0000FF00&
            Height          =   255
            Left            =   3600
            TabIndex        =   21
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label lbl_usr_ans 
            BackColor       =   &H0000FF00&
            Height          =   255
            Left            =   3600
            TabIndex        =   20
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label25 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Caption         =   "Indx"
            Height          =   255
            Left            =   2280
            TabIndex        =   19
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label bookmark 
            BackColor       =   &H00FF8080&
            DataField       =   "BOOKMRK"
            DataSource      =   "Adodc2"
            Height          =   255
            Left            =   3600
            TabIndex        =   18
            Top             =   2160
            Width           =   735
         End
         Begin VB.Label Label26 
            Alignment       =   2  'Center
            BackColor       =   &H0080C0FF&
            Caption         =   "Bookmark"
            Height          =   255
            Left            =   2280
            TabIndex        =   17
            Top             =   2160
            Width           =   1095
         End
      End
      Begin VB.OptionButton options 
         BackColor       =   &H00FFFFFF&
         Caption         =   "30% jgjbjgygybtbftyftf ty t t yv t tyttv yv yvyvvvvyvtvt vtvtvuvtytv tvyvyv vtvt vtvytvtvtvtuvtuvutvytvtvyvyvtyvtvtvt yvy vyv"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Index           =   3
         Left            =   960
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   3000
         Width           =   5775
      End
      Begin VB.OptionButton options 
         BackColor       =   &H00FFFFFF&
         Caption         =   "30% jgjbjgygybtbftyftf ty t t yv t tyttv yv yvyvvvvyvtvt vtvtvuvtytv tvyvyv vtvt vtvytvtvtvtuvtuvutvytvtvyvyvtyvtvtvt yvy vyv"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Index           =   2
         Left            =   960
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   2400
         Width           =   5775
      End
      Begin VB.OptionButton options 
         BackColor       =   &H00FFFFFF&
         Caption         =   "30% jgjbjgygybtbftyftf ty t t yv t tyttv yv yvyvvvvyvtvt vtvtvuvtytv tvyvyv vtvt vtvytvtvtvtuvtuvutvytvtvyvyvtyvtvtvt yvy vyv"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Index           =   1
         Left            =   960
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   1800
         Width           =   5775
      End
      Begin VB.OptionButton options 
         BackColor       =   &H00FFFFFF&
         Caption         =   "30% jgjbjgygybtbftyftf ty t t yv t tyttv yv yvyvvvvyvtvt vtvtvuvtytv tvyvyv vtvt vtvytvtvtvtuvtuvutvytvtvyvyvtyvtvtvt yvy vyv"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Index           =   0
         Left            =   960
         TabIndex        =   10
         Top             =   1200
         Width           =   5895
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
         TabIndex        =   9
         Top             =   3165
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
         TabIndex        =   8
         Top             =   2565
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
         TabIndex        =   7
         Top             =   1965
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
         TabIndex        =   6
         Top             =   1365
         Width           =   375
      End
      Begin VB.Label qtext 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   $"mS_Exam_Interface.frx":15F15
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
         TabIndex        =   5
         Top             =   360
         Width           =   14655
      End
   End
   Begin VB.PictureBox vkFrame11 
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
      Left            =   -30
      ScaleHeight     =   825
      ScaleWidth      =   16065
      TabIndex        =   139
      Top             =   -120
      Width           =   16125
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
         Left            =   14400
         MouseIcon       =   "mS_Exam_Interface.frx":15FB8
         MousePointer    =   99  'Custom
         Picture         =   "mS_Exam_Interface.frx":1610A
         Style           =   1  'Graphical
         TabIndex        =   140
         ToolTipText     =   "Terminate Exam Abnormally"
         Top             =   240
         Width           =   1365
      End
      Begin VB.Shape Shape8 
         Height          =   735
         Index           =   0
         Left            =   0
         Top             =   90
         Width           =   15975
      End
      Begin VB.Label Label1 
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
         Left            =   200
         TabIndex        =   147
         Top             =   240
         Width           =   1515
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
         TabIndex        =   146
         Top             =   240
         Width           =   5115
      End
      Begin VB.Label Timer_remain 
         BackColor       =   &H00D1E0D0&
         Caption         =   "04 : 36"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   11085
         TabIndex        =   142
         Top             =   330
         Width           =   1085
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
         TabIndex        =   138
         Top             =   360
         Visible         =   0   'False
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
         TabIndex        =   145
         Top             =   360
         Visible         =   0   'False
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
         TabIndex        =   144
         Top             =   360
         Visible         =   0   'False
         Width           =   300
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
         TabIndex        =   143
         Top             =   360
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time Left -"
         BeginProperty Font 
            Name            =   "Bangkok"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9840
         TabIndex        =   141
         Top             =   360
         Width           =   1155
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00D1E0D0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   495
         Left            =   9530
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame4Five 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   535
      Left            =   1440
      TabIndex        =   40
      Top             =   840
      Width           =   14505
      Begin ChamaleonButton.ChameleonBtn sub51 
         Height          =   465
         Left            =   0
         TabIndex        =   42
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
         MICON           =   "mS_Exam_Interface.frx":16989
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
         TabIndex        =   43
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
         MICON           =   "mS_Exam_Interface.frx":169A5
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
         TabIndex        =   44
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
         MICON           =   "mS_Exam_Interface.frx":169C1
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
         TabIndex        =   45
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
         MICON           =   "mS_Exam_Interface.frx":169DD
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
         TabIndex        =   41
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
         MICON           =   "mS_Exam_Interface.frx":169F9
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
      TabIndex        =   39
      Top             =   840
      Width           =   14505
      Begin ChamaleonButton.ChameleonBtn sub44 
         Height          =   465
         Left            =   10800
         TabIndex        =   49
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
         MICON           =   "mS_Exam_Interface.frx":16A15
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
         TabIndex        =   48
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
         MICON           =   "mS_Exam_Interface.frx":16A31
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
         TabIndex        =   47
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
         MICON           =   "mS_Exam_Interface.frx":16A4D
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
         TabIndex        =   46
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
         MICON           =   "mS_Exam_Interface.frx":16A69
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
      TabIndex        =   38
      Top             =   840
      Width           =   14505
      Begin ChamaleonButton.ChameleonBtn sub33 
         Height          =   465
         Left            =   9600
         TabIndex        =   52
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
         MICON           =   "mS_Exam_Interface.frx":16A85
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
         TabIndex        =   51
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
         MICON           =   "mS_Exam_Interface.frx":16AA1
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
         TabIndex        =   50
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
         MICON           =   "mS_Exam_Interface.frx":16ABD
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
      TabIndex        =   37
      Top             =   840
      Width           =   14505
      Begin ChamaleonButton.ChameleonBtn sub21 
         Height          =   465
         Left            =   0
         TabIndex        =   53
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
         MICON           =   "mS_Exam_Interface.frx":16AD9
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
         TabIndex        =   54
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
         MICON           =   "mS_Exam_Interface.frx":16AF5
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
   Begin VB.Frame Frame4One 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   535
      Left            =   1440
      TabIndex        =   36
      Top             =   840
      Width           =   14505
      Begin ChamaleonButton.ChameleonBtn Sub11 
         Height          =   465
         Left            =   0
         TabIndex        =   55
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
         MICON           =   "mS_Exam_Interface.frx":16B11
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
   Begin VB.Shape Shape8 
      Height          =   650
      Index           =   1
      Left            =   0
      Top             =   780
      Width           =   15975
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   525
      Left            =   9480
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   3495
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
      MouseIcon       =   "mS_Exam_Interface.frx":16B2D
      MousePointer    =   99  'Custom
      TabIndex        =   35
      Top             =   1965
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
      MouseIcon       =   "mS_Exam_Interface.frx":16C7F
      MousePointer    =   99  'Custom
      TabIndex        =   34
      Top             =   1950
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
      Left            =   14650
      MouseIcon       =   "mS_Exam_Interface.frx":16DD1
      MousePointer    =   99  'Custom
      TabIndex        =   33
      Top             =   1965
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
      TabIndex        =   32
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
      TabIndex        =   31
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
      TabIndex        =   30
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
      MouseIcon       =   "mS_Exam_Interface.frx":16F23
      MousePointer    =   99  'Custom
      Picture         =   "mS_Exam_Interface.frx":177ED
      ToolTipText     =   "Use Rough page here"
      Top             =   1455
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   15285
      MouseIcon       =   "mS_Exam_Interface.frx":180B7
      MousePointer    =   99  'Custom
      Picture         =   "mS_Exam_Interface.frx":18209
      Stretch         =   -1  'True
      ToolTipText     =   "Click To Listen Question"
      Top             =   1485
      Width           =   525
   End
   Begin VB.Image Command1 
      Height          =   450
      Left            =   14640
      MouseIcon       =   "mS_Exam_Interface.frx":19A32
      MousePointer    =   99  'Custom
      Picture         =   "mS_Exam_Interface.frx":19B84
      Stretch         =   -1  'True
      ToolTipText     =   "Show Current Status"
      Top             =   1500
      Width           =   480
   End
   Begin VB.Label t_time_min 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Elapsed Time - "
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
      TabIndex        =   3
      Top             =   1700
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
      TabIndex        =   2
      Top             =   -240
      Width           =   1470
   End
   Begin VB.Label q_id 
      AutoSize        =   -1  'True
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
      TabIndex        =   1
      Top             =   1695
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
      TabIndex        =   0
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
      TabIndex        =   14
      Top             =   1695
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
ChkOption    'Matching option
ChkUserANS   'matching answer
CHKbookmark  'checking bookmark
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
On Error Resume Next
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
   TestTypeInfo.Caption = " Subject Wise Test :          [Subject]:- " & CurrentSub
  Else
   TestTypeInfo.Caption = " Topic Wise Test :            [Sub]:- " & CurrentSub & "        [Topic] :- " & CurrentTopic
  End If
End If

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
  StuPic.Picture = LoadPicture(img_pth)
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
 options(i).Value = False
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
On Error Resume Next
mcqTestRunPPr.Orientation = rptOrientPortrait
mcqTestRunPPr.Show vbModal, MDI
End Sub

Private Sub vkCommand3_Click() 'Show Instructions
Instruction_test.Show vbModal, MDI
End Sub

