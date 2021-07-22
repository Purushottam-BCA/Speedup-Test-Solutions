VERSION 5.00
Object = "{08654D78-6636-11D3-87BF-B4980CC10374}#2.0#0"; "MyEllipticButton.ocx"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14655
   LinkTopic       =   "Form2"
   ScaleHeight     =   10950
   ScaleWidth      =   14655
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox vkFrame1 
      Height          =   885
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   16035
      TabIndex        =   82
      Top             =   4320
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
         MouseIcon       =   "#DEMO.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "#DEMO.frx":0152
         Style           =   1  'Graphical
         TabIndex        =   83
         ToolTipText     =   "Terminate Exam "
         Top             =   260
         Width           =   1485
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
         TabIndex        =   92
         Top             =   1440
         Width           =   855
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
         TabIndex        =   91
         Top             =   210
         Width           =   5115
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
         TabIndex        =   90
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
         TabIndex        =   89
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
         TabIndex        =   88
         Top             =   360
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
         TabIndex        =   87
         Top             =   360
         Width           =   75
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
         TabIndex        =   86
         Top             =   330
         Width           =   1215
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
         TabIndex        =   85
         Top             =   360
         Width           =   1155
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
         TabIndex        =   84
         Top             =   210
         Width           =   1515
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   855
      Left            =   0
      Picture         =   "#DEMO.frx":09D1
      ScaleHeight     =   795
      ScaleWidth      =   15915
      TabIndex        =   74
      Top             =   1080
      Width           =   15975
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
         MouseIcon       =   "#DEMO.frx":2D37
         MousePointer    =   99  'Custom
         TabIndex        =   80
         ToolTipText     =   "Mark this Question"
         Top             =   215
         Width           =   2655
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
         MouseIcon       =   "#DEMO.frx":2E89
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "Clear Answer"
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
         MouseIcon       =   "#DEMO.frx":2FDB
         MousePointer    =   99  'Custom
         Picture         =   "#DEMO.frx":312D
         Style           =   1  'Graphical
         TabIndex        =   78
         ToolTipText     =   "Move to Previous Question"
         Top             =   150
         Width           =   1845
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
         MouseIcon       =   "#DEMO.frx":39CB
         MousePointer    =   99  'Custom
         Picture         =   "#DEMO.frx":3B1D
         Style           =   1  'Graphical
         TabIndex        =   77
         ToolTipText     =   "Move to Next Question"
         Top             =   150
         Width           =   1845
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
         MouseIcon       =   "#DEMO.frx":43B8
         MousePointer    =   99  'Custom
         Picture         =   "#DEMO.frx":450A
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   "Save and Move to Next Question"
         Top             =   150
         Width           =   1965
      End
      Begin VB.TextBox Text1 
         Height          =   330
         Left            =   6120
         TabIndex        =   75
         Text            =   "Text1"
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   0
         Left            =   6840
         TabIndex        =   81
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
         Picture         =   "#DEMO.frx":4F6B
         DisabledPicture =   "#DEMO.frx":4F87
         DownPicture     =   "#DEMO.frx":4FA3
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":4FBF
         Caption         =   "1"
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   465
         Left            =   150
         Top             =   150
         Width           =   2790
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
   End
   Begin VB.PictureBox Picture1 
      Height          =   1350
      Left            =   5280
      Picture         =   "#DEMO.frx":5121
      ScaleHeight     =   1290
      ScaleWidth      =   4395
      TabIndex        =   70
      Top             =   9240
      Width           =   4455
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
         MouseIcon       =   "#DEMO.frx":69D3
         MousePointer    =   99  'Custom
         Picture         =   "#DEMO.frx":6B25
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "See All Questions at a glance"
         Top             =   120
         Width           =   2000
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
         MouseIcon       =   "#DEMO.frx":7684
         MousePointer    =   99  'Custom
         Picture         =   "#DEMO.frx":77D6
         Style           =   1  'Graphical
         TabIndex        =   72
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
         Height          =   495
         Left            =   200
         MouseIcon       =   "#DEMO.frx":8105
         MousePointer    =   99  'Custom
         Picture         =   "#DEMO.frx":8257
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Submit The Test"
         Top             =   700
         Width           =   4020
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   9165
      Left            =   5280
      Picture         =   "#DEMO.frx":8E7B
      ScaleHeight     =   9105
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   1
         Left            =   120
         TabIndex        =   1
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
         Picture         =   "#DEMO.frx":A72D
         DisabledPicture =   "#DEMO.frx":A749
         DownPicture     =   "#DEMO.frx":A765
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":A781
         Caption         =   "1"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   2
         Left            =   840
         TabIndex        =   2
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
         Picture         =   "#DEMO.frx":A8E3
         DisabledPicture =   "#DEMO.frx":A8FF
         DownPicture     =   "#DEMO.frx":A91B
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":A937
         Caption         =   "2"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   3
         Left            =   1560
         TabIndex        =   3
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
         Picture         =   "#DEMO.frx":AA99
         DisabledPicture =   "#DEMO.frx":AAB5
         DownPicture     =   "#DEMO.frx":AAD1
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":AAED
         Caption         =   "3"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   4
         Left            =   2280
         TabIndex        =   4
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
         Picture         =   "#DEMO.frx":AC4F
         DisabledPicture =   "#DEMO.frx":AC6B
         DownPicture     =   "#DEMO.frx":AC87
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":ACA3
         Caption         =   "4"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   5
         Left            =   3000
         TabIndex        =   5
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
         Picture         =   "#DEMO.frx":AE05
         DisabledPicture =   "#DEMO.frx":AE21
         DownPicture     =   "#DEMO.frx":AE3D
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":AE59
         Caption         =   "5"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   6
         Left            =   3720
         TabIndex        =   6
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
         Picture         =   "#DEMO.frx":AFBB
         DisabledPicture =   "#DEMO.frx":AFD7
         DownPicture     =   "#DEMO.frx":AFF3
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":B00F
         Caption         =   "6"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   7
         Left            =   120
         TabIndex        =   7
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
         Picture         =   "#DEMO.frx":B171
         DisabledPicture =   "#DEMO.frx":B18D
         DownPicture     =   "#DEMO.frx":B1A9
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":B1C5
         Caption         =   "7"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   8
         Left            =   840
         TabIndex        =   8
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
         Picture         =   "#DEMO.frx":B327
         DisabledPicture =   "#DEMO.frx":B343
         DownPicture     =   "#DEMO.frx":B35F
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":B37B
         Caption         =   "8"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   9
         Left            =   1560
         TabIndex        =   9
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
         Picture         =   "#DEMO.frx":B4DD
         DisabledPicture =   "#DEMO.frx":B4F9
         DownPicture     =   "#DEMO.frx":B515
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":B531
         Caption         =   "9"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   10
         Left            =   2280
         TabIndex        =   10
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
         Picture         =   "#DEMO.frx":B693
         DisabledPicture =   "#DEMO.frx":B6AF
         DownPicture     =   "#DEMO.frx":B6CB
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":B6E7
         Caption         =   "10"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   11
         Left            =   3000
         TabIndex        =   11
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
         Picture         =   "#DEMO.frx":B849
         DisabledPicture =   "#DEMO.frx":B865
         DownPicture     =   "#DEMO.frx":B881
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":B89D
         Caption         =   "11"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   12
         Left            =   3720
         TabIndex        =   12
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
         Picture         =   "#DEMO.frx":B9FF
         DisabledPicture =   "#DEMO.frx":BA1B
         DownPicture     =   "#DEMO.frx":BA37
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":BA53
         Caption         =   "12"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   13
         Left            =   120
         TabIndex        =   13
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
         Picture         =   "#DEMO.frx":BBB5
         DisabledPicture =   "#DEMO.frx":BBD1
         DownPicture     =   "#DEMO.frx":BBED
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":BC09
         Caption         =   "13"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   14
         Left            =   840
         TabIndex        =   14
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
         Picture         =   "#DEMO.frx":BD6B
         DisabledPicture =   "#DEMO.frx":BD87
         DownPicture     =   "#DEMO.frx":BDA3
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":BDBF
         Caption         =   "14"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   15
         Left            =   1560
         TabIndex        =   15
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
         Picture         =   "#DEMO.frx":BF21
         DisabledPicture =   "#DEMO.frx":BF3D
         DownPicture     =   "#DEMO.frx":BF59
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":BF75
         Caption         =   "15"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   16
         Left            =   2280
         TabIndex        =   16
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
         Picture         =   "#DEMO.frx":C0D7
         DisabledPicture =   "#DEMO.frx":C0F3
         DownPicture     =   "#DEMO.frx":C10F
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":C12B
         Caption         =   "16"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   17
         Left            =   3000
         TabIndex        =   17
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
         Picture         =   "#DEMO.frx":C28D
         DisabledPicture =   "#DEMO.frx":C2A9
         DownPicture     =   "#DEMO.frx":C2C5
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":C2E1
         Caption         =   "17"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   18
         Left            =   3720
         TabIndex        =   18
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
         Picture         =   "#DEMO.frx":C443
         DisabledPicture =   "#DEMO.frx":C45F
         DownPicture     =   "#DEMO.frx":C47B
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":C497
         Caption         =   "18"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   19
         Left            =   120
         TabIndex        =   19
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
         Picture         =   "#DEMO.frx":C5F9
         DisabledPicture =   "#DEMO.frx":C615
         DownPicture     =   "#DEMO.frx":C631
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":C64D
         Caption         =   "19"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   20
         Left            =   840
         TabIndex        =   20
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
         Picture         =   "#DEMO.frx":C7AF
         DisabledPicture =   "#DEMO.frx":C7CB
         DownPicture     =   "#DEMO.frx":C7E7
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":C803
         Caption         =   "20"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   21
         Left            =   1560
         TabIndex        =   21
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
         Picture         =   "#DEMO.frx":C965
         DisabledPicture =   "#DEMO.frx":C981
         DownPicture     =   "#DEMO.frx":C99D
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":C9B9
         Caption         =   "21"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   22
         Left            =   2280
         TabIndex        =   22
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
         Picture         =   "#DEMO.frx":CB1B
         DisabledPicture =   "#DEMO.frx":CB37
         DownPicture     =   "#DEMO.frx":CB53
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":CB6F
         Caption         =   "22"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   23
         Left            =   3000
         TabIndex        =   23
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
         Picture         =   "#DEMO.frx":CCD1
         DisabledPicture =   "#DEMO.frx":CCED
         DownPicture     =   "#DEMO.frx":CD09
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":CD25
         Caption         =   "23"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   24
         Left            =   3720
         TabIndex        =   24
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
         Picture         =   "#DEMO.frx":CE87
         DisabledPicture =   "#DEMO.frx":CEA3
         DownPicture     =   "#DEMO.frx":CEBF
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":CEDB
         Caption         =   "24"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   25
         Left            =   120
         TabIndex        =   25
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
         Picture         =   "#DEMO.frx":D03D
         DisabledPicture =   "#DEMO.frx":D059
         DownPicture     =   "#DEMO.frx":D075
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":D091
         Caption         =   "25"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   26
         Left            =   840
         TabIndex        =   26
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
         Picture         =   "#DEMO.frx":D1F3
         DisabledPicture =   "#DEMO.frx":D20F
         DownPicture     =   "#DEMO.frx":D22B
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":D247
         Caption         =   "26"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   27
         Left            =   1560
         TabIndex        =   27
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
         Picture         =   "#DEMO.frx":D3A9
         DisabledPicture =   "#DEMO.frx":D3C5
         DownPicture     =   "#DEMO.frx":D3E1
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":D3FD
         Caption         =   "27"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   28
         Left            =   2280
         TabIndex        =   28
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
         Picture         =   "#DEMO.frx":D55F
         DisabledPicture =   "#DEMO.frx":D57B
         DownPicture     =   "#DEMO.frx":D597
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":D5B3
         Caption         =   "28"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   29
         Left            =   3000
         TabIndex        =   29
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
         Picture         =   "#DEMO.frx":D715
         DisabledPicture =   "#DEMO.frx":D731
         DownPicture     =   "#DEMO.frx":D74D
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":D769
         Caption         =   "29"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   30
         Left            =   3720
         TabIndex        =   30
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
         Picture         =   "#DEMO.frx":D8CB
         DisabledPicture =   "#DEMO.frx":D8E7
         DownPicture     =   "#DEMO.frx":D903
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":D91F
         Caption         =   "30"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   31
         Left            =   120
         TabIndex        =   31
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
         Picture         =   "#DEMO.frx":DA81
         DisabledPicture =   "#DEMO.frx":DA9D
         DownPicture     =   "#DEMO.frx":DAB9
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":DAD5
         Caption         =   "31"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   32
         Left            =   840
         TabIndex        =   32
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
         Picture         =   "#DEMO.frx":DC37
         DisabledPicture =   "#DEMO.frx":DC53
         DownPicture     =   "#DEMO.frx":DC6F
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":DC8B
         Caption         =   "32"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   33
         Left            =   1560
         TabIndex        =   33
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
         Picture         =   "#DEMO.frx":DDED
         DisabledPicture =   "#DEMO.frx":DE09
         DownPicture     =   "#DEMO.frx":DE25
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":DE41
         Caption         =   "33"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   34
         Left            =   2280
         TabIndex        =   34
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
         Picture         =   "#DEMO.frx":DFA3
         DisabledPicture =   "#DEMO.frx":DFBF
         DownPicture     =   "#DEMO.frx":DFDB
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":DFF7
         Caption         =   "34"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   35
         Left            =   3000
         TabIndex        =   35
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
         Picture         =   "#DEMO.frx":E159
         DisabledPicture =   "#DEMO.frx":E175
         DownPicture     =   "#DEMO.frx":E191
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":E1AD
         Caption         =   "35"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   36
         Left            =   3720
         TabIndex        =   36
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
         Picture         =   "#DEMO.frx":E30F
         DisabledPicture =   "#DEMO.frx":E32B
         DownPicture     =   "#DEMO.frx":E347
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":E363
         Caption         =   "36"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   37
         Left            =   120
         TabIndex        =   37
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
         Picture         =   "#DEMO.frx":E4C5
         DisabledPicture =   "#DEMO.frx":E4E1
         DownPicture     =   "#DEMO.frx":E4FD
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":E519
         Caption         =   "37"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   38
         Left            =   840
         TabIndex        =   38
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
         Picture         =   "#DEMO.frx":E67B
         DisabledPicture =   "#DEMO.frx":E697
         DownPicture     =   "#DEMO.frx":E6B3
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":E6CF
         Caption         =   "38"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   39
         Left            =   1560
         TabIndex        =   39
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
         Picture         =   "#DEMO.frx":E831
         DisabledPicture =   "#DEMO.frx":E84D
         DownPicture     =   "#DEMO.frx":E869
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":E885
         Caption         =   "39"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   40
         Left            =   2280
         TabIndex        =   40
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
         Picture         =   "#DEMO.frx":E9E7
         DisabledPicture =   "#DEMO.frx":EA03
         DownPicture     =   "#DEMO.frx":EA1F
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":EA3B
         Caption         =   "40"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   41
         Left            =   3000
         TabIndex        =   41
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
         Picture         =   "#DEMO.frx":EB9D
         DisabledPicture =   "#DEMO.frx":EBB9
         DownPicture     =   "#DEMO.frx":EBD5
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":EBF1
         Caption         =   "41"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   42
         Left            =   3720
         TabIndex        =   42
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
         Picture         =   "#DEMO.frx":ED53
         DisabledPicture =   "#DEMO.frx":ED6F
         DownPicture     =   "#DEMO.frx":ED8B
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":EDA7
         Caption         =   "42"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   43
         Left            =   120
         TabIndex        =   43
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
         Picture         =   "#DEMO.frx":EF09
         DisabledPicture =   "#DEMO.frx":EF25
         DownPicture     =   "#DEMO.frx":EF41
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":EF5D
         Caption         =   "43"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   44
         Left            =   840
         TabIndex        =   44
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
         Picture         =   "#DEMO.frx":F0BF
         DisabledPicture =   "#DEMO.frx":F0DB
         DownPicture     =   "#DEMO.frx":F0F7
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":F113
         Caption         =   "44"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   45
         Left            =   1560
         TabIndex        =   45
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
         Picture         =   "#DEMO.frx":F275
         DisabledPicture =   "#DEMO.frx":F291
         DownPicture     =   "#DEMO.frx":F2AD
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":F2C9
         Caption         =   "45"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   46
         Left            =   2280
         TabIndex        =   46
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
         Picture         =   "#DEMO.frx":F42B
         DisabledPicture =   "#DEMO.frx":F447
         DownPicture     =   "#DEMO.frx":F463
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":F47F
         Caption         =   "46"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   47
         Left            =   3000
         TabIndex        =   47
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
         Picture         =   "#DEMO.frx":F5E1
         DisabledPicture =   "#DEMO.frx":F5FD
         DownPicture     =   "#DEMO.frx":F619
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":F635
         Caption         =   "47"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   48
         Left            =   3720
         TabIndex        =   48
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
         Picture         =   "#DEMO.frx":F797
         DisabledPicture =   "#DEMO.frx":F7B3
         DownPicture     =   "#DEMO.frx":F7CF
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":F7EB
         Caption         =   "48"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   49
         Left            =   120
         TabIndex        =   49
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
         Picture         =   "#DEMO.frx":F94D
         DisabledPicture =   "#DEMO.frx":F969
         DownPicture     =   "#DEMO.frx":F985
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":F9A1
         Caption         =   "49"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   50
         Left            =   840
         TabIndex        =   50
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
         Picture         =   "#DEMO.frx":FB03
         DisabledPicture =   "#DEMO.frx":FB1F
         DownPicture     =   "#DEMO.frx":FB3B
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":FB57
         Caption         =   "50"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   51
         Left            =   1560
         TabIndex        =   51
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
         Picture         =   "#DEMO.frx":FCB9
         DisabledPicture =   "#DEMO.frx":FCD5
         DownPicture     =   "#DEMO.frx":FCF1
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":FD0D
         Caption         =   "51"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   52
         Left            =   2280
         TabIndex        =   52
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
         Picture         =   "#DEMO.frx":FE6F
         DisabledPicture =   "#DEMO.frx":FE8B
         DownPicture     =   "#DEMO.frx":FEA7
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":FEC3
         Caption         =   "52"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   53
         Left            =   3000
         TabIndex        =   53
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
         Picture         =   "#DEMO.frx":10025
         DisabledPicture =   "#DEMO.frx":10041
         DownPicture     =   "#DEMO.frx":1005D
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":10079
         Caption         =   "53"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   54
         Left            =   3720
         TabIndex        =   54
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
         Picture         =   "#DEMO.frx":101DB
         DisabledPicture =   "#DEMO.frx":101F7
         DownPicture     =   "#DEMO.frx":10213
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":1022F
         Caption         =   "54"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   55
         Left            =   120
         TabIndex        =   55
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
         Picture         =   "#DEMO.frx":10391
         DisabledPicture =   "#DEMO.frx":103AD
         DownPicture     =   "#DEMO.frx":103C9
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":103E5
         Caption         =   "55"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   56
         Left            =   840
         TabIndex        =   56
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
         Picture         =   "#DEMO.frx":10547
         DisabledPicture =   "#DEMO.frx":10563
         DownPicture     =   "#DEMO.frx":1057F
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":1059B
         Caption         =   "56"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   57
         Left            =   1560
         TabIndex        =   57
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
         Picture         =   "#DEMO.frx":106FD
         DisabledPicture =   "#DEMO.frx":10719
         DownPicture     =   "#DEMO.frx":10735
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":10751
         Caption         =   "57"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   58
         Left            =   2280
         TabIndex        =   58
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
         Picture         =   "#DEMO.frx":108B3
         DisabledPicture =   "#DEMO.frx":108CF
         DownPicture     =   "#DEMO.frx":108EB
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":10907
         Caption         =   "58"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   59
         Left            =   3000
         TabIndex        =   59
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
         Picture         =   "#DEMO.frx":10A69
         DisabledPicture =   "#DEMO.frx":10A85
         DownPicture     =   "#DEMO.frx":10AA1
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":10ABD
         Caption         =   "59"
      End
      Begin MyEllipticButton.EllipticButton btn 
         Height          =   645
         Index           =   60
         Left            =   3720
         TabIndex        =   60
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
         Picture         =   "#DEMO.frx":10C1F
         DisabledPicture =   "#DEMO.frx":10C3B
         DownPicture     =   "#DEMO.frx":10C57
         MousePointer    =   99
         MouseIcon       =   "#DEMO.frx":10C73
         Caption         =   "60"
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
         TabIndex        =   69
         Top             =   1800
         Width           =   900
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
         TabIndex        =   68
         Top             =   1800
         Width           =   1650
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
         TabIndex        =   67
         Top             =   140
         Width           =   1260
      End
      Begin VB.Image stuPIC 
         Height          =   600
         Left            =   15
         Picture         =   "#DEMO.frx":10DD5
         Stretch         =   -1  'True
         Top             =   0
         Width           =   615
      End
      Begin VB.Line Line3 
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
         TabIndex        =   66
         Top             =   135
         Width           =   2250
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
         TabIndex        =   65
         Top             =   1275
         Width           =   1335
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
         TabIndex        =   64
         Top             =   1275
         Width           =   1890
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
         TabIndex        =   63
         Top             =   765
         Width           =   1005
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
         TabIndex        =   62
         Top             =   765
         Width           =   690
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
         TabIndex        =   61
         Top             =   765
         Width           =   945
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
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

