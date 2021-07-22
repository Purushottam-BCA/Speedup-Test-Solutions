VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{EAFDAFBF-1D88-41DD-B117-60ECBC4B8441}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form QpaprSetup 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Question Paper Properties"
   ClientHeight    =   9285
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   18645
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9285
   ScaleWidth      =   18645
   Begin vkUserContolsXP.vkFrame fram3 
      Height          =   5020
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   8864
      BackColor1      =   14737632
      BackColor2      =   16777215
      Caption         =   "Questions Selection"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleHeight     =   300
      Enabled         =   0   'False
      BorderColor     =   12632256
      RoundAngle      =   5
      Begin VB.ListBox List1 
         Height          =   780
         Left            =   2880
         Style           =   1  'Checkbox
         TabIndex        =   88
         Top             =   3000
         Width           =   3015
      End
      Begin VB.CommandButton ClearLeval 
         Caption         =   "Clear"
         Height          =   315
         Left            =   6000
         TabIndex        =   72
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin vkUserContolsXP.vkCommand Command1 
         Height          =   390
         Left            =   5795
         TabIndex        =   20
         Top             =   4560
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   688
         Caption         =   ">>>>"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         BorderColor     =   16744576
         CustomStyle     =   0
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1680
         Width           =   3015
      End
      Begin VB.ComboBox Combo2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2400
         Width           =   3015
      End
      Begin VB.ComboBox Combo4 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   4080
         Width           =   3015
      End
      Begin VB.OptionButton opt6 
         Caption         =   "Topic Wise Question"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   600
         Width           =   2415
      End
      Begin VB.OptionButton opt5 
         Caption         =   "Subject Wise Question"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label info5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*  Select Any format"
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2520
         TabIndex        =   60
         Top             =   360
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label lbl2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   19
         Top             =   2400
         Width           =   705
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Course"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   18
         Top             =   1680
         Width           =   675
      End
      Begin VB.Label lbl4 
         AutoSize        =   -1  'True
         Caption         =   "Difficulti Level"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   17
         Top             =   4080
         Width           =   1410
      End
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Topic"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   16
         Top             =   3000
         Width           =   525
      End
      Begin VB.Label star1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   2640
         TabIndex        =   15
         Top             =   1680
         Width           =   105
      End
      Begin VB.Label star2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   2640
         TabIndex        =   14
         Top             =   2400
         Width           =   105
      End
      Begin VB.Label star3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   2640
         TabIndex        =   13
         Top             =   3120
         Width           =   105
      End
   End
   Begin vkUserContolsXP.vkFrame fram2 
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   3413
      BackColor1      =   14737632
      BackColor2      =   16777215
      Caption         =   "Question Catogary"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleHeight     =   300
      Enabled         =   0   'False
      BorderColor     =   12632256
      RoundAngle      =   5
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton opt4 
         Caption         =   "Random Questions"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   600
         Width           =   2415
      End
      Begin VB.OptionButton opt3 
         Caption         =   "Selected Questions"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label info2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*  For Selecting  Question Manually , Select The  "" Selected Questions "" option."
         Enabled         =   0   'False
         ForeColor       =   &H008080FF&
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1570
         Width           =   5625
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   6960
      Top             =   120
   End
   Begin vkUserContolsXP.vkFrame fram1 
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   3413
      BackColor1      =   14737632
      BackColor2      =   16777215
      Caption         =   "Purpose"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleHeight     =   300
      BorderColor     =   12632256
      RoundAngle      =   5
      Begin vkUserContolsXP.vkCommand browse 
         Height          =   300
         Left            =   5880
         TabIndex        =   64
         ToolTipText     =   "Select Order Number"
         Top             =   1200
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   529
         Caption         =   "...."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         BorderColor     =   5454592
         CustomStyle     =   0
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton opt1 
         Caption         =   "General Use"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1080
         Picture         =   "Form5.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton opt2 
         Caption         =   "Client Order"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   3960
         Picture         =   "Form5.frx":13D4
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   500
         Width           =   1695
      End
      Begin VB.Label info1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*  For Individual && Common Purpose Select The ""General Use"" option."
         ForeColor       =   &H008080FF&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   1570
         Width           =   4950
      End
   End
   Begin VB.Frame fram4 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   9030
      Left            =   6945
      TabIndex        =   0
      Top             =   120
      Width           =   5430
      Begin VB.CommandButton nextbtn 
         Caption         =   "Next >>"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   91
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Timer Timer2 
         Interval        =   30
         Left            =   360
         Top             =   0
      End
      Begin vkUserContolsXP.vkFrame fram5 
         Height          =   6285
         Left            =   0
         TabIndex        =   30
         Top             =   2760
         Width           =   5445
         _ExtentX        =   9604
         _ExtentY        =   11086
         BackColor1      =   14474460
         BackColor2      =   16777215
         Caption         =   "Organisation && Paper Information"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TitleHeight     =   350
         RoundAngle      =   4
         Begin VB.TextBox txt9 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2535
            Locked          =   -1  'True
            TabIndex        =   61
            Top             =   4920
            Width           =   2040
         End
         Begin vkUserContolsXP.vkCheck chk1 
            Height          =   375
            Left            =   1200
            TabIndex        =   58
            Top             =   5520
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Include Answer Key"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox txt8 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3425
            MaxLength       =   2
            TabIndex        =   54
            Top             =   4320
            Width           =   480
         End
         Begin vkUserContolsXP.vkCommand Command2 
            Height          =   390
            Left            =   4530
            TabIndex        =   46
            Top             =   5840
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   688
            Caption         =   ">>>>"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   16744576
            CustomStyle     =   0
         End
         Begin VB.TextBox txt7 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2535
            MaxLength       =   2
            TabIndex        =   45
            Top             =   4320
            Width           =   480
         End
         Begin VB.TextBox txt6 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2535
            TabIndex        =   43
            Top             =   3720
            Width           =   2040
         End
         Begin VB.TextBox txt5 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2535
            TabIndex        =   41
            Top             =   3120
            Width           =   2040
         End
         Begin VB.TextBox txt4 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2535
            TabIndex        =   39
            Top             =   2520
            Width           =   2040
         End
         Begin VB.TextBox txt3 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2535
            TabIndex        =   37
            Top             =   1920
            Width           =   2520
         End
         Begin VB.TextBox txt2 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2535
            TabIndex        =   35
            Top             =   1320
            Width           =   2520
         End
         Begin VB.TextBox txt1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2535
            TabIndex        =   33
            Top             =   755
            Width           =   2520
         End
         Begin VB.Label Label26 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Eg :- Unit Test 1, Annual exam 2019"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   2610
            TabIndex        =   92
            Top             =   2280
            Visible         =   0   'False
            Width           =   2595
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   240
            MouseIcon       =   "Form5.frx":1985
            MousePointer    =   99  'Custom
            Picture         =   "Form5.frx":1AD7
            Stretch         =   -1  'True
            ToolTipText     =   "Click Here to See Demo"
            Top             =   5640
            Width           =   480
         End
         Begin VB.Label ANS_Key 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            Height          =   255
            Left            =   240
            TabIndex        =   71
            Top             =   5880
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label info52 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* Only Number Allow "
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   2400
            TabIndex        =   69
            Top             =   4680
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.Label info42 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* Maximum 9 hrs"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   2415
            TabIndex        =   68
            Top             =   4680
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.Label info32 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* Invalid Time"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   3780
            TabIndex        =   67
            Top             =   4680
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label info41 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* Maximum 100 Marks"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   3750
            TabIndex        =   66
            Top             =   4080
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label info31 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* Only Number Allow "
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   3840
            TabIndex        =   65
            Top             =   4080
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Marks per Question"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00733C00&
            Height          =   270
            Left            =   300
            TabIndex        =   63
            Top             =   4900
            Width           =   1980
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   270
            Left            =   2350
            TabIndex        =   62
            Top             =   4965
            Width           =   105
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "min"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00733C00&
            Height          =   270
            Left            =   3950
            TabIndex        =   56
            Top             =   4320
            Width           =   375
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "hr"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00733C00&
            Height          =   270
            Left            =   3110
            TabIndex        =   55
            Top             =   4320
            Width           =   225
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   270
            Left            =   2350
            TabIndex        =   53
            Top             =   4400
            Width           =   105
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   270
            Left            =   2350
            TabIndex        =   52
            Top             =   3760
            Width           =   105
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   270
            Left            =   2350
            TabIndex        =   51
            Top             =   3120
            Width           =   105
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   270
            Left            =   2350
            TabIndex        =   50
            Top             =   2575
            Width           =   105
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   270
            Left            =   2350
            TabIndex        =   49
            Top             =   2000
            Width           =   105
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   270
            Left            =   2350
            TabIndex        =   48
            Top             =   1400
            Width           =   105
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   270
            Left            =   2350
            TabIndex        =   47
            Top             =   720
            Width           =   105
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00733C00&
            Height          =   270
            Left            =   300
            TabIndex        =   44
            Top             =   4400
            Width           =   525
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Max. Marks"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00733C00&
            Height          =   270
            Left            =   300
            TabIndex        =   42
            Top             =   3810
            Width           =   1185
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Subject"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00733C00&
            Height          =   270
            Left            =   300
            TabIndex        =   40
            Top             =   3210
            Width           =   765
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Class"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00733C00&
            Height          =   270
            Left            =   300
            TabIndex        =   38
            Top             =   2610
            Width           =   585
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Test Name "
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00733C00&
            Height          =   270
            Left            =   300
            TabIndex        =   36
            Top             =   2010
            Width           =   1185
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Org. Address"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00733C00&
            Height          =   270
            Left            =   300
            TabIndex        =   34
            Top             =   1350
            Width           =   1335
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Organisation Name "
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00733C00&
            Height          =   270
            Left            =   300
            TabIndex        =   32
            Top             =   810
            Width           =   1995
         End
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2670
         TabIndex        =   29
         Top             =   1440
         Width           =   2280
      End
      Begin VB.Label info0 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "* Cannot be Zero."
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   3480
         TabIndex        =   70
         Top             =   1875
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label info4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "* Only Number Allow "
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   3300
         TabIndex        =   59
         Top             =   1890
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Label info3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "* Too Much Questions "
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   3360
         TabIndex        =   31
         Top             =   1880
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Total Questions : "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00733C00&
         Height          =   270
         Left            =   375
         TabIndex        =   28
         Top             =   1530
         Width           =   2355
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Available Questions : "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00733C00&
         Height          =   270
         Left            =   435
         TabIndex        =   27
         Top             =   600
         Width           =   2145
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   26
         Top             =   550
         Width           =   2295
      End
   End
   Begin vkUserContolsXP.vkFrame Frame5 
      Height          =   9180
      Left            =   12480
      TabIndex        =   57
      Top             =   50
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   16193
      Caption         =   "Additional Information"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleHeight     =   380
      Begin VB.CommandButton GoToPaper 
         Caption         =   "Go to Question Paper"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         MouseIcon       =   "Form5.frx":261C
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   8700
         Width           =   2055
      End
      Begin RichTextLib.RichTextBox rtfbox1 
         Height          =   4335
         Left            =   360
         TabIndex        =   74
         Top             =   2640
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   7646
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         TextRTF         =   $"Form5.frx":276E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vkUserContolsXP.vkFrame LogoFrame 
         Height          =   1335
         Left            =   240
         TabIndex        =   84
         Top             =   7635
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   2355
         BackColor1      =   14474460
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowTitle       =   0   'False
         Enabled         =   0   'False
         BorderWidth     =   0
         Begin VB.CommandButton BrowseLogo 
            Caption         =   "Browse"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   87
            Top             =   0
            Width           =   975
         End
         Begin VB.CommandButton ClearLOGO 
            Caption         =   "Clear"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   86
            Top             =   420
            Width           =   975
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Default"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   85
            Top             =   840
            Width           =   975
         End
         Begin VB.Image Image1 
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   1125
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1620
         End
      End
      Begin MSComDlg.CommonDialog cd1 
         Left            =   3840
         Top             =   5760
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin vkUserContolsXP.vkCheck logoCHK 
         Height          =   375
         Left            =   240
         TabIndex        =   83
         Top             =   7080
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Used School Logo (if any) "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   1
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   81
         Top             =   1200
         Width           =   960
      End
      Begin vkUserContolsXP.vkCheck NegMark 
         Height          =   375
         Left            =   200
         TabIndex        =   80
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Include Negative Marking "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   1
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   255
         Left            =   4620
         TabIndex        =   78
         Top             =   2950
         Width           =   150
         Begin VB.Label award_mark 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   -15
            TabIndex        =   79
            Top             =   0
            Width           =   135
         End
      End
      Begin vkUserContolsXP.vkCheck vkCheck1 
         Height          =   375
         Left            =   200
         TabIndex        =   73
         Top             =   2040
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Include Instruction "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   1
      End
      Begin VB.Label Label27 
         Caption         =   "Label27"
         Height          =   255
         Left            =   3240
         TabIndex        =   93
         Top             =   7560
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Decducted marks for each wrong answer "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00733C00&
         Height          =   615
         Left            =   240
         TabIndex        =   82
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2."
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   77
         Top             =   2970
         Width           =   180
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1."
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   76
         Top             =   2700
         Width           =   180
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "3. 4. 5.    6.   7."
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3525
         Left            =   120
         TabIndex        =   75
         Top             =   3270
         Width           =   180
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         FillColor       =   &H00FFFFFF&
         Height          =   4320
         Left            =   75
         Top             =   2655
         Width           =   315
      End
   End
   Begin VB.CommandButton Command5 
      Height          =   8415
      Left            =   0
      Picture         =   "Form5.frx":27EA
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   0
      Visible         =   0   'False
      Width           =   12495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      Height          =   9135
      Left            =   6890
      Top             =   80
      Width           =   5535
   End
End
Attribute VB_Name = "QpaprSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pic_name As String, PicAvail As Byte
Private Sub BrowseLogo_Click()
On Error Resume Next
cd1.Filter = "all picture files *.jpg,.gif,.bmp,.ico jpeg image"
cd1.ShowOpen
If cd1.FileName <> "" Then
 Image1.Picture = LoadPicture(cd1.FileName)
pic_name = cd1.FileName
school_pic = pic_name
PicAvail = 1
Else
PicAvail = 0
Exit Sub
End If
End Sub

Private Sub ClearLOGO_Click()
On Error Resume Next
    Set Image1.Picture = Nothing
    PicAvail = 0
    pic_name = ""
    school_pic = App.Path & "\Graphics\Images\Logo\mmf.gif"
End Sub

Private Sub Command4_Click() 'Default Picture button
If logoCHK.Value = vbChecked Then
Image1.Picture = LoadPicture(App.Path & "\Graphics\Images\Logo\mmf.gif")
school_pic = App.Path & "\Graphics\Gifs\mmf.gif"
pic_name = school_pic
PicAvail = 1
End If
End Sub
Private Sub browse_Click()
FrmClient4.Show vbModal, MDI
 pic_name = Label27.Caption
 school_pic = Label27.Caption
 PicAvail = 1
End Sub

Private Sub chk1_Click()
If chk1.Value = vbChecked Then
ANS_Key.Caption = 1
Else
ANS_Key.Caption = 0
End If
End Sub

Private Sub ClearLeval_Click()
'''''''''''''''''''''''''''''''''''''''''''''''''''PPPKKK
' Set r = c1.Execute("select count(*) from quesms where sub_id =(select sub_id from sub where sub_nm ='" & Combo2.Text & "' and c_id=(select c_id from course where c_nm='" & Combo1.Text & "'))and c_id=(select c_id from course where c_nm='" & Combo1.Text & "') ")

Combo4.Clear
Ques_diff_leval = ""
Combo4.AddItem "EASY"
Combo4.AddItem "MEDIUM"
Combo4.AddItem "HARD"
End Sub

Private Sub Combo1_Click()
List1.Clear
Combo2.Clear
Set r1 = New ADODB.Recordset
Set r1 = c1.Execute("select sub_nm from sub where c_id =(select c_id from course where upper(c_nm)='" & UCase(Combo1.Text) & "') ")
If IsNull(r1.Fields(0)) = False Then
While r1.EOF = False
 Combo2.AddItem r1.Fields(0)
 r1.MoveNext
Wend
If Text1.Text = opt2.Caption Then
Else
txt4.Text = Combo1.Text
'txt4.Locked = True
End If
Else
End If
Set r = New ADODB.Recordset
Set r = c.Execute("select count(*)from quesms where c_id=(select c_id from course where c_nm='" & Combo1.Text & "') ")
If IsNull(r.Fields(0)) = False Then
   Label1.Caption = r.Fields(0)
  Else
   r.Close
  End If
End Sub

Private Sub combo2_Click()
List1.Clear
If Combo1.Text = "" Then
Combo1.SetFocus
Exit Sub
End If
Set r = New ADODB.Recordset
Set r = c.Execute("select count(*)from quesms where c_id=(select c_id from course where c_nm='" & Combo1.Text & "') and sub_id=(select  sub_id from sub where sub_nm='" & Combo2.Text & "' and c_id=(select c_id from course where c_nm='" & Combo1.Text & "')) ")
If IsNull(r.Fields(0)) = False Then
   Label1.Caption = r.Fields(0)
  Else
  r.Close
End If
Set r1 = New ADODB.Recordset
Set r1 = c1.Execute("select TP_NM from topic where sub_id =(select sub_id from sub where sub_nm='" & Combo2.Text & "' and c_id=(select c_id from course where c_nm='" & Combo1.Text & "'))and c_id =(select c_id from course where c_nm='" & Combo1.Text & "') ")
If r1.EOF = False Then
 While r1.EOF = False
  List1.AddItem r1.Fields(0)
  r1.MoveNext
 Wend
End If
 If Text1.Text = opt2.Caption Then
 Else
   txt5.Text = Combo2.Text
  'txt5.Locked = True
 End If
End Sub

Private Sub combo3_Click()
If Combo1.Text = "" Then
Combo1.SetFocus
ElseIf Combo2.Text = "" Then
Combo2.SetFocus
Else
Set r = New ADODB.Recordset
Set r = c.Execute("select count(*)from quesms where c_id=(select c_id from course where c_nm='" & Combo1.Text & "') and sub_id=(select  sub_id from sub where sub_nm='" & Combo2.Text & "' and c_id=(select c_id from course where c_nm='" & Combo1.Text & "'))and tp_id=(select tp_id from topic where tp_nm='" & Combo3.Text & "' and c_id=(select c_id from course where c_nm='" & Combo1.Text & "')and sub_id=(select  sub_id from sub where sub_nm='" & Combo2.Text & "' and c_id=(select c_id from course where c_nm='" & Combo1.Text & "'))) ")
If IsNull(r.Fields(0)) = False Then
   Label1.Caption = r.Fields(0)
  Else
   r.Close
End If
End If
End Sub

Private Sub Combo4_Click()
If Text3.Text = opt5.Caption Then 'Subject Wise
 If Combo1.Text = "" Then
  Combo1.SetFocus
 ElseIf Combo2.Text = "" Then
  Combo2.SetFocus
 Else
   info5.Visible = False
   Set r = New ADODB.Recordset
   Set r = c1.Execute("select count(*) from quesms where sub_id =(select sub_id from sub where sub_nm ='" & Combo2.Text & "' and c_id=(select c_id from course where c_nm='" & Combo1.Text & "'))and c_id=(select c_id from course where c_nm='" & Combo1.Text & "') and Q_DIF_LVL='" & Combo4.Text & "'  ")
   Label1.Caption = r.Fields(0)
   Exit Sub
 End If
ElseIf Text3.Text = opt6.Caption Then 'Topic Wise
 If Combo1.Text = "" Then
   Combo1.SetFocus
  ElseIf Combo2.Text = "" Then
   Combo2.SetFocus
  ElseIf Combo3.Text = "" Then
   Combo3.SetFocus
 Else
 Dim i As Integer
 Label1.Caption = 0
 Set r = New ADODB.Recordset
For i = 0 To List1.ListCount - 1
 If List1.Selected(i) Then
  Set r = c1.Execute("select count(*) from quesms where sub_id =(select sub_id from sub where sub_nm ='" & Combo2.Text & "' and c_id=(select c_id from course where c_nm='" & Combo1.Text & "'))and c_id=(select c_id from course where c_nm='" & Combo1.Text & "') and Q_DIF_LVL='" & Combo4.Text & "' and tp_id=(select tp_id from topic where tp_nm='" & List1.list(i) & "' and c_id=(select c_id from course where c_nm='" & Combo1.Text & "')and sub_id=(select  sub_id from sub where sub_nm='" & Combo2.Text & "' and c_id=(select c_id from course where c_nm='" & Combo1.Text & "'))) ")
  Label1.Caption = Val(Label1.Caption) + r.Fields(0)
 End If
Next i
 End If
ElseIf Text3.Text = "" Then
 info5.Visible = True
 Exit Sub
End If
End Sub

Private Sub Command1_Click() '1st Next
If Text3.Text = "" Then
 info5.Visible = True
 Exit Sub
ElseIf Combo1.Text = "" Then
 Combo1.SetFocus
 info5.Visible = False
 Exit Sub
ElseIf Combo2.Text = "" Then
 Combo2.SetFocus
 info5.Visible = False
 Exit Sub
ElseIf Text3.Text = opt6.Caption Then
 If List1.SelCount = 0 Then
  MsgBox "Select Topic ", vbInformation + vbOKOnly, "Topic Not Selected"
  List1.SetFocus
 'If Combo3.Text = "" Then
 'Combo3.SetFocus
 info5.Visible = False
 Exit Sub
 End If
Else
info5.Visible = False
End If
 Timer1.Enabled = True
End Sub

Private Sub Command5_Click()
Command5.Visible = False
End Sub

Private Sub Form_Load()
Me.Width = 7100
conn
Me.Top = 200
Me.Left = 200

While rs_course.EOF = False  'Adding Course
 Combo1.AddItem rs_course.Fields(0)
 rs_course.MoveNext
Wend
Combo4.Clear
Combo4.AddItem "EASY"
Combo4.AddItem "MEDIUM"
Combo4.AddItem "HARD"

Shape1.Visible = False
fram4.Visible = False
fram5.Visible = False

GoToPaper.Enabled = False

Timer1.Enabled = False
Timer2.Enabled = False
End Sub



Private Sub Form_Unload(cancel As Integer)
QuestionPPRdashboard.Show
End Sub


Private Sub GoToPaper_Click()
If logoCHK.Value = vbChecked Then 'For Logo
 If PicAvail = 1 Then
   school_pic = pic_path
 Else
  MsgBox "Select The Logo for Organisation or click on the Default Logo", vbInformation + vbOKOnly, "Logo Unavailable"
  Exit Sub
 End If
End If

If NegMark.Value = vbChecked Then 'For Negative Mark
 If Trim(Text5.Text) = "" Then
  MsgBox "Enter  Negative Mark value", vbExclamation + vbOKOnly, "Empty Value"
  Text5.SetFocus
  Exit Sub
  Else
  End If
End If

If vkCheck1.Value = vbChecked Then
 WantInstruction = 1
Else
 WantInstruction = 2
End If
'Storing Data In Variables
Ques_Purpose = Text1.Text
purposedash = UCase(Text1.Text) 'For DashBoard
Ques_Cat = Text2.Text
Ques_selection = Text3.Text
Ques_Course = Combo1.Text
classdash = UCase(Combo1.Text)  'For DashBoard
Ques_Subject = Combo2.Text
subdash = UCase(Combo2.Text) 'For Dashboard
Ques_diff_leval = Combo4.Text
tot_Ques_in_bank = Val(Label1.Caption)
delivrdt = Date 'for Dashboard
'For Paper Display
 school_nm = UCase(Trim(txt1.Text)) 'Schoool Name
 school_add = Trim(txt2.Text)
 Test_nm = Trim(txt3.Text)
 tstTypDash = UCase(Test_nm) 'For Dashboard
 testSUB_nm = txt5.Text
 testclass_nm = txt4.Text
 testFULLmrk = Int(Val(txt6.Text))
 totmrkdash = testFULLmrk  'For dashboard
 testTOTALtime = Format(Int(Val(txt7.Text)), "00") & ":" & Format(Int(Val(txt8.Text)), "00")
 testTotQues = Val(Text4.Text)
 totqsdash = testTotQues   'For Dashboard
 testCorrectMRK = Val(txt9.Text)
 testWrongMRK = Val(Text5.Text)
 ppr_Time_hr = Int(txt7.Text)
 ppr_time_min = Int(txt8.Text)
school_pic = pic_name

 rtfbox1.Text = "The Question Paper contains total " & testTotQues & " Questions ." & vbCrLf & "All Questions are compulsory." & vbCrLf & "Each Question will have 4 choices ,out of which only one choice is correct." & vbCrLf & "Darken the Options (A/B/C/D) with Ball Pen (Blue or Black) ONLY. " & vbCrLf & "For Each Question , you will be awarded " & testCorrectMRK & " marks if you have darkened only one bubble corresponding to the right answer." & vbCrLf & "In case you have not darkened any option ,you will be awarded 0 mark for that question." & vbCrLf & "In Case if you answered wrong or if you have darkened more than one options, " & Val(Text5.Text) & " mark will be deducted for each wrong answer"
 instructionSET = rtfbox1.Text
 
 Ques_Include_Ans = Int(ANS_Key.Caption) '1 if answer required

QpaprSetup.Hide
Question_PPR.Show
End Sub

Private Sub Image2_Click() 'Help
If Command5.Visible = True Then
  Command5.Visible = False
Else
 Command5.Visible = True
End If
End Sub

Private Sub List1_Click()
Dim i As Integer
Label1.Caption = 0
Set r = New ADODB.Recordset
For i = 0 To List1.ListCount - 1
If List1.Selected(i) Then
 Set r = c.Execute("select count(*)from quesms where c_id=(select c_id from course where c_nm='" & Combo1.Text & "') and sub_id=(select  sub_id from sub where sub_nm='" & Combo2.Text & "' and c_id=(select c_id from course where c_nm='" & Combo1.Text & "'))and tp_id=(select tp_id from topic where tp_nm='" & List1.list(i) & "' and c_id=(select c_id from course where c_nm='" & Combo1.Text & "')and sub_id=(select  sub_id from sub where sub_nm='" & Combo2.Text & "' and c_id=(select c_id from course where c_nm='" & Combo1.Text & "'))) ")
If IsNull(r.Fields(0)) = False Then
   Label1.Caption = Val(Label1.Caption) + r.Fields(0)
End If
End If
Next i
End Sub

Private Sub logoCHK_Click() 'Apllying School Logo
If logoCHK.Value = vbChecked Then
 LogoFrame.Enabled = True
Else
 LogoFrame.Enabled = False
 End If
End Sub

Private Sub nextbtn_Click()
If Text4.Text = "" Or Val(Text4.Text) > Val(Label1.Caption) Then
 Text4.SetFocus
 Frame5.Visible = False
 MsgBox "Cannot Enter More Questions than available in Database..", vbInformation + vbOKOnly, ""
ElseIf Text4.Text <> "" Then
 fram5.Visible = True
End If
End Sub

Private Sub opt1_Click()
Text1.Text = opt1.Caption
browse.Enabled = False
fram2.Enabled = True
ordrdt = Date 'for Dashboard
End Sub

Private Sub opt2_Click()
Text1.Text = opt2.Caption
browse.Enabled = True
fram2.Enabled = False
End Sub

Private Sub opt3_Click()
Ques_Cat = opt4.Caption
Text2.Text = opt3.Caption
fram3.Enabled = True
End Sub

Private Sub opt4_Click()
Ques_Cat = opt4.Caption
Text2.Text = opt4.Caption
fram3.Enabled = True
End Sub

Private Sub opt5_Click()
Text3.Text = opt5.Caption
List1.Enabled = False
lbl3.Enabled = False
star3.Enabled = False
info5.Visible = False
End Sub

Private Sub opt6_Click()
Text3.Text = opt6.Caption
List1.Enabled = True
lbl3.Enabled = True
star3.Enabled = True
info5.Visible = False
End Sub


Private Sub Text4_Change()
If Text4.Text = "" Then
 nextbtn.Enabled = False
Else
 nextbtn.Enabled = True
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer) 'Total Question required
 If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Then
        Text4.SetFocus
       If Val(Text4.Text & Chr(KeyAscii)) <= Val(Label1.Caption) And Val(Text4.Text & Chr(KeyAscii)) > 0 Then
          info3.Visible = False
          info4.Visible = False
          info0.Visible = False
          nextbtn.Enabled = True
    ElseIf Val(Text4.Text & Chr(KeyAscii)) > Val(Label1.Caption) Then
        info3.Visible = True
       info4.Visible = False
          info0.Visible = False
          nextbtn.Enabled = False
        KeyAscii = 0
    ElseIf Val(Text4.Text & Chr(KeyAscii)) <= 0 Then
        info3.Visible = False
        info4.Visible = False
        info0.Visible = True
        nextbtn.Enabled = False
       End If
    Else
        KeyAscii = 0
        info4.Visible = True
        info3.Visible = False
        info0.Visible = False
        nextbtn.Enabled = False
    End If
End Sub



Private Sub Text5_LostFocus()
If NegMark.Value = vbChecked Then
 rtfbox1.Text = "The Question Paper contains total " & testTotQues & " Questions ." & vbCrLf & "All Questions are compulsory." & vbCrLf & "Each Question will have 4 choices ,out of which only one choice is correct." & vbCrLf & "Darken the Options (A/B/C/D) with Ball Pen (Blue or Black) ONLY. " & vbCrLf & "For Each Question , you will be awarded " & testCorrectMRK & " marks if you have darkened only one bubble corresponding to the right answer." & vbCrLf & "In case you have not darkened any option ,you will be awarded 0 mark for that question." & vbCrLf & "In Case if you answered wrong or if you have darkened more than one options, " & Val(Text5.Text) & " mark will be deducted for each wrong answer"
End If
End Sub

Private Sub Timer1_Timer()
If Command1.Caption = ">>>>" Then
Me.Width = Me.Width + 100
If Me.Width >= 12555 Then
 Timer1.Enabled = False
 fram4.Visible = True
 'fram5.Visible = True
 Shape1.Visible = True
 Command1.Caption = "<<<<"
 End If
Else
 Me.Width = Me.Width - 150
 If Me.Width <= 7100 Then
  Timer1.Enabled = False
  fram4.Visible = False
  fram5.Visible = False
  Shape1.Visible = False
 Command1.Caption = ">>>>"
  End If
End If
End Sub

Private Sub Command2_Click()
If txt1.Text = "" Then
GoToPaper.Enabled = False
txt1.SetFocus
 ElseIf txt2.Text = "" Then
 GoToPaper.Enabled = False
txt2.SetFocus
 ElseIf txt3.Text = "" Then
 GoToPaper.Enabled = False
 txt3.SetFocus
 ElseIf txt6.Text = "" Then
 GoToPaper.Enabled = False
 txt6.SetFocus
 ElseIf txt7.Text = "" Then
 GoToPaper.Enabled = False
 txt7.SetFocus
 ElseIf txt8.Text = "" Then
 GoToPaper.Enabled = False
 txt8.SetFocus
 ElseIf txt9.Text = "" Then
 GoToPaper.Enabled = False
 txt9.SetFocus
 Else
Timer2.Enabled = True
 tot_Ques_in_bank = Val(Label1.Caption)
 testTotQues = Val(Text4.Text)
 testCorrectMRK = Val(txt9.Text)
 testWrongMRK = Val(Text5.Text)
rtfbox1.Text = "The Question Paper contains total " & testTotQues & " Questions ." & vbCrLf & "All Questions are compulsory." & vbCrLf & "Each Question will have 4 choices ,out of which only one choice is correct." & vbCrLf & "Darken the Options (A/B/C/D) with Ball Pen (Blue or Black) ONLY. " & vbCrLf & "For Each Question , you will be awarded " & testCorrectMRK & " marks if you have darkened only one bubble corresponding to the right answer." & vbCrLf & "In case you have not darkened any option ,you will be awarded 0 mark for that question." & vbCrLf & "In Case if you answered wrong or if you have darkened more than one options, " & Val(Text5.Text) & " mark will be deducted for each wrong answer"
GoToPaper.Enabled = True
End If
 Command5.Visible = False
 Image1.Picture = LoadPicture(App.Path & "\Graphics\Images\Logo\mmf.gif")
school_pic = App.Path & "\Graphics\Gifs\mmf.gif"
pic_name = school_pic
PicAvail = 1
Text5.Text = 0
End Sub

Private Sub Timer2_Timer()

If Command2.Caption = ">>>>" Then
Me.Width = Me.Width + 90
If Me.Width >= 18735 Then
 Timer2.Enabled = False
 Frame5.Visible = True
 Command2.Caption = "<<<<"
 End If
Else
 Me.Width = Me.Width - 150
 If Me.Width <= 12555 Then
  Timer2.Enabled = False
  Frame5.Visible = False
 Command2.Caption = ">>>>"
  End If
  End If
End Sub

Private Sub txt1_KeyPress(KeyAscii As Integer)
   If (((KeyAscii >= 65) And (KeyAscii <= 90)) Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or (KeyAscii = 8) Or (KeyAscii = 32)) Then
        txt1.SetFocus
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt2.SetFocus
     Else
        KeyAscii = 0
    End If
End Sub

Private Sub txt2_KeyPress(KeyAscii As Integer)
   If (((KeyAscii >= 65) And (KeyAscii <= 90)) Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or (KeyAscii = 8) Or (KeyAscii = 32) Or ((KeyAscii >= 48) And (KeyAscii <= 57))) Then
        txt2.SetFocus
            ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt3.SetFocus
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txt3_KeyPress(KeyAscii As Integer)
 If (((KeyAscii >= 65) And (KeyAscii <= 90)) Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or (KeyAscii = 8) Or (KeyAscii = 32) Or ((KeyAscii >= 48) And (KeyAscii <= 57))) Then
        txt3.SetFocus
        ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt4.SetFocus
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txt4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        KeyAscii = 0
        txt5.SetFocus
        End If
End Sub
Private Sub txt5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        KeyAscii = 0
        txt6.SetFocus
        End If
End Sub

Private Sub txt6_Change()
If Val(Text4.Text) <> 0 Then
txt9.Text = Format(Val(txt6.Text) / Val(Text4.Text), "00.0") 'A perfect format for number like 3.3 not 3.33333
txt9.Locked = True
Else
MsgBox "Select Correct Numbers"
End If
End Sub

Private Sub txt6_KeyPress(KeyAscii As Integer)
If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Then
        txt6.SetFocus
       If Val(txt6.Text & Chr(KeyAscii)) <= 100 And Val(txt6.Text & Chr(KeyAscii)) > 0 Then
          info31.Visible = False
          info41.Visible = False
       Else
       info31.Visible = False
       info41.Visible = True
        KeyAscii = 0
       End If
   ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt7.SetFocus
    Else
        KeyAscii = 0
        info41.Visible = False
        info31.Visible = True
    End If
End Sub

Private Sub txt7_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        KeyAscii = 0
        txt7.SetFocus
        Exit Sub
        End If
If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Then
        txt7.SetFocus
          If Val(txt7.Text & Chr(KeyAscii)) < 10 And Val(txt7.Text & Chr(KeyAscii)) > 0 Then
            info32.Visible = False
            info52.Visible = False
            info42.Visible = False
           ElseIf Val(txt7.Text & Chr(KeyAscii)) = 0 Then
            If Val(txt8.Text) = 0 And Trim(txt8.Text) <> "" Then
             info32.Visible = True
             info52.Visible = False
             info42.Visible = False
             KeyAscii = 0
            Else
             info32.Visible = False
             info52.Visible = False
             info42.Visible = False
            End If
           Else
            info52.Visible = False
            info42.Visible = True
            info32.Visible = False
            KeyAscii = 0
         End If
      Else
        info52.Visible = True
        info32.Visible = False
        info42.Visible = False
        KeyAscii = 0
      End If
End Sub

Private Sub txt8_KeyPress(KeyAscii As Integer)
If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Then
        txt8.SetFocus
        If (Val(txt7.Text) > 0) Then
         If Val(txt8.Text & Chr(KeyAscii)) < 60 And Val(txt8.Text & Chr(KeyAscii)) >= 0 Then
            info32.Visible = False
            info42.Visible = False
            info52.Visible = False
         Else
            info42.Visible = False
            info32.Visible = True
            info52.Visible = False
            KeyAscii = 0
         End If
        ElseIf Val(txt7.Text) = 0 Or txt7.Text = "" Then
        If Val(txt8.Text & Chr(KeyAscii)) < 60 And Val(txt8.Text & Chr(KeyAscii)) > 0 Then
            info32.Visible = False
            info42.Visible = False
            info52.Visible = False
        Else
            info42.Visible = False
            info32.Visible = True
            info52.Visible = False
           KeyAscii = 0
         End If
       End If
    Else
    info42.Visible = False
    info32.Visible = False
    info52.Visible = True
    KeyAscii = 0
End If
End Sub

Private Sub txt9_Change()
If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Then
        txt9.SetFocus
        Else
    KeyAscii = 0
End If
End Sub

