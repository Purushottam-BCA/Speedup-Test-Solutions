VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Begin VB.Form QuesBank 
   BackColor       =   &H00808080&
   ClientHeight    =   10245
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18960
   Icon            =   "QuesBank.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10245
   ScaleWidth      =   18960
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   11280
      ScaleHeight     =   330
      ScaleWidth      =   2055
      TabIndex        =   72
      Top             =   45
      Width           =   2055
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Question Property"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   73
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8760
      MouseIcon       =   "QuesBank.frx":0EE2
      MousePointer    =   99  'Custom
      Picture         =   "QuesBank.frx":1034
      Style           =   1  'Graphical
      TabIndex        =   67
      ToolTipText     =   "Search Here"
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Refresh All"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      MouseIcon       =   "QuesBank.frx":1C76
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   66
      ToolTipText     =   "Refresh"
      Top             =   240
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1560
      MouseIcon       =   "QuesBank.frx":1DC8
      MousePointer    =   99  'Custom
      Style           =   2  'Dropdown List
      TabIndex        =   51
      Top             =   1485
      Width           =   3255
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1545
      MouseIcon       =   "QuesBank.frx":1F1A
      MousePointer    =   99  'Custom
      Style           =   2  'Dropdown List
      TabIndex        =   50
      Top             =   865
      Width           =   2895
   End
   Begin VB.ComboBox Combo4 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1560
      MouseIcon       =   "QuesBank.frx":206C
      MousePointer    =   99  'Custom
      Style           =   2  'Dropdown List
      TabIndex        =   49
      Top             =   270
      Width           =   2535
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5730
      MouseIcon       =   "QuesBank.frx":21BE
      MousePointer    =   99  'Custom
      Style           =   2  'Dropdown List
      TabIndex        =   48
      Top             =   1500
      Width           =   2055
   End
   Begin VB.ComboBox Combo9 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5760
      MouseIcon       =   "QuesBank.frx":2310
      MousePointer    =   99  'Custom
      Style           =   2  'Dropdown List
      TabIndex        =   47
      Top             =   540
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillColor       =   &H00C0C0C0&
      Height          =   495
      Index           =   1
      Left            =   17160
      ScaleHeight     =   495
      ScaleWidth      =   375
      TabIndex        =   46
      Top             =   9480
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillColor       =   &H00C0C0C0&
      Height          =   495
      Index           =   0
      Left            =   14040
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   45
      Top             =   9480
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      MouseIcon       =   "QuesBank.frx":2462
      MousePointer    =   99  'Custom
      Picture         =   "QuesBank.frx":25B4
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Delete Question"
      Top             =   9480
      Width           =   2895
   End
   Begin MSComctlLib.ListView Lvl1 
      Height          =   8075
      Left            =   0
      TabIndex        =   0
      Top             =   2160
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   14235
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "QuesBank.frx":2E77
      NumItems        =   0
   End
   Begin VB.CommandButton Command3 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   17400
      MouseIcon       =   "QuesBank.frx":2FD9
      MousePointer    =   99  'Custom
      Picture         =   "QuesBank.frx":312B
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Exit From Here"
      Top             =   9480
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14400
      MouseIcon       =   "QuesBank.frx":394D
      MousePointer    =   99  'Custom
      Picture         =   "QuesBank.frx":3A9F
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Delete Question"
      Top             =   9480
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FC9090&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   11115
      TabIndex        =   3
      Top             =   4935
      Width           =   9150
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   2415
         Left            =   720
         TabIndex        =   4
         Top             =   480
         Width           =   14970
         Begin VB.OptionButton btnopt 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   360
            MouseIcon       =   "QuesBank.frx":423A
            MousePointer    =   99  'Custom
            TabIndex        =   71
            Top             =   200
            Width           =   255
         End
         Begin VB.OptionButton btnopt 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   360
            MouseIcon       =   "QuesBank.frx":438C
            MousePointer    =   99  'Custom
            TabIndex        =   70
            Top             =   2000
            Width           =   255
         End
         Begin VB.OptionButton btnopt 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   360
            MouseIcon       =   "QuesBank.frx":44DE
            MousePointer    =   99  'Custom
            TabIndex        =   69
            Top             =   1400
            Width           =   255
         End
         Begin VB.OptionButton btnopt 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   360
            MouseIcon       =   "QuesBank.frx":4630
            MousePointer    =   99  'Custom
            TabIndex        =   68
            Top             =   800
            Width           =   255
         End
         Begin RichTextLib.RichTextBox opt1 
            Height          =   585
            Index           =   0
            Left            =   1080
            TabIndex        =   5
            Top             =   15
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   1032
            _Version        =   393217
            BackColor       =   14737632
            BorderStyle     =   0
            Enabled         =   -1  'True
            MaxLength       =   200
            Appearance      =   0
            TextRTF         =   $"QuesBank.frx":4782
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin RichTextLib.RichTextBox opt1 
            Height          =   585
            Index           =   1
            Left            =   1080
            TabIndex        =   6
            Top             =   615
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   1032
            _Version        =   393217
            BackColor       =   14737632
            BorderStyle     =   0
            Enabled         =   -1  'True
            MaxLength       =   200
            Appearance      =   0
            TextRTF         =   $"QuesBank.frx":48AE
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin RichTextLib.RichTextBox opt1 
            Height          =   585
            Index           =   2
            Left            =   1080
            TabIndex        =   7
            Top             =   1215
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   1032
            _Version        =   393217
            BackColor       =   14737632
            BorderStyle     =   0
            Enabled         =   -1  'True
            MaxLength       =   200
            Appearance      =   0
            TextRTF         =   $"QuesBank.frx":49DA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin RichTextLib.RichTextBox opt1 
            Height          =   585
            Index           =   3
            Left            =   1080
            TabIndex        =   8
            Top             =   1815
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   1032
            _Version        =   393217
            BackColor       =   14737632
            BorderStyle     =   0
            Enabled         =   -1  'True
            MaxLength       =   200
            Appearance      =   0
            TextRTF         =   $"QuesBank.frx":4B06
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Line Line16 
            X1              =   8415
            X2              =   8415
            Y1              =   0
            Y2              =   2400
         End
         Begin VB.Line Line8 
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   2400
         End
         Begin VB.Line Line14 
            X1              =   0
            X2              =   15000
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Line Line13 
            X1              =   0
            X2              =   15000
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Line Line12 
            X1              =   0
            X2              =   15000
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Line Line9 
            X1              =   0
            X2              =   15000
            Y1              =   2400
            Y2              =   2400
         End
         Begin VB.Line Line11 
            X1              =   0
            X2              =   15000
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line Line23 
            X1              =   960
            X2              =   960
            Y1              =   0
            Y2              =   2400
         End
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Choices"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1920
         TabIndex        =   15
         Top             =   120
         Width           =   840
      End
      Begin VB.Line Line17 
         X1              =   720
         X2              =   720
         Y1              =   0
         Y2              =   480
      End
      Begin VB.Line Line15 
         X1              =   0
         X2              =   720
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line10 
         X1              =   0
         X2              =   720
         Y1              =   1685
         Y2              =   1685
      End
      Begin VB.Line Line7 
         X1              =   0
         X2              =   720
         Y1              =   1085
         Y2              =   1085
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   240
         TabIndex        =   14
         Top             =   2325
         Width           =   180
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   240
         TabIndex        =   13
         Top             =   1800
         Width           =   165
      End
      Begin VB.Line Line6 
         X1              =   0
         X2              =   14160
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   2880
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   15750
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   14160
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line5 
         X1              =   12550
         X2              =   15690
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No."
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   210
         TabIndex        =   12
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   180
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   165
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Correct"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   825
         TabIndex        =   9
         Top             =   120
         Width           =   810
      End
      Begin VB.Line Line24 
         X1              =   1680
         X2              =   1680
         Y1              =   0
         Y2              =   480
      End
   End
   Begin RichTextLib.RichTextBox qtext_mcqs 
      Height          =   1440
      Left            =   11235
      TabIndex        =   1
      Top             =   2880
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   2540
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"QuesBank.frx":4C32
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Question Property"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   11160
      TabIndex        =   19
      Top             =   240
      Width           =   8775
      Begin VB.ComboBox Combo11 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6480
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   300
         Width           =   2055
      End
      Begin VB.ComboBox Combo8 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1800
         MouseIcon       =   "QuesBank.frx":4CAE
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1365
         Width           =   3255
      End
      Begin VB.ComboBox Combo7 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1785
         MouseIcon       =   "QuesBank.frx":4E00
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   750
         Width           =   2895
      End
      Begin VB.ComboBox Combo6 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1800
         MouseIcon       =   "QuesBank.frx":4F52
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   150
         Width           =   2535
      End
      Begin VB.ComboBox Combo5 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6450
         MouseIcon       =   "QuesBank.frx":50A4
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1380
         Width           =   2055
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Aardvark"
            Size            =   9.75
            Charset         =   0
            Weight          =   800
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   7560
         TabIndex        =   39
         Top             =   -120
         Width           =   105
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TYPE"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6960
         TabIndex        =   38
         Top             =   0
         Width           =   435
      End
      Begin VB.Line Line18 
         BorderColor     =   &H8000000A&
         BorderWidth     =   2
         X1              =   5640
         X2              =   5640
         Y1              =   0
         Y2              =   1910
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Aardvark"
            Size            =   9.75
            Charset         =   0
            Weight          =   800
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   1560
         TabIndex        =   31
         Top             =   765
         Width           =   105
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Aardvark"
            Size            =   9.75
            Charset         =   0
            Weight          =   800
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   1560
         TabIndex        =   30
         Top             =   120
         Width           =   105
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Aardvark"
            Size            =   9.75
            Charset         =   0
            Weight          =   800
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   1560
         TabIndex        =   29
         Top             =   1440
         Width           =   105
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "SUBJECT"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   480
         TabIndex        =   28
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "CHAPTER"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   " COURSE"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   480
         TabIndex        =   26
         Top             =   220
         Width           =   1095
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Difficulti Level"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   270
         Left            =   6675
         TabIndex        =   25
         Top             =   960
         Width           =   1365
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Aardvark"
            Size            =   9.75
            Charset         =   0
            Weight          =   800
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   8250
         TabIndex        =   24
         Top             =   1200
         Width           =   105
      End
   End
   Begin RichTextLib.RichTextBox expn_mcq 
      Height          =   720
      Left            =   11205
      TabIndex        =   33
      Top             =   8340
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   1270
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"QuesBank.frx":51F6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00808080&
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   2295
      Index           =   1
      Left            =   11100
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   9135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      Index           =   1
      X1              =   8280
      X2              =   8295
      Y1              =   15
      Y2              =   2000
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Aardvark"
         Size            =   9.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   1320
      TabIndex        =   65
      Top             =   885
      Width           =   105
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Aardvark"
         Size            =   9.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   1320
      TabIndex        =   64
      Top             =   240
      Width           =   105
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Aardvark"
         Size            =   9.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   1320
      TabIndex        =   63
      Top             =   1440
      Width           =   105
   End
   Begin VB.Label Label4 
      Caption         =   "SUBJECT"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   62
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "CHAPTER"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   61
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   " COURSE"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   60
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Difficulti Leval"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5925
      TabIndex        =   59
      Top             =   1080
      Width           =   1545
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Aardvark"
         Size            =   9.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   7530
      TabIndex        =   58
      Top             =   1080
      Width           =   105
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      Index           =   0
      X1              =   5160
      X2              =   5175
      Y1              =   15
      Y2              =   2000
   End
   Begin VB.Label lblcid 
      BackColor       =   &H8000000D&
      Height          =   255
      Left            =   5280
      TabIndex        =   57
      Top             =   1440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblSid 
      BackColor       =   &H8000000A&
      Height          =   255
      Left            =   6720
      TabIndex        =   56
      Top             =   1440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblQtyp 
      BackColor       =   &H0080FFFF&
      Height          =   255
      Left            =   6720
      TabIndex        =   55
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lbltpid 
      BackColor       =   &H008080FF&
      Height          =   255
      Left            =   5280
      TabIndex        =   54
      Top             =   1800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TYPE"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   53
      Top             =   240
      Width           =   555
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Aardvark"
         Size            =   9.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   6840
      TabIndex        =   52
      Top             =   120
      Width           =   105
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00808080&
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   2055
      Index           =   0
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   10935
   End
   Begin VB.Label lblcid1 
      BackColor       =   &H8000000D&
      Height          =   255
      Left            =   12360
      TabIndex        =   43
      Top             =   2520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblsid1 
      BackColor       =   &H8000000A&
      Height          =   255
      Left            =   13800
      TabIndex        =   42
      Top             =   2520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblqtyp1 
      BackColor       =   &H0080FFFF&
      Height          =   255
      Left            =   16560
      TabIndex        =   41
      Top             =   2520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lbltpid1 
      BackColor       =   &H008080FF&
      Height          =   255
      Left            =   15120
      TabIndex        =   40
      Top             =   2520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label aNswer_txt 
      Caption         =   "ans_txt"
      Height          =   255
      Left            =   15600
      TabIndex        =   36
      Top             =   4560
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label ans_num 
      Caption         =   "ans_Num"
      Height          =   255
      Left            =   14400
      TabIndex        =   35
      Top             =   4560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   11115
      Shape           =   4  'Rounded Rectangle
      Top             =   8295
      Width           =   9255
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Explanation : "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   11115
      TabIndex        =   34
      Top             =   7920
      Width           =   1410
   End
   Begin VB.Label qnoMS 
      Caption         =   "qnoMS"
      Height          =   255
      Left            =   17640
      TabIndex        =   32
      Top             =   4560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000006&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   11040
      Shape           =   4  'Rounded Rectangle
      Top             =   9240
      Width           =   9255
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Choices : "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   11220
      TabIndex        =   16
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Question :"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   11160
      TabIndex        =   2
      Top             =   2355
      Width           =   1080
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1695
      Left            =   11115
      Shape           =   4  'Rounded Rectangle
      Top             =   2760
      Width           =   9135
   End
End
Attribute VB_Name = "QuesBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnopt_Click(Index As Integer)
If Trim(opt1(Index).Text) = "" Then
 MsgBox "Enter Option value First in Option box " & vbCrLf & "Then select the correct answer No.", vbInformation + vbOKOnly, "Invalid Answer"
 opt1(Index).SetFocus
 btnopt(Index).Value = vbUnchecked
 Exit Sub
End If
ans_num.Caption = Index + 1
aNswer_txt.Caption = opt1(Index).Text
End Sub

Private Sub combo3_Click()
Combo1.Clear
Set r1 = New ADODB.Recordset
Set r1 = c1.Execute("select tp_nm from topic where sub_id =(select sub_id from sub where upper(sub_nm)='" & UCase(Combo3.Text) & "')and c_id =(select c_id from course where upper(c_nm)='" + UCase(Combo4.Text) + "') ")
While r1.EOF = False
 Combo1.AddItem r1.Fields(0)
 r1.MoveNext
Wend
End Sub

Private Sub Combo4_Click() 'Course
Combo3.Clear
Combo1.Clear
Set r1 = New ADODB.Recordset
Set r1 = c1.Execute("select initcap(sub_nm) from sub where c_id =(select c_id from course where upper(c_nm)='" & UCase(Combo4.Text) & "') ")
While r1.EOF = False
 Combo3.AddItem r1.Fields(0)
 r1.MoveNext
Wend
End Sub

Private Sub Combo7_Click()
Combo8.Clear
Set r1 = New ADODB.Recordset
Set r1 = c1.Execute("select tp_nm from topic where sub_id =(select sub_id from sub where sub_nm='" & Combo7.Text & "')and c_id =(select c_id from course where c_nm='" + Combo6.Text + "') ")
While r1.EOF = False
 Combo8.AddItem r1.Fields(0)
 r1.MoveNext
Wend
End Sub

Private Sub Combo6_Click()
Combo8.Clear
Combo7.Clear
Set r1 = New ADODB.Recordset
Set r1 = c1.Execute("select sub_nm from sub where c_id =(select c_id from course where upper(c_nm)='" & UCase(Combo6.Text) & "') ")
While r1.EOF = False
 Combo7.AddItem r1.Fields(0)
 r1.MoveNext
Wend
End Sub

Private Sub Command1_Click() 'Update
If qnoMS.Caption = "" Then
 MsgBox "Select Questions From Left Side Panel Then Update ...", vbInformation + vbOKOnly, "Not Selected"
 Exit Sub
End If
If Trim(qtext_mcqs.Text) = "" Then
MsgBox "Enter The Question ??? ", vbInformation + vbOKOnly, "Question Empty"
qtext_mcqs.SetFocus
Exit Sub
ElseIf Trim(opt1(0).Text) = "" Or Trim(opt1(1).Text) = "" Or Trim(opt1(2).Text) = "" Or Trim(opt1(3).Text) = "" Then
MsgBox "Enter All The 4 Options ??? ", vbInformation + vbOKOnly, "Fill All Options"
opt1(0).SetFocus
Exit Sub
ElseIf btnopt(0).Value = vbUnchecked And btnopt(1).Value = vbUnchecked And btnopt(2).Value = vbUnchecked And btnopt(3).Value = vbUnchecked Then
MsgBox "Select the Correct Answer !!! ", vbInformation + vbOKOnly, "Select Correct Answer"
opt1(0).SetFocus
Exit Sub
ElseIf Combo6.Text = "" Then
MsgBox "Select Course First ", vbInformation + vbOKOnly, "Course Missing"
Combo6.SetFocus
Exit Sub
ElseIf Combo7.Text = "" Then
MsgBox "Select Subject Correspondent to the Course", vbInformation + vbOKOnly, "Subject Missing"
Combo7.SetFocus
Exit Sub
ElseIf Combo8.Text = "" Then
MsgBox "Select Topic Correspondent to the Subject of the course", vbInformation + vbOKOnly, "Topic Missing"
Combo8.SetFocus
Exit Sub
ElseIf Combo11.Text = "" Then
MsgBox "Select Question Type", vbInformation + vbOKOnly, "Question type Missing"
Combo11.SetFocus
Exit Sub
ElseIf Combo5.Text = "" Then
MsgBox "Choose Difficulty Level of Question", vbInformation + vbOKOnly, "Select Difficulty"
Combo5.SetFocus
Exit Sub
ElseIf Trim(expn_mcq.Text) = "" Then
MsgBox "Please Enter Some Explanation at least Write the exact Answer.. ??? ", vbInformation + vbOKOnly, "Question Explanation"
expn_mcq.SetFocus
Exit Sub
End If
If MsgBox("Are You Sure to Update the question .. ?", vbQuestion + vbYesNo, "Are You Sure") = vbYes Then
updtDA2  'Calling Function To Load data
c1.Execute ("update quesMS set c_id='" & lblcid1.Caption & "',sub_id='" & lblsid1.Caption & "',tp_id='" & lbltpid1.Caption & "',Q_TYP_ID='" & lblqtyp1.Caption & "',q_txt='" & qtext_mcqs.Text & "',opt1='" & opt1(0).Text & "',opt2='" & opt1(1).Text & "',opt3='" & opt1(2).Text & "',opt4='" & opt1(3).Text & "',ANS_TXT='" & aNswer_txt.Caption & "',ANS_NO=" & Val(ans_num.Caption) & ",Q_DIF_LVL='" & Combo5.Text & "',Q_EXPLN='" & expn_mcq.Text & "' where q_id='" & qnoMS.Caption & "' ")
MsgBox "Question SuccessFully Updated...", vbInformation + vbOKOnly, "Update Success"
qtext_mcqs.Text = ""
For i = 0 To 3
 btnopt(i).Value = vbUnchecked
 opt1(i).Text = ""
Next i
expn_mcq.Text = ""
aNswer_txt.Caption = ""
ans_num.Caption = ""
qnoMS.Caption = ""
Combo4.Clear
Combo9.Clear
Combo3.Clear
Combo1.Clear
Combo2.Clear
Combo5.Clear
Combo6.Clear
Combo7.Clear
Combo8.Clear
Combo11.Clear
conn
If rs_course.EOF = False Then
While rs_course.EOF = False
 Combo4.AddItem rs_course(0)
 Combo6.AddItem rs_course(0)
 rs_course.MoveNext
Wend
Else
Exit Sub
End If

Set r = New ADODB.Recordset
Set r = c.Execute("select initcap(Q_TYP_NM) from q_typ")
While r.EOF = False
 Combo9.AddItem r.Fields(0)
 Combo11.AddItem r.Fields(0)
r.MoveNext
Wend

Combo2.AddItem "Easy"
Combo5.AddItem "Easy"
Combo2.AddItem "Medium"
Combo2.AddItem "Hard"
Combo5.AddItem "Medium"
Combo5.AddItem "Hard"
Combo2.AddItem "All"
Combo5.AddItem "All"
loaddata
Else
End If
End Sub

Sub updtDA() 'Important to load before save
Set r4 = New ADODB.Recordset
Set r1 = New ADODB.Recordset
Set r2 = New ADODB.Recordset
Set r3 = New ADODB.Recordset
If Combo9.Text <> "" Then
Set r4 = c1.Execute("SELECT * FROM Q_TYP WHERE" _
 & " upper(Q_TYP_NM)='" & UCase(Combo9.Text) & "'")
If IsNull(r4.Fields(0)) = False Then
lblQtyp.Caption = r4.Fields(0)
Else
End If
End If
If Combo4.Text <> "" Then
 Set r1 = c1.Execute("SELECT c_id FROM COURSE WHERE upper(C_NM)='" & UCase(Combo4.Text) & "'")
If IsNull(r1.Fields(0)) = False Then
lblcid.Caption = r1.Fields(0)
Else
End If
End If
If Combo3.Text <> "" Then
 Set r2 = c1.Execute("SELECT sub_id from sub where upper(sub_nm)='" & UCase(Combo3.Text) & "' AND upper(C_ID)=(SELECT C_ID FROM COURSE WHERE upper(C_NM)='" & UCase(Combo4.Text) & "')")
If IsNull(r2.Fields(0)) = False Then
lblSid.Caption = r2.Fields(0)
Else
End If
End If
If Combo1.Text <> "" Then
 Set r3 = c1.Execute("SELECT tp_id from TOPIC where upper(TP_NM) ='" & UCase(Combo1.Text) & "' AND upper(SUB_ID)=(SELECT SUB_ID FROM SUB WHERE upper(sub_nm)='" & UCase(Combo3.Text) & "' AND C_ID=(SELECT C_ID FROM COURSE WHERE upper(C_NM)='" & UCase(Combo4.Text) & "')) AND C_ID=(SELECT C_ID FROM COURSE WHERE upper(C_NM)='" & UCase(Combo4.Text) & "') ")
If IsNull(r3.Fields(0)) = False Then
lbltpid.Caption = r3.Fields(0)
Else
End If
End If

End Sub

Sub updtDA2() 'Important to load before save
Set r4 = New ADODB.Recordset
Set r1 = New ADODB.Recordset
Set r2 = New ADODB.Recordset
Set r3 = New ADODB.Recordset
Set r4 = c1.Execute("SELECT * FROM Q_TYP WHERE" _
& " upper(Q_TYP_NM)='" & UCase(Combo11.Text) & "'")
Set r1 = c1.Execute("SELECT c_id FROM COURSE WHERE upper(C_NM)='" & UCase(Combo6.Text) & "'")
Set r2 = c1.Execute("SELECT sub_id from sub where upper(sub_nm)='" & UCase(Combo7.Text) & "' AND C_ID=(SELECT C_ID FROM COURSE WHERE upper(C_NM)='" & UCase(Combo6.Text) & "')")
Set r3 = c1.Execute("SELECT tp_id from TOPIC where upper(TP_NM) ='" & UCase(Combo8.Text) & "' AND SUB_ID=(SELECT SUB_ID FROM SUB WHERE upper(sub_nm)='" & UCase(Combo7.Text) & "' AND C_ID=(SELECT C_ID FROM COURSE WHERE upper(C_NM)='" & UCase(Combo6.Text) & "')) AND C_ID=(SELECT C_ID FROM COURSE WHERE upper(C_NM)='" & UCase(Combo6.Text) & "') ")
If IsNull(r4.Fields(0)) = False Then
lblqtyp1.Caption = r4.Fields(0)
Else
End If
If IsNull(r1.Fields(0)) = False Then
lblcid1.Caption = r1.Fields(0)
Else
End If
If IsNull(r2.Fields(0)) = False Then
lblsid1.Caption = r2.Fields(0)
Else
End If
If IsNull(r.Fields(0)) = False Then
lbltpid1.Caption = r3.Fields(0)
Else
End If
End Sub

Private Sub Command2_Click() 'Delete
If Trim(qnoMS.Caption) = "" Then
 MsgBox "Question Not Selected. Select Question From Left side Panel..", vbCritical + vbOKOnly, "Not Select"
 Exit Sub
End If
If MsgBox("Are You Sure to Delete the question .. ?", vbCritical + vbYesNo, "Are You Sure") = vbYes Then
 c1.Execute ("delete from quesMS where q_id='" & qnoMS.Caption & "' ")
 MsgBox "Question SuccessFully Deleted...", vbInformation + vbOKOnly, "Deleted"
qtext_mcqs.Text = ""
For i = 0 To 3
 btnopt(i).Value = vbUnchecked
 opt1(i).Text = ""
Next i
expn_mcq.Text = ""
aNswer_txt.Caption = ""
ans_num.Caption = ""
qnoMS.Caption = ""
Combo4.Clear
Combo9.Clear
Combo3.Clear
Combo1.Clear
Combo2.Clear
Combo5.Clear
Combo6.Clear
Combo7.Clear
Combo8.Clear
Combo11.Clear
conn
If rs_course.EOF = False Then
While rs_course.EOF = False
 Combo4.AddItem rs_course(0)
 Combo6.AddItem rs_course(0)
 rs_course.MoveNext
Wend
Else
Exit Sub
End If

Set r = New ADODB.Recordset
Set r = c.Execute("select initcap(Q_TYP_NM) from q_typ")
While r.EOF = False
 Combo9.AddItem r.Fields(0)
 Combo11.AddItem r.Fields(0)
r.MoveNext
Wend

Combo2.AddItem "Easy"
Combo5.AddItem "Easy"
Combo2.AddItem "Medium"
Combo2.AddItem "Hard"
Combo5.AddItem "Medium"
Combo5.AddItem "Hard"
Combo2.AddItem "All"
Combo5.AddItem "All"
loaddata
Exit Sub
End If

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
If Combo4.Text = "" Then
 MsgBox "Select Course First To  search ", vbInformation + vbOKOnly, "No course Selected"
 Combo4.SetFocus
Exit Sub
End If
If Combo9.Text = "" Then
 MsgBox "Select Question type.. ", vbInformation + vbOKOnly, "No course Selected"
Combo9.SetFocus
Exit Sub
End If
If Combo2.Text = "" Then
 MsgBox "Select Difficulti Leval for Questions.. ", vbInformation + vbOKOnly, "No course Selected"
Combo2.SetFocus
Exit Sub
End If
 updtDA
 Dim r As New ADODB.Recordset
 Dim list As ListItem
 Lvl1.ListItems.Clear
If Combo3.Text = "" And Combo1.Text = "" Then 'No topic  No Subject
 If Combo2.ListIndex <> 3 Then 'Easy/Medium/Hard
  Set r = c1.Execute("select upper(q_id),initcap(q_txt) from quesms where c_id='" & lblcid.Caption & "' and Q_TYP_ID='" & lblQtyp.Caption & "' and upper(Q_DIF_LVL)='" & UCase(Combo2.Text) & "' ")
 ElseIf Combo2.ListIndex = 3 Then 'All
  Set r = c1.Execute("select upper(q_id),initcap(q_txt) from quesms where c_id='" & lblcid.Caption & "' and Q_TYP_ID='" & lblQtyp.Caption & "' ")
 End If
ElseIf Combo1.Text = "" Then 'No topic
 If Combo2.ListIndex <> 3 Then
  Set r = c1.Execute("select upper(q_id),initcap(q_txt) from quesms where sub_id='" & lblSid.Caption & "' and c_id='" & lblcid.Caption & "' and Q_TYP_ID='" & lblQtyp.Caption & "' and upper(Q_DIF_LVL)='" & UCase(Combo2.Text) & "' ")
 ElseIf Combo2.ListIndex = 3 Then
  Set r = c1.Execute("select upper(q_id),initcap(q_txt) from quesms where sub_id='" & lblSid.Caption & "' and c_id='" & lblcid.Caption & "' and Q_TYP_ID='" & lblQtyp.Caption & "' ")
 End If
Else
 If Combo2.ListIndex <> 3 Then
  Set r = c1.Execute("select upper(q_id),initcap(q_txt) from quesms where tp_id='" & lbltpid.Caption & "' and sub_id='" & lblSid.Caption & "' and c_id='" & lblcid.Caption & "' and Q_TYP_ID='" & lblQtyp.Caption & "' and upper(Q_DIF_LVL)='" & UCase(Combo2.Text) & "' ")
 ElseIf Combo2.ListIndex = 3 Then
  Set r = c1.Execute("select upper(q_id),initcap(q_txt) from quesms where tp_id='" & lbltpid.Caption & "' and sub_id='" & lblSid.Caption & "' and c_id='" & lblcid.Caption & "' and Q_TYP_ID='" & lblQtyp.Caption & "' ")
 End If
End If

  While r.EOF = False
   Set list = Lvl1.ListItems.add(, , r.Fields(0))
       list.SubItems(1) = r.Fields(1)
   r.MoveNext
   Wend
End Sub

Private Sub Command5_Click()
qtext_mcqs.Text = ""
For i = 0 To 3
 btnopt(i).Value = vbUnchecked
 opt1(i).Text = ""
Next i
expn_mcq.Text = ""
aNswer_txt.Caption = ""
ans_num.Caption = ""
qnoMS.Caption = ""
Combo4.Clear
Combo9.Clear
Combo3.Clear
Combo1.Clear
Combo2.Clear
Combo5.Clear
Combo6.Clear
Combo7.Clear
Combo8.Clear
Combo11.Clear
conn
If rs_course.EOF = False Then
While rs_course.EOF = False
 Combo4.AddItem rs_course(0)
 Combo6.AddItem rs_course(0)
 rs_course.MoveNext
Wend
Else
Exit Sub
End If

Set r = New ADODB.Recordset
Set r = c.Execute("select initcap(Q_TYP_NM) from q_typ")
While r.EOF = False
 Combo9.AddItem r.Fields(0)
 Combo11.AddItem r.Fields(0)
r.MoveNext
Wend

Combo2.AddItem "Easy"
Combo5.AddItem "Easy"
Combo2.AddItem "Medium"
Combo2.AddItem "Hard"
Combo5.AddItem "Medium"
Combo5.AddItem "Hard"
Combo2.AddItem "All"
Combo5.AddItem "All"
loaddata
End Sub

Private Sub Form_Load()
qtext_mcqs.Text = ""
For i = 0 To 3
 btnopt(i).Value = vbUnchecked
 opt1(i).Text = ""
Next i
expn_mcq.Text = ""
aNswer_txt.Caption = ""
ans_num.Caption = ""
qnoMS.Caption = ""
Combo4.Clear
Combo9.Clear
Combo3.Clear
Combo1.Clear
Combo2.Clear
Combo5.Clear
Combo6.Clear
Combo7.Clear
Combo8.Clear
Combo11.Clear
conn
If rs_course.EOF = False Then
While rs_course.EOF = False
 Combo4.AddItem rs_course(0)
 Combo6.AddItem rs_course(0)
 rs_course.MoveNext
Wend
Else
Exit Sub
End If

Set r = New ADODB.Recordset
Set r = c.Execute("select initcap(Q_TYP_NM) from q_typ")
While r.EOF = False
 Combo9.AddItem r.Fields(0)
 Combo11.AddItem r.Fields(0)
r.MoveNext
Wend

Combo2.AddItem "Easy"
Combo5.AddItem "Easy"
Combo2.AddItem "Medium"
Combo2.AddItem "Hard"
Combo5.AddItem "Medium"
Combo5.AddItem "Hard"
Combo2.AddItem "All"
Combo5.AddItem "All"

With Lvl1.ColumnHeaders
.Clear
.add , "", "Question ID", Width / 14, lvwColumnLeft
.add , "", " Question", Width / 1.74, lvwColumnLeft
End With
loaddata
End Sub

 Public Sub loaddata()
 Dim r As New ADODB.Recordset
 Dim list As ListItem
 Lvl1.ListItems.Clear
 Set r = c1.Execute("select upper(q_id),initcap(q_txt) from quesms")
 While r.EOF = False
  Set list = Lvl1.ListItems.add(, , r.Fields(0))
  list.SubItems(1) = r.Fields(1)
  r.MoveNext
  Wend
 End Sub

Private Sub Lvl1_Click()
On Error Resume Next
Set r = New ADODB.Recordset
Set r = c.Execute("select * from quesms where upper(Q_ID) ='" & UCase(Lvl1.SelectedItem) & "' ")
If r.EOF = False Then
 qnoMS.Caption = r.Fields(0)
 lblcid1.Caption = r.Fields(2)
 lblsid1.Caption = r.Fields(3)
 lbltpid1.Caption = r.Fields(4)
 lblqtyp1.Caption = r.Fields(5)
 qtext_mcqs.Text = r.Fields(6)
 opt1(0).Text = r.Fields(7)
 opt1(1).Text = r.Fields(8)
 opt1(2).Text = r.Fields(9)
 opt1(3).Text = r.Fields(10)
aNswer_txt.Caption = r.Fields(11)
ans_num.Caption = r.Fields(12)
btnopt(r.Fields(12) - 1).Value = True
  Combo5.Text = r.Fields(13)
 If IsNull(r.Fields(14)) = False Then
  expn_mcq.Text = r.Fields(14)
 Else
  expn_mcq.Text = ""
 End If
End If
Set r1 = c.Execute("select c_nm from course where c_id='" & lblcid1.Caption & "' ")
 Combo6.Text = r1.Fields(0)
Combo7.Clear
Set r1 = New ADODB.Recordset
Set r1 = c1.Execute("select sub_nm from sub where c_id =(select c_id from course where upper(c_nm)='" & UCase(Combo6.Text) & "') ")
While r1.EOF = False
 Combo7.AddItem r1.Fields(0)
 r1.MoveNext
Wend
Set r2 = c.Execute("select sub_nm from sub where sub_id='" & lblsid1.Caption & "' ")
 Combo7.Text = r2.Fields(0)
Combo8.Clear
Set r1 = New ADODB.Recordset
Set r1 = c1.Execute("select tp_nm from topic where sub_id =(select sub_id from sub where sub_nm='" & Combo7.Text & "')and c_id =(select c_id from course where c_nm='" + Combo6.Text + "') ")
While r1.EOF = False
 Combo8.AddItem r1.Fields(0)
 r1.MoveNext
Wend
Set r3 = c.Execute("select tp_nm from topic where tp_id='" & lbltpid1.Caption & "' ")
 Combo8.Text = r3.Fields(0)
Set r4 = c.Execute("select Q_TYP_NM from Q_typ where Q_TYP_ID='" & lblqtyp1.Caption & "' ")
 Combo11.Text = r4.Fields(0)
End Sub
