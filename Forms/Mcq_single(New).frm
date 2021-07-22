VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{EAFDAFBF-1D88-41DD-B117-60ECBC4B8441}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form mcq_s 
   BorderStyle     =   1  'Fixed Single
   Caption         =   $"Mcq_single(New).frx":0000
   ClientHeight    =   10050
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   19170
   Icon            =   "Mcq_single(New).frx":00CF
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Mcq_single(New).frx":0459
   ScaleHeight     =   10050
   ScaleWidth      =   19170
   Begin VB.CommandButton vkCommand4 
      Height          =   450
      Left            =   17760
      MouseIcon       =   "Mcq_single(New).frx":AED7
      MousePointer    =   99  'Custom
      Picture         =   "Mcq_single(New).frx":B029
      Style           =   1  'Graphical
      TabIndex        =   91
      Top             =   9480
      Width           =   1250
   End
   Begin VB.CommandButton NewQs 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4200
      MouseIcon       =   "Mcq_single(New).frx":B75C
      MousePointer    =   99  'Custom
      Picture         =   "Mcq_single(New).frx":B8AE
      Style           =   1  'Graphical
      TabIndex        =   90
      ToolTipText     =   "Save and Move to Next Question"
      Top             =   9480
      Width           =   1965
   End
   Begin VB.CommandButton pquesBTN 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   240
      MouseIcon       =   "Mcq_single(New).frx":C30F
      MousePointer    =   99  'Custom
      Picture         =   "Mcq_single(New).frx":C461
      Style           =   1  'Graphical
      TabIndex        =   89
      ToolTipText     =   "Move to Previous Question"
      Top             =   9480
      Width           =   1845
   End
   Begin VB.CommandButton nquesBTN 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2205
      MouseIcon       =   "Mcq_single(New).frx":CCFF
      MousePointer    =   99  'Custom
      Picture         =   "Mcq_single(New).frx":CE51
      Style           =   1  'Graphical
      TabIndex        =   88
      ToolTipText     =   "Move to Next Question"
      Top             =   9480
      Width           =   1845
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
      Left            =   120
      TabIndex        =   3
      Top             =   4440
      Width           =   18945
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   2415
         Left            =   720
         TabIndex        =   4
         Top             =   480
         Width           =   18210
         Begin vkUserContolsXP.vkOptionButton btnOpt 
            Height          =   345
            Index           =   0
            Left            =   360
            TabIndex        =   5
            Top             =   120
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   609
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   1
            Alignment       =   2
         End
         Begin RichTextLib.RichTextBox opt1 
            Height          =   345
            Index           =   0
            Left            =   1080
            TabIndex        =   6
            Top             =   135
            Width           =   17055
            _ExtentX        =   30083
            _ExtentY        =   609
            _Version        =   393217
            BackColor       =   14737632
            BorderStyle     =   0
            MaxLength       =   200
            Appearance      =   0
            TextRTF         =   $"Mcq_single(New).frx":D6EC
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
            Height          =   345
            Index           =   1
            Left            =   1080
            TabIndex        =   7
            Top             =   735
            Width           =   17055
            _ExtentX        =   30083
            _ExtentY        =   609
            _Version        =   393217
            BackColor       =   14737632
            BorderStyle     =   0
            MaxLength       =   200
            Appearance      =   0
            TextRTF         =   $"Mcq_single(New).frx":D818
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
            Height          =   345
            Index           =   2
            Left            =   1080
            TabIndex        =   8
            Top             =   1335
            Width           =   17055
            _ExtentX        =   30083
            _ExtentY        =   609
            _Version        =   393217
            BackColor       =   14737632
            BorderStyle     =   0
            MaxLength       =   200
            Appearance      =   0
            TextRTF         =   $"Mcq_single(New).frx":D944
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
            Height          =   345
            Index           =   3
            Left            =   1080
            TabIndex        =   9
            Top             =   1935
            Width           =   17055
            _ExtentX        =   30083
            _ExtentY        =   609
            _Version        =   393217
            BackColor       =   14737632
            BorderStyle     =   0
            MaxLength       =   200
            Appearance      =   0
            TextRTF         =   $"Mcq_single(New).frx":DA70
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
         Begin vkUserContolsXP.vkOptionButton btnOpt 
            Height          =   345
            Index           =   1
            Left            =   360
            TabIndex        =   10
            Top             =   720
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   609
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   1
            Alignment       =   2
         End
         Begin vkUserContolsXP.vkOptionButton btnOpt 
            Height          =   345
            Index           =   3
            Left            =   360
            TabIndex        =   11
            Top             =   1920
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   609
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   1
            Alignment       =   2
         End
         Begin vkUserContolsXP.vkOptionButton btnOpt 
            Height          =   345
            Index           =   2
            Left            =   360
            TabIndex        =   12
            Top             =   1335
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   609
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   1
            Alignment       =   2
         End
         Begin VB.Line Line8 
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   2400
         End
         Begin VB.Line Line14 
            X1              =   0
            X2              =   19000
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Line Line13 
            X1              =   0
            X2              =   19000
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Line Line12 
            X1              =   0
            X2              =   19000
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Line Line9 
            X1              =   0
            X2              =   19000
            Y1              =   2400
            Y2              =   2400
         End
         Begin VB.Line Line11 
            X1              =   0
            X2              =   19000
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
      Begin VB.Label Label4 
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
         TabIndex        =   19
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
      Begin VB.Line Line7 
         X1              =   0
         X2              =   720
         Y1              =   1685
         Y2              =   1685
      End
      Begin VB.Line Line6 
         X1              =   0
         X2              =   720
         Y1              =   1085
         Y2              =   1085
      End
      Begin VB.Label Label9 
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
         TabIndex        =   18
         Top             =   2325
         Width           =   180
      End
      Begin VB.Label Label8 
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
         TabIndex        =   17
         Top             =   1800
         Width           =   165
      End
      Begin VB.Line Line1 
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
         X2              =   19000
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
         X1              =   18930
         X2              =   18930
         Y1              =   0
         Y2              =   2880
      End
      Begin VB.Label Label2 
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
         TabIndex        =   16
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label5 
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
         TabIndex        =   15
         Top             =   600
         Width           =   180
      End
      Begin VB.Label Label6 
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
         TabIndex        =   14
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
         TabIndex        =   13
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
   Begin VB.TextBox QID 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   15945
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   2
      Text            =   "MS0005"
      Top             =   1950
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   15360
      Top             =   7440
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   17640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3945
      Width           =   1455
   End
   Begin vkUserContolsXP.vkCheck vkCheck1 
      Height          =   255
      Left            =   1920
      TabIndex        =   0
      Top             =   7710
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      BackColor       =   8421504
      Caption         =   "Same as Answer"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin MSAdodcLib.Adodc adoInterface 
      Height          =   330
      Left            =   13440
      Top             =   7560
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      RecordSource    =   "select * from quesMS"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cdb 
      Left            =   14760
      Top             =   7440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin vkUserContolsXP.vkCommand pic 
      Height          =   1455
      Left            =   17640
      TabIndex        =   20
      Top             =   2400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2566
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Mcq_single(New).frx":DB9C
   End
   Begin vkUserContolsXP.vkFrame vkFrame6 
      Height          =   1095
      Left            =   5910
      TabIndex        =   21
      Top             =   315
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1931
      BackColor1      =   14737632
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowBackGround  =   0   'False
      ShowTitle       =   0   'False
      TextPosition    =   0
      BorderColor     =   16761024
      PictureAlignment=   1
      DisplayPicture  =   0   'False
      Begin vkUserContolsXP.vkCommand img 
         Height          =   895
         Left            =   150
         TabIndex        =   22
         Top             =   120
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   1588
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   12632256
         Picture         =   "Mcq_single(New).frx":E47F
         CustomStyle     =   0
      End
   End
   Begin vkUserContolsXP.vkFrame vkFrame4 
      Height          =   1095
      Left            =   4200
      TabIndex        =   23
      Top             =   315
      Width           =   1665
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowBackGround  =   0   'False
      ShowTitle       =   0   'False
      TextPosition    =   0
      BorderColor     =   16761024
      Begin vkUserContolsXP.vkCommand align 
         Height          =   375
         Index           =   1
         Left            =   630
         TabIndex        =   28
         Top             =   120
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   12632256
         Picture         =   "Mcq_single(New).frx":EBC0
         CustomStyle     =   0
      End
      Begin vkUserContolsXP.vkCommand align 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   120
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   12632256
         Picture         =   "Mcq_single(New).frx":F02C
         CustomStyle     =   0
      End
      Begin vkUserContolsXP.vkCommand align 
         Height          =   375
         Index           =   2
         Left            =   1120
         TabIndex        =   26
         Top             =   120
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   12632256
         Picture         =   "Mcq_single(New).frx":F4B3
         CustomStyle     =   0
      End
      Begin vkUserContolsXP.vkToggleButton vkToggleButton2 
         Height          =   375
         Left            =   1120
         TabIndex        =   25
         Top             =   600
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   "a"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         BorderColor     =   12632256
         CustomStyle     =   0
      End
      Begin vkUserContolsXP.vkToggleButton vkToggleButton3 
         Height          =   375
         Left            =   630
         TabIndex        =   24
         Top             =   600
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   "A"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         BorderColor     =   12632256
         CustomStyle     =   0
      End
   End
   Begin vkUserContolsXP.vkFrame vkFrame3 
      Height          =   1095
      Left            =   1155
      TabIndex        =   29
      Top             =   315
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1931
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowBackGround  =   0   'False
      ShowTitle       =   0   'False
      TextPosition    =   0
      BorderColor     =   16761024
      Begin vkUserContolsXP.vkCommand fnt_colr 
         Height          =   375
         Left            =   2250
         TabIndex        =   30
         Top             =   600
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   12632256
         Picture         =   "Mcq_single(New).frx":F90D
         CustomStyle     =   0
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   600
         Width           =   495
      End
      Begin vkUserContolsXP.vkToggleButton undrln_btn 
         Height          =   345
         Left            =   1080
         TabIndex        =   36
         Top             =   615
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   609
         Caption         =   "U"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         BorderColor     =   12632256
         CustomStyle     =   0
      End
      Begin vkUserContolsXP.vkToggleButton itlc_btn 
         Height          =   345
         Left            =   600
         TabIndex        =   35
         Top             =   600
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   609
         Caption         =   "I"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         BorderColor     =   12632256
         CustomStyle     =   0
      End
      Begin vkUserContolsXP.vkToggleButton bold_btn 
         Height          =   345
         Left            =   120
         TabIndex        =   34
         Top             =   615
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   609
         Caption         =   "B"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         BorderColor     =   12632256
         CustomStyle     =   0
      End
      Begin VB.ComboBox fnt_size 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2160
         TabIndex        =   33
         Text            =   "10"
         Top             =   120
         Width           =   735
      End
      Begin VB.ComboBox fnt_styl 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Mcq_single(New).frx":FDBE
         Left            =   120
         List            =   "Mcq_single(New).frx":FE28
         TabIndex        =   32
         Text            =   "Calibiri"
         Top             =   120
         Width           =   1815
      End
      Begin vkUserContolsXP.vkToggleButton vkToggleButton1 
         Height          =   345
         Left            =   1560
         TabIndex        =   31
         Top             =   615
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   609
         Caption         =   "U"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   -1  'True
         EndProperty
         ForeColor       =   0
         BorderColor     =   12632256
         CustomStyle     =   0
      End
   End
   Begin RichTextLib.RichTextBox qtext_mcqs 
      Height          =   1200
      Left            =   240
      TabIndex        =   38
      Top             =   2520
      Width           =   17055
      _ExtentX        =   30083
      _ExtentY        =   2117
      _Version        =   393217
      BorderStyle     =   0
      Appearance      =   0
      TextRTF         =   $"Mcq_single(New).frx":FFFC
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
   Begin vkUserContolsXP.vkFrame vkFrame1 
      Height          =   1800
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Width           =   19290
      _ExtentX        =   34025
      _ExtentY        =   3175
      BackColor1      =   14737632
      BackColor2      =   16777215
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleColor1     =   16761024
      TitleGradient   =   2
      TitleHeight     =   300
      BorderColor     =   16761024
      BorderWidth     =   0
      Begin vkUserContolsXP.vkFrame vkFrame13 
         Height          =   345
         Left            =   18060
         TabIndex        =   68
         Top             =   1425
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   609
         Caption         =   "Help"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210752
         ShowBackGround  =   0   'False
         TitleColor1     =   16744576
         TitleColor2     =   12632256
         TitleGradient   =   2
         TitleHeight     =   340
         RoundAngle      =   4
      End
      Begin vkUserContolsXP.vkFrame vkFrame12 
         Height          =   1095
         Left            =   18045
         TabIndex        =   66
         Top             =   315
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1931
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowBackGround  =   0   'False
         ShowTitle       =   0   'False
         TextPosition    =   0
         BorderColor     =   16761024
         Begin vkUserContolsXP.vkCommand vkCommand7 
            Height          =   750
            Left            =   150
            TabIndex        =   67
            Top             =   180
            Width           =   760
            _ExtentX        =   1349
            _ExtentY        =   1323
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   12632256
            Picture         =   "Mcq_single(New).frx":10078
            PictureOffsetX  =   360
            PictureOffsetY  =   -200
            CustomStyle     =   0
         End
      End
      Begin vkUserContolsXP.vkFrame vkFrame11 
         Height          =   345
         Left            =   6795
         TabIndex        =   65
         Top             =   1425
         Width           =   11220
         _ExtentX        =   19791
         _ExtentY        =   609
         Caption         =   "Question info"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210752
         ShowBackGround  =   0   'False
         TitleColor1     =   16744576
         TitleColor2     =   12632256
         TitleGradient   =   2
         TitleHeight     =   340
         RoundAngle      =   4
      End
      Begin vkUserContolsXP.vkFrame vkFrame10 
         Height          =   345
         Left            =   5910
         TabIndex        =   64
         Top             =   1425
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   609
         Caption         =   "Picture"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210752
         ShowBackGround  =   0   'False
         TitleColor1     =   16744576
         TitleColor2     =   12632256
         TitleGradient   =   2
         TitleHeight     =   340
         RoundAngle      =   4
      End
      Begin vkUserContolsXP.vkFrame vkFrame9 
         Height          =   345
         Left            =   4200
         TabIndex        =   63
         Top             =   1425
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   609
         Caption         =   "Format"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210752
         ShowBackGround  =   0   'False
         TitleColor1     =   16744576
         TitleColor2     =   12632256
         TitleGradient   =   2
         TitleHeight     =   340
         RoundAngle      =   4
      End
      Begin vkUserContolsXP.vkFrame vkFrame7 
         Height          =   345
         Left            =   1150
         TabIndex        =   62
         Top             =   1425
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   609
         Caption         =   "Font"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210752
         ShowBackGround  =   0   'False
         TitleColor1     =   16744576
         TitleColor2     =   12632256
         TitleGradient   =   2
         TitleHeight     =   340
         RoundAngle      =   4
      End
      Begin vkUserContolsXP.vkFrame vkFrame5 
         Height          =   345
         Left            =   40
         TabIndex        =   61
         Top             =   1425
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "Clipboard"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210752
         ShowBackGround  =   0   'False
         TitleColor1     =   16744576
         TitleColor2     =   12632256
         TitleGradient   =   2
         TitleHeight     =   340
         RoundAngle      =   4
      End
      Begin vkUserContolsXP.vkFrame vkFrame2 
         Height          =   1095
         Left            =   40
         TabIndex        =   57
         Top             =   315
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1931
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowBackGround  =   0   'False
         ShowTitle       =   0   'False
         TextPosition    =   0
         BorderColor     =   16761024
         Begin VB.Image Image3 
            Height          =   195
            Left            =   220
            MouseIcon       =   "Mcq_single(New).frx":10BCD
            MousePointer    =   99  'Custom
            Picture         =   "Mcq_single(New).frx":11497
            Stretch         =   -1  'True
            Top             =   735
            Width           =   195
         End
         Begin VB.Image Image2 
            Height          =   180
            Left            =   240
            MouseIcon       =   "Mcq_single(New).frx":118D6
            MousePointer    =   99  'Custom
            Picture         =   "Mcq_single(New).frx":121A0
            Stretch         =   -1  'True
            Top             =   450
            Width           =   180
         End
         Begin VB.Image Image1 
            Height          =   180
            Left            =   240
            MouseIcon       =   "Mcq_single(New).frx":125D6
            MousePointer    =   99  'Custom
            Picture         =   "Mcq_single(New).frx":12EA0
            Stretch         =   -1  'True
            Top             =   130
            Width           =   180
         End
         Begin VB.Label cutMCQs 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cut"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   480
            MouseIcon       =   "Mcq_single(New).frx":13252
            MousePointer    =   99  'Custom
            TabIndex        =   60
            Top             =   120
            Width           =   345
         End
         Begin VB.Label copyMCQs 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Copy"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   480
            MouseIcon       =   "Mcq_single(New).frx":13B1C
            MousePointer    =   99  'Custom
            TabIndex        =   59
            Top             =   435
            Width           =   465
         End
         Begin VB.Label pasteMCQs 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Paste"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   480
            MouseIcon       =   "Mcq_single(New).frx":143E6
            MousePointer    =   99  'Custom
            TabIndex        =   58
            Top             =   735
            Width           =   465
         End
      End
      Begin vkUserContolsXP.vkFrame vkFrame17 
         Height          =   1095
         Left            =   6795
         TabIndex        =   40
         Top             =   315
         Width           =   11205
         _ExtentX        =   19764
         _ExtentY        =   1931
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowBackGround  =   0   'False
         ShowTitle       =   0   'False
         TextPosition    =   0
         BorderColor     =   16761024
         Begin VB.ComboBox Combo7 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1080
            MouseIcon       =   "Mcq_single(New).frx":14CB0
            MousePointer    =   99  'Custom
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   180
            Width           =   2775
         End
         Begin VB.ComboBox Combo6 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   5145
            MouseIcon       =   "Mcq_single(New).frx":14E02
            MousePointer    =   99  'Custom
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Top             =   180
            Width           =   3015
         End
         Begin VB.ComboBox Combo2 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1080
            MouseIcon       =   "Mcq_single(New).frx":14F54
            MousePointer    =   99  'Custom
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   645
            Width           =   3735
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   6120
            MouseIcon       =   "Mcq_single(New).frx":150A6
            MousePointer    =   99  'Custom
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   660
            Width           =   2055
         End
         Begin VB.ComboBox Combo3 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   9000
            MouseIcon       =   "Mcq_single(New).frx":151F8
            MousePointer    =   99  'Custom
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   180
            Width           =   2055
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Remember Me"
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
            Left            =   8880
            MouseIcon       =   "Mcq_single(New).frx":1534A
            MousePointer    =   99  'Custom
            TabIndex        =   41
            Top             =   720
            Width           =   3735
         End
         Begin VB.Label Label18 
            Caption         =   "COURSE"
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
            Left            =   120
            TabIndex        =   56
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label17 
            Caption         =   "TOPIC "
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
            Left            =   120
            TabIndex        =   55
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label16 
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
            Left            =   4200
            TabIndex        =   54
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   270
            Left            =   960
            TabIndex        =   53
            Top             =   600
            Width           =   105
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   270
            Left            =   960
            TabIndex        =   52
            Top             =   120
            Width           =   105
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   270
            Left            =   4995
            TabIndex        =   51
            Top             =   165
            Width           =   105
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   270
            Left            =   8880
            TabIndex        =   50
            Top             =   120
            Width           =   105
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LEVAL"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   5400
            TabIndex        =   49
            Top             =   720
            Width           =   555
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   270
            Left            =   6000
            TabIndex        =   48
            Top             =   600
            Width           =   105
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TYPE"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   8400
            TabIndex        =   47
            Top             =   240
            Width           =   435
         End
      End
      Begin VB.Line Line21 
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1800
      End
   End
   Begin RichTextLib.RichTextBox expn_mcq 
      Height          =   840
      Left            =   240
      TabIndex        =   69
      Top             =   8085
      Width           =   18735
      _ExtentX        =   33046
      _ExtentY        =   1482
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      Appearance      =   0
      TextRTF         =   $"Mcq_single(New).frx":1549C
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
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Height          =   735
      Left            =   0
      TabIndex        =   87
      Top             =   9330
      Width           =   19215
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   3840
      Top             =   9480
      Width           =   615
   End
   Begin VB.Label qnoMS 
      Alignment       =   2  'Center
      Caption         =   "QNO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   86
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label ans_num 
      Caption         =   "ansNO"
      Height          =   375
      Left            =   4200
      TabIndex        =   85
      Top             =   6840
      Width           =   495
   End
   Begin VB.Label answer_txt 
      Caption         =   "ANSWER_text"
      Height          =   375
      Left            =   4800
      TabIndex        =   84
      Top             =   6840
      Width           =   4935
   End
   Begin VB.Label lblSID 
      Height          =   375
      Left            =   13800
      TabIndex        =   83
      Top             =   6840
      Width           =   855
   End
   Begin VB.Label lblCID 
      Height          =   375
      Left            =   12840
      TabIndex        =   82
      Top             =   6840
      Width           =   855
   End
   Begin VB.Label lbltpid 
      Height          =   375
      Left            =   14760
      TabIndex        =   81
      Top             =   6840
      Width           =   855
   End
   Begin VB.Label lblQTYP 
      Height          =   375
      Left            =   11880
      TabIndex        =   80
      Top             =   6840
      Width           =   855
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BorderColor     =   &H00E0E0E0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   15900
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   2400
      Width           =   17295
   End
   Begin VB.Line Line22 
      X1              =   0
      X2              =   13560
      Y1              =   -360
      Y2              =   -360
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the Question :"
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
      Index           =   0
      Left            =   180
      TabIndex        =   79
      Top             =   1995
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the Choices : "
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
      Left            =   180
      TabIndex        =   78
      Top             =   4065
      Width           =   1995
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Question ID"
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
      Index           =   1
      Left            =   14460
      TabIndex        =   77
      Top             =   1950
      Width           =   1215
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   8040
      Width           =   18975
   End
   Begin VB.Label Label12 
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
      Left            =   195
      TabIndex        =   76
      Top             =   7665
      Width           =   1410
   End
   Begin VB.Line Line20 
      BorderWidth     =   2
      X1              =   0
      X2              =   0
      Y1              =   -360
      Y2              =   10640
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      X1              =   0
      X2              =   19150
      Y1              =   10095
      Y2              =   10095
   End
   Begin VB.Image Qimg 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Left            =   17640
      Stretch         =   -1  'True
      Top             =   2400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label23 
      Height          =   375
      Left            =   3120
      TabIndex        =   75
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image6 
      Height          =   345
      Left            =   15555
      MouseIcon       =   "Mcq_single(New).frx":15518
      MousePointer    =   99  'Custom
      Picture         =   "Mcq_single(New).frx":1566A
      Stretch         =   -1  'True
      Top             =   -345
      Width           =   360
   End
   Begin VB.Line Line18 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      X1              =   0
      X2              =   19150
      Y1              =   9315
      Y2              =   9315
   End
   Begin VB.Label Tpic 
      Height          =   375
      Left            =   7080
      TabIndex        =   74
      Top             =   1920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Subj 
      Height          =   375
      Left            =   5640
      TabIndex        =   73
      Top             =   1920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Course 
      Height          =   375
      Left            =   4080
      TabIndex        =   72
      Top             =   1920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label QuTyp 
      Height          =   375
      Left            =   8760
      TabIndex        =   71
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label QLevel 
      Height          =   375
      Left            =   10080
      TabIndex        =   70
      Top             =   1920
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "mcq_s"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CHKSelRTF As Byte
Dim imgpath As String, tempvr As String, tempvr2 As String
Dim pic_name As String, pic_ext As String, pic_changed As Boolean
Dim a As Integer

Private Sub align_Click(Index As Integer) 'Alignment
If qtext_mcqs.SelLength > 0 Then
qtext_mcqs.Find Trim(qtext_mcqs.Text)
qtext_mcqs.SelAlignment = Index
End If

If expn_mcq.SelLength > 0 Then
expn_mcq.Find Trim(expn_mcq.TextRTF)
expn_mcq.SelAlignment = Index
End If
If opt1(0).SelLength > 0 Then
opt1(0).Find Trim(opt1(0).TextRTF)
opt1(0).SelAlignment = Index
End If

If opt1(1).SelLength > 0 Then
opt1(1).Find Trim(opt1(1).TextRTF)
opt1(1).SelAlignment = Index
End If

If opt1(2).SelLength > 0 Then
opt1(2).Find Trim(opt1(2).TextRTF)
opt1(2).SelAlignment = Index
End If

If opt1(3).SelLength > 0 Then
opt1(3).Find Trim(opt1(3).TextRTF)
opt1(3).SelAlignment = Index
End If
End Sub

Private Sub bold_btn_Click()
qtext_mcqs.SelBold = Not qtext_mcqs.SelBold
opt1(0).SelBold = Not opt1(0).SelBold
opt1(1).SelBold = Not opt1(1).SelBold
opt1(2).SelBold = Not opt1(2).SelBold
opt1(3).SelBold = Not opt1(3).SelBold
expn_mcq.SelBold = Not expn_mcq.SelBold
End Sub

Private Sub btnOpt_MouseDblClick(Index As Integer, Button As MouseButtonConstants, Shift As Integer, Control As Integer, X As Long, Y As Long)
If Trim(opt1(Index).Text) = "" Then
 MsgBox "Enter Option value First in Option box " & vbCrLf & "Then select the correct answer No.", vbInformation + vbOKOnly, "Invalid Answer"
 opt1(Index).SetFocus
 btnopt(Index).Value = vbUnchecked
 Exit Sub
End If
ans_num.Caption = Index + 1
aNswer_txt.Caption = opt1(Index).Text
End Sub

Private Sub btnOpt_MouseDown(Index As Integer, Button As MouseButtonConstants, Shift As Integer, Control As Integer, X As Long, Y As Long)
If Trim(opt1(Index).Text) = "" Then
 MsgBox "Enter Option value First in Option box " & vbCrLf & "Then select the correct answer No.", vbInformation + vbOKOnly, "Invalid Answer"
 opt1(Index).SetFocus
 btnopt(Index).Value = vbUnchecked
 Exit Sub
End If
If Button = vbLeftButton Then
ans_num.Caption = Index + 1
aNswer_txt.Caption = opt1(Index).Text
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
qtext_mcqs.SetFocus
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
Combo3.SetFocus
End If
End Sub
Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
Combo1.SetFocus
End If
End Sub

Private Sub Combo4_Click()
fnt_colr_Click
End Sub
Private Sub Combo4_GotFocus()
fnt_colr_Click
End Sub

Private Sub Combo6_Click() 'Subject Combo Click
Combo2.Clear
Set r1 = New ADODB.Recordset
Set r1 = c1.Execute("select TP_NM from topic where sub_id =(select sub_id from sub where sub_nm='" & Combo6.Text & "' and c_id=(select c_id from course where c_nm='" & Combo7.Text & "'))and c_id =(select c_id from course where c_nm='" & Combo7.Text & "') ")
If IsNull(r1.Fields(0)) = False Then
 While r1.EOF = False
  Combo2.AddItem r1.Fields(0)
  r1.MoveNext
 Wend
Else
End If
End Sub

Private Sub Combo6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
Combo2.SetFocus
End If
End Sub

Private Sub Combo7_Click() 'Course Click
Combo6.Clear
Combo2.Clear
Set r1 = New ADODB.Recordset
Set r1 = c1.Execute("select sub_nm from sub where c_id =(select c_id from course where upper(c_nm)='" & UCase(Combo7.Text) & "') ")
If IsNull(r1.Fields(0)) = False Then
While r1.EOF = False
 Combo6.AddItem r1.Fields(0)
 r1.MoveNext
Wend
Else
End If
End Sub

Private Sub Combo7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
Combo6.SetFocus
End If
End Sub

Private Sub Command1_Click() 'Clear
On Error Resume Next
    Set Qimg.Picture = Nothing
    pic_name = ""
    pic_changed = True
End Sub

Private Sub copyMCQs_Click()
Clipboard.Clear
If qtext_mcqs.SelLength > 0 Then
Clipboard.SetText qtext_mcqs.SelText
Exit Sub
End If

If expn_mcq.SelLength > 0 Then
Clipboard.SetText expn_mcq.SelText
Exit Sub
End If

If opt1(0).SelLength > 0 Then
Clipboard.SetText opt1(0).SelText
Exit Sub
End If

If opt1(1).SelLength > 0 Then
Clipboard.SetText opt1(1).SelText
Exit Sub
End If

If opt1(2).SelLength > 0 Then
Clipboard.SetText opt1(2).SelText
Exit Sub
End If

If opt1(3).SelLength > 0 Then
Clipboard.SetText opt1(3).SelText
Exit Sub
End If

End Sub

Private Sub copyMCQs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
copyMCQs.FontUnderline = True
cutMCQs.FontUnderline = False
pasteMCQs.FontUnderline = False
End Sub

Private Sub cutMCQs_Click() 'Cut Method
 Clipboard.Clear
If qtext_mcqs.SelLength > 0 Then
Clipboard.SetText (qtext_mcqs.SelText)
qtext_mcqs.SelText = ""
Exit Sub
End If

If expn_mcq.SelLength > 0 Then
Clipboard.SetText expn_mcq.SelText
expn_mcq.SelText = ""
Exit Sub
End If

If opt1(0).SelLength > 0 Then
Clipboard.SetText opt1(0).SelText
opt1(0).SelText = ""
Exit Sub
End If

If opt1(1).SelLength > 0 Then
Clipboard.SetText opt1(1).SelText
opt1(1).SelText = ""
Exit Sub
End If

If opt1(2).SelLength > 0 Then
Clipboard.SetText opt1(2).SelText
opt1(2).SelText = ""
Exit Sub
End If

If opt1(3).SelLength > 0 Then
Clipboard.SetText opt1(3).SelText
opt1(3).SelText = ""
Exit Sub
End If
End Sub

Private Sub cutMCQs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cutMCQs.FontUnderline = True
copyMCQs.FontUnderline = False
pasteMCQs.FontUnderline = False
End Sub
Private Sub expn_mcq_GotFocus()
CHKSelRTF = 6
End Sub

Private Sub fnt_colr_Click()
cdb.ShowColor
If qtext_mcqs.SelLength > 0 Then
qtext_mcqs.SelColor = cdb.Color
End If

If expn_mcq.SelLength > 0 Then
expn_mcq.SelColor = cdb.Color
End If

If opt1(0).SelLength > 0 Then
opt1(0).SelColor = cdb.Color
End If
If opt1(1).SelLength > 0 Then
opt1(1).SelColor = cdb.Color
End If
If opt1(2).SelLength > 0 Then
opt1(2).SelColor = cdb.Color
End If
If opt1(3).SelLength > 0 Then
opt1(3).SelColor = cdb.Color
End If
End Sub

Private Sub fnt_size_Click()
If qtext_mcqs.SelLength > 0 Then
qtext_mcqs.SelFontSize = fnt_size.Text
End If

If expn_mcq.SelLength > 0 Then
expn_mcq.SelFontSize = fnt_size.Text
End If

If opt1(0).SelLength > 0 Then
opt1(0).SelFontName = fnt_size.Text
End If
If opt1(1).SelLength > 0 Then
opt1(1).SelFontName = fnt_size.Text
End If
If opt1(2).SelLength > 0 Then
opt1(2).SelFontName = fnt_size.Text
End If
If opt1(3).SelLength > 0 Then
opt1(3).SelFontName = fnt_size.Text
End If
End Sub

Private Sub fnt_styl_Click()
If qtext_mcqs.SelLength > 0 Then
qtext_mcqs.SelFontName = fnt_styl.Text
End If

If expn_mcq.SelLength > 0 Then
expn_mcq.SelFontName = fnt_styl.Text
End If

If opt1(0).SelLength > 0 Then
opt1(0).SelFontName = fnt_styl.Text
End If
If opt1(1).SelLength > 0 Then
opt1(1).SelFontName = fnt_styl.Text
End If
If opt1(2).SelLength > 0 Then
opt1(2).SelFontName = fnt_styl.Text
End If
If opt1(3).SelLength > 0 Then
opt1(3).SelFontName = fnt_styl.Text
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
conn
Command1_Click
Me.Top = 50 '100
Me.Left = 1300 '4400
Combo7.Clear
Combo6.Clear
Combo2.Clear
CHKSelRTF = 0
Combo3.Clear
If rs_qtyp.EOF = False Then
rs_qtyp.MoveFirst
End If
If IsNull(rs_qtyp.EOF) = False Then
While rs_qtyp.EOF = False
 Combo3.AddItem rs_qtyp(1)
 rs_qtyp.MoveNext
Wend
Else
End If
If rs_course.EOF = False Then
 rs_course.MoveFirst
End If
If IsNull(rs_course.EOF) = False Then
While rs_course.EOF = False
 Combo7.AddItem rs_course(0)
 rs_course.MoveNext
Wend
Else
End If
Combo1.Clear
Combo1.AddItem "EASY"
Combo1.AddItem "MEDIUM"
Combo1.AddItem "HARD"

nquesBTN.Enabled = True
pquesBTN.Enabled = True

For i = 5 To 12 Step 1
fnt_size.AddItem i
Next

For i = 13 To 36 Step 2
fnt_size.AddItem i
Next

'Clear  All Filled Text
qtext_mcqs.Text = ""
For i = 0 To 3
 btnopt(i).Value = vbUnchecked
 opt1(i).Text = ""
Next i
expn_mcq.Text = ""
aNswer_txt.Caption = ""
ans_num.Caption = ""
qnoMS.Caption = ""
pic_name = ""
vkCheck1.Value = vbUnchecked 'Same as answer

q_autoID 'New ID GENERATED
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cutMCQs.FontUnderline = False
copyMCQs.FontUnderline = False
pasteMCQs.FontUnderline = False
End Sub

Private Sub Form_Unload(cancel As Integer)
ques_entry_dash.Enabled = True
End Sub

Private Sub Image1_Click()
cutMCQs_Click
End Sub

Private Sub Image2_Click()
copyMCQs_Click
End Sub

Private Sub Image3_Click()
pasteMCQs_Click
End Sub

Private Sub Image5_Click()
nquesBTN_Click
End Sub

Private Sub Image6_Click() 'Close
Unload Me
End Sub

Private Sub img_Click()
On Error Resume Next
   cdb.Filter = "All Picture Files *.jpg,.gif,.bmp,.ico JPEG Image"
   cdb.ShowOpen
   If cdb.FileName <> "" Then
      Qimg.Visible = True
      pic.Visible = False
      Qimg.Picture = LoadPicture(cdb.FileName)
        pic_name = cdb.FileName
        pic_ext = Right(cbd.FileTitle, 4)
        pic_changed = True
     End If
End Sub

Private Sub itlc_btn_Click()
qtext_mcqs.SelItalic = Not qtext_mcqs.SelItalic
opt1(0).SelItalic = Not opt1(0).SelItalic
opt1(1).SelItalic = Not opt1(1).SelItalic
opt1(2).SelItalic = Not opt1(2).SelItalic
opt1(3).SelItalic = Not opt1(3).SelItalic
expn_mcq.SelItalic = Not expn_mcq.SelItalic
End Sub

Private Sub newQues_Click()
q_autoID
End Sub

Sub updtDA() 'Important to load before save
'On Error Resume Next
Set r4 = New ADODB.Recordset
Set r1 = New ADODB.Recordset
Set r2 = New ADODB.Recordset
Set r3 = New ADODB.Recordset
Set r4 = c1.Execute("SELECT * FROM Q_TYP WHERE" _
& " Q_TYP_NM='" & Combo3.Text & "'")
Set r1 = c1.Execute("SELECT * FROM COURSE WHERE C_NM='" & Combo7.Text & "'")
Set r2 = c1.Execute("SELECT * from sub where sub_nm='" & Combo6.Text & "' AND C_ID=(SELECT C_ID FROM COURSE WHERE C_NM='" & Combo7.Text & "')")
Set r3 = c1.Execute("SELECT * from TOPIC where TP_NM ='" & Combo2.Text & "' AND SUB_ID=(SELECT SUB_ID FROM SUB WHERE sub_nm='" & Combo6.Text & "' AND C_ID=(SELECT C_ID FROM COURSE WHERE C_NM='" & Combo7.Text & "')) AND C_ID=(SELECT C_ID FROM COURSE WHERE C_NM='" & Combo7.Text & "') ")
If IsNull(r4.Fields(0)) = False Then
lblQtyp.Caption = r4.Fields(0)
Else
End If
If IsNull(r1.Fields(0)) = False Then
lblcid.Caption = r1.Fields(0)
Else
End If
If IsNull(r2.Fields(0)) = False Then
lblSid.Caption = r2.Fields(0)
Else
End If
If IsNull(r.Fields(0)) = False Then
lbltpid.Caption = r3.Fields(0)
Else
End If
End Sub
Private Sub NewQs_Click()
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
ElseIf Combo7.Text = "" Then
MsgBox "Select Course First ", vbInformation + vbOKOnly, "Course Missing"
Combo7.SetFocus
Exit Sub
ElseIf Combo6.Text = "" Then
MsgBox "Select Subject Correspondent to the Course", vbInformation + vbOKOnly, "Subject Missing"
Combo6.SetFocus
Exit Sub
ElseIf Combo2.Text = "" Then
MsgBox "Select Topic Correspondent to the Subject of the course", vbInformation + vbOKOnly, "Topic Missing"
Combo2.SetFocus
Exit Sub
ElseIf Combo3.Text = "" Then
MsgBox "Select Question Type", vbInformation + vbOKOnly, "Question type Missing"
Combo3.SetFocus
Exit Sub
ElseIf Combo1.Text = "" Then
MsgBox "Choose Difficulty Level of Question", vbInformation + vbOKOnly, "Select Difficulty"
Combo1.SetFocus
Exit Sub
ElseIf Trim(expn_mcq.Text) = "" Then
MsgBox "Please Enter Some Explanation at least Select Same as answer answer ??? ", vbInformation + vbOKOnly, "Question Explanation"
expn_mcq.SetFocus
Exit Sub
End If
updtDA  'Calling Function To Load data
If pic_name = "" Then
 c1.Execute ("insert into quesMS values('" & QID.Text & "'," & Val(qnoMS.Caption) & ",'" & lblcid.Caption & "','" & lblSid.Caption & "','" & lbltpid.Caption & "','" & lblQtyp.Caption & "','" & qtext_mcqs.Text & "','" & opt1(0).Text & "','" & opt1(1).Text & "','" & opt1(2).Text & "','" & opt1(3).Text & "','" & aNswer_txt.Caption & "'," & Val(ans_num.Caption) & ",'" & Combo1.Text & "','" & expn_mcq.Text & "',NULL)")
ElseIf pic_name <> "" And pic.Visible = False And Qimg.Visible = True Then
 c1.Execute ("insert into quesMS values('" & QID.Text & "'," & Val(qnoMS.Caption) & ",'" & lblcid.Caption & "','" & lblSid.Caption & "','" & lbltpid.Caption & "','" & lblQtyp.Caption & "','" & qtext_mcqs.Text & "','" & opt1(0).Text & "','" & opt1(1).Text & "','" & opt1(2).Text & "','" & opt1(3).Text & "','" & aNswer_txt.Caption & "'," & Val(ans_num.Caption) & ",'" & Combo1.Text & "','" & expn_mcq.Text & "','" & pic_name & "')")
End If
formConfirm.Show vbModal, MDI
If Check1.Value = vbChecked Then
Command1_Click
Me.Top = 50 '100
Me.Left = 900 '4400
CHKSelRTF = 0
nquesBTN.Enabled = True
pquesBTN.Enabled = True
qtext_mcqs.Text = ""
For i = 0 To 3
 btnopt(i).Value = vbUnchecked
 opt1(i).Text = ""
Next i
vkCheck1.Value = vbUnchecked
expn_mcq.Text = ""
aNswer_txt.Caption = ""
ans_num.Caption = ""
qnoMS.Caption = ""
pic_name = ""
vkCheck1.Value = vbUnchecked 'Same as answer
q_autoID 'New ID GENERATED
Else
Form_Load
End If
End Sub

Private Sub nquesBTN_Click()
On Error Resume Next
Timer1.Enabled = False
Command1.Visible = False
tempvr = ""
Set r = New ADODB.Recordset
tempvr2 = "MS" & Format(Val(Right(QID.Text, 4)) + 1, "0000")
Set r = c.Execute("select * from quesms where upper(Q_ID) ='" & UCase(tempvr2) & "' ")
If r.EOF = False Then
 QID.Text = r.Fields(0)
 Course.Caption = r.Fields(2)
 Subj.Caption = r.Fields(3)
 Tpic.Caption = r.Fields(4)
 QuTyp.Caption = r.Fields(5)
 qtext_mcqs.Text = r.Fields(6)
 opt1(0).Text = r.Fields(7)
 opt1(1).Text = r.Fields(8)
 opt1(2).Text = r.Fields(9)
 opt1(3).Text = r.Fields(10)
 QLevel.Caption = r.Fields(13)
 btnopt(r.Fields(12) - 1).Value = vbChecked
 If IsNull(r.Fields(14)) = False Then
  expn_mcq.Text = r.Fields(14)
 Else
  expn_mcq.Text = ""
 End If
 If IsNull(r.Fields(15)) = False Then
  tempvr = r.Fields(15)
 Else
  tempvr = ""
 End If
 If tempvr <> "" Then
  Qimg.Picture = LoadPicture(tempvr)
  pic.Visible = False
  Qimg.Visible = True
  Else
  pic.Visible = True
  Qimg.Visible = False
 End If
  img.Enabled = False
  pic.Enabled = False
Set r1 = c.Execute("select c_nm from course where c_id='" & Course.Caption & "' ")
 Combo7.Text = r1.Fields(0)
Set r1 = c.Execute("select sub_nm from sub where sub_id='" & Subj.Caption & "' ")
 Combo6.Text = r1.Fields(0)
Set r1 = c.Execute("select tp_nm from topic where tp_id='" & Tpic.Caption & "' ")
 Combo2.Text = r1.Fields(0)
Set r1 = c.Execute("select Q_TYP_NM from Q_typ where Q_TYP_ID='" & QuTyp.Caption & "' ")
 Combo3.Text = r1.Fields(0)
 Combo1.Text = UCase(QLevel.Caption)
 NewQs.Enabled = False
Exit Sub
End If
'Clear  All Filled Text
Combo7.Clear
Combo6.Clear
Combo2.Clear
CHKSelRTF = 0
Combo3.Clear
rs_qtyp.MoveFirst
If IsNull(rs_qtyp.EOF) = False Then
While rs_qtyp.EOF = False
 Combo3.AddItem rs_qtyp(1)
 rs_qtyp.MoveNext
Wend
Else
End If
rs_course.MoveFirst
If IsNull(rs_course.EOF) = False Then
While rs_course.EOF = False
 Combo7.AddItem rs_course(0)
 rs_course.MoveNext
Wend
Else
End If
Combo1.Clear
Combo1.AddItem "EASY"
Combo1.AddItem "MEDIUM"
Combo1.AddItem "HARD"
 qtext_mcqs.Text = ""
 For i = 0 To 3
  btnopt(i).Value = vbUnchecked
  opt1(i).Text = ""
 Next i
 expn_mcq.Text = ""
 aNswer_txt.Caption = ""
 ans_num.Caption = ""
 qnoMS.Caption = ""
 pic_name = ""
 Qimg.Visible = True
 pic.Visible = True
 vkCheck1.Value = vbUnchecked 'Same as answer
 q_autoID 'New ID GENERATED
 NewQs.Enabled = True
  Command1.Visible = False
  Timer1.Enabled = True
  Combo7.SetFocus
  Combo6.Text = ""
  Combo2.Text = ""
  Combo3.Text = ""
  Combo1.Text = ""
  img.Enabled = True
  pic.Enabled = True
  Command1_Click
 MsgBox "No More Question ahead, You Can Add new Question Here", vbInformation + vbOKOnly, "No Question Ahead"
End Sub

Private Sub opt1_GotFocus(Index As Integer)
CHKSelRTF = Index + 2
End Sub

Private Sub opt1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
If Index < 3 Then
opt1(Index + 1).SetFocus
ElseIf Index = 3 Then
expn_mcq.SetFocus
End If
End If
End Sub

Private Sub pasteMCQs_Click()
If CHKSelRTF = 1 Then
qtext_mcqs.SelText = Clipboard.GetText()
ElseIf CHKSelRTF = 2 Then
opt1(0).SelText = Clipboard.GetText()
ElseIf CHKSelRTF = 3 Then
opt1(1).SelText = Clipboard.GetText()
ElseIf CHKSelRTF = 4 Then
opt1(2).SelText = Clipboard.GetText()
ElseIf CHKSelRTF = 5 Then
opt1(3).SelText = Clipboard.GetText()
ElseIf CHKSelRTF = 6 Then
expn_mcq.SelText = Clipboard.GetText()
End If
End Sub

Private Sub pasteMCQs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cutMCQs.FontUnderline = False
copyMCQs.FontUnderline = False
pasteMCQs.FontUnderline = True
End Sub

Private Sub pic_Click() 'Upload
On Error Resume Next
   cdb.Filter = "All Picture Files *.jpg,.gif,.bmp,.ico JPEG Image"
   cdb.ShowOpen
   If cdb.FileName <> "" Then
       pic.Visible = False
       Qimg.Visible = True
       Qimg.Picture = LoadPicture(cdb.FileName)
        pic_name = cdb.FileName
        pic_ext = Right(cdb.FileTitle, 4)
        pic_changed = True
     End If
End Sub

Private Sub pquesBTN_Click()
On Error Resume Next
Timer1.Enabled = False
Command1.Visible = False
tempvr = ""
Set r = New ADODB.Recordset
tempvr2 = "MS" & Format(Val(Right(QID.Text, 4)) - 1, "0000")
Set r = c.Execute("select * from quesms where upper(Q_ID) ='" & UCase(tempvr2) & "' ")
If r.EOF = False Then
 QID.Text = r.Fields(0)
 Course.Caption = r.Fields(2)
 Subj.Caption = r.Fields(3)
 Tpic.Caption = r.Fields(4)
 QuTyp.Caption = r.Fields(5)
 qtext_mcqs.Text = r.Fields(6)
 opt1(0).Text = r.Fields(7)
 opt1(1).Text = r.Fields(8)
 opt1(2).Text = r.Fields(9)
 opt1(3).Text = r.Fields(10)
 QLevel.Caption = r.Fields(13)
 btnopt(r.Fields(12) - 1).Value = vbChecked
 If IsNull(r.Fields(14)) = False Then
  expn_mcq.Text = r.Fields(14)
 Else
  expn_mcq.Text = ""
 End If
 img.Enabled = False
  pic.Enabled = False
 If IsNull(r.Fields(15)) = False Then
  tempvr = r.Fields(15)
 Else
  tempvr = ""
 End If
 If tempvr <> "" Then
  Qimg.Picture = LoadPicture(tempvr)
  pic.Visible = False
  Qimg.Visible = True
 Else
  Qimg.Visible = False
  pic.Visible = True
 End If
Set r1 = c.Execute("select c_nm from course where c_id='" & Course.Caption & "' ")
 Combo7.Text = r1.Fields(0)
Set r1 = c.Execute("select sub_nm from sub where sub_id='" & Subj.Caption & "' ")
 Combo6.Text = r1.Fields(0)
Set r1 = c.Execute("select tp_nm from topic where tp_id='" & Tpic.Caption & "' ")
 Combo2.Text = r1.Fields(0)
Set r1 = c.Execute("select Q_TYP_NM from Q_typ where Q_TYP_ID='" & QuTyp.Caption & "' ")
 Combo3.Text = r1.Fields(0)
 Combo1.Text = UCase(QLevel.Caption)
 NewQs.Enabled = False
Exit Sub
End If
 MsgBox "No More Question Left.", vbInformation + vbOKOnly, "No Question Ahead"
 Command1.Visible = False
 End Sub

Private Sub qtext_mcqs_GotFocus()
CHKSelRTF = 1
End Sub

Private Sub qtext_mcqs_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
opt1(0).SetFocus
End If
End Sub

Private Sub Timer1_Timer()
If Qimg.Picture <> 0 Then
pic.Visible = False
Command1.Visible = True
Else
Command1.Visible = False
pic.Visible = True
End If
End Sub

Private Sub undrln_btn_Click()
qtext_mcqs.SelUnderline = Not qtext_mcqs.SelUnderline
opt1(0).SelUnderline = Not opt1(0).SelUnderline
opt1(1).SelUnderline = Not opt1(1).SelUnderline
opt1(2).SelUnderline = Not opt1(2).SelUnderline
opt1(3).SelUnderline = Not opt1(3).SelUnderline
expn_mcq.SelUnderline = Not expn_mcq.SelUnderline
End Sub

Private Sub vkCommand1_Click() 'Left  Alignment
qtext_mcqs.SelIndent = Not qtext_mcqs.SelIndent
End Sub

Private Sub vkCheck1_Click()
If vkCheck1.Value = vbChecked Then
expn_mcq.Text = ""
expn_mcq.Text = aNswer_txt.Caption
ElseIf vkCheck1.Value = vbUnchecked Then
expn_mcq.Text = ""
End If

End Sub
Private Sub vkCommand4_Click()
Unload Me
End Sub

Private Sub vkFrame2_MouseHover()
cutMCQs.FontUnderline = False
copyMCQs.FontUnderline = False
pasteMCQs.FontUnderline = False
End Sub

Public Function q_autoID()
Set r = New ADODB.Recordset
sql = "select MAX(to_number(substr(q_id,3,length(q_id))))from quesMs"
Set r = c1.Execute(sql)
If IsNull(r.Fields(0)) Then
QID.Text = "MS000" & 1
qnoMS.Caption = 1
Else
 t = r.Fields(0)
 If t > 0 And t < 9 Then
  QID.Text = "MS000" & (t + 1)
  qnoMS.Caption = t + 1
 ElseIf t < 99 Then
  QID.Text = "MS00" & (t + 1)
  qnoMS.Caption = t + 1
ElseIf t < 999 Then
 QID.Text = "MS0" & (t + 1)
 qnoMS.Caption = t + 1
 End If
End If
End Function

Private Sub vkToggleButton1_Click() 'Strike through
qtext_mcqs.SelStrikeThru = Not qtext_mcqs.SelStrikeThru
opt1(0).SelStrikeThru = Not opt1(0).SelStrikeThru
opt1(1).SelStrikeThru = Not opt1(1).SelStrikeThru
opt1(2).SelStrikeThru = Not opt1(2).SelStrikeThru
opt1(3).SelStrikeThru = Not opt1(3).SelStrikeThru
expn_mcq.SelStrikeThru = Not expn_mcq.SelStrikeThru
End Sub

Private Sub vkToggleButton2_Click() 'LowerCase
On Error Resume Next
If qtext_mcqs.SelLength > 0 Then
qtext_mcqs.SelRTF = LCase(qtext_mcqs.SelRTF)
opt1(0).SelLength = 0
opt1(1).SelLength = 0
opt1(2).SelLength = 0
opt1(3).SelLength = 0
expn_mcq.SelLength = 0
End If

If expn_mcq.SelLength > 0 Then
expn_mcq.SelText = LCase(expn_mcq.SelText)
opt1(0).SelLength = 0
opt1(1).SelLength = 0
opt1(2).SelLength = 0
opt1(3).SelLength = 0
qtext_mcqs.SelLength = 0
End If

If opt1(0).SelLength > 0 Then
opt1(0).SelRTF = LCase(opt1(0).SelRTF)
End If
If opt1(1).SelLength > 0 Then
opt1(1).SelRTF = LCase(opt1(1).SelRTF)
End If
If opt1(2).SelLength > 0 Then
opt1(2).SelRTF = LCase(opt1(2).SelRTF)
End If
If opt1(3).SelLength > 0 Then
opt1(3).SelRTF = LCase(opt1(3).SelRTF)
End If
End Sub

Private Sub vkToggleButton3_Click() 'Capital Letter
If qtext_mcqs.SelLength > 0 Then
qtext_mcqs.SelRTF = UCase(qtext_mcqs.SelRTF)
opt1(0).SelLength = 0
opt1(1).SelLength = 0
opt1(2).SelLength = 0
opt1(3).SelLength = 0
expn_mcq.SelLength = 0
End If

If expn_mcq.SelLength > 0 Then
expn_mcq.SelText = UCase(expn_mcq.SelText)
opt1(0).SelLength = 0
opt1(1).SelLength = 0
opt1(2).SelLength = 0
opt1(3).SelLength = 0
qtext_mcqs.SelLength = 0
End If

If opt1(0).SelLength > 0 Then
opt1(0).SelRTF = UCase(opt1(0).SelRTF)
End If
If opt1(1).SelLength > 0 Then
opt1(1).SelRTF = UCase(opt1(1).SelRTF)
End If
If opt1(2).SelLength > 0 Then
opt1(2).SelRTF = UCase(opt1(2).SelRTF)
End If
If opt1(3).SelLength > 0 Then
opt1(3).SelRTF = UCase(opt1(3).SelRTF)
End If
End Sub


