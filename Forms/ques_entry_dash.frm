VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{EAFDAFBF-1D88-41DD-B117-60ECBC4B8441}#1.0#0"; "vkUserControlsXP.ocx"
Object = "{F5E116E1-0563-11D8-AA80-000B6A0D10CB}#1.0#0"; "HookMenu.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form ques_entry_dash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   10530
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   20400
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10530
   ScaleWidth      =   20400
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnLVL 
      Height          =   495
      Left            =   12975
      MouseIcon       =   "ques_entry_dash.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "ques_entry_dash.frx":0152
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   1545
      Width           =   1305
   End
   Begin VB.CommandButton btnans 
      Height          =   495
      Left            =   10170
      MouseIcon       =   "ques_entry_dash.frx":087E
      MousePointer    =   99  'Custom
      Picture         =   "ques_entry_dash.frx":09D0
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   1545
      Width           =   2775
   End
   Begin VB.CommandButton btnCHapter 
      Height          =   495
      Left            =   17040
      MouseIcon       =   "ques_entry_dash.frx":130E
      MousePointer    =   99  'Custom
      Picture         =   "ques_entry_dash.frx":1460
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   1545
      Width           =   3350
   End
   Begin VB.CommandButton btnCOURSE 
      Height          =   495
      Left            =   14310
      MouseIcon       =   "ques_entry_dash.frx":1F41
      MousePointer    =   99  'Custom
      Picture         =   "ques_entry_dash.frx":2093
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   1545
      Width           =   1350
   End
   Begin VB.CommandButton btnSub 
      Height          =   495
      Left            =   15690
      MouseIcon       =   "ques_entry_dash.frx":27B6
      MousePointer    =   99  'Custom
      Picture         =   "ques_entry_dash.frx":2908
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   1545
      Width           =   1335
   End
   Begin VB.CommandButton btnQUES 
      Height          =   495
      Left            =   3360
      MouseIcon       =   "ques_entry_dash.frx":30CD
      MousePointer    =   99  'Custom
      Picture         =   "ques_entry_dash.frx":321F
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   1545
      Width           =   6780
   End
   Begin VB.CommandButton btnID 
      Height          =   495
      Left            =   2355
      MouseIcon       =   "ques_entry_dash.frx":4215
      MousePointer    =   99  'Custom
      Picture         =   "ques_entry_dash.frx":4367
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   1545
      Width           =   1000
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   5760
      Top             =   8880
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6240
      Top             =   8880
   End
   Begin MSAdodcLib.Adodc Ado2 
      Height          =   330
      Left            =   8400
      Top             =   8880
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      Connect         =   "Provider=MSDAORA.1;Password=sts;User ID=sts;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;Password=sts;User ID=sts;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"ques_entry_dash.frx":49A2
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
   Begin vkUserContolsXP.vkFrame fram6 
      Height          =   7215
      Left            =   150
      TabIndex        =   70
      Top             =   2520
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   12726
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
      Begin RichTextLib.RichTextBox rtf 
         Height          =   2775
         Left            =   60
         TabIndex        =   72
         Top             =   480
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   4895
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"ques_entry_dash.frx":4A6B
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
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Type Question "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   2
         Left            =   0
         TabIndex        =   71
         Top             =   0
         Width           =   2055
      End
   End
   Begin vkUserContolsXP.vkFrame fram5 
      Height          =   7215
      Left            =   150
      TabIndex        =   67
      Top             =   2520
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   12726
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
      Begin VB.ComboBox QType 
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
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Select Question Type"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   1
         Left            =   0
         TabIndex        =   69
         Top             =   0
         Width           =   2055
      End
   End
   Begin vkUserContolsXP.vkFrame fram3 
      Height          =   7215
      Left            =   150
      TabIndex        =   58
      Top             =   2520
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   12726
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
      Begin VB.ComboBox Combo8 
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
         Left            =   45
         TabIndex        =   66
         Text            =   "Combo8"
         Top             =   3280
         Width           =   1935
      End
      Begin VB.CheckBox chk2 
         Caption         =   "Include Subject"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   2360
         Width           =   1815
      End
      Begin VB.CheckBox chk1 
         Caption         =   "Include Course"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   960
         Width           =   1815
      End
      Begin VB.ComboBox Combo7 
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
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   1920
         Width           =   1935
      End
      Begin VB.ComboBox Combo5 
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
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Type / Select Topic"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   0
         TabIndex        =   65
         Top             =   2790
         Width           =   2055
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Select Course"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   0
         TabIndex        =   63
         Top             =   0
         Width           =   2055
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Select Subject"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   0
         TabIndex        =   62
         Top             =   1440
         Width           =   2055
      End
   End
   Begin vkUserContolsXP.vkFrame Fram2 
      Height          =   7215
      Left            =   150
      TabIndex        =   52
      Top             =   2520
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   12726
      BackColor1      =   14737632
      BackColor2      =   12632256
      BackGradient    =   0
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
      Begin VB.ComboBox cmbCRSfrm2 
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
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   480
         Width           =   1935
      End
      Begin VB.ComboBox fram2SUB 
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
         Left            =   55
         TabIndex        =   54
         Top             =   1890
         Width           =   1935
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Include Course"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   960
         Width           =   1815
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
         Left            =   120
         TabIndex        =   73
         Top             =   300
         Width           =   105
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Type / Select Subject"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   0
         TabIndex        =   57
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Select Course"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   0
         TabIndex        =   56
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000010&
      Height          =   855
      Left            =   0
      TabIndex        =   47
      Top             =   9720
      Width           =   2295
      Begin vkUserContolsXP.vkCommand vkCommand2 
         Height          =   450
         Left            =   240
         TabIndex        =   48
         Top             =   255
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   794
         BackColor1      =   16777215
         BackColor2      =   12632256
         Caption         =   "Back"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   16761024
         Picture         =   "ques_entry_dash.frx":4AF4
         CustomStyle     =   0
         MouseHoverPicture=   "ques_entry_dash.frx":4E58
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000F&
         BorderWidth     =   2
         X1              =   0
         X2              =   2280
         Y1              =   120
         Y2              =   120
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      TabIndex        =   29
      Top             =   1560
      Width           =   2130
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         Caption         =   "-: Search By :-"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   11.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   320
         Left            =   0
         TabIndex        =   30
         Top             =   120
         Width           =   2050
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   135
      Left            =   3480
      TabIndex        =   28
      Top             =   4200
      Width           =   15
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid1 
      Bindings        =   "ques_entry_dash.frx":51BC
      Height          =   8000
      Left            =   2355
      TabIndex        =   24
      Top             =   2075
      Width           =   18025
      _ExtentX        =   31803
      _ExtentY        =   14102
      _Version        =   393216
      BackColor       =   14869218
      Cols            =   7
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   450
      BackColorBkg    =   16761024
      GridColor       =   -2147483632
      ScrollTrack     =   -1  'True
      HighLight       =   2
      FillStyle       =   1
      GridLinesFixed  =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BorderStyle     =   0
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
      _Band(0).GridLinesBand=   3
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0)._NumMapCols=   7
      _Band(0)._MapCol(0)._Name=   "Q_ID"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(1)._Name=   "INITCAP(Q.Q_TXT)"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(0)._MapCol(2)._Name=   "ANS_TXT"
      _Band(0)._MapCol(2)._RSIndex=   2
      _Band(0)._MapCol(3)._Name=   "Q_DIF_LVL"
      _Band(0)._MapCol(3)._RSIndex=   3
      _Band(0)._MapCol(4)._Name=   "C_NM"
      _Band(0)._MapCol(4)._RSIndex=   4
      _Band(0)._MapCol(5)._Name=   "SUB_NM"
      _Band(0)._MapCol(5)._RSIndex=   5
      _Band(0)._MapCol(6)._Name=   "TP_NM"
      _Band(0)._MapCol(6)._RSIndex=   6
   End
   Begin VB.ComboBox Combo6 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      MouseIcon       =   "ques_entry_dash.frx":51CF
      MousePointer    =   99  'Custom
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2040
      Width           =   2055
   End
   Begin vkUserContolsXP.vkFrame Fram1 
      Height          =   7215
      Left            =   150
      TabIndex        =   10
      Top             =   2520
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   12726
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
      Begin VB.ComboBox srhCourse 
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
         Left            =   60
         TabIndex        =   50
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Type / Select Course"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   0
         Left            =   0
         TabIndex        =   51
         Top             =   0
         Width           =   2055
      End
   End
   Begin vkUserContolsXP.vkFrame vkFrame8 
      Height          =   1095
      Left            =   3960
      TabIndex        =   0
      Top             =   75
      Width           =   6135
      _ExtentX        =   10821
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
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1080
         MouseIcon       =   "ques_entry_dash.frx":5321
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   645
         Width           =   3255
      End
      Begin VB.CommandButton search 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   415
         Left            =   4680
         Picture         =   "ques_entry_dash.frx":5473
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   600
         Width           =   990
      End
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4305
         MouseIcon       =   "ques_entry_dash.frx":5BEE
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   180
         Width           =   1695
      End
      Begin VB.ComboBox Combo4 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1080
         MouseIcon       =   "ques_entry_dash.frx":5D40
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   180
         Width           =   1695
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
         Left            =   4150
         TabIndex        =   27
         Top             =   160
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
         Left            =   960
         TabIndex        =   26
         Top             =   120
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
         Left            =   960
         TabIndex        =   25
         Top             =   600
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
         Left            =   3360
         TabIndex        =   23
         Top             =   240
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
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label2 
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
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
   End
   Begin vkUserContolsXP.vkFrame vkFrame1 
      Height          =   1800
      Left            =   0
      TabIndex        =   3
      Top             =   -240
      Width           =   20475
      _ExtentX        =   36116
      _ExtentY        =   3175
      BackColor1      =   12632256
      BackColor2      =   14737632
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
      ShowTitle       =   0   'False
      TitleColor1     =   16761024
      TitleGradient   =   2
      TitleHeight     =   300
      BorderColor     =   16761024
      BorderWidth     =   0
      Begin vkUserContolsXP.vkFrame vkFrame26 
         Height          =   1065
         Left            =   15885
         TabIndex        =   45
         ToolTipText     =   "Display All Questions"
         Top             =   345
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   1879
         BackColor1      =   16777215
         BackColor2      =   16777215
         BackGradient    =   0
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
         TextPosition    =   0
         TitleColor1     =   16777215
         TitleColor2     =   16777215
         BorderColor     =   16761024
         DisplayPicture  =   0   'False
         Begin VB.Image btnShowAll 
            Height          =   945
            Left            =   90
            MouseIcon       =   "ques_entry_dash.frx":5E92
            MousePointer    =   99  'Custom
            Picture         =   "ques_entry_dash.frx":5FE4
            Top             =   45
            Width           =   960
         End
      End
      Begin vkUserContolsXP.vkFrame vkFrame25 
         Height          =   345
         Left            =   15890
         TabIndex        =   44
         Top             =   1440
         Width           =   1000
         _ExtentX        =   1773
         _ExtentY        =   609
         Caption         =   "Show All"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowBackGround  =   0   'False
         TitleColor1     =   64
         TitleColor2     =   4210816
         TitleGradient   =   2
         TitleHeight     =   340
         BorderColor     =   12648384
         RoundAngle      =   4
         BorderWidth     =   0
      End
      Begin vkUserContolsXP.vkFrame vkFrame24 
         Height          =   1080
         Left            =   18110
         TabIndex        =   42
         Top             =   345
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1905
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
         Begin vkUserContolsXP.vkCommand vkCommand1 
            Height          =   885
            Left            =   120
            TabIndex        =   43
            Top             =   105
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   1561
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
            Picture         =   "ques_entry_dash.frx":6D54
            CustomStyle     =   0
            MouseHoverPicture=   "ques_entry_dash.frx":7C5B
         End
      End
      Begin vkUserContolsXP.vkFrame vkFrame23 
         Height          =   345
         Left            =   18110
         TabIndex        =   41
         Top             =   1440
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         Caption         =   "Update"
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
      Begin vkUserContolsXP.vkFrame vkFrame22 
         Height          =   1080
         Left            =   13250
         TabIndex        =   37
         Top             =   330
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   1905
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
         Begin VB.CommandButton btnSrhID 
            BackColor       =   &H00E0E0E0&
            Caption         =   "GO"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            MouseIcon       =   "ques_entry_dash.frx":8200
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Bookman Old Style"
               Size            =   9.75
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   120
            MaxLength       =   15
            TabIndex        =   38
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label12 
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
            Left            =   2160
            TabIndex        =   46
            Top             =   100
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Enter Question ID "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   315
            TabIndex        =   40
            Top             =   120
            Width           =   1905
         End
      End
      Begin vkUserContolsXP.vkFrame vkFrame21 
         Height          =   345
         Left            =   13250
         TabIndex        =   36
         Top             =   1440
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   609
         Caption         =   "Search  Question"
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
      Begin vkUserContolsXP.vkFrame vkFrame20 
         Height          =   1095
         Left            =   10220
         TabIndex        =   32
         Top             =   315
         Width           =   2925
         _ExtentX        =   5159
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
         Begin VB.ComboBox Combo2 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   360
            MouseIcon       =   "ques_entry_dash.frx":8352
            MousePointer    =   99  'Custom
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   540
            Width           =   2055
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Choose Difficulti Leval"
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
            Left            =   150
            TabIndex        =   35
            Top             =   120
            Width           =   2355
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
            Left            =   2520
            TabIndex        =   34
            Top             =   120
            Width           =   105
         End
      End
      Begin vkUserContolsXP.vkFrame vkFrame19 
         Height          =   345
         Left            =   10220
         TabIndex        =   31
         Top             =   1440
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   609
         Caption         =   "Difficulti  Leval"
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
      Begin vkUserContolsXP.vkFrame vkFrame14 
         Height          =   345
         Left            =   16920
         TabIndex        =   19
         Top             =   1440
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         Caption         =   "Delete"
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
         Height          =   1065
         Left            =   16920
         TabIndex        =   18
         Top             =   345
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1879
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
         Begin VB.Image btnDelQues 
            Height          =   840
            Left            =   120
            MouseIcon       =   "ques_entry_dash.frx":84A4
            MousePointer    =   99  'Custom
            Picture         =   "ques_entry_dash.frx":85F6
            Stretch         =   -1  'True
            Top             =   120
            Width           =   840
         End
      End
      Begin vkUserContolsXP.vkFrame vkFrame9 
         Height          =   345
         Left            =   2750
         TabIndex        =   17
         Top             =   1425
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         Caption         =   "Export"
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
      Begin vkUserContolsXP.vkFrame vkFrame6 
         Height          =   1095
         Left            =   60
         TabIndex        =   15
         Top             =   315
         Width           =   1125
         _ExtentX        =   1984
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
         Begin HookMenu.XpMenu XpMenu1 
            Left            =   480
            Top             =   240
            _ExtentX        =   900
            _ExtentY        =   900
            BmpCount        =   1
            CheckBorderColor=   7021576
            SelMenuBorder   =   7021576
            SelMenuBackColor=   14073525
            SelMenuForeColor=   16646297
            SelCheckBackColor=   14791828
            MenuBorderColor =   6956042
            SeparatorColor  =   -2147483632
            MenuBackColor   =   14609903
            MenuForeColor   =   0
            CheckBackColor  =   15326939
            CheckForeColor  =   10027263
            DisabledMenuBorderColor=   -2147483632
            DisabledMenuBackColor=   15660791
            DisabledMenuForeColor=   -2147483631
            MenuBarBackColor=   15790320
            MenuPopupBackColor=   16777215
            ShortCutNormalColor=   0
            ShortCutSelectColor=   16646297
            ArrowNormalColor=   10027263
            ArrowSelectColor=   12484864
            ShadowColor     =   0
            Bmp:1           =   "ques_entry_dash.frx":9215
            Mask:1          =   16777215
            Key:1           =   "#mnuExt"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Image newQues 
            Height          =   825
            Left            =   120
            MouseIcon       =   "ques_entry_dash.frx":94CF
            MousePointer    =   99  'Custom
            Picture         =   "ques_entry_dash.frx":9621
            Stretch         =   -1  'True
            ToolTipText     =   "Click To Add Questions"
            Top             =   120
            Width           =   840
         End
      End
      Begin vkUserContolsXP.vkFrame vkFrame17 
         Height          =   1095
         Left            =   2750
         TabIndex        =   14
         Top             =   315
         Width           =   1125
         _ExtentX        =   1984
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
         Begin VB.CommandButton btnExport 
            Height          =   880
            Left            =   120
            MouseIcon       =   "ques_entry_dash.frx":9E74
            MousePointer    =   99  'Custom
            Picture         =   "ques_entry_dash.frx":9FC6
            Style           =   1  'Graphical
            TabIndex        =   78
            Top             =   120
            Width           =   880
         End
      End
      Begin vkUserContolsXP.vkFrame vkFrame4 
         Height          =   1095
         Left            =   1270
         TabIndex        =   13
         Top             =   315
         Width           =   1365
         _ExtentX        =   2408
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
         Begin VB.Image btnImport 
            Height          =   850
            Left            =   225
            MouseIcon       =   "ques_entry_dash.frx":AA0B
            MousePointer    =   99  'Custom
            Picture         =   "ques_entry_dash.frx":AB5D
            Stretch         =   -1  'True
            ToolTipText     =   "Import Questions From Excel File"
            Top             =   105
            Width           =   900
         End
      End
      Begin vkUserContolsXP.vkFrame vkFrame3 
         Height          =   345
         Left            =   1270
         TabIndex        =   12
         Top             =   1425
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   609
         Caption         =   "Import"
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
      Begin vkUserContolsXP.vkFrame vkFrame18 
         Height          =   345
         Left            =   50
         TabIndex        =   8
         Top             =   1425
         Width           =   1130
         _ExtentX        =   1984
         _ExtentY        =   609
         Caption         =   "Add New"
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
      Begin vkUserContolsXP.vkFrame vkFrame11 
         Height          =   345
         Left            =   3990
         TabIndex        =   6
         Top             =   1425
         Width           =   6135
         _ExtentX        =   10821
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
      Begin vkUserContolsXP.vkFrame vkFrame12 
         Height          =   1095
         Left            =   19295
         TabIndex        =   5
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
         Begin VB.CommandButton Command1 
            Height          =   855
            Left            =   120
            MouseIcon       =   "ques_entry_dash.frx":B3F9
            MousePointer    =   99  'Custom
            Picture         =   "ques_entry_dash.frx":B54B
            Style           =   1  'Graphical
            TabIndex        =   74
            ToolTipText     =   "Close"
            Top             =   120
            Width           =   855
         End
      End
      Begin vkUserContolsXP.vkFrame vkFrame13 
         Height          =   345
         Left            =   19295
         TabIndex        =   4
         Top             =   1425
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "Close"
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
   End
   Begin MSComctlLib.StatusBar sbar 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   10035
      Width           =   20400
      _ExtentX        =   35983
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4066
            MinWidth        =   4066
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   7056
            MinWidth        =   7056
            Text            =   "Speedup Test Solutions"
            TextSave        =   "Speedup Test Solutions"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "Total Questions : "
            TextSave        =   "Total Questions : "
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4057
            MinWidth        =   4057
            Text            =   "Courses : "
            TextSave        =   "Courses : "
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4057
            MinWidth        =   4057
            Text            =   "Subjects : "
            TextSave        =   "Subjects : "
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5115
            MinWidth        =   5115
            Text            =   "Date : "
            TextSave        =   "Date : "
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6103
            MinWidth        =   6103
            Text            =   "Time : "
            TextSave        =   "Time : "
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   8175
      Left            =   0
      ScaleHeight     =   8175
      ScaleWidth      =   2325
      TabIndex        =   77
      Top             =   1560
      Width           =   2320
   End
   Begin VB.Label holdRowCol 
      BackColor       =   &H0000FF00&
      Height          =   375
      Left            =   6720
      TabIndex        =   49
      Top             =   8880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      Height          =   1455
      Left            =   0
      Top             =   -240
      Width           =   20055
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fill In The Blanks"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   6255
      TabIndex        =   7
      Top             =   -525
      Width           =   1545
   End
   Begin VB.Line Line22 
      X1              =   120
      X2              =   13680
      Y1              =   -600
      Y2              =   -600
   End
   Begin VB.Menu mnuQues 
      Caption         =   "Question"
      Visible         =   0   'False
      Begin VB.Menu wqdd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuadd_new 
         Caption         =   "Add Question "
      End
      Begin VB.Menu wsd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelQues 
         Caption         =   "Delete Question"
      End
      Begin VB.Menu dffef 
         Caption         =   "-"
      End
      Begin VB.Menu mnuupdtQues 
         Caption         =   "Update Question"
      End
      Begin VB.Menu ddwd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExt 
         Caption         =   "Exit"
      End
      Begin VB.Menu shgushstftsgvgsv 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "ques_entry_dash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'USED ONLY FOR SORTING While Clicking
Dim chkIDsorting As Integer
Dim chkQUESsorting As Integer
Dim chkANSsorting As Integer
Dim chkLEVALsorting As Integer
Dim chkCOURSEsorting As Integer
Dim chkTOPICsorting As Integer
Dim chkSUBJECTsorting As Integer
Public Sub INITIALISE_ALL_SORTING_VALUES()
chkIDsorting = 0
chkQUESsorting = 0
chkANSsorting = 0
chkLEVALsorting = 0
chkCOURSEsorting = 0
chkTOPICsorting = 0
chkSUBJECTsorting = 0
End Sub
Private Sub add_ques_Click()
mcq_s.Show
End Sub

Private Sub btnANS_Click()
If chkANSsorting = 0 Then
 Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  C.c_id=Q.c_id and S.sub_id=Q.Sub_id and T.tp_id=Q.tp_id order by Q.ans_txt"
  chkANSsorting = 1
Else
 Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  C.c_id=Q.c_id and S.sub_id=Q.Sub_id and T.tp_id=Q.tp_id order by Q.ans_txt desc"
 chkANSsorting = 0
End If
Ado2.Refresh
call_grid_function
End Sub

Private Sub btnCHapter_Click()
If chkTOPICsorting = 0 Then
Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  C.c_id=Q.c_id and S.sub_id=Q.Sub_id and T.tp_id=Q.tp_id order by T.tp_nm"
chkTOPICsorting = 1
Else
Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  C.c_id=Q.c_id and S.sub_id=Q.Sub_id and T.tp_id=Q.tp_id order by T.tp_nm desc"
chkTOPICsorting = 0
End If
Ado2.Refresh
call_grid_function
End Sub

Private Sub btnCOURSE_Click()
If chkCOURSEsorting = 0 Then
Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  C.c_id=Q.c_id and S.sub_id=Q.Sub_id and T.tp_id=Q.tp_id order by C.c_nm"
chkCOURSEsorting = 1
Else
Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  C.c_id=Q.c_id and S.sub_id=Q.Sub_id and T.tp_id=Q.tp_id order by C.c_nm desc"
chkCOURSEsorting = 0
End If
Ado2.Refresh
call_grid_function
End Sub

Private Sub btnDelQues_Click()
QuesBank.Show
End Sub

Private Sub btnExport_Click() 'Export Button
FrmExportQues.Show
End Sub

Private Sub btnID_Click()
If chkIDsorting = 0 Then
Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  C.c_id=Q.c_id and S.sub_id=Q.Sub_id and T.tp_id=Q.tp_id order by q.q_id desc"
chkIDsorting = 1
Else
Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  C.c_id=Q.c_id and S.sub_id=Q.Sub_id and T.tp_id=Q.tp_id order by q.q_id "
chkIDsorting = 0
End If
Ado2.Refresh
call_grid_function
End Sub

Private Sub btnImport_Click() 'Import Button
FrmImportQues.Show
End Sub

Private Sub btnLVL_Click()
If chkLEVALsorting = 0 Then
 Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  C.c_id=Q.c_id and S.sub_id=Q.Sub_id and T.tp_id=Q.tp_id order by Q.Q_DIF_LVL"
chkLEVALsorting = 1
Else
Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  C.c_id=Q.c_id and S.sub_id=Q.Sub_id and T.tp_id=Q.tp_id order by Q.Q_DIF_LVL desc"
chkLEVALsorting = 0
End If
Ado2.Refresh
call_grid_function
End Sub

Private Sub btnQUES_Click()
If chkQUESsorting = 0 Then
Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  C.c_id=Q.c_id and S.sub_id=Q.Sub_id and T.tp_id=Q.tp_id order by Q.q_txt"
chkQUESsorting = 1
Else
Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  C.c_id=Q.c_id and S.sub_id=Q.Sub_id and T.tp_id=Q.tp_id order by Q.q_txt desc"
chkQUESsorting = 0
End If
Ado2.Refresh
call_grid_function
End Sub

Private Sub btnShowAll_Click()
Text1.Text = ""
Label12.Visible = False
Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  C.c_id=Q.c_id and S.sub_id=Q.Sub_id and T.tp_id=Q.tp_id order by Q.q_id"
Ado2.Refresh
call_grid_function
End Sub

Private Sub btnSrhID_Click()
If Text1.Text = "" Then
Label12.Visible = True
Text1.SetFocus
Else
Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  C.c_id=Q.c_id and S.sub_id=Q.Sub_id and T.tp_id=Q.tp_id and upper(Q.q_id)='" & UCase(Text1.Text) & "' order by S.sub_nm"
Ado2.Refresh
call_grid_function
Label12.Visible = False
End If
End Sub

Private Sub btnSub_Click()
If chkSUBJECTsorting = 0 Then
Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  C.c_id=Q.c_id and S.sub_id=Q.Sub_id and T.tp_id=Q.tp_id order by S.sub_nm"
chkSUBJECTsorting = 1
Else
Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  C.c_id=Q.c_id and S.sub_id=Q.Sub_id and T.tp_id=Q.tp_id order by S.sub_nm desc"
chkSUBJECTsorting = 0
End If
Ado2.Refresh
call_grid_function
End Sub

Private Sub Check2_Click()
If Check2.Value = Unchecked Then
cmbCRSfrm2.Enabled = False
Set r = New ADODB.Recordset
Set r = c1.Execute("select sub_nm from sub")
fram2SUB.Clear
 While r.EOF = False
  fram2SUB.AddItem r.Fields(0)
  r.MoveNext
Wend
 Else
 cmbCRSfrm2.Enabled = True
 Set r = New ADODB.Recordset
 Set r = c1.Execute("select sub_nm from sub where c_id=(select c_id from course where c_nm='" & cmbCRSfrm2.Text & "') ")
 If IsNull(r.Fields(0)) = False Then
  fram2SUB.Clear
  While r.EOF = False
   fram2SUB.AddItem r.Fields(0)
   r.MoveNext
   Wend
   End If
 End If
End Sub

Private Sub chk1_Click()
Set r = New ADODB.Recordset
 If chk1.Value = Checked Then
  Combo5.Enabled = True
'  Set r = New ADODB.Recordset ' ADDING TO COURSE
' Set r = c1.Execute("select sub_nm from sub where c_id=(select c_id from course where c_nm='" & Combo5.Text & "') ")
' If IsNull(r.Fields(0)) = False Then
'  Combo7.clear
'  While r.EOF = False
'   Combo7.AddItem r.Fields(0)
'   r.MoveNext
'   Wend
'   End If
'  If chk2.Enabled = False Then 'If Second Option is disabled
'   Set r = New ADODB.Recordset 'ADDING TO TOPIC
'   Set r = c1.Execute("select tp_nm from topic where c_id=(select c_id from course where c_nm='" & Combo5.Text & "') ")
'  If IsNull(r.Fields(0)) = False Then
'  Combo8.clear
'  While r.EOF = False
'   Combo8.AddItem r.Fields(0)
'   r.MoveNext
'   Wend
'   End If
'ElseIf chk2.Enabled = True And Combo7.Text <> "" Then
'   Set r = New ADODB.Recordset 'ADDING TO TOPIC
'   Set r = c1.Execute("select tp_nm from topic where sub_id=(select sub_id from sub where sub_nm='" & Combo7.Text & "' and c_id=(select c_id from course where c_nm='" & Combo5.Text & "')) and c_id=(select c_id from course where c_nm='" & Combo5.Text & "') ")
'  If IsNull(r.Fields(0)) = False Then
'  Combo8.clear
'  While r.EOF = False
'   Combo8.AddItem r.Fields(0)
'   r.MoveNext
'   Wend
'   End If
'End If
 Else
  Combo5.Enabled = False
'   Set r = c1.Execute("select sub_nm from sub")
'    Combo7.clear
'    While r.EOF = False
'     Combo7.AddItem r.Fields(0)
'     r.MoveNext
'    Wend

'Set r = New ADODB.Recordset
' Set r = c1.Execute("select tp_nm from topic")
' If IsNull(r.Fields(0)) = False Then
'  Combo8.clear
'  While r.EOF = False
'   Combo8.AddItem r.Fields(0)
'   r.MoveNext
'   Wend
'   End If
 End If
End Sub

Private Sub chk2_Click()
If chk2.Value = Checked Then
Combo7.Enabled = True
Combo8.Clear 'CLEANING TOPIC COMBO
'Set r = New ADODB.Recordset
' Set r = c1.Execute("select tp_nm from topic where sub_id=(select sub_id from sub where sub_nm='" & Combo7.Text & "') ")
' If IsNull(r.Fields(0)) = False Then
'  Combo7.clear
'  While r.EOF = False
'   Combo7.AddItem r.Fields(0)
'   r.MoveNext
'   Wend
'   End If
Else
Combo7.Enabled = False
If chk1.Value = Checked And Combo5.Text <> "" Then
Set r = c1.Execute("select tp_nm from topic where c_id=(select c_id from course where c_nm='" & Combo5.Text & "')")
    Combo8.Clear
    While r.EOF = False
     Combo8.AddItem r.Fields(0)
     r.MoveNext
    Wend
ElseIf chk2.Value = Unchecked Then
If chk1.Value = Checked And Combo5.Text <> "" Then
Set r = c1.Execute("select tp_nm from topic ")
    Combo8.Clear
    While r.EOF = False
     Combo8.AddItem r.Fields(0)
     r.MoveNext
    Wend
End If
End If
End If
End Sub

Private Sub cmbCRSfrm2_Click()
Set r1 = New ADODB.Recordset
Set r1 = c1.Execute("select sub_nm from sub where c_id =(select c_id from course where upper(c_nm)='" & UCase(cmbCRSfrm2.Text) & "') ")
If IsNull(r1.Fields(0)) Then
Else
fram2SUB.Clear
While r1.EOF = False
 fram2SUB.AddItem r1.Fields(0)
 r1.MoveNext
Wend
End If
End Sub

Private Sub combo2_Click()
Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  C.c_id=Q.c_id and S.sub_id=Q.Sub_id and T.tp_id=Q.tp_id and upper(Q.Q_DIF_LVL)='" & Combo2.Text & "' order by q.q_id"
Ado2.Refresh
call_grid_function
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

Private Sub Combo4_Click()
Combo3.Clear
Combo1.Clear
Set r1 = New ADODB.Recordset
Set r1 = c1.Execute("select sub_nm from sub where c_id =(select c_id from course where upper(c_nm)='" & UCase(Combo4.Text) & "') ")
While r1.EOF = False
 Combo3.AddItem r1.Fields(0)
 r1.MoveNext
Wend
End Sub

Private Sub Combo5_Click() 'Course frame 3
Combo7.Clear
Set r1 = New ADODB.Recordset
Set r1 = c1.Execute("select sub_nm from sub where c_id =(select c_id from course where upper(c_nm)='" & UCase(Combo5.Text) & "') ")
While r1.EOF = False
 Combo7.AddItem r1.Fields(0)
 r1.MoveNext
Wend

Combo8.Clear
 Set r1 = New ADODB.Recordset
  Set r1 = c1.Execute("select tp_nm from topic where c_id=(select c_id from course where c_nm='" & Combo5.Text & "') ")
  While r1.EOF = False
 Combo8.AddItem r1.Fields(0) 'r1!
 r1.MoveNext
Wend
End Sub

Private Sub Combo6_Click()
Label19.Visible = False
If Combo6.ListIndex = 0 Then 'COURSE WISE
Fram1.Visible = True
Fram2.Visible = False
Fram3.Visible = False
fram5.Visible = False
fram6.Visible = False
Set r = New ADODB.Recordset
Set r = c1.Execute("select c_nm from course")
srhCourse.Clear
 While r.EOF = False
  srhCourse.AddItem r.Fields(0)
  r.MoveNext
Wend
ElseIf Combo6.ListIndex = 1 Then 'SUBJECT WISE
Fram1.Visible = False
Fram2.Visible = True
Fram3.Visible = False
fram5.Visible = False
fram6.Visible = False
Set r = New ADODB.Recordset
Set r = c1.Execute("select c_nm from course")
cmbCRSfrm2.Clear
 While r.EOF = False
  cmbCRSfrm2.AddItem r.Fields(0)
  r.MoveNext
Wend
If Check2.Value = Unchecked Then
 cmbCRSfrm2.Enabled = False
 Set r = New ADODB.Recordset
Set r = c1.Execute("select sub_nm from sub")
fram2SUB.Clear
 While r.EOF = False
  fram2SUB.AddItem r.Fields(0)
  r.MoveNext
Wend
 Else
 cmbCRSfrm2.Enabled = True
  Set r = New ADODB.Recordset
Set r = c1.Execute("select sub_nm from sub where c_id=(select c_id from course where c_nm='" & cmbCRSfrm2 & "'")
 fram2SUB.Clear
 While r.EOF = False
  fram2SUB.AddItem r.Fields(0)
  r.MoveNext
Wend
End If
ElseIf Combo6.ListIndex = 2 Then 'Search By Topics
Fram1.Visible = False
Fram2.Visible = False
Fram3.Visible = True
fram5.Visible = False
fram6.Visible = False
Set r = New ADODB.Recordset
Set r = c1.Execute("select c_nm from course")
Combo5.Clear
 While r.EOF = False
  Combo5.AddItem r.Fields(0)
  r.MoveNext
Wend
r.Close
'''''''''''''Checking Option is checked or not
If chk1.Value = Unchecked Then
 Combo5.Enabled = False
 Else
 Combo5.Enabled = True
End If

If chk2.Value = Unchecked Then
 Combo7.Enabled = False
 Else
 Combo7.Enabled = True
End If
 '++++++ Now adding subject +++++++'
  Set r = New ADODB.Recordset
  Set r = c1.Execute("select sub_nm from sub ")
   Combo7.Clear
  While r.EOF = False
   Combo7.AddItem r.Fields(0)
   r.MoveNext
  Wend
r.Close

Set r = New ADODB.Recordset
Set r = c1.Execute("select tp_nm from topic ")
 Combo8.Clear
 While r.EOF = False
  Combo8.AddItem r.Fields(0)
  r.MoveNext
 Wend

ElseIf Combo6.ListIndex = 3 Then 'Difficulti Leval
Fram1.Visible = False
Fram2.Visible = False
Fram3.Visible = False
fram5.Visible = False
fram6.Visible = False
Combo2.SetFocus
ElseIf Combo6.ListIndex = 4 Then 'Question type
Fram1.Visible = False
Fram2.Visible = False
Fram3.Visible = False
fram5.Visible = True
fram6.Visible = False
Set r = New ADODB.Recordset
Set r = c1.Execute("select q_typ_nm from q_typ")
While r.EOF = False
 QType.AddItem r.Fields(0)
 r.MoveNext
Wend
ElseIf Combo6.ListIndex = 5 Then 'Question ID
Fram1.Visible = False
Fram2.Visible = False
Fram3.Visible = False
fram5.Visible = False
fram6.Visible = False
Text1.SetFocus
ElseIf Combo6.ListIndex = 6 Then 'Question Search
Fram1.Visible = False
Fram2.Visible = False
Fram3.Visible = False
fram5.Visible = False
fram6.Visible = True
rtf.SetFocus
End If
End Sub '++++++++++++++++++++HERE I WORKING NOW

Private Sub Combo7_Click()
If chk1.Value = Checked Then 'Course Is enabled but not select
  If Combo5.Text = "" Then
   'Label19.Visible = True
   Combo5.SetFocus
  Else 'Course is selected
  Set r1 = New ADODB.Recordset
   Combo8.Clear
   Set r1 = c1.Execute("select tp_nm from topic where c_id =(select c_id from course where upper(c_nm)='" & UCase(Combo5.Text) & "')and sub_id =(select sub_id from sub where sub_nm='" & Combo7.Text & "' and c_id=(select c_id from course where c_nm='" & Combo5.Text & "')) ")
   While r1.EOF = False
    Combo8.AddItem r1.Fields(0)
    r1.MoveNext
   Wend
  End If
Else 'Course is disabled
Set r1 = New ADODB.Recordset
 Combo8.Clear
 Set r1 = c1.Execute("select tp_nm from topic where sub_id =(select sub_id from sub where sub_nm='" & Combo7.Text & "'")
While r1.EOF = False
 Combo8.AddItem r1.Fields(0)
 r1.MoveNext
Wend
End If
End Sub

Private Sub Combo8_Change()
If chk1.Value = Checked And Combo5.Text <> "" Then
 If chk2.Value = Checked And Combo7.Text = "" Then
   Combo7.SetFocus
 ElseIf chk2.Value = Checked And Combo7.Text <> "" Then 'Subject Is Enabled
   Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  C.c_id=Q.c_id and S.sub_id=Q.Sub_id and T.tp_id=Q.tp_id and upper(T.tp_nm) like '" & UCase(Trim(Combo8.Text)) & "%' and T.c_id=(select c_id from course where c_nm='" & Combo5.Text & "')and T.sub_id =(select sub_id from sub where sub_nm='" & Combo7.Text & "' and c_id=(select c_id from course where c_nm='" & Combo5.Text & "')) order by Q.q_id"
 ElseIf chk2.Value = Unchecked Then 'Subject Not Enabled
   Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  C.c_id=Q.c_id and S.sub_id=Q.Sub_id and T.tp_id=Q.tp_id and upper(T.tp_nm) like '" & UCase(Trim(Combo8.Text)) & "%' and T.c_id=(select c_id from course where c_nm='" & Combo5.Text & "') order by Q.q_id"
 End If
ElseIf chk1.Value = Checked And Combo5.Text = "" Then
Combo5.SetFocus
ElseIf chk1.Value = Unchecked Then 'Course is disabled
 If chk2.Value = Checked And Combo7.Text = "" Then
   Combo7.SetFocus
 ElseIf chk2.Value = Checked And Combo7.Text <> "" Then 'Subject is enabled
   Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  C.c_id=Q.c_id and S.sub_id=Q.Sub_id and T.tp_id=Q.tp_id and upper(T.tp_nm) like '" & UCase(Trim(Combo8.Text)) & "%' and T.sub_id =(select sub_id from sub where sub_nm='" & Combo7.Text & "') order by Q.q_id"
 ElseIf chk2.Value = Unchecked Then 'Subject Not Enabled
   Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  C.c_id=Q.c_id and S.sub_id=Q.Sub_id and T.tp_id=Q.tp_id and upper(T.tp_nm) like '" & UCase(Trim(Combo8.Text)) & "%' order by Q.q_id"
 End If
End If
Ado2.Refresh
call_grid_function
End Sub

Private Sub Combo8_Click() 'Click on topic
If chk1.Value = Checked And Combo5.Text <> "" Then 'Course selected
 If chk2.Value = Checked And Combo7.Text = "" Then 'Subject Selected
    Combo7.SetFocus
    Exit Sub
 ElseIf chk2.Value = Checked And Combo7.Text <> "" Then
   Set r = New ADODB.Recordset
   Set r = c1.Execute("select tp_id from topic where tp_nm='" & Combo8.Text & "' and c_id =(select c_id from course where c_nm='" & Combo5.Text & "')and sub_id =(select sub_id from sub where sub_nm='" & Combo7.Text & "' and c_id=(select c_id from course where  c_nm='" & Combo5.Text & "'))")
   Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  C.c_id=Q.c_id and S.sub_id=Q.Sub_id and T.tp_id=Q.tp_id and upper(T.tp_id) = '" & UCase(Trim(r.Fields(0))) & "' and T.sub_id=(select sub_id from sub where sub_nm='" & Combo7.Text & "' and c_id=(select c_id from course where c_nm='" & Combo5.Text & "'))and T.c_id=(select c_id from course where c_nm='" & Combo5.Text & "') order by Q.q_id"
   Ado2.Refresh
 ElseIf chk2.Value = Unchecked Then 'Only Course is selected
   Set r = New ADODB.Recordset
   Set r = c1.Execute("select tp_id from topic where tp_nm='" & Combo8.Text & "' and c_id =(select c_id from course where c_nm='" & Combo5.Text & "')")
   Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  C.c_id=Q.c_id and S.sub_id=Q.Sub_id and T.tp_id=Q.tp_id and upper(T.tp_id) = '" & UCase(Trim(r.Fields(0))) & "' and T.c_id=(select c_id from course where c_nm='" & Combo5.Text & "') order by Q.q_id"
   Ado2.Refresh
 End If
ElseIf chk1.Value = Unchecked Then
 If chk2.Value = Checked And Combo7.Text = "" Then
   Combo7.SetFocus
   Exit Sub
 ElseIf chk2.Value = Checked And Combo7.Text <> "" Then 'Subject Selected
   Set r = New ADODB.Recordset
   Set r = c1.Execute("select tp_id from topic where tp_nm='" & Combo8.Text & "' and sub_id =(select sub_id from sub where sub_nm='" & Combo7.Text & "')")
   Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  C.c_id=Q.c_id and S.sub_id=Q.Sub_id and T.tp_id=Q.tp_id and upper(T.tp_id) = '" & UCase(Trim(r.Fields(0))) & "' and T.sub_id=(select sub_id from sub where sub_nm='" & Combo7.Text & "')  order by Q.q_id"
   Ado2.Refresh
 ElseIf chk2.Value = Unchecked Then 'No Course and No subject
   Set r = New ADODB.Recordset
   Set r = c1.Execute("select tp_id from topic where tp_nm='" & Combo8.Text & "' ")
   Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  C.c_id=Q.c_id and S.sub_id=Q.Sub_id and T.tp_id=Q.tp_id and upper(T.tp_id) = '" & UCase(Trim(r.Fields(0))) & "')order by Q.q_id"
   Ado2.Refresh
 End If
End If
call_grid_function
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Activate()
 Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  C.c_id=Q.c_id and S.sub_id=Q.Sub_id and T.tp_id=Q.tp_id order by Q.q_id"
 Ado2.Refresh
 call_grid_function
End Sub

Private Sub Form_Load()
conn
Me.Top = 0
Me.Left = 0
QUESTIONRight = ""
Timer1.Enabled = True
INITIALISE_ALL_SORTING_VALUES 'Function For Storing Values
Combo4.Clear
If IsNull(rs_course.EOF) = False Then
While rs_course.EOF = False
 Combo4.AddItem rs_course(0)
 rs_course.MoveNext
Wend
Else
End If
Combo6.AddItem "Course"
Combo6.AddItem "Subject"
Combo6.AddItem "Topic"
Combo6.AddItem "Difficulti Leval"
Combo6.AddItem "Question Type"
Combo6.AddItem "Question ID"
Combo6.AddItem "Question Name"

Combo2.AddItem "EASY"
Combo2.AddItem "MEDIUM"
Combo2.AddItem "HARD"

call_grid_function 'Function For Arranging Grids
End Sub

Private Sub fram2SUB_Change()
If Check2.Value = Checked Then
Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  C.c_id=Q.c_id and S.sub_id=Q.Sub_id and T.tp_id=Q.tp_id and upper(S.sub_nm) like '" & UCase(Trim(fram2SUB.Text)) & "%' and S.c_id=(select c_id from course where c_nm='" & cmbCRSfrm2.Text & "') order by Q.q_id"
Else
Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  C.c_id=Q.c_id and S.sub_id=Q.Sub_id and T.tp_id=Q.tp_id and upper(S.sub_nm) like '" & UCase(Trim(fram2SUB.Text)) & "%' order by Q.q_id"
End If
Ado2.Refresh
call_grid_function
End Sub

Private Sub fram2SUB_Click()
If Check2.Value = Checked Then
 If cmbCRSfrm2.Text = "" Then
  Label19.Visible = True
  cmbCRSfrm2.SetFocus
  Else
  Set r = New ADODB.Recordset
Set r = c1.Execute("select sub_id from sub where sub_nm='" & fram2SUB.Text & "' and c_id =(select c_id from course where c_nm='" & cmbCRSfrm2.Text & "')")
   Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  C.c_id=Q.c_id and S.sub_id=Q.Sub_id and T.tp_id=Q.tp_id and upper(S.sub_id) = '" & UCase(Trim(r.Fields(0))) & "' and S.c_id=(select c_id from course where c_nm='" & cmbCRSfrm2.Text & "') order by Q.q_id"
  End If
Else
Set r = New ADODB.Recordset
Set r = c1.Execute("select sub_id from sub where sub_nm='" & fram2SUB.Text & "'")
Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  C.c_id=Q.c_id and S.sub_id=Q.Sub_id and T.tp_id=Q.tp_id and upper(S.sub_id) = '" & UCase(Trim(r.Fields(0))) & "' order by Q.q_id"
End If
Ado2.Refresh
call_grid_function
End Sub

Private Sub grid1_Click()
With grid1
.Col = 0
.ColSel = .Cols - 1
holdRowCol.Caption = .TextMatrix(.Row, .Col)
QUESTIONRight = holdRowCol.Caption
End With
End Sub

Private Sub grid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu mnuQues
End If
End Sub

Private Sub mnuadd_new_Click()
mcq_s.Show
End Sub

Private Sub mnuDelQues_Click() 'Delete by Right Click
Dim temp As String
On Error GoTo k:
If QUESTIONRight = "" Then
GoTo k:
Else
temp = MsgBox("Are You Sure !!!!", vbCritical + vbYesNo, "Delete Question")
If temp = vbYes Then
 c1.Execute ("delete from quesMS where q_id='" & QUESTIONRight & "' ")
 MsgBox "Question SuccessFully Deleted...", vbInformation + vbOKOnly, "Deleted"
 Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  C.c_id=Q.c_id and S.sub_id=Q.Sub_id and T.tp_id=Q.tp_id order by Q.q_id"
 Ado2.Refresh
 call_grid_function
 QUESTIONRight = ""
Else
End If
Exit Sub
End If
k:
MsgBox "Right Click On To The Question First then Click on delete..", vbCritical + vbOKOnly, "No Selection"
End Sub

Private Sub mnuExt_Click()
Unload Me
End Sub

Private Sub mnuupdtQues_Click()
'Form4.Text2.Text = grid1.TextMatrix(grid1.Row, grid1.Col)
Dim temp As String
On Error GoTo k:
If QUESTIONRight = "" Then
GoTo k:
Else
 ques_entry_dash.Enabled = False
 FrmQuesUpdate.Show
Exit Sub
End If
k:
MsgBox "Select Question First then Click on Update..", vbCritical + vbOKOnly, "No Selection"
End Sub

Private Sub newQues_Click()
Unload admin_dash
Me.Enabled = False
mcq_s.Show
End Sub

Public Function call_grid_function()
Dim i As Integer
grid1.ColWidth(0) = 1000
grid1.ColWidth(1) = 6800
grid1.ColWidth(2) = 2825
grid1.ColWidth(3) = 1300
grid1.ColWidth(4) = 1400
grid1.ColWidth(5) = 1350
grid1.ColWidth(6) = 3060 '3200
For i = 0 To 6
If i <> 1 Then
 grid1.ColAlignment(i) = flexAlignCenterCenter
End If
Next i
End Function

Private Sub QType_Click()
Set r = New ADODB.Recordset
Set r = c1.Execute("select q_typ_id from q_typ where q_typ_nm='" & QType.Text & "'")
Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  C.c_id=Q.c_id and S.sub_id=Q.Sub_id and T.tp_id=Q.tp_id and upper(Q.q_typ_id) ='" & UCase(Trim(r.Fields(0))) & "' order by Q.q_id"
Ado2.Refresh
call_grid_function
End Sub

Private Sub rtf_Change()
Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  C.c_id=Q.c_id and S.sub_id=Q.Sub_id and T.tp_id=Q.tp_id and upper(Q.q_txt) like '%" & UCase(Trim(rtf.Text)) & "%' order by Q.q_id"
Ado2.Refresh
call_grid_function
End Sub

Private Sub search_Click()
If Combo4.Text <> "" And Combo3.Text <> "" And Combo1.Text <> "" Then
Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  Q.c_id=(select c_id from course where c_nm='" & Combo4.Text & "') and Q.sub_id=(select sub_id from sub where sub_nm='" & Combo3.Text & "' and c_id=(select c_id from course where c_nm='" & Combo4.Text & "'))" _
& "and Q.tp_id=(select tp_id from topic where tp_nm='" & Combo1.Text & "' and sub_id=(select sub_id from sub where upper(sub_nm)='" & UCase(Combo3.Text) & "' and c_id=(select c_id from course where c_nm='" & Combo4.Text & "')) AND C_ID=(SELECT C_ID FROM COURSE WHERE  C_NM='" & Combo4.Text & "')) and S.SUB_ID =(select sub_id from sub where sub_nm='" & Combo3.Text & "' and c_id=(select c_id from course where c_nm='" & Combo4.Text & "')) and  T.TP_ID=(SELECT TP_ID FROM TOPIC WHERE TP_NM='" & Combo1.Text & "' AND SUB_ID =(select sub_id from sub where sub_nm='" & Combo3.Text & "' and c_id=(select c_id from course where c_nm = '" & Combo4.Text & "'))" _
& "AND c_id =(select c_id from course where c_nm='" & Combo4.Text & "')) and C.c_id=(select c_id from course where upper(c_nm)='" & UCase(Combo4.Text) & "')"
Ado2.Refresh
call_grid_function
End If
End Sub

Private Sub srhCourse_Change()
Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  C.c_id=Q.c_id and S.sub_id=Q.Sub_id and T.tp_id=Q.tp_id and upper(C.c_nm) like '" & UCase(Trim(srhCourse.Text)) & "%' order by Q.q_id"
Ado2.Refresh
call_grid_function
End Sub

Private Sub srhCourse_Click()
Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  C.c_id=Q.c_id and S.sub_id=Q.Sub_id and T.tp_id=Q.tp_id and upper(C.c_nm) ='" & UCase(Trim(srhCourse.Text)) & "' order by Q.q_id"
Ado2.Refresh
call_grid_function
End Sub

Private Sub Text1_Change()
Ado2.RecordSource = "select Q.q_id, Q.q_txt, Q.ans_txt, Q.Q_DIF_LVL, C.c_nm,S.sub_nm,T.tp_nm from quesms Q, Course C,Sub S,Topic T where  C.c_id=Q.c_id and S.sub_id=Q.Sub_id and T.tp_id=Q.tp_id and upper(Q.q_id) like '" & UCase(Trim(Text1.Text)) & "%' order by Q.q_id"
Ado2.Refresh
call_grid_function
If Ado2.Recordset.RecordCount > 0 Then
Label12.Visible = False
Else
Label12.Visible = True
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
btnSrhID_Click
End If
End Sub

Private Sub Timer1_Timer()
Set r = New ADODB.Recordset
Set r = c1.Execute("select count(*) from quesms")
sbar.Panels(3).Text = "Total Questions :  " & r.Fields(0)
Set r = New ADODB.Recordset
Set r = c1.Execute("select count(*) from course")
sbar.Panels(4).Text = "Courses :    " & r.Fields(0)
Set r = New ADODB.Recordset
Set r = c1.Execute("select count(*) from sub")
sbar.Panels(5).Text = "Subjects :    " & r.Fields(0)
End Sub

Private Sub Timer2_Timer()
sbar.Panels(7).Text = "Time : " & Format$(Time, "hh:mm:ss  AM/PM")
sbar.Panels(6).Text = "Date :  " & Format$(Date, "dd-mm-yyyy")
End Sub

Private Sub vkCommand1_Click()
QuesBank.Show
End Sub

Private Sub vkCommand2_Click()
Unload Me
End Sub

Private Sub Form_Unload(cancel As Integer)
If EMP_login_reg_no = "" Then
admin_dash.Enabled = True
Else
emp_dash.Enabled = True
End If
End Sub

