VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamalButton.ocx"
Object = "{EAFDAFBF-1D88-41DD-B117-60ECBC4B8441}#1.0#0"; "vkUserControlsXP.ocx"
Object = "{08654D78-6636-11D3-87BF-B4980CC10374}#2.0#0"; "MyEllipticButton.ocx"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form FrmReportMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Report"
   ClientHeight    =   10710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
   Icon            =   "REPORT_MASTER.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10710
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   820
      Left            =   2760
      Picture         =   "REPORT_MASTER.frx":076A
      ScaleHeight     =   825
      ScaleWidth      =   17655
      TabIndex        =   218
      Top             =   0
      Width           =   17655
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   2775
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   1080
         MouseIcon       =   "REPORT_MASTER.frx":2E8A
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":2FDC
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   100
         Width           =   650
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   80
      Left            =   2880
      Top             =   240
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000011&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000011&
      Height          =   9855
      Left            =   0
      TabIndex        =   0
      Top             =   780
      Width           =   2775
      Begin vkUserContolsXP.vkCommand rptbtn 
         Height          =   1335
         Index           =   0
         Left            =   0
         TabIndex        =   1
         ToolTipText     =   "Report of All Master Entries"
         Top             =   360
         Width           =   2450
         _ExtentX        =   4313
         _ExtentY        =   2355
         BackColor1      =   7171437
         BackColor2      =   7171437
         BackColorPushed1=   -2147483632
         BackColorPushed2=   16777215
         BackGradient    =   0
         Caption         =   "Master Entry"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
         BorderColor     =   10526880
         DrawFocus       =   0   'False
         DrawMouseInRect =   0   'False
         DisabledBackColor=   15070196
         CustomStyle     =   0
      End
      Begin vkUserContolsXP.vkCommand rptbtn 
         Height          =   1335
         Index           =   1
         Left            =   0
         TabIndex        =   2
         ToolTipText     =   "This Option Will Allow to print Student Details on multiple Criteria."
         Top             =   1920
         Width           =   2450
         _ExtentX        =   4313
         _ExtentY        =   2355
         BackColor1      =   7171437
         BackColor2      =   7171437
         BackColorPushed1=   -2147483632
         BackColorPushed2=   16777215
         BackGradient    =   0
         Caption         =   "Students"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
         BorderColor     =   10526880
         DrawFocus       =   0   'False
         DrawMouseInRect =   0   'False
         DisabledBackColor=   15070196
         CustomStyle     =   0
      End
      Begin vkUserContolsXP.vkCommand rptbtn 
         Height          =   1335
         Index           =   2
         Left            =   0
         TabIndex        =   3
         ToolTipText     =   "It Contains All Details about Account Department of Company."
         Top             =   3480
         Width           =   2450
         _ExtentX        =   4313
         _ExtentY        =   2355
         BackColor1      =   7171437
         BackColor2      =   7171437
         BackColorPushed1=   -2147483632
         BackColorPushed2=   16777215
         BackGradient    =   0
         Caption         =   "Accounts"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
         BorderColor     =   10526880
         DrawFocus       =   0   'False
         DrawMouseInRect =   0   'False
         DisabledBackColor=   15070196
         CustomStyle     =   0
      End
      Begin vkUserContolsXP.vkCommand rptbtn 
         Height          =   1335
         Index           =   3
         Left            =   0
         TabIndex        =   4
         ToolTipText     =   "Many Important Options Related To User & Clients Order."
         Top             =   5040
         Width           =   2450
         _ExtentX        =   4313
         _ExtentY        =   2355
         BackColor1      =   7171437
         BackColor2      =   7171437
         BackColorPushed1=   -2147483632
         BackColorPushed2=   16777215
         BackGradient    =   0
         Caption         =   "Users && Clients"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
         BorderColor     =   8421504
         DrawFocus       =   0   'False
         DrawMouseInRect =   0   'False
         DisabledBackColor=   15070196
         CustomStyle     =   0
      End
      Begin vkUserContolsXP.vkCommand rptbtn 
         Height          =   1335
         Index           =   4
         Left            =   0
         TabIndex        =   5
         ToolTipText     =   "Major Important Work Can Be done Quickly from Here."
         Top             =   6600
         Width           =   2450
         _ExtentX        =   4313
         _ExtentY        =   2355
         BackColor1      =   7171437
         BackColor2      =   7171437
         BackColorPushed1=   -2147483632
         BackColorPushed2=   16777215
         BackGradient    =   0
         Caption         =   "Tools"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
         BorderColor     =   10526880
         DrawFocus       =   0   'False
         DrawMouseInRect =   0   'False
         DisabledBackColor=   15070196
         CustomStyle     =   0
      End
      Begin vkUserContolsXP.vkCommand rptbtn 
         Height          =   1350
         Index           =   5
         Left            =   0
         TabIndex        =   6
         Top             =   8160
         Width           =   2450
         _ExtentX        =   4313
         _ExtentY        =   2381
         BackColor1      =   7171437
         BackColor2      =   7171437
         BackColorPushed1=   -2147483632
         BackColorPushed2=   16777215
         BackGradient    =   0
         Caption         =   "Exit"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
         BorderColor     =   10526880
         DrawFocus       =   0   'False
         DrawMouseInRect =   0   'False
         DisabledBackColor=   15070196
         CustomStyle     =   0
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   9695
      Left            =   2880
      TabIndex        =   45
      Top             =   840
      Width           =   17535
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   6
         Left            =   14040
         MouseIcon       =   "REPORT_MASTER.frx":344A
         MousePointer    =   99  'Custom
         ScaleHeight     =   375
         ScaleWidth      =   615
         TabIndex        =   207
         Top             =   750
         Width           =   615
         Begin VB.Label a1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "15"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   255
            Index           =   6
            Left            =   0
            TabIndex        =   214
            ToolTipText     =   "List Of Students whose package is expiring Today"
            Top             =   30
            Width           =   615
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   7
         Left            =   16560
         MouseIcon       =   "REPORT_MASTER.frx":359C
         MousePointer    =   99  'Custom
         ScaleHeight     =   375
         ScaleWidth      =   615
         TabIndex        =   206
         Top             =   750
         Width           =   615
         Begin VB.Label a1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "15"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   255
            Index           =   7
            Left            =   0
            TabIndex        =   215
            ToolTipText     =   "List Of Student Who Has Requested For Updating Package."
            Top             =   30
            Width           =   615
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   4
         Left            =   8880
         MouseIcon       =   "REPORT_MASTER.frx":36EE
         MousePointer    =   99  'Custom
         ScaleHeight     =   375
         ScaleWidth      =   615
         TabIndex        =   205
         Top             =   750
         Width           =   615
         Begin VB.Label a1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "15"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   212
            ToolTipText     =   "List Of Individual Students"
            Top             =   30
            Width           =   615
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   5
         Left            =   11520
         MouseIcon       =   "REPORT_MASTER.frx":3840
         MousePointer    =   99  'Custom
         ScaleHeight     =   375
         ScaleWidth      =   615
         TabIndex        =   204
         Top             =   750
         Width           =   615
         Begin VB.Label a1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "15"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   213
            ToolTipText     =   "List of Students Who Have Enrolled In Today."
            Top             =   30
            Width           =   615
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   2
         Left            =   4800
         MouseIcon       =   "REPORT_MASTER.frx":3992
         MousePointer    =   99  'Custom
         ScaleHeight     =   375
         ScaleWidth      =   735
         TabIndex        =   203
         Top             =   750
         Width           =   735
         Begin VB.Label a1 
            BackStyle       =   0  'Transparent
            Caption         =   "15"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   210
            ToolTipText     =   "List of Girls Students"
            Top             =   30
            Width           =   615
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   3
         Left            =   6840
         MouseIcon       =   "REPORT_MASTER.frx":3AE4
         MousePointer    =   99  'Custom
         ScaleHeight     =   375
         ScaleWidth      =   615
         TabIndex        =   202
         Top             =   750
         Width           =   615
         Begin VB.Label a1 
            BackStyle       =   0  'Transparent
            Caption         =   "15"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   211
            ToolTipText     =   "List of Students Having Packages Facility"
            Top             =   30
            Width           =   615
         End
      End
      Begin VB.CommandButton studnt 
         BackColor       =   &H8000000E&
         Height          =   1120
         Index           =   7
         Left            =   15345
         MouseIcon       =   "REPORT_MASTER.frx":3C36
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":3D88
         Style           =   1  'Graphical
         TabIndex        =   199
         Top             =   120
         Width           =   2145
      End
      Begin VB.CommandButton studnt 
         BackColor       =   &H8000000E&
         Height          =   1120
         Index           =   6
         Left            =   12705
         MouseIcon       =   "REPORT_MASTER.frx":553A
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":568C
         Style           =   1  'Graphical
         TabIndex        =   198
         Top             =   120
         Width           =   2655
      End
      Begin VB.CommandButton studnt 
         BackColor       =   &H8000000E&
         Height          =   1120
         Index           =   5
         Left            =   10065
         MouseIcon       =   "REPORT_MASTER.frx":7052
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":71A4
         Style           =   1  'Graphical
         TabIndex        =   197
         Top             =   120
         Width           =   2655
      End
      Begin VB.CommandButton studnt 
         BackColor       =   &H8000000E&
         Height          =   1120
         Index           =   4
         Left            =   7800
         MouseIcon       =   "REPORT_MASTER.frx":899C
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":8AEE
         Style           =   1  'Graphical
         TabIndex        =   196
         Top             =   120
         Width           =   2280
      End
      Begin VB.CommandButton studnt 
         BackColor       =   &H8000000E&
         Height          =   1120
         Index           =   3
         Left            =   5640
         MouseIcon       =   "REPORT_MASTER.frx":9FA6
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":A0F8
         Style           =   1  'Graphical
         TabIndex        =   195
         Top             =   120
         Width           =   2175
      End
      Begin VB.CommandButton studnt 
         BackColor       =   &H8000000E&
         Height          =   1120
         Index           =   2
         Left            =   3840
         MouseIcon       =   "REPORT_MASTER.frx":B65E
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":B7B0
         Style           =   1  'Graphical
         TabIndex        =   194
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00B0D0D0&
         Caption         =   "REFRESH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   13440
         MouseIcon       =   "REPORT_MASTER.frx":C815
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Refresh All Records"
         Top             =   9015
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00B0D0D0&
         Height          =   675
         Left            =   16440
         MouseIcon       =   "REPORT_MASTER.frx":C967
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":CAB9
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Print Report"
         Top             =   8880
         Width           =   975
      End
      Begin VB.TextBox txtStud2 
         Height          =   375
         Left            =   6360
         TabIndex        =   72
         Text            =   "Text4"
         Top             =   5160
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Txtstud 
         Height          =   375
         Left            =   6360
         TabIndex        =   71
         Text            =   "Text4"
         Top             =   5160
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00B0D0D0&
         Caption         =   "PRINT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   11880
         MouseIcon       =   "REPORT_MASTER.frx":DA5B
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Print Report"
         Top             =   9015
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
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
         Left            =   2040
         MouseIcon       =   "REPORT_MASTER.frx":DBAD
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   9000
         Width           =   2070
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00B0D0D0&
         Caption         =   "SEARCH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   10320
         MouseIcon       =   "REPORT_MASTER.frx":DCFF
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Click to Search"
         Top             =   9015
         Width           =   1335
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   495
         Left            =   6360
         Top             =   5160
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
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
         Connect         =   "Provider=MSDAORA.1;Password=STS;User ID=STS;Persist Security Info=True"
         OLEDBString     =   "Provider=MSDAORA.1;Password=STS;User ID=STS;Persist Security Info=True"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   $"REPORT_MASTER.frx":DE51
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
      Begin MSACAL.Calendar Calendar2 
         Height          =   3015
         Left            =   6960
         TabIndex        =   46
         Top             =   6000
         Width           =   3135
         _Version        =   524288
         _ExtentX        =   5530
         _ExtentY        =   5318
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2019
         Month           =   6
         Day             =   4
         DayLength       =   1
         MonthLength     =   2
         DayFontColor    =   0
         FirstDay        =   1
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   0   'False
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSACAL.Calendar Calendar1 
         Height          =   3015
         Left            =   4320
         TabIndex        =   47
         Top             =   6000
         Width           =   3255
         _Version        =   524288
         _ExtentX        =   5741
         _ExtentY        =   5318
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2019
         Month           =   6
         Day             =   4
         DayLength       =   1
         MonthLength     =   2
         DayFontColor    =   0
         FirstDay        =   1
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   0   'False
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Dateframe 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   4440
         TabIndex        =   52
         Top             =   8880
         Visible         =   0   'False
         Width           =   5175
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   54
            Top             =   120
            Width           =   1575
         End
         Begin VB.TextBox Text6 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   53
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "From :"
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
            Index           =   0
            Left            =   120
            TabIndex        =   56
            Top             =   135
            Width           =   690
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To :"
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
            Index           =   1
            Left            =   2880
            TabIndex        =   55
            Top             =   120
            Width           =   390
         End
      End
      Begin VB.Frame NotDateFrame 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   4320
         TabIndex        =   48
         Top             =   8905
         Width           =   5475
         Begin VB.ComboBox Combo2 
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
            Left            =   2520
            TabIndex        =   49
            Text            =   "Combo2"
            Top             =   105
            Width           =   3000
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Enter or select Value :"
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
            Index           =   3
            Left            =   0
            TabIndex        =   50
            Top             =   120
            Width           =   2280
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "REPORT_MASTER.frx":DF55
         Height          =   7035
         Left            =   30
         TabIndex        =   59
         Top             =   1755
         Width           =   17460
         _ExtentX        =   30798
         _ExtentY        =   12409
         _Version        =   393216
         AllowUpdate     =   -1  'True
         ColumnHeaders   =   0   'False
         HeadLines       =   1
         RowHeight       =   23
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "RSTUD_REG_NO"
            Caption         =   "RSTUD_REG_NO"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "RSTUD_NM"
            Caption         =   "RSTUD_NM"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "RSTUD_FATHER_NM"
            Caption         =   "RSTUD_FATHER_NM"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "RSTUD_STATUS"
            Caption         =   "RSTUD_STATUS"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "RSTUD_MOB"
            Caption         =   "RSTUD_MOB"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "RSTUD_DOJ"
            Caption         =   "RSTUD_DOJ"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "C_NM"
            Caption         =   "C_NM"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "PKG_NM"
            Caption         =   "PKG_NM"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "SCH_TIMING"
            Caption         =   "SCH_TIMING"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   5
            ScrollBars      =   2
            Locked          =   -1  'True
            RecordSelectors =   0   'False
            BeginProperty Column00 
               Alignment       =   2
               ColumnWidth     =   1785.26
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2894.74
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2624.882
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   1844.787
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnWidth     =   1500.095
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               ColumnWidth     =   1574.929
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               ColumnWidth     =   1620.284
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               ColumnWidth     =   1604.976
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   1200
         MouseIcon       =   "REPORT_MASTER.frx":DF6A
         MousePointer    =   99  'Custom
         ScaleHeight     =   375
         ScaleWidth      =   615
         TabIndex        =   201
         Top             =   800
         Width           =   615
         Begin VB.Label a1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "15"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   270
            Index           =   0
            Left            =   0
            TabIndex        =   208
            ToolTipText     =   "List Of All Students .( Whether Package or Without Package )"
            Top             =   45
            Width           =   600
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   1
         Left            =   3000
         MouseIcon       =   "REPORT_MASTER.frx":E0BC
         MousePointer    =   99  'Custom
         ScaleHeight     =   375
         ScaleWidth      =   735
         TabIndex        =   200
         Top             =   750
         Width           =   735
         Begin VB.Label a1 
            BackStyle       =   0  'Transparent
            Caption         =   "15"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   209
            ToolTipText     =   "List Of Boys Students."
            Top             =   30
            Width           =   615
         End
      End
      Begin VB.CommandButton studnt 
         BackColor       =   &H8000000E&
         Height          =   1120
         Index           =   1
         Left            =   2055
         MouseIcon       =   "REPORT_MASTER.frx":E20E
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":E360
         Style           =   1  'Graphical
         TabIndex        =   193
         Top             =   120
         Width           =   1800
      End
      Begin VB.CommandButton studnt 
         BackColor       =   &H8000000E&
         Height          =   1120
         Index           =   0
         Left            =   45
         MouseIcon       =   "REPORT_MASTER.frx":F391
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":F4E3
         Style           =   1  'Graphical
         TabIndex        =   192
         Top             =   120
         Width           =   2030
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1095
         Left            =   30
         Top             =   120
         Width           =   17460
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   2040
         X2              =   2040
         Y1              =   120
         Y2              =   1200
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         X1              =   3840
         X2              =   3840
         Y1              =   120
         Y2              =   1200
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         X1              =   5640
         X2              =   5640
         Y1              =   120
         Y2              =   1200
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00808080&
         X1              =   7800
         X2              =   7800
         Y1              =   120
         Y2              =   1200
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00808080&
         X1              =   10080
         X2              =   10080
         Y1              =   120
         Y2              =   1200
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00808080&
         X1              =   12720
         X2              =   12720
         Y1              =   120
         Y2              =   1200
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00808080&
         X1              =   15360
         X2              =   15360
         Y1              =   120
         Y2              =   1200
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search By : "
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
         Index           =   2
         Left            =   720
         TabIndex        =   69
         Top             =   9000
         Width           =   1200
      End
      Begin VB.Shape Shape9 
         BorderColor     =   &H00808080&
         Height          =   495
         Left            =   30
         Top             =   1245
         Width           =   17460
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registration No"
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
         Index           =   4
         Left            =   180
         TabIndex        =   68
         Top             =   1350
         Width           =   1440
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student Name"
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
         Index           =   5
         Left            =   1900
         TabIndex        =   67
         Top             =   1350
         Width           =   1335
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Father's Name"
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
         Index           =   6
         Left            =   4755
         TabIndex        =   66
         Top             =   1350
         Width           =   1335
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student Type"
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
         Index           =   7
         Left            =   7635
         TabIndex        =   65
         Top             =   1350
         Width           =   1230
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Schedule "
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
         Index           =   8
         Left            =   15990
         TabIndex        =   64
         Top             =   1350
         Width           =   900
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Package"
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
         Index           =   9
         Left            =   14365
         TabIndex        =   63
         Top             =   1350
         Width           =   750
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Course"
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
         Index           =   10
         Left            =   12850
         TabIndex        =   62
         Top             =   1350
         Width           =   645
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Join Date"
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
         Index           =   11
         Left            =   11050
         TabIndex        =   61
         Top             =   1350
         Width           =   855
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile No"
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
         Index           =   12
         Left            =   9415
         TabIndex        =   60
         Top             =   1350
         Width           =   975
      End
      Begin VB.Line Line10 
         Index           =   0
         X1              =   7335
         X2              =   7335
         Y1              =   1250
         Y2              =   1730
      End
      Begin VB.Line Line10 
         Index           =   1
         X1              =   4710
         X2              =   4710
         Y1              =   1250
         Y2              =   1730
      End
      Begin VB.Line Line10 
         Index           =   2
         X1              =   10680
         X2              =   10680
         Y1              =   1250
         Y2              =   1730
      End
      Begin VB.Line Line10 
         Index           =   3
         X1              =   9185
         X2              =   9185
         Y1              =   1250
         Y2              =   1730
      End
      Begin VB.Line Line10 
         Index           =   4
         X1              =   12412
         X2              =   12415
         Y1              =   1250
         Y2              =   1730
      End
      Begin VB.Line Line10 
         Index           =   5
         X1              =   13995
         X2              =   13995
         Y1              =   1250
         Y2              =   1730
      End
      Begin VB.Line Line10 
         Index           =   6
         X1              =   15615
         X2              =   15615
         Y1              =   1250
         Y2              =   1730
      End
      Begin VB.Line Line10 
         Index           =   7
         X1              =   1820
         X2              =   1820
         Y1              =   1250
         Y2              =   1750
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H00808080&
         FillColor       =   &H8000000F&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   0
         Top             =   8805
         Width           =   17520
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Height          =   9695
      Left            =   2880
      TabIndex        =   141
      Top             =   840
      Width           =   17535
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0E0E0&
         Height          =   4815
         Left            =   360
         TabIndex        =   149
         Top             =   3600
         Visible         =   0   'False
         Width           =   3320
         Begin VB.CommandButton ChameleonBtn2 
            Height          =   415
            Left            =   480
            MouseIcon       =   "REPORT_MASTER.frx":10AD0
            MousePointer    =   99  'Custom
            Picture         =   "REPORT_MASTER.frx":10C22
            Style           =   1  'Graphical
            TabIndex        =   217
            Top             =   3900
            Width           =   2415
         End
         Begin VB.ComboBox Combo14 
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
            Left            =   480
            MouseIcon       =   "REPORT_MASTER.frx":11840
            MousePointer    =   99  'Custom
            Style           =   2  'Dropdown List
            TabIndex        =   150
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label lb6 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Student Name :"
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
            Left            =   1200
            TabIndex        =   159
            Top             =   165
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lb5 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   360
            TabIndex        =   158
            Top             =   120
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lb3 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Student Name :"
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
            TabIndex        =   157
            Top             =   4395
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lb4 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Student Name :"
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
            Left            =   840
            TabIndex        =   156
            Top             =   4440
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lb2 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   480
            TabIndex        =   155
            Top             =   3200
            Width           =   2415
         End
         Begin VB.Label lb1 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   480
            TabIndex        =   154
            Top             =   2160
            Width           =   2415
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Father Name :"
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
            Index           =   2
            Left            =   480
            TabIndex        =   153
            Top             =   2760
            Width           =   2415
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Student Name :"
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
            Index           =   1
            Left            =   480
            TabIndex        =   152
            Top             =   1755
            Width           =   2415
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Select Student Reg. No"
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
            Index           =   0
            Left            =   480
            TabIndex        =   151
            Top             =   600
            Width           =   2415
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00E0E0E0&
         Height          =   4815
         Left            =   10680
         TabIndex        =   181
         Top             =   3600
         Visible         =   0   'False
         Width           =   3320
         Begin VB.CommandButton ChameleonBtn4 
            Height          =   435
            Left            =   240
            MouseIcon       =   "REPORT_MASTER.frx":11992
            MousePointer    =   99  'Custom
            Picture         =   "REPORT_MASTER.frx":11AE4
            Style           =   1  'Graphical
            TabIndex        =   216
            Top             =   3900
            Width           =   2775
         End
         Begin VB.ComboBox Combo18 
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
            Left            =   480
            MouseIcon       =   "REPORT_MASTER.frx":129C3
            MousePointer    =   99  'Custom
            Style           =   2  'Dropdown List
            TabIndex        =   182
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Father Name :"
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
            Index           =   3
            Left            =   480
            TabIndex        =   191
            Top             =   2760
            Width           =   2415
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Select Student Reg. No"
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
            Index           =   5
            Left            =   480
            TabIndex        =   190
            Top             =   600
            Width           =   2415
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Student Name :"
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
            Index           =   4
            Left            =   480
            TabIndex        =   189
            Top             =   1755
            Width           =   2415
         End
         Begin VB.Label lb8 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   480
            TabIndex        =   188
            Top             =   2160
            Width           =   2415
         End
         Begin VB.Label lb9 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   480
            TabIndex        =   187
            Top             =   3200
            Width           =   2415
         End
         Begin VB.Label lb11 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Student Name :"
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
            Left            =   840
            TabIndex        =   186
            Top             =   4440
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lb10 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Student Name :"
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
            TabIndex        =   185
            Top             =   4395
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lb12 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   360
            TabIndex        =   184
            Top             =   120
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lb13 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Student Name :"
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
            Left            =   1200
            TabIndex        =   183
            Top             =   165
            Visible         =   0   'False
            Width           =   975
         End
      End
      Begin VB.CommandButton Command53 
         BackColor       =   &H8000000E&
         Caption         =   "Student Progress Report Card"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   11040
         MouseIcon       =   "REPORT_MASTER.frx":12B15
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":12C67
         Style           =   1  'Graphical
         TabIndex        =   180
         ToolTipText     =   "Details of Student  Progress report"
         Top             =   960
         Width           =   2535
      End
      Begin VB.CommandButton Command41 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   4320
         MouseIcon       =   "REPORT_MASTER.frx":14559
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":146AB
         Style           =   1  'Graphical
         TabIndex        =   161
         ToolTipText     =   "Import Questions From Xls or CSV Files , Directly Into database."
         Top             =   4440
         Width           =   2535
      End
      Begin VB.CommandButton Command40 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   7800
         MouseIcon       =   "REPORT_MASTER.frx":1674D
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":1689F
         Style           =   1  'Graphical
         TabIndex        =   160
         ToolTipText     =   "Export Questions in Format of Xls or Csv file."
         Top             =   4440
         Width           =   2535
      End
      Begin VB.CommandButton Command39 
         BackColor       =   &H8000000E&
         Caption         =   "       Student Ranking              "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   11040
         MouseIcon       =   "REPORT_MASTER.frx":18A47
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":18B99
         Style           =   1  'Graphical
         TabIndex        =   148
         ToolTipText     =   "Details of Student  rankwise.."
         Top             =   4440
         Width           =   2535
      End
      Begin VB.CommandButton Command38 
         BackColor       =   &H8000000E&
         Caption         =   "      Set Test Properties             "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   14280
         MouseIcon       =   "REPORT_MASTER.frx":19BF9
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":19D4B
         Style           =   1  'Graphical
         TabIndex        =   147
         ToolTipText     =   "Set The Property of Different Test Types"
         Top             =   4440
         Width           =   2535
      End
      Begin VB.CommandButton Command37 
         BackColor       =   &H8000000E&
         Caption         =   "      Organistaion Info.             "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   14280
         MouseIcon       =   "REPORT_MASTER.frx":1BD5F
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":1BEB1
         Style           =   1  'Graphical
         TabIndex        =   146
         ToolTipText     =   "Details About SpeedUp test Solutions."
         Top             =   960
         Width           =   2535
      End
      Begin VB.CommandButton Command36 
         BackColor       =   &H8000000E&
         Caption         =   "     Security Questions          "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   840
         MouseIcon       =   "REPORT_MASTER.frx":1DD0B
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":1DE5D
         Style           =   1  'Graphical
         TabIndex        =   145
         ToolTipText     =   "Security Questions are used While Recovering Password, In Case of user forget Login  password."
         Top             =   4440
         Width           =   2535
      End
      Begin VB.CommandButton Command35 
         BackColor       =   &H8000000E&
         Caption         =   "       Restore Backup       "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   7800
         MouseIcon       =   "REPORT_MASTER.frx":1FB97
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":1FCE9
         Style           =   1  'Graphical
         TabIndex        =   144
         ToolTipText     =   "Restore The Backup From backup File."
         Top             =   960
         Width           =   2535
      End
      Begin VB.CommandButton Command34 
         BackColor       =   &H8000000E&
         Caption         =   "       Create Backup                 "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   4320
         MouseIcon       =   "REPORT_MASTER.frx":211AF
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":21301
         Style           =   1  'Graphical
         TabIndex        =   143
         ToolTipText     =   "Create & Save Backup File on the Disk."
         Top             =   960
         Width           =   2535
      End
      Begin VB.CommandButton Command33 
         BackColor       =   &H8000000E&
         Caption         =   " Create Student ID Card   "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   840
         MouseIcon       =   "REPORT_MASTER.frx":22673
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":227C5
         Style           =   1  'Graphical
         TabIndex        =   142
         ToolTipText     =   "Generate Student ID Card easily."
         Top             =   960
         Width           =   2535
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   9695
      Left            =   2880
      TabIndex        =   135
      Top             =   840
      Width           =   17535
      Begin VB.Frame Frame9 
         BackColor       =   &H00E0E0E0&
         Height          =   4335
         Left            =   10365
         TabIndex        =   162
         Top             =   3300
         Visible         =   0   'False
         Width           =   3375
         Begin ChamaleonButton.ChameleonBtn ChameleonBtn3 
            Height          =   430
            Left            =   480
            TabIndex        =   170
            Top             =   3600
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   767
            BTYPE           =   1
            TX              =   "Generate Report"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "REPORT_MASTER.frx":24100
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.ComboBox Combo17 
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
            Left            =   480
            MouseIcon       =   "REPORT_MASTER.frx":2411C
            MousePointer    =   99  'Custom
            Style           =   2  'Dropdown List
            TabIndex        =   169
            Top             =   2820
            Width           =   2415
         End
         Begin VB.ComboBox Combo16 
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
            Left            =   480
            MouseIcon       =   "REPORT_MASTER.frx":2426E
            MousePointer    =   99  'Custom
            Style           =   2  'Dropdown List
            TabIndex        =   167
            Top             =   1860
            Width           =   2415
         End
         Begin VB.ComboBox Combo15 
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
            Left            =   480
            MouseIcon       =   "REPORT_MASTER.frx":243C0
            MousePointer    =   99  'Custom
            Style           =   2  'Dropdown List
            TabIndex        =   164
            Top             =   840
            Width           =   2415
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Select Month :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   168
            Top             =   2400
            Width           =   1935
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Select Year :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   166
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Order Status :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   165
            Top             =   360
            Width           =   1935
         End
         Begin VB.Shape Shape11 
            Height          =   4165
            Index           =   0
            Left            =   30
            Top             =   120
            Width           =   3295
         End
      End
      Begin VB.CommandButton Command51 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Delete Completed Order Lists"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   14160
         MouseIcon       =   "REPORT_MASTER.frx":24512
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":24664
         Style           =   1  'Graphical
         TabIndex        =   179
         ToolTipText     =   "Delete Those Order Details Whose Status Has  been Completed or Payment has Been Done."
         Top             =   3720
         Width           =   2655
      End
      Begin VB.CommandButton Command50 
         BackColor       =   &H00FFFFFF&
         Caption         =   "       Add New Subject                "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   3960
         MouseIcon       =   "REPORT_MASTER.frx":26088
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":261DA
         Style           =   1  'Graphical
         TabIndex        =   178
         ToolTipText     =   "Click Here To Add A New Subject."
         Top             =   6720
         Width           =   2655
      End
      Begin VB.CommandButton Command49 
         BackColor       =   &H00FFFFFF&
         Caption         =   "       Add New Package              "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   7320
         MouseIcon       =   "REPORT_MASTER.frx":27336
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":27488
         Style           =   1  'Graphical
         TabIndex        =   177
         ToolTipText     =   "Click Here To Add A New Package."
         Top             =   6720
         Width           =   2655
      End
      Begin VB.CommandButton Command48 
         BackColor       =   &H00FFFFFF&
         Caption         =   "       Add New Course         "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   600
         MouseIcon       =   "REPORT_MASTER.frx":2950B
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":2965D
         Style           =   1  'Graphical
         TabIndex        =   176
         ToolTipText     =   "Click Here To Add A New Course."
         Top             =   6720
         Width           =   2655
      End
      Begin VB.CommandButton Command47 
         BackColor       =   &H00FFFFFF&
         Caption         =   "      Add New Schedule             "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   10800
         MouseIcon       =   "REPORT_MASTER.frx":2A5B9
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":2A70B
         Style           =   1  'Graphical
         TabIndex        =   175
         ToolTipText     =   "Click Here To Add A New Schedule."
         Top             =   6720
         Width           =   2655
      End
      Begin VB.CommandButton Command46 
         BackColor       =   &H00FFFFFF&
         Caption         =   "      Delete All Students            "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   3960
         MouseIcon       =   "REPORT_MASTER.frx":2B3DC
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":2B52E
         Style           =   1  'Graphical
         TabIndex        =   174
         ToolTipText     =   "Delete all Student record on a Single Click."
         Top             =   3720
         Width           =   2655
      End
      Begin VB.CommandButton Command45 
         BackColor       =   &H00FFFFFF&
         Caption         =   "     Delete All Courses             "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   600
         MouseIcon       =   "REPORT_MASTER.frx":2CB45
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":2CC97
         Style           =   1  'Graphical
         TabIndex        =   173
         ToolTipText     =   "Delete All Courses on a Single Click."
         Top             =   3720
         Width           =   2655
      End
      Begin VB.CommandButton Command44 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Remove Expired Non Package Student"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   7320
         MouseIcon       =   "REPORT_MASTER.frx":2E1EE
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":2E340
         Style           =   1  'Graphical
         TabIndex        =   172
         ToolTipText     =   "Delete All Those Non Package Student Whose Validity is expired."
         Top             =   3720
         Width           =   2655
      End
      Begin VB.CommandButton Command43 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Remove All Expired Package Student"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   10800
         MouseIcon       =   "REPORT_MASTER.frx":30DB9
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":30F0B
         Style           =   1  'Graphical
         TabIndex        =   171
         ToolTipText     =   "Remove All Those Package Benifited Student Whose Package Validity Has Been Expired."
         Top             =   3720
         Width           =   2655
      End
      Begin VB.CommandButton Command42 
         BackColor       =   &H00FFFFFF&
         Caption         =   "     Student's Login Info           "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   14160
         MouseIcon       =   "REPORT_MASTER.frx":33C07
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":33D59
         Style           =   1  'Graphical
         TabIndex        =   163
         ToolTipText     =   "All Login Details of Student"
         Top             =   720
         Width           =   2655
      End
      Begin VB.CommandButton Command32 
         BackColor       =   &H00FFFFFF&
         Caption         =   "      Client Order  Lists        "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   10800
         MouseIcon       =   "REPORT_MASTER.frx":35266
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":353B8
         Style           =   1  'Graphical
         TabIndex        =   140
         ToolTipText     =   "Order List given By Client ."
         Top             =   720
         Width           =   2655
      End
      Begin VB.CommandButton Command31 
         BackColor       =   &H80000014&
         Caption         =   "     All Client's  Records           "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   7320
         MouseIcon       =   "REPORT_MASTER.frx":3691B
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":36A6D
         Style           =   1  'Graphical
         TabIndex        =   139
         ToolTipText     =   "List of All clients."
         Top             =   720
         Width           =   2655
      End
      Begin VB.CommandButton Command30 
         BackColor       =   &H00FFFFFF&
         Caption         =   "       User's Login Info                "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   14160
         MouseIcon       =   "REPORT_MASTER.frx":37EC5
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":38017
         Style           =   1  'Graphical
         TabIndex        =   138
         ToolTipText     =   "All Login Details of Users"
         Top             =   6720
         Width           =   2655
      End
      Begin VB.CommandButton Command29 
         BackColor       =   &H00FFFFFF&
         Caption         =   "       All User's Records             "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   3960
         MouseIcon       =   "REPORT_MASTER.frx":398CE
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":39A20
         Style           =   1  'Graphical
         TabIndex        =   137
         ToolTipText     =   "List of All Allowed users."
         Top             =   720
         Width           =   2655
      End
      Begin VB.CommandButton Command28 
         BackColor       =   &H00FFFFFF&
         Caption         =   "    All Admin's Records           "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   600
         MouseIcon       =   "REPORT_MASTER.frx":3BA7A
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":3BBCC
         Style           =   1  'Graphical
         TabIndex        =   136
         ToolTipText     =   "List Of All Admins for This System."
         Top             =   720
         Width           =   2655
      End
   End
   Begin VB.Frame AccountFrame 
      BackColor       =   &H00E0E0E0&
      Height          =   9695
      Left            =   2880
      TabIndex        =   70
      Top             =   840
      Width           =   17535
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Height          =   2175
         Left            =   13200
         TabIndex        =   132
         Top             =   3960
         Visible         =   0   'False
         Width           =   3255
         Begin VB.CommandButton ChameleonBtn1 
            Height          =   425
            Left            =   360
            MouseIcon       =   "REPORT_MASTER.frx":3C967
            MousePointer    =   99  'Custom
            Picture         =   "REPORT_MASTER.frx":3CAB9
            Style           =   1  'Graphical
            TabIndex        =   220
            Top             =   1440
            Width           =   2535
         End
         Begin VB.Label Label35 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   134
            Top             =   780
            Width           =   1815
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Due Amount :"
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
            Index           =   7
            Left            =   600
            TabIndex        =   133
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Height          =   4170
         Left            =   9000
         TabIndex        =   125
         Top             =   3960
         Visible         =   0   'False
         Width           =   3495
         Begin VB.CommandButton ovrlbtn 
            Height          =   425
            Left            =   480
            MouseIcon       =   "REPORT_MASTER.frx":3D654
            MousePointer    =   99  'Custom
            Picture         =   "REPORT_MASTER.frx":3D7A6
            Style           =   1  'Graphical
            TabIndex        =   219
            Top             =   3360
            Width           =   2535
         End
         Begin VB.ComboBox Combo13 
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
            Left            =   480
            MouseIcon       =   "REPORT_MASTER.frx":3E341
            MousePointer    =   99  'Custom
            Style           =   2  'Dropdown List
            TabIndex        =   130
            Top             =   720
            Width           =   1695
         End
         Begin VB.ComboBox Combo12 
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
            Left            =   480
            MouseIcon       =   "REPORT_MASTER.frx":3E493
            MousePointer    =   99  'Custom
            Style           =   2  'Dropdown List
            TabIndex        =   127
            Top             =   1695
            Width           =   2535
         End
         Begin VB.ComboBox Combo11 
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
            Left            =   480
            MouseIcon       =   "REPORT_MASTER.frx":3E5E5
            MousePointer    =   99  'Custom
            Style           =   2  'Dropdown List
            TabIndex        =   126
            Top             =   2655
            Width           =   2535
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "SEARCH BY  :"
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
            Index           =   6
            Left            =   480
            TabIndex        =   131
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label lbl12 
            BackStyle       =   0  'Transparent
            Caption         =   "Select Year :"
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
            Left            =   480
            TabIndex        =   129
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label lbl11 
            BackStyle       =   0  'Transparent
            Caption         =   "Select month :"
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
            Left            =   480
            TabIndex        =   128
            Top             =   2160
            Width           =   1935
         End
      End
      Begin MSACAL.Calendar cld3 
         Height          =   2895
         Left            =   5880
         TabIndex        =   123
         Top             =   6120
         Visible         =   0   'False
         Width           =   3135
         _Version        =   524288
         _ExtentX        =   5530
         _ExtentY        =   5106
         _StockProps     =   1
         BackColor       =   14737632
         Year            =   2019
         Month           =   7
         Day             =   8
         DayLength       =   1
         MonthLength     =   2
         DayFontColor    =   0
         FirstDay        =   1
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   0   'False
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSACAL.Calendar cld2 
         Height          =   2895
         Left            =   5880
         TabIndex        =   124
         Top             =   5520
         Visible         =   0   'False
         Width           =   3135
         _Version        =   524288
         _ExtentX        =   5530
         _ExtentY        =   5106
         _StockProps     =   1
         BackColor       =   14737632
         Year            =   2019
         Month           =   7
         Day             =   8
         DayLength       =   1
         MonthLength     =   2
         DayFontColor    =   0
         FirstDay        =   1
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   0   'False
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame ExpenseFrame 
         BackColor       =   &H00E0E0E0&
         Height          =   5300
         Left            =   4560
         TabIndex        =   103
         Top             =   3960
         Visible         =   0   'False
         Width           =   4335
         Begin VB.CommandButton Command27 
            Height          =   425
            Left            =   480
            MouseIcon       =   "REPORT_MASTER.frx":3E737
            MousePointer    =   99  'Custom
            Picture         =   "REPORT_MASTER.frx":3E889
            Style           =   1  'Graphical
            TabIndex        =   221
            Top             =   3000
            Width           =   3495
         End
         Begin VB.ComboBox Combo8 
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
            Left            =   2040
            MouseIcon       =   "REPORT_MASTER.frx":3F5A7
            MousePointer    =   99  'Custom
            Style           =   2  'Dropdown List
            TabIndex        =   116
            Top             =   480
            Width           =   1695
         End
         Begin VB.CommandButton Command26 
            BackColor       =   &H8000000E&
            Caption         =   "This Year Expense  Report"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   580
            Left            =   2160
            MouseIcon       =   "REPORT_MASTER.frx":3F6F9
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   112
            ToolTipText     =   "This Year Expense Report"
            Top             =   4320
            Width           =   1815
         End
         Begin VB.CommandButton Command25 
            BackColor       =   &H8000000E&
            Caption         =   "This Month Expense Report"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   580
            Left            =   480
            MouseIcon       =   "REPORT_MASTER.frx":3F84B
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   111
            ToolTipText     =   "This Month expense Report"
            Top             =   4320
            Width           =   1695
         End
         Begin VB.CommandButton Command24 
            BackColor       =   &H8000000E&
            Caption         =   "This Week Expense  Report"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   580
            Left            =   2160
            MouseIcon       =   "REPORT_MASTER.frx":3F99D
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   110
            ToolTipText     =   "This Week Expense Report"
            Top             =   3750
            Width           =   1815
         End
         Begin VB.CommandButton Command23 
            BackColor       =   &H8000000E&
            Caption         =   "Today  Expense Report"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   580
            Left            =   480
            MouseIcon       =   "REPORT_MASTER.frx":3FAEF
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   109
            ToolTipText     =   "Today Expense Report"
            Top             =   3750
            Width           =   1695
         End
         Begin VB.Frame Exp1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   1815
            Left            =   600
            TabIndex        =   104
            Top             =   840
            Width           =   3255
            Begin VB.TextBox Text8 
               Alignment       =   2  'Center
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
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   106
               Text            =   "Text4"
               Top             =   240
               Width           =   1695
            End
            Begin VB.TextBox Text7 
               Alignment       =   2  'Center
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
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   105
               Text            =   "Text4"
               Top             =   960
               Width           =   1695
            End
            Begin VB.Label Label22 
               BackStyle       =   0  'Transparent
               Caption         =   "From     :"
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
               Index           =   4
               Left            =   450
               TabIndex        =   108
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label22 
               BackStyle       =   0  'Transparent
               Caption         =   "To         :"
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
               Index           =   3
               Left            =   450
               TabIndex        =   107
               Top             =   960
               Width           =   975
            End
            Begin VB.Line Line11 
               BorderColor     =   &H00808080&
               BorderWidth     =   2
               Index           =   1
               X1              =   260
               X2              =   260
               Y1              =   1100
               Y2              =   340
            End
            Begin VB.Line Line12 
               BorderColor     =   &H00808080&
               BorderWidth     =   2
               Index           =   3
               X1              =   240
               X2              =   400
               Y1              =   360
               Y2              =   360
            End
            Begin VB.Line Line12 
               BorderColor     =   &H00808080&
               BorderWidth     =   2
               Index           =   2
               X1              =   240
               X2              =   400
               Y1              =   1115
               Y2              =   1115
            End
         End
         Begin VB.Frame Exp2 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   1770
            Left            =   600
            TabIndex        =   117
            Top             =   960
            Visible         =   0   'False
            Width           =   3255
            Begin VB.ComboBox Combo10 
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
               Left            =   0
               MouseIcon       =   "REPORT_MASTER.frx":3FC41
               MousePointer    =   99  'Custom
               Style           =   2  'Dropdown List
               TabIndex        =   119
               Top             =   1335
               Width           =   3135
            End
            Begin VB.ComboBox Combo9 
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
               Left            =   0
               MouseIcon       =   "REPORT_MASTER.frx":3FD93
               MousePointer    =   99  'Custom
               Style           =   2  'Dropdown List
               TabIndex        =   118
               Top             =   495
               Width           =   3135
            End
            Begin VB.Label Label34 
               BackStyle       =   0  'Transparent
               Caption         =   "Select month :"
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
               TabIndex        =   121
               Top             =   960
               Width           =   1935
            End
            Begin VB.Label Label33 
               BackStyle       =   0  'Transparent
               Caption         =   "Select Year :"
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
               TabIndex        =   120
               Top             =   120
               Width           =   1335
            End
         End
         Begin VB.Frame Exp3 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   925
            Left            =   600
            TabIndex        =   113
            Top             =   960
            Visible         =   0   'False
            Width           =   3255
            Begin VB.ComboBox Combo7 
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
               Left            =   0
               MouseIcon       =   "REPORT_MASTER.frx":3FEE5
               MousePointer    =   99  'Custom
               Style           =   2  'Dropdown List
               TabIndex        =   114
               Top             =   500
               Width           =   3135
            End
            Begin VB.Label Label32 
               BackStyle       =   0  'Transparent
               Caption         =   "Select Year :"
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
               TabIndex        =   115
               Top             =   120
               Width           =   1215
            End
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "SEARCH BY  :"
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
            Index           =   5
            Left            =   600
            TabIndex        =   122
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.TextBox Text5a 
         Height          =   285
         Left            =   7800
         TabIndex        =   102
         Text            =   "Text7"
         Top             =   4200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text4a 
         Height          =   285
         Left            =   7800
         TabIndex        =   101
         Text            =   "Text7"
         Top             =   3720
         Visible         =   0   'False
         Width           =   1695
      End
      Begin MSACAL.Calendar cld1 
         Height          =   2775
         Left            =   1800
         TabIndex        =   89
         Top             =   6120
         Visible         =   0   'False
         Width           =   3135
         _Version        =   524288
         _ExtentX        =   5530
         _ExtentY        =   4895
         _StockProps     =   1
         BackColor       =   14737632
         Year            =   2019
         Month           =   7
         Day             =   8
         DayLength       =   1
         MonthLength     =   2
         DayFontColor    =   0
         FirstDay        =   1
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   0   'False
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSACAL.Calendar cld 
         Height          =   3255
         Left            =   1800
         TabIndex        =   96
         Top             =   5400
         Visible         =   0   'False
         Width           =   3135
         _Version        =   524288
         _ExtentX        =   5530
         _ExtentY        =   5741
         _StockProps     =   1
         BackColor       =   14737632
         Year            =   2019
         Month           =   7
         Day             =   8
         DayLength       =   1
         MonthLength     =   2
         DayFontColor    =   0
         FirstDay        =   1
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   0   'False
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton Command17 
         BackColor       =   &H8000000E&
         Caption         =   "DUE   PAYMENT  INFO."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   13320
         MouseIcon       =   "REPORT_MASTER.frx":40037
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":40189
         Style           =   1  'Graphical
         TabIndex        =   78
         ToolTipText     =   "Due Payment Info"
         Top             =   1080
         Width           =   3015
      End
      Begin VB.CommandButton Command16 
         BackColor       =   &H8000000E&
         Caption         =   "OVERALL   CALCULATION"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   9240
         MouseIcon       =   "REPORT_MASTER.frx":4166D
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":417BF
         Style           =   1  'Graphical
         TabIndex        =   77
         ToolTipText     =   "Overall Calculation (Including Income & Expense )"
         Top             =   1080
         Width           =   3015
      End
      Begin VB.CommandButton Command15 
         BackColor       =   &H8000000E&
         Caption         =   "EXPENSE"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   5160
         MouseIcon       =   "REPORT_MASTER.frx":4394B
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":43A9D
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   "Expense Details Of Orgnisation"
         Top             =   1080
         Width           =   3015
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H8000000E&
         Caption         =   "INCOME"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   1080
         MouseIcon       =   "REPORT_MASTER.frx":45FCB
         MousePointer    =   99  'Custom
         Picture         =   "REPORT_MASTER.frx":4611D
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Income details Of Organistaion"
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Frame IncomeFrame 
         BackColor       =   &H00E0E0E0&
         Height          =   5300
         Left            =   480
         TabIndex        =   79
         Top             =   3960
         Visible         =   0   'False
         Width           =   4335
         Begin VB.CommandButton Command20 
            Height          =   450
            Left            =   480
            MouseIcon       =   "REPORT_MASTER.frx":48180
            MousePointer    =   99  'Custom
            Picture         =   "REPORT_MASTER.frx":482D2
            Style           =   1  'Graphical
            TabIndex        =   222
            Top             =   3000
            Width           =   3495
         End
         Begin VB.Frame inc1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   1815
            Left            =   600
            TabIndex        =   84
            Top             =   840
            Width           =   3255
            Begin VB.TextBox Text5 
               Alignment       =   2  'Center
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
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   86
               Text            =   "Text4"
               Top             =   960
               Width           =   1695
            End
            Begin VB.TextBox Text4 
               Alignment       =   2  'Center
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
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   85
               Text            =   "Text4"
               Top             =   240
               Width           =   1695
            End
            Begin VB.Line Line12 
               BorderColor     =   &H00808080&
               BorderWidth     =   2
               Index           =   1
               X1              =   240
               X2              =   400
               Y1              =   1115
               Y2              =   1115
            End
            Begin VB.Line Line12 
               BorderColor     =   &H00808080&
               BorderWidth     =   2
               Index           =   0
               X1              =   240
               X2              =   400
               Y1              =   360
               Y2              =   360
            End
            Begin VB.Line Line11 
               BorderColor     =   &H00808080&
               BorderWidth     =   2
               Index           =   0
               X1              =   240
               X2              =   240
               Y1              =   1100
               Y2              =   340
            End
            Begin VB.Label Label22 
               BackStyle       =   0  'Transparent
               Caption         =   "To         :"
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
               Index           =   2
               Left            =   435
               TabIndex        =   88
               Top             =   960
               Width           =   975
            End
            Begin VB.Label Label22 
               BackStyle       =   0  'Transparent
               Caption         =   "From     :"
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
               Index           =   1
               Left            =   435
               TabIndex        =   87
               Top             =   240
               Width           =   855
            End
         End
         Begin VB.CommandButton Command18 
            BackColor       =   &H8000000E&
            Caption         =   "Today income Report"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   580
            Left            =   480
            MouseIcon       =   "REPORT_MASTER.frx":48FF0
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   80
            ToolTipText     =   "Today Income Report"
            Top             =   3750
            Width           =   1695
         End
         Begin VB.CommandButton Command19 
            BackColor       =   &H8000000E&
            Caption         =   "This Week Income Report"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   580
            Left            =   2160
            MouseIcon       =   "REPORT_MASTER.frx":49142
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   81
            Top             =   3750
            Width           =   1815
         End
         Begin VB.CommandButton Command22 
            BackColor       =   &H8000000E&
            Caption         =   "This Month Income Report"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   580
            Left            =   480
            MouseIcon       =   "REPORT_MASTER.frx":49294
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   100
            ToolTipText     =   "Income report of This Month"
            Top             =   4320
            Width           =   1695
         End
         Begin VB.CommandButton Command21 
            BackColor       =   &H8000000E&
            Caption         =   "This Year Income Report"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   580
            Left            =   2160
            MouseIcon       =   "REPORT_MASTER.frx":493E6
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   99
            ToolTipText     =   "Income Report of This Year"
            Top             =   4320
            Width           =   1815
         End
         Begin VB.Frame inc2 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   925
            Left            =   600
            TabIndex        =   93
            Top             =   960
            Visible         =   0   'False
            Width           =   3255
            Begin VB.ComboBox Combo6 
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
               Left            =   0
               MouseIcon       =   "REPORT_MASTER.frx":49538
               MousePointer    =   99  'Custom
               Style           =   2  'Dropdown List
               TabIndex        =   94
               Top             =   500
               Width           =   3135
            End
            Begin VB.Label Label30 
               BackStyle       =   0  'Transparent
               Caption         =   "Select Year :"
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
               TabIndex        =   95
               Top             =   120
               Width           =   1215
            End
         End
         Begin VB.ComboBox Combo3 
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
            Left            =   2040
            MouseIcon       =   "REPORT_MASTER.frx":4968A
            MousePointer    =   99  'Custom
            Style           =   2  'Dropdown List
            TabIndex        =   82
            Top             =   480
            Width           =   1695
         End
         Begin VB.Frame inc3 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   1770
            Left            =   600
            TabIndex        =   90
            Top             =   960
            Visible         =   0   'False
            Width           =   3255
            Begin VB.ComboBox ComboYr 
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
               Left            =   0
               MouseIcon       =   "REPORT_MASTER.frx":497DC
               MousePointer    =   99  'Custom
               Style           =   2  'Dropdown List
               TabIndex        =   97
               Top             =   495
               Width           =   3135
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
               Left            =   0
               MouseIcon       =   "REPORT_MASTER.frx":4992E
               MousePointer    =   99  'Custom
               Style           =   2  'Dropdown List
               TabIndex        =   91
               Top             =   1335
               Width           =   3135
            End
            Begin VB.Label Label31 
               BackStyle       =   0  'Transparent
               Caption         =   "Select Year :"
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
               TabIndex        =   98
               Top             =   120
               Width           =   1335
            End
            Begin VB.Label MAINJ 
               BackStyle       =   0  'Transparent
               Caption         =   "Select month :"
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
               TabIndex        =   92
               Top             =   960
               Width           =   1935
            End
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "SEARCH BY  :"
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
            Index           =   0
            Left            =   600
            TabIndex        =   83
            Top             =   480
            Width           =   1335
         End
      End
   End
   Begin VB.Frame frm 
      BackColor       =   &H00E0E0E0&
      Height          =   9695
      Index           =   0
      Left            =   2880
      TabIndex        =   9
      Top             =   840
      Width           =   17535
      Begin VB.TextBox Text3 
         Height          =   525
         Left            =   8400
         TabIndex        =   44
         Text            =   "Text3"
         Top             =   3480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   525
         Left            =   7080
         TabIndex        =   43
         Text            =   "Text2"
         Top             =   3480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame frm 
         BackColor       =   &H00E0E0E0&
         Height          =   3050
         Index           =   2
         Left            =   2640
         TabIndex        =   35
         Top             =   6480
         Width           =   2295
         Begin VB.ComboBox Combo5 
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
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   1440
            Width           =   1695
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Print All Topics"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            MouseIcon       =   "REPORT_MASTER.frx":49A80
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   2400
            Width           =   1815
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Print"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            MouseIcon       =   "REPORT_MASTER.frx":49BD2
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   1920
            Width           =   1335
         End
         Begin VB.ComboBox Combo 
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
            Index           =   1
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Subject"
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
            Left            =   360
            TabIndex        =   41
            Top             =   1080
            Width           =   1350
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Course"
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
            Left            =   360
            TabIndex        =   39
            Top             =   240
            Width           =   1320
         End
      End
      Begin VB.Frame frm 
         BackColor       =   &H00E0E0E0&
         Height          =   2415
         Index           =   1
         Left            =   1800
         TabIndex        =   30
         Top             =   2640
         Width           =   2295
         Begin VB.CommandButton Command9 
            Caption         =   "Print All Subjects"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            MouseIcon       =   "REPORT_MASTER.frx":49D24
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   1680
            Width           =   1815
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Print"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            MouseIcon       =   "REPORT_MASTER.frx":49E76
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   1200
            Width           =   1335
         End
         Begin VB.ComboBox Combo 
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
            Index           =   0
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Course"
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
            Left            =   360
            TabIndex        =   34
            Top             =   240
            Width           =   1320
         End
      End
      Begin VB.Frame frm 
         BackColor       =   &H00E0E0E0&
         Height          =   2415
         Index           =   4
         Left            =   12850
         TabIndex        =   25
         Top             =   2640
         Width           =   2295
         Begin VB.ComboBox Combo 
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
            Index           =   3
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   600
            Width           =   1695
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Print"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            MouseIcon       =   "REPORT_MASTER.frx":49FC8
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   1200
            Width           =   1335
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Print All Packages"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            MouseIcon       =   "REPORT_MASTER.frx":4A11A
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   1680
            Width           =   1815
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Course"
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
            Left            =   360
            TabIndex        =   29
            Top             =   240
            Width           =   1320
         End
      End
      Begin VB.Frame frm 
         BackColor       =   &H00E0E0E0&
         Height          =   2415
         Index           =   3
         Left            =   12000
         TabIndex        =   20
         Top             =   6480
         Width           =   2295
         Begin VB.CommandButton Command5 
            Caption         =   "Print All Schedules"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            MouseIcon       =   "REPORT_MASTER.frx":4A26C
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   1680
            Width           =   1815
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Print"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            MouseIcon       =   "REPORT_MASTER.frx":4A3BE
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   1200
            Width           =   1335
         End
         Begin VB.ComboBox Combo 
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
            Index           =   2
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Course"
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
            Left            =   360
            TabIndex        =   22
            Top             =   240
            Width           =   1320
         End
      End
      Begin MyEllipticButton.EllipticButton cmd1 
         Height          =   2055
         Index           =   1
         Left            =   4080
         TabIndex        =   10
         ToolTipText     =   "Subjects of a particular class can be shown here."
         Top             =   2760
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   3625
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
         Picture         =   "REPORT_MASTER.frx":4A510
         DisabledPicture =   "REPORT_MASTER.frx":4B0CB
         DownPicture     =   "REPORT_MASTER.frx":4B0E7
         MousePointer    =   99
         MouseIcon       =   "REPORT_MASTER.frx":4B103
         Caption         =   ""
      End
      Begin MyEllipticButton.EllipticButton cmd1 
         Height          =   2055
         Index           =   0
         Left            =   7500
         TabIndex        =   11
         ToolTipText     =   "See all Courses."
         Top             =   15
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   3625
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
         Picture         =   "REPORT_MASTER.frx":4B265
         DisabledPicture =   "REPORT_MASTER.frx":4BE69
         DownPicture     =   "REPORT_MASTER.frx":4BE85
         MousePointer    =   99
         MouseIcon       =   "REPORT_MASTER.frx":4BEA1
         Caption         =   ""
      End
      Begin MyEllipticButton.EllipticButton cmd1 
         Height          =   2055
         Index           =   2
         Left            =   4920
         TabIndex        =   12
         ToolTipText     =   "Topic came under Subjects"
         Top             =   6480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   3625
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
         Picture         =   "REPORT_MASTER.frx":4C003
         DisabledPicture =   "REPORT_MASTER.frx":4CD63
         DownPicture     =   "REPORT_MASTER.frx":4CD7F
         MousePointer    =   99
         MouseIcon       =   "REPORT_MASTER.frx":4CD9B
         Caption         =   ""
      End
      Begin MyEllipticButton.EllipticButton cmd1 
         Height          =   2055
         Index           =   4
         Left            =   10800
         TabIndex        =   13
         ToolTipText     =   "List Of All Packages of a Course"
         Top             =   2760
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   3625
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
         Picture         =   "REPORT_MASTER.frx":4CEFD
         DisabledPicture =   "REPORT_MASTER.frx":4DE5C
         DownPicture     =   "REPORT_MASTER.frx":4DE78
         MousePointer    =   99
         MouseIcon       =   "REPORT_MASTER.frx":4DE94
         Caption         =   ""
      End
      Begin MyEllipticButton.EllipticButton cmd1 
         Height          =   2055
         Index           =   3
         Left            =   9960
         TabIndex        =   14
         Top             =   6480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   3625
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
         Picture         =   "REPORT_MASTER.frx":4DFF6
         DisabledPicture =   "REPORT_MASTER.frx":4ECD7
         DownPicture     =   "REPORT_MASTER.frx":4ECF3
         MousePointer    =   99
         MouseIcon       =   "REPORT_MASTER.frx":4ED0F
         Caption         =   ""
      End
      Begin MyEllipticButton.EllipticButton cmd 
         Height          =   1575
         Left            =   7680
         TabIndex        =   42
         ToolTipText     =   "Refresh Again"
         Top             =   4080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   2778
         BackColor       =   -2147483637
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "REPORT_MASTER.frx":4EE71
         DisabledPicture =   "REPORT_MASTER.frx":4EE8D
         DownPicture     =   "REPORT_MASTER.frx":4EEA9
         MousePointer    =   99
         MouseIcon       =   "REPORT_MASTER.frx":4EEC5
         Caption         =   " Master Entry"
      End
      Begin VB.Shape Shape10 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         Height          =   1455
         Left            =   7530
         Shape           =   3  'Circle
         Top             =   4120
         Width           =   1845
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Height          =   2175
         Left            =   10000
         Shape           =   3  'Circle
         Top             =   6415
         Width           =   1935
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Height          =   2175
         Left            =   10850
         Shape           =   3  'Circle
         Top             =   2690
         Width           =   1935
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Height          =   2175
         Left            =   7540
         Shape           =   3  'Circle
         Top             =   -45
         Width           =   1935
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Height          =   2175
         Left            =   4970
         Shape           =   3  'Circle
         Top             =   6410
         Width           =   1935
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Height          =   2175
         Left            =   4120
         Shape           =   3  'Circle
         Top             =   2700
         Width           =   1935
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Topics"
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
         Left            =   5640
         TabIndex        =   19
         ToolTipText     =   "Topic came under Subjects"
         Top             =   8520
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Courses"
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
         Left            =   8100
         TabIndex        =   18
         ToolTipText     =   "See all Courses."
         Top             =   1960
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subjects"
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
         Left            =   4920
         TabIndex        =   17
         ToolTipText     =   "Subjects of a particular class can be shown here."
         Top             =   4800
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Packages"
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
         Left            =   11280
         TabIndex        =   16
         ToolTipText     =   "List Of All Packages of a Course"
         Top             =   4800
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Schedules"
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
         Left            =   10560
         TabIndex        =   15
         Top             =   8520
         Width           =   975
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Height          =   7575
         Left            =   4560
         Shape           =   3  'Circle
         Top             =   1200
         Width           =   7935
      End
   End
End
Attribute VB_Name = "FrmReportMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim navigation_counter As Integer, i As Integer, CrentScene As Integer

Private Sub a1_Click(Index As Integer)
studnt_Click (Index)
End Sub

Private Sub AccountFrame_Click()
Frame5.Visible = False
Frame4.Visible = False
IncomeFrame.Visible = False
ExpenseFrame.Visible = False
End Sub

Private Sub ChameleonBtn1_Click()
If Val(Label35.Caption) = 0 Or Label35.Caption = "" Then
 MsgBox "No Dues Amount..", vbInformation + vbOKOnly, ""
Else
 RptDue.Show 1, MDI
End If
End Sub

Private Sub ChameleonBtn2_Click()
If Combo14.Text = "" Or lb1.Caption = "" Then
 MsgBox "Select Student Reg.No First,", vbInformation + vbOKOnly, ""
 Combo14.SetFocus
 Exit Sub
End If
DV.Stud_Id_Gen Combo14.Text
Set idcard.Sections("section1").Controls.Item("stphoto").Picture = LoadPicture(lb6.Caption)
idcard.Sections("section1").Controls("text2").Caption = Combo14.Text
idcard.Sections("section1").Controls("text3").Caption = lb1.Caption
idcard.Sections("section1").Controls("text4").Caption = lb2.Caption
idcard.Sections("section1").Controls("text5").Caption = lb3.Caption
idcard.Sections("section1").Controls("text6").Caption = lb4.Caption
idcard.Sections("section1").Controls("text7").Caption = lb5.Caption
idcard.Show vbModal, MDI
idcard.Refresh
DV.rsStud_Id_Gen.Close
End Sub

Private Sub ChameleonBtn3_Click()
If Combo15.Text = "" Then
MsgBox "Select Order status..", vbInformation + vbOKOnly, ""
Combo15.SetFocus
Exit Sub
End If
If Combo16.Text = "" Then
MsgBox "Select Year from List..", vbInformation + vbOKOnly, ""
Combo16.SetFocus
Exit Sub
End If
If Combo17.Text = "" Then
MsgBox "Select Month from list.", vbInformation + vbOKOnly, ""
Combo17.SetFocus
Exit Sub
End If
Text4a.Text = "1-" & Left(Combo17.Text, 3) & "-" & Combo16.Text
 If Combo17.ListIndex = 0 Or Combo17.ListIndex Mod 2 = 0 Then '31
    Text5a.Text = "31-" & Left(Combo17.Text, 3) & "-" & Combo16.Text
  ElseIf Combo17.ListIndex = 1 Then '30 / 28
   If Val(Combo16.Text) Mod 4 = 0 Then
    Text5a.Text = "29-" & Left(Combo17.Text, 3) & "-" & Combo16.Text
   Else
    Text5a.Text = "28-" & Left(Combo17.Text, 3) & "-" & Combo16.Text
   End If
  Else
    Text5a.Text = "30-" & Left(Combo17.Text, 3) & "-" & Combo16.Text
  End If
  Text4.Text = "[ Order List ] - " & Left(Combo17.Text, 3) & " " & Combo16.Text
DV.CmdOrderList UCase(Combo15.Text), Text4a.Text, Text5a.Text
 RptOrderList.Sections("section4").Controls("label4").Caption = Text4.Text
 RptOrderList.Show 1, MDI
DV.rsCmdOrderList.Close
End Sub

Private Sub ChameleonBtn4_Click() 'Student Progress Report
If Combo18.Text = "" Or lb8.Caption = "" Then
 MsgBox "Select Student Reg.No First,", vbInformation + vbOKOnly, ""
 Combo18.SetFocus
 Exit Sub
End If
DV.CmdStudProgress Combo18.Text
Set RptStudProgress.Sections("section4").Controls.Item("Image2").Picture = LoadPicture(lb13.Caption)
RptStudProgress.Sections("section4").Controls("n1").Caption = lb8.Caption
RptStudProgress.Sections("section4").Controls("n2").Caption = lb9.Caption
RptStudProgress.Sections("section4").Controls("n3").Caption = lb10.Caption
RptStudProgress.Sections("section4").Controls("n4").Caption = lb11.Caption
RptStudProgress.Sections("section4").Controls("n5").Caption = lb12.Caption
RptStudProgress.Show vbModal, MDI
RptStudProgress.Refresh
DV.rsCmdStudProgress.Close
End Sub

Private Sub cld_Click()
Text4.Text = cld.Value
End Sub

Private Sub cld1_Click()
If cld1.Value < cld.Value Then
 MsgBox "Cannot Select Date Older Than First Date.Select Date After Previous Date.", vbInformation + vbOKOnly, ""
 Text5.SetFocus
Else
 Text5.Text = cld1.Value
End If
End Sub

Private Sub cld2_Click()
Text8.Text = cld2.Value
End Sub

Private Sub cld3_Click()
If cld2.Value > cld3.Value Then
 MsgBox "Cannot Select Date Older Than First Date.Select Date After Previous Date.", vbInformation + vbOKOnly, ""
 Text7.SetFocus
Else
 Text7.Text = cld3.Value
End If
End Sub

Private Sub cmd_Click()
 Text2.Text = ""
 Text3.Text = ""
 For i = 1 To 4
  frm(i).Visible = False
 Next i
End Sub

Private Sub cmd1_Click(Index As Integer)
cmd_Click
For i = 1 To 4
 If i = Index Then
  frm(Index).Visible = True
 ElseIf Index > 0 Then
  frm(i).Visible = False
 End If
Next i
If Index = 0 Then
 RptCourse.Show 1, MDI
End If
End Sub

Private Sub Combo_Click(Index As Integer)
Set r = New ADODB.Recordset
Set r = c.Execute("select c_id from course where c_nm='" & Combo(Index).Text & "' ")
If r.EOF = False Then
 Text2.Text = r.Fields(0)
End If
If Index = 1 Then
 Combo5.Clear
 Set r1 = c.Execute("select sub_nm from sub where c_id='" & Text2.Text & "' ")
 While r1.EOF = False
  Combo5.AddItem r1.Fields(0)
 r1.MoveNext
 Wend
End If
End Sub

Private Sub Combo1_Click()
CrentScene = 9
If Combo1.ListIndex <> 3 Then
 NotDateFrame.Visible = True
 Dateframe.Visible = False
Else
 NotDateFrame.Visible = False
 Dateframe.Visible = True
End If
Combo2.Clear
Set r = New ADODB.Recordset
If Combo1.ListIndex = 0 Then
 Set r = c.Execute("select rstud_reg_no from rstud")
 While r.EOF = False
  Combo2.AddItem r.Fields(0)
 r.MoveNext
 Wend
ElseIf Combo1.ListIndex = 1 Then
 Set r = c.Execute("select rstud_nm from rstud")
 While r.EOF = False
  Combo2.AddItem r.Fields(0)
 r.MoveNext
 Wend
ElseIf Combo1.ListIndex = 2 Then
 Set r = c.Execute("select c_nm from course")
 While r.EOF = False
  Combo2.AddItem r.Fields(0)
 r.MoveNext
 Wend
End If
End Sub

Private Sub Combo11_Click()
Text4a.Text = "1-" & Left(Combo11.Text, 3) & "-" & Combo12.Text
  If Combo11.ListIndex = 0 Or Combo11.ListIndex Mod 2 = 0 Then '31
    Text5a.Text = "31-" & Left(Combo11.Text, 3) & "-" & Combo12.Text
  ElseIf Combo11.ListIndex = 1 Then '30 / 28
   If Val(Combo11.Text) Mod 4 = 0 Then
    Text5a.Text = "29-" & Left(Combo11.Text, 3) & "-" & Combo12.Text
   Else
    Text5a.Text = "28-" & Left(Combo11.Text, 3) & "-" & Combo12.Text
   End If
  Else
    Text5a.Text = "30-" & Left(Combo11.Text, 3) & "-" & Combo12.Text
  End If
End Sub

Private Sub Combo12_Click()
Text4a.Text = "1-Jan-" & Combo12.Text
Text5a.Text = "31-Dec-" & Combo12.Text
End Sub

Private Sub Combo13_Click()
If Combo13.ListIndex = 1 Then
 lbl11.Visible = False
 Combo11.Visible = False
Else
 lbl11.Visible = True
 Combo11.Visible = True
End If
End Sub

Private Sub Combo14_Click()
Set r = New ADODB.Recordset
Set r = c.Execute("select initcap(R.rstud_nm),initcap(R.RSTUD_FATHER_NM),initcap(C.c_nm),R.rstud_doj,S.sch_timing,R.rstud_pic from rstud R,course C,Schdl S where R.c_id=C.c_id and R.sch_id=S.sch_id and upper(R.rstud_status)='REGISTERED' and rstud_reg_no='" & Combo14.Text & "' ")
If r.EOF = False Then
 lb1.Caption = r.Fields(0)
 lb2.Caption = r.Fields(1)
 lb3.Caption = r.Fields(2)
 lb4.Caption = r.Fields(3)
 lb5.Caption = r.Fields(4)
 lb6.Caption = r.Fields(5)
End If
End Sub

Private Sub Combo18_Click()
Set r = New ADODB.Recordset
Set r = c.Execute("select initcap(R.rstud_nm),initcap(R.RSTUD_FATHER_NM),initcap(C.c_nm),R.rstud_doj,S.sch_timing,R.rstud_pic from rstud R,course C,Schdl S where R.c_id=C.c_id and R.sch_id=S.sch_id and upper(R.rstud_status)='REGISTERED' and rstud_reg_no='" & Combo18.Text & "' ")
If r.EOF = False Then
 lb8.Caption = r.Fields(0) 'Name
 lb9.Caption = r.Fields(1) 'Father
 lb10.Caption = r.Fields(2) 'course
 lb11.Caption = r.Fields(3) 'Join date
 lb12.Caption = r.Fields(4) 'Schedule
 lb13.Caption = r.Fields(5) 'Picture
End If
End Sub

Private Sub combo3_Click()
If Combo3.ListIndex = 0 Then
Text4.Text = ""
Text5.Text = ""
 inc1.Visible = True
 inc2.Visible = False
 inc3.Visible = False
ElseIf Combo3.ListIndex = 2 Then 'Month
inc1.Visible = False
 inc2.Visible = True
 inc3.Visible = False
ElseIf Combo3.ListIndex = 1 Then 'Year
 inc1.Visible = False
 inc2.Visible = False
 inc3.Visible = True
End If
End Sub

Private Sub Combo5_Click()
Set r1 = New ADODB.Recordset
Set r1 = c.Execute("select sub_id from sub where c_id='" & Text2.Text & "' and sub_nm='" & Combo5.Text & "' ")
If r1.EOF = False Then
 Text3.Text = r1.Fields(0)
End If
End Sub

Private Sub Combo8_Click()
If Combo8.ListIndex = 0 Then
Text8.Text = ""
Text7.Text = ""
 Exp1.Visible = True
 Exp2.Visible = False
 Exp3.Visible = False
ElseIf Combo8.ListIndex = 1 Then 'Month
Exp1.Visible = False
 Exp2.Visible = True
 Exp3.Visible = False
ElseIf Combo8.ListIndex = 2 Then 'Year
 Exp1.Visible = False
 Exp2.Visible = False
 Exp3.Visible = True
End If

End Sub

Private Sub Command1_Click()
Timer1.Enabled = True
End Sub

Private Sub Command10_Click()
On Error Resume Next
If Trim(Text2.Text) <> "" Or Trim(Text3.Text) <> "" Then
 Set r = c.Execute("select count(*) from topic where c_id='" & Text2.Text & "' and sub_id='" & Text3.Text & "' ")
 If r.EOF = False Then
    DV.rsCmdTopic.Close
    DV.CmdTopic Text2.Text, Text3.Text, ""
    RptTopic.Show 1, MDI
  Else
  MsgBox "No Topic Available in Database..", vbInformation + vbOKOnly, ""
 End If
Else
 MsgBox "Select Course and Subject Both, Then click on Print..", vbInformation + vbOKOnly, ""
End If
cmd_Click
End Sub

Private Sub Command11_Click()
On Error Resume Next
Set r = c.Execute("select count(*) from topic")
If r.EOF = False Then
 Text2.Text = r.Fields(0) + 1
 DV.rsCmdTopic.Close
 DV.CmdTopic "", "", Text2.Text
 RptTopic.Show 1, MDI
Else
 MsgBox "No Topic Available in Database..", vbInformation + vbOKOnly, ""
End If
cmd_Click
End Sub

Private Sub Command12_Click()
If Combo1.ListIndex = 0 Then
  If Trim(Combo2.Text) <> "" Then
   Adodc1.RecordSource = "select R.rstud_reg_no, R.rstud_nm, R.rstud_father_nm ,R.RSTUD_STATUS,  R.rstud_mob, R.RSTUD_DOJ, C.c_nm, P.Pkg_nm, S.SCH_TIMING  from rstud R, pkg P, Course C, schdl S where R.c_id=C.c_id and R.pkg_id=P.pkg_id and  R.sch_id=S.sch_id and upper(R.rstud_reg_no)='" & UCase(Trim(Combo2.Text)) & "'"
   Adodc1.Refresh
   Else
   MsgBox "No Record Available for This Search", vbInformation + vbOKOnly, "Empty search"
  End If
ElseIf Combo1.ListIndex = 1 Then
  If Trim(Combo2.Text) <> "" Then
   Adodc1.RecordSource = "select R.rstud_reg_no, R.rstud_nm, R.rstud_father_nm ,R.RSTUD_STATUS,  R.rstud_mob, R.RSTUD_DOJ, C.c_nm, P.Pkg_nm, S.SCH_TIMING  from rstud R, pkg P, Course C, schdl S where R.c_id=C.c_id and R.pkg_id=P.pkg_id and  R.sch_id=S.sch_id and upper(R.rstud_nm)='" & UCase(Trim(Combo2.Text)) & "' "
   Adodc1.Refresh
   Else
   MsgBox "No Record Available for This Search", vbInformation + vbOKOnly, "Empty search"
  End If
ElseIf Combo1.ListIndex = 2 Then
  If Trim(Combo2.Text) <> "" Then
   Adodc1.RecordSource = "select R.rstud_reg_no, R.rstud_nm, R.rstud_father_nm ,R.RSTUD_STATUS,  R.rstud_mob, R.RSTUD_DOJ, C.c_nm, P.Pkg_nm, S.SCH_TIMING  from rstud R, pkg P, Course C, schdl S where R.c_id=C.c_id and R.pkg_id=P.pkg_id and  R.sch_id=S.sch_id and R.c_id=(select c_id from course where upper(c_nm)= '" & UCase(Trim(Combo2.Text)) & "') "
   Adodc1.Refresh
   Else
   MsgBox "No Record Available for This Search", vbInformation + vbOKOnly, "Empty search"
  End If
ElseIf Combo1.ListIndex = 3 Then
 If Trim(Text1.Text) = "" Or Trim(Text1.Text) = "" Then
  MsgBox "No Record Available for This Search", vbInformation + vbOKOnly, "Empty search"
 Else
  Adodc1.RecordSource = "select R.rstud_reg_no, R.rstud_nm, R.rstud_father_nm ,R.RSTUD_STATUS,  R.rstud_mob, R.RSTUD_DOJ, C.c_nm, P.Pkg_nm, S.SCH_TIMING  from rstud R, pkg P, Course C, schdl S where R.c_id=C.c_id and R.pkg_id=P.pkg_id and  R.sch_id=S.sch_id and R.rstud_doj between '" & Format(Text1.Text, "DD-MMM-YYYY") & "' and '" & Format(Text6.Text, "DD-MMM-YYYY") & "' "
  Adodc1.Refresh
 End If
End If
End Sub

Private Sub Command13_Click()
  Adodc1.RecordSource = "select R.rstud_reg_no, R.rstud_nm, R.rstud_father_nm ,R.RSTUD_STATUS,  R.rstud_mob, R.RSTUD_DOJ, C.c_nm, P.Pkg_nm, S.SCH_TIMING  from rstud R, pkg P, Course C, schdl S where R.c_id=C.c_id and R.pkg_id=P.pkg_id and  R.sch_id=S.sch_id order by R.rstud_reg_no"
  Adodc1.Refresh
End Sub

Private Sub Command14_Click()
Frame5.Visible = False
IncomeFrame.Visible = True
ExpenseFrame.Visible = False
Frame4.Visible = False
End Sub

Private Sub Command15_Click()
Frame5.Visible = False
IncomeFrame.Visible = False
ExpenseFrame.Visible = True
Frame4.Visible = False
End Sub

Private Sub Command16_Click()
Frame5.Visible = False
Frame4.Visible = True
IncomeFrame.Visible = False
ExpenseFrame.Visible = False
End Sub

Private Sub Command17_Click()
Frame5.Visible = True
Frame4.Visible = False
IncomeFrame.Visible = False
ExpenseFrame.Visible = False
Set r = New ADODB.Recordset
Set r = c.Execute("Select sum(CL_DAMT) from CLIENT_PMT ")
If r.Fields(0) > 0 Then
 Label35.Caption = r.Fields(0)
Else
 Label35.Caption = 0
End If
End Sub

Private Sub Command18_Click()
On Error Resume Next
DV.CmdIncome Format(Date, "DD-MMM-YYYY"), Format(Date, "DD-MMM-YYYY")
RptIncm.Show 1, MDI
DV.rsCmdIncome.Close
End Sub

Private Sub Command19_Click() 'This Week
On Error Resume Next
Text4a.Text = Format(Date - (Weekday(Date, vbSunday) - 1), "DD-MMM-YYYY")
Text5a.Text = Format(Date, "dd-mmm-yyyy")
DV.CmdIncome Text4a.Text, Text5a.Text
RptIncm.Show 1, MDI
End Sub

Private Sub Command2_Click()
On Error Resume Next
If CrentScene = 9 Then
 If Combo1.Text <> "" Or Trim(Combo2.Text) <> "" Then
  If Combo1.ListIndex = 0 Then
   Set r = c.Execute("select rstud_pic from rstud where upper(rstud_reg_no)='" & UCase(Trim(Combo2.Text)) & "' ")
   If r.EOF = False Then
    DV.CmsStudSingle UCase(Combo2.Text)
    Set StudSingle.Sections("section1").Controls.Item("image1").Picture = LoadPicture(r.Fields(0))
    StudSingle.Refresh
    StudSingle.Show 1, MDI
    DV.rsCmsStudSingle.Close
   Else
    MsgBox "No Record AVailable for This Search..Try Again", vbCritical + vbOKOnly, ""
   Exit Sub
   End If
  ElseIf Combo1.ListIndex = 1 Then
    Set r = c.Execute("select count(*) from rstud where upper(rstud_nm)='" & UCase(Trim(Combo2.Text)) & "' ")
    If r.Fields(0) > 0 Then
     DV.CmdStudRep UCase(Combo2.Text), "", "", "", "", "", "", ""
     StudReport.Sections("section4").Controls("TotStu").Caption = r.Fields(0)
     StudReport.Show 1, MDI
     DV.rsCmdStudRep.Close
    Else
     MsgBox "No Record AVailable for This Search..Try Again", vbCritical + vbOKOnly, ""
    Exit Sub
    End If
 ElseIf Combo1.ListIndex = 2 Then
  Set r = c.Execute("select c_id from course where upper(c_nm)='" & UCase(Trim(Combo2.Text)) & "' ")
  If r.EOF = False Then
   Txtstud.Text = r.Fields(0)
   Set r1 = c.Execute("select count(*) from rstud where c_id='" & Txtstud.Text & "' ")
    If r1.Fields(0) > 0 Then
     DV.CmdStudRep "", "", "", Txtstud.Text, "", "", "", ""
     StudReport.Sections("section4").Controls("TotStu").Caption = r1.Fields(0)
     StudReport.Show 1, MDI
     DV.rsCmdStudRep.Close
    Else
     MsgBox "No Record AVailable for This Search..Try Again", vbCritical + vbOKOnly, ""
    Exit Sub
    End If
  End If
 ElseIf Combo1.ListIndex = 3 Then
  If Trim(Text1.Text) = "" Or Trim(Text6.Text) = "" Then
    MsgBox "Enter  Range of date ...", vbInformation + vbOKOnly, ""
  Exit Sub
  Else
   Set r1 = c.Execute("select count(*) from rstud where rstud_doj between '" & Format(Text1.Text, "dd-mmm-yyyy") & "' and '" & Format(Text6.Text, "dd-mmm-yyyy") & "' ")
    If r1.Fields(0) > 0 Then
     DV.CmdStudRep "", "", "", "", Format(Text1.Text, "DD-MMM-YYYY"), Format(Text6.Text, "DD-MMM-YYYY"), "", ""
     StudReport.Sections("Section4").Controls("TotStu").Caption = r1.Fields(0)
     StudReport.Show 1, MDI
     DV.rsCmdStudRep.Close
    Else
     MsgBox "No Record AVailable for This Search..Try Again", vbCritical + vbOKOnly, ""
    Exit Sub
    End If
  End If
End If
Else
  MsgBox "First Choose Select By Option then Enter Value ,Then Go for search or Print..", vbInformation + vbOKOnly, ""
End If
ElseIf CrentScene >= 1 And CrentScene <= 8 Then
 Set r1 = New ADODB.Recordset
 If CrentScene = 1 Then
    Set r1 = c.Execute("select count(*) from rstud ")
    If r1.Fields(0) > 0 Then
     DV.CmdStudRep "", "", "", "", "", "", "", Txtstud.Text
     StudReport.Sections("section4").Controls("TotStu").Caption = r1.Fields(0)
     StudReport.Show 1, MDI
     DV.rsCmdStudRep.Close
    End If
 ElseIf CrentScene = 2 Then
    Set r1 = c.Execute("select count(*) from rstud where upper(rstud_gndr)='MALE' ")
    If r1.Fields(0) > 0 Then
     DV.CmdStudRep "", Txtstud.Text, "", "", "", "", "", ""
     StudReport.Sections("section4").Controls("TotStu").Caption = r1.Fields(0)
     StudReport.Show 1, MDI
     DV.rsCmdStudRep.Close
    End If
 ElseIf CrentScene = 3 Then 'Girl
    Set r1 = c.Execute("select count(*) from rstud where upper(rstud_gndr)='FEMALE' ")
    If r1.Fields(0) > 0 Then
     DV.CmdStudRep "", Txtstud.Text, "", "", "", "", "", ""
     StudReport.Sections("section4").Controls("TotStu").Caption = r1.Fields(0)
     StudReport.Show 1, MDI
     DV.rsCmdStudRep.Close
    End If
 ElseIf CrentScene = 4 Then 'Package
    Set r1 = c.Execute("select count(*) from rstud where upper(RSTUD_STATUS)='REGISTERED'  ")
    If r1.Fields(0) > 0 Then
     DV.CmdStudRep "", "", Txtstud.Text, "", "", "", "", ""
     StudReport.Sections("section4").Controls("TotStu").Caption = r1.Fields(0)
     StudReport.Show 1, MDI
     DV.rsCmdStudRep.Close
    End If
 ElseIf CrentScene = 5 Then 'NonPakage
    Set r1 = c.Execute("select count(*) from rstud where upper(RSTUD_STATUS)='UNREGISTERED' ")
    If r1.Fields(0) > 0 Then
     DV.CmdStudRep "", "", Txtstud.Text, "", "", "", "", ""
     StudReport.Sections("section4").Controls("TotStu").Caption = r1.Fields(0)
     StudReport.Show 1, MDI
     DV.rsCmdStudRep.Close
    End If
 ElseIf CrentScene = 6 Then 'Tday Enrolled
    Set r1 = c.Execute("select count(*) from rstud where rstud_doj='" & Format(Date, "dd-mmm-yyyy") & "' ")
    If r1.Fields(0) > 0 Then
     DV.CmdStudRep "", "", "", "", Txtstud.Text, txtStud2.Text, "", ""
     StudReport.Sections("section4").Controls("TotStu").Caption = r1.Fields(0)
     StudReport.Show 1, MDI
     DV.rsCmdStudRep.Close
    End If
 ElseIf CrentScene = 7 Then 'Today expired
    Set r1 = c.Execute("select count(*) from rstud where rstud_doe='" & Format(Date, "dd-mmm-yyyy") & "' ")
    If r1.Fields(0) > 0 Then
     DV.CmdStudRep "", "", "", "", "", "", Txtstud.Text, ""
     StudReport.Sections("section4").Controls("TotStu").Caption = r1.Fields(0)
     StudReport.Show 1, MDI
     DV.rsCmdStudRep.Close
    End If
 ElseIf CrentScene = 8 Then
    MsgBox "Cannot Be Printed Now. But Can be Viewed.", vbInformation + vbOKOnly, ""
 End If
End If
End Sub

Private Sub Command20_Click() 'Income Report
On Error Resume Next
If Combo3.Text = "" Then
MsgBox "Select search by Option..", vbCritical + vbOKOnly, "No Selection"
Exit Sub
End If
If Combo3.ListIndex = 0 Then
 If Trim(Text4.Text) = "" Or Trim(Text5.Text) = "" Then
  MsgBox "Enter date Range First..", vbInformation + vbOKOnly, ""
  Exit Sub
 End If
Text4a.Text = Format(Text4.Text, "DD-MMM-YYYY")
Text5a.Text = Format(Text5.Text, "DD-MMM-YYYY")
DV.CmdIncome Text4a.Text, Text5a.Text
RptIncm.Show 1, MDI
DV.rsCmdIncome.Close
ElseIf Combo3.ListIndex = 1 Then
 If ComboYr.Text = "" Then
  MsgBox "Select Year First..", vbInformation + vbOKOnly, ""
  Exit Sub
 ElseIf Combo4.Text = "" Then
  MsgBox "Select the Month..", vbInformation + vbOKOnly, ""
  Exit Sub
 End If
   Text4a.Text = "01-" & Left(Combo4.Text, 3) & "-" & ComboYr.Text
  If Combo4.ListIndex = 0 Or Combo4.ListIndex Mod 2 = 0 Then '31
   Text5a.Text = "31-" & Left(Combo4.Text, 3) & "-" & ComboYr.Text
  ElseIf Combo4.ListIndex = 1 Then '30 / 28
   If Val(Combo3.Text) Mod 4 = 0 Then
    Text5a.Text = "29-" & Left(Combo4.Text, 3) & "-" & ComboYr.Text
   Else
    Text5a.Text = "28-" & Left(Combo4.Text, 3) & "-" & ComboYr.Text
   End If
  Else
   Text5a.Text = "30-" & Left(Combo4.Text, 3) & "-" & ComboYr.Text
  End If
 DV.CmdIncome Text4a.Text, Text5a.Text
 RptIncm.Show 1, MDI
 DV.rsCmdIncome.Close
ElseIf Combo3.ListIndex = 2 Then
If Combo6.Text = "" Then
  MsgBox "Select the Year..", vbInformation + vbOKOnly, ""
  Exit Sub
End If
Text4a.Text = "01-Jan-" & Combo6.Text
Text5a.Text = "31-Dec-" & Combo6.Text
DV.CmdIncome Text4a.Text, Text5a.Text
RptIncm.Show 1, MDI
DV.rsCmdIncome.Close
End If
End Sub

Private Sub Command21_Click()
On Error Resume Next
Text4a.Text = "01-Jan-" & Format(Date, "yyyy")
 Text5a.Text = Format(Date, "dd-mmm-yyyy")
DV.CmdIncome Text4a.Text, Text5a.Text
RptIncm.Show 1, MDI
DV.rsCmdIncome.Close
End Sub

Private Sub Command22_Click() 'This Month
On Error Resume Next
Text4a.Text = "01-" & Format(Date, "mmm-yyyy")
 Text5a.Text = Format(Date, "dd-mmm-yyyy")
DV.CmdIncome Text4a.Text, Text5a.Text
RptIncm.Show 1, MDI
DV.rsCmdIncome.Close
End Sub

Private Sub Command23_Click()
On Error Resume Next
DV.CmdExpense Format(Date, "DD-MMM-YYYY"), Format(Date, "DD-MMM-YYYY")
RptExpns.Show 1, MDI
DV.rsCmdIncome.Close
End Sub

Private Sub Command24_Click()
On Error Resume Next
Text4a.Text = Format(Date - (Weekday(Date, vbSunday) - 1), "DD-MMM-YYYY")
Text5a.Text = Format(Date, "dd-mmm-yyyy")
DV.CmdExpense Text4a.Text, Text5a.Text
RptExpns.Show 1, MDI
DV.rsCmdExpense.Close
End Sub

Private Sub Command25_Click()
On Error Resume Next
Text4a.Text = "01-" & Format(Date, "mmm-yyyy")
 Text5a.Text = Format(Date, "dd-mmm-yyyy")
DV.CmdExpense Text4a.Text, Text5a.Text
RptExpns.Show 1, MDI
DV.rsCmdExpense.Close
End Sub

Private Sub Command26_Click()
On Error Resume Next
Text4a.Text = "01-Jan-" & Format(Date, "yyyy")
 Text5a.Text = Format(Date, "dd-mmm-yyyy")
DV.CmdExpense Text4a.Text, Text5a.Text
RptExpns.Show 1, MDI
DV.rsCmdExpense.Close
End Sub

Private Sub Command27_Click()
On Error Resume Next
If Combo8.Text = "" Then
MsgBox "Select search by Option..", vbCritical + vbOKOnly, "No Selection"
Exit Sub
End If
If Combo8.ListIndex = 0 Then
 If Trim(Text8.Text) = "" Or Trim(Text7.Text) = "" Then
  MsgBox "Enter date Range First..", vbInformation + vbOKOnly, ""
  Exit Sub
 End If
Text4a.Text = Format(Text8.Text, "DD-MMM-YYYY")
Text5a.Text = Format(Text7.Text, "DD-MMM-YYYY")
DV.CmdExpense Text4a.Text, Text5a.Text
RptExpns.Show 1, MDI
DV.rsCmdExpense.Close
ElseIf Combo8.ListIndex = 1 Then
 If Combo9.Text = "" Then
  MsgBox "Select Year First..", vbInformation + vbOKOnly, ""
  Exit Sub
 ElseIf Combo10.Text = "" Then
  MsgBox "Select the Month..", vbInformation + vbOKOnly, ""
  Exit Sub
 End If
   Text4a.Text = "01-" & Left(Combo10.Text, 3) & "-" & Combo9.Text
  If Combo10.ListIndex = 0 Or Combo10.ListIndex Mod 2 = 0 Then '31
   Text5a.Text = "31-" & Left(Combo10.Text, 3) & "-" & Combo9.Text
  ElseIf Combo10.ListIndex = 1 Then '30 / 28
   If Val(Combo9.Text) Mod 4 = 0 Then
    Text5a.Text = "29-" & Left(Combo10.Text, 3) & "-" & Combo9.Text
   Else
    Text5a.Text = "28-" & Left(Combo10.Text, 3) & "-" & Combo9.Text
   End If
  Else
   Text5a.Text = "30-" & Left(Combo10.Text, 3) & "-" & Combo9.Text
  End If
 DV.CmdExpense Text4a.Text, Text5a.Text
 RptExpns.Show 1, MDI
 DV.rsCmdExpense.Close
ElseIf Combo8.ListIndex = 2 Then
 If Combo7.Text = "" Then
  MsgBox "Select the Year..", vbInformation + vbOKOnly, ""
  Exit Sub
 End If
Text4a.Text = "01-Jan-" & Combo7.Text
Text5a.Text = "31-Dec-" & Combo7.Text
DV.CmdExpense Text4a.Text, Text5a.Text
RptExpns.Show 1, MDI
DV.rsCmdExpense.Close
End If
End Sub

Private Sub Command28_Click()
Frame9.Visible = False
RptAdminAll.Show 1, MDI
End Sub

Private Sub Command29_Click()
Frame9.Visible = False
RptUserAll.Show 1, MDI
End Sub

Private Sub Command3_Click()
On Error Resume Next
If CrentScene = 9 Then
 If Combo1.Text <> "" Or Trim(Combo2.Text) <> "" Then
  If Combo1.ListIndex = 0 Then
   Set r = c.Execute("select rstud_pic from rstud where upper(rstud_reg_no)='" & UCase(Trim(Combo2.Text)) & "' ")
   If r.EOF = False Then
    DV.CmsStudSingle UCase(Combo2.Text)
    Set StudSingle.Sections("section1").Controls.Item("image1").Picture = LoadPicture(r.Fields(0))
    StudSingle.Refresh
    StudSingle.Show 1, MDI
    DV.rsCmsStudSingle.Close
   Else
    MsgBox "No Record AVailable for This Search..Try Again", vbCritical + vbOKOnly, ""
   Exit Sub
   End If
  ElseIf Combo1.ListIndex = 1 Then
    Set r = c.Execute("select count(*) from rstud where upper(rstud_nm)='" & UCase(Trim(Combo2.Text)) & "' ")
    If r.Fields(0) > 0 Then
     DV.CmdStudRep UCase(Combo2.Text), "", "", "", "", "", "", ""
     StudReport.Sections("section4").Controls("TotStu").Caption = r.Fields(0)
     StudReport.Show 1, MDI
     DV.rsCmdStudRep.Close
    Else
     MsgBox "No Record AVailable for This Search..Try Again", vbCritical + vbOKOnly, ""
    Exit Sub
    End If
 ElseIf Combo1.ListIndex = 2 Then
  Set r = c.Execute("select c_id from course where upper(c_nm)='" & UCase(Trim(Combo2.Text)) & "' ")
  If r.EOF = False Then
   Txtstud.Text = r.Fields(0)
   Set r1 = c.Execute("select count(*) from rstud where c_id='" & Txtstud.Text & "' ")
    If r1.Fields(0) > 0 Then
     DV.CmdStudRep "", "", "", Txtstud.Text, "", "", "", ""
     StudReport.Sections("section4").Controls("TotStu").Caption = r1.Fields(0)
     StudReport.Show 1, MDI
     DV.rsCmdStudRep.Close
    Else
     MsgBox "No Record AVailable for This Search..Try Again", vbCritical + vbOKOnly, ""
    Exit Sub
    End If
  End If
 ElseIf Combo1.ListIndex = 3 Then
  If Trim(Text1.Text) = "" Or Trim(Text6.Text) = "" Then
    MsgBox "Enter  Range of date ...", vbInformation + vbOKOnly, ""
  Exit Sub
  Else
   Set r1 = c.Execute("select count(*) from rstud where rstud_doj between '" & Format(Text1.Text, "dd-mmm-yyyy") & "' and '" & Format(Text6.Text, "dd-mmm-yyyy") & "' ")
    If r1.Fields(0) > 0 Then
     DV.CmdStudRep "", "", "", "", Format(Text1.Text, "DD-MMM-YYYY"), Format(Text6.Text, "DD-MMM-YYYY"), "", ""
     StudReport.Sections("Section4").Controls("TotStu").Caption = r1.Fields(0)
     StudReport.Show 1, MDI
     DV.rsCmdStudRep.Close
    Else
     MsgBox "No Record AVailable for This Search..Try Again", vbCritical + vbOKOnly, ""
    Exit Sub
    End If
  End If
End If
Else
  MsgBox "First Choose Select By Option then Enter Value ,Then Go for search or Print..", vbInformation + vbOKOnly, ""
End If
ElseIf CrentScene >= 1 And CrentScene <= 8 Then
 Set r1 = New ADODB.Recordset
 If CrentScene = 1 Then
    Set r1 = c.Execute("select count(*) from rstud ")
    If r1.Fields(0) > 0 Then
     DV.CmdStudRep "", "", "", "", "", "", "", Txtstud.Text
     StudReport.Sections("section4").Controls("TotStu").Caption = r1.Fields(0)
     StudReport.Show 1, MDI
     DV.rsCmdStudRep.Close
    End If
 ElseIf CrentScene = 2 Then
    Set r1 = c.Execute("select count(*) from rstud where upper(rstud_gndr)='MALE' ")
    If r1.Fields(0) > 0 Then
     DV.CmdStudRep "", Txtstud.Text, "", "", "", "", "", ""
     StudReport.Sections("section4").Controls("TotStu").Caption = r1.Fields(0)
     StudReport.Show 1, MDI
     DV.rsCmdStudRep.Close
    End If
 ElseIf CrentScene = 3 Then 'Girl
    Set r1 = c.Execute("select count(*) from rstud where upper(rstud_gndr)='FEMALE' ")
    If r1.Fields(0) > 0 Then
     DV.CmdStudRep "", Txtstud.Text, "", "", "", "", "", ""
     StudReport.Sections("section4").Controls("TotStu").Caption = r1.Fields(0)
     StudReport.Show 1, MDI
     DV.rsCmdStudRep.Close
    End If
 ElseIf CrentScene = 4 Then 'Package
    Set r1 = c.Execute("select count(*) from rstud where upper(RSTUD_STATUS)='REGISTERED'  ")
    If r1.Fields(0) > 0 Then
     DV.CmdStudRep "", "", Txtstud.Text, "", "", "", "", ""
     StudReport.Sections("section4").Controls("TotStu").Caption = r1.Fields(0)
     StudReport.Show 1, MDI
     DV.rsCmdStudRep.Close
    End If
 ElseIf CrentScene = 5 Then 'NonPakage
    Set r1 = c.Execute("select count(*) from rstud where upper(RSTUD_STATUS)='UNREGISTERED' ")
    If r1.Fields(0) > 0 Then
     DV.CmdStudRep "", "", Txtstud.Text, "", "", "", "", ""
     StudReport.Sections("section4").Controls("TotStu").Caption = r1.Fields(0)
     StudReport.Show 1, MDI
     DV.rsCmdStudRep.Close
    End If
 ElseIf CrentScene = 6 Then 'Tday Enrolled
    Set r1 = c.Execute("select count(*) from rstud where rstud_doj='" & Format(Date, "dd-mmm-yyyy") & "' ")
    If r1.Fields(0) > 0 Then
     DV.CmdStudRep "", "", "", "", Txtstud.Text, txtStud2.Text, "", ""
     StudReport.Sections("section4").Controls("TotStu").Caption = r1.Fields(0)
     StudReport.Show 1, MDI
     DV.rsCmdStudRep.Close
    End If
 ElseIf CrentScene = 7 Then 'Today expired
    Set r1 = c.Execute("select count(*) from rstud where rstud_doe='" & Format(Date, "dd-mmm-yyyy") & "' ")
    If r1.Fields(0) > 0 Then
     DV.CmdStudRep "", "", "", "", "", "", Txtstud.Text, ""
     StudReport.Sections("section4").Controls("TotStu").Caption = r1.Fields(0)
     StudReport.Show 1, MDI
     DV.rsCmdStudRep.Close
    End If
 ElseIf CrentScene = 8 Then
    MsgBox "Cannot Be Printed Now. But Can be Viewed.", vbInformation + vbOKOnly, ""
 End If
End If
End Sub

Private Sub Command30_Click()
Frame9.Visible = False
RptUserLoginInfo.Show 1, MDI
End Sub

Private Sub Command31_Click()
Frame9.Visible = False
RptClientAll.Show 1, MDI
End Sub

Private Sub Command32_Click()
Frame9.Visible = True
End Sub

Private Sub Command33_Click()
Frame10.Visible = False
Frame8.Visible = True
End Sub

Private Sub Command34_Click()
Frame10.Visible = False
Frame8.Visible = False
frmbackup.Show 1, MDI
End Sub

Private Sub Command35_Click()
Frame10.Visible = False
Frame8.Visible = False
FrmRestore.Show 1, MDI
End Sub

Private Sub Command36_Click()
Frame10.Visible = False
Frame8.Visible = False
Security_Question.Show 1, MDI
End Sub

Private Sub Command37_Click()
Frame10.Visible = False
Frame8.Visible = False
about_org.Show 1, MDI
End Sub

Private Sub Command38_Click()
Frame8.Visible = False
Frame10.Visible = False
FrmTestPrpt1.Show 1, MDI
End Sub

Private Sub Command39_Click()
Frame10.Visible = False
Frame8.Visible = False
Stud_Ranking.Show
End Sub

Private Sub Command4_Click()
On Error Resume Next
If Trim(Text2.Text) <> "" Then
 Set r = c.Execute("select count(*) from schdl where c_id='" & Text2.Text & "' ")
 If r.EOF = False Then
   DV.rsCmdSchedule.Close
   DV.CmdSchedule Text2.Text, ""
   RptSchedule.Show 1, MDI
 Else
  MsgBox "No Schedule Available in Database..", vbInformation + vbOKOnly, ""
 End If
  Else
   MsgBox "Select Course First..", vbInformation + vbOKOnly, ""
End If
cmd_Click
End Sub

Private Sub Command40_Click()
Frame10.Visible = False
Frame8.Visible = False
FrmExportQues.Show
End Sub

Private Sub Command41_Click()
Frame10.Visible = False
Frame8.Visible = False
FrmImportQues.Show
End Sub

Private Sub Command42_Click()
Frame9.Visible = False
RptStudLoginInfo.Show 1, MDI
End Sub

Private Sub Command43_Click()
Frame9.Visible = False
If MsgBox("Once Deleted Cannot be Recovered. Are You Sure To Delete ?", vbCritical + vbYesNo, "Delete") = vbYes Then
 c.Execute ("delete from rstud where upper(rstud_status)='REGISTERED'and RSTUD_DOE < '" & Format(Date, "dd-mmm-yyyy") & "' ")
MsgBox "All Student Records (With Package) SuccessFully Deleted.", vbInformation + vbOKOnly, "Delete Student"
End If
setstudrecord
End Sub

Private Sub Command44_Click()
Frame9.Visible = False
If MsgBox("Once Deleted Cannot be Recovered. Are You Sure To Delete ?", vbCritical + vbYesNo, "Delete") = vbYes Then
 c.Execute ("delete from rstud where upper(rstud_status)='UNREGISTERED' and RSTUD_DOE < '" & Format(Date, "dd-mmm-yyyy") & "' ")
 MsgBox "All Student Records (Without Package) SuccessFully Deleted.", vbInformation + vbOKOnly, "Delete Student"
End If
setstudrecord
End Sub

Private Sub Command45_Click()
Frame9.Visible = False
If MsgBox("Once Deleted Cannot be Recovered. Are You Sure To Delete ?", vbCritical + vbYesNo, "Delete") = vbYes Then
 c.Execute ("delete from course")
 MsgBox "All Courses SuccessFully Deleted.", vbInformation + vbOKOnly, "Delete Course"
End If
End Sub

Private Sub Command46_Click()
Frame9.Visible = False
If MsgBox("Once Deleted Cannot be Recovered. Are You Sure To Delete ?", vbCritical + vbYesNo, "Delete") = vbYes Then
 c.Execute ("delete from rstud")
 MsgBox "All Student Records SuccessFully Deleted.", vbInformation + vbOKOnly, "Delete Student"
End If
setstudrecord
End Sub

Private Sub Command47_Click()
Frame9.Visible = False
FrmSchedule.Show 1, MDI
End Sub

Private Sub Command48_Click()
Frame9.Visible = False
frmCourseMaster.Show 1, MDI
End Sub

Private Sub Command49_Click()
Frame9.Visible = False
FrmPackage.Show 1, MDI
End Sub

Private Sub Command5_Click()
On Error Resume Next
Set r = c.Execute("select count(*) from Schdl")
If r.EOF = False Then
Text2.Text = r.Fields(0) + 1
 DV.rsCmdSchedule.Close
 DV.CmdSchedule "", Text2.Text
 RptSchedule.Show 1, MDI
Else
 MsgBox "No Schedule Available in Database..", vbInformation + vbOKOnly, ""
End If
cmd_Click
End Sub

Private Sub Command50_Click()
Frame9.Visible = False
frmSubMaster.Show
End Sub

Private Sub Command51_Click()
Frame9.Visible = False
If MsgBox("Once Deleted Cannot be Recovered. Are You Sure To Delete ?", vbCritical + vbYesNo, "Delete") = vbYes Then
c.Execute ("delete from clnt_ordr_chln where upper(CSTATUS)='COMPLETED'")
MsgBox "All Completed Orders List Deleted.", vbInformation + vbOKOnly, "Order List"
End If
End Sub

Private Sub Command53_Click()
Frame10.Visible = True
Frame8.Visible = False


End Sub

Private Sub Command6_Click()
On Error Resume Next
Set r = c.Execute("select count(*) from pkg")
If r.EOF = False Then
Text2.Text = r.Fields(0) + 1
 DV.rsCmdPkg.Close
 DV.CmdPkg "", Text2.Text ', Int(Val(r.Fields(0) + 1))
 RptPackage.Show 1, MDI
Else
 MsgBox "No Package Available in Database..", vbInformation + vbOKOnly, ""
End If
cmd_Click
End Sub

Private Sub Command7_Click()
On Error Resume Next
If Trim(Text2.Text) <> "" Then
 Set r = c.Execute("select count(*) from pkg where c_id='" & Text2.Text & "' ")
 If r.EOF = False Then
   DV.rsCmdPkg.Close
   DV.CmdPkg Text2.Text, ""
   RptPackage.Show 1, MDI
 Else
  MsgBox "No Package Available in Database..", vbInformation + vbOKOnly, ""
 End If
  Else
   MsgBox "Select Course First..", vbInformation + vbOKOnly, ""
End If
cmd_Click
End Sub

Private Sub Command8_Click()
On Error Resume Next
If Trim(Text2.Text) <> "" Then
 Set r = c.Execute("select count(*) from sub where c_id='" & Text2.Text & "' ")
 If r.EOF = False Then
   DV.rsCmdSubject.Close
   DV.CmdSubject Text2.Text, "" ', Int(Val(r.Fields(0) + 1))
   RptSubject.Show 1, MDI
 Else
  MsgBox "No Subject Available in Database..", vbInformation + vbOKOnly, ""
 End If
 Else
   MsgBox "Select Course First..", vbInformation + vbOKOnly, ""
End If
cmd_Click
End Sub

Private Sub Command9_Click()
On Error Resume Next
Set r = c.Execute("select count(*) from sub")
If r.EOF = False Then
Text2.Text = r.Fields(0) + 1
 DV.rsCmdSubject.Close
 DV.CmdSubject "", Text2.Text ', Int(Val(r.Fields(0) + 1))
 RptSubject.Show 1, MDI
Else
 MsgBox "No Subject Available in Database..", vbInformation + vbOKOnly, ""
End If
cmd_Click
End Sub
Private Sub Calendar1_Click()
Text1.Text = Calendar1.Day & "-" & Calendar1.Month & "-" & Calendar1.Year
Calendar1.Visible = False
If Calendar1.Value > Date Then
MsgBox "Cannot Select Future Dates !!!", vbInformation + vbOKOnly, "Invalid Date"
Text1.SetFocus
End If
End Sub

Private Sub Calendar2_Click()
If Calendar2.Value < Calendar1.Value Then
 Text6.Text = ""
 Calendar2.Visible = True
Else
 Text6.Text = Calendar2.Day & "-" & Calendar2.Month & "-" & Calendar2.Year
 Calendar2.Visible = False
End If
End Sub

Private Sub Form_Load()
conn
Combo18.Clear
Combo14.Clear
Set r = c.Execute("select rstud_reg_no from rstud where UPPER(rstud_status)='REGISTERED' ")
While r.EOF = False
 Combo14.AddItem r.Fields(0)
 Combo18.AddItem r.Fields(0)
r.MoveNext
Wend
Calendar1.Value = Format(Date, "DD-MMM-YY")
Calendar2.Value = Format(Date, "DD-MMM-YY")
cld.Value = Format(Date, "DD-MMM-YY")
cld1.Value = Format(Date, "DD-MMM-YY")
cld2.Value = Format(Date, "DD-MMM-YY")
cld3.Value = Format(Date, "DD-MMM-YY")
Combo15.Clear
Combo15.AddItem "COMPLETED"
Combo15.AddItem "PENDING"
Frame2.Height = 795
Label35.Caption = ""
Text8.Text = ""
Text7.Text = ""
Text4.Text = ""
Text5.Text = ""
Combo8.Clear
Combo3.Clear
Combo13.Clear
Combo13.AddItem "Month Wise"
Combo13.AddItem "Year Wise"
Combo3.AddItem "Date"
Combo3.AddItem "Month"
Combo3.AddItem "Year"
Combo8.AddItem "Date"
Combo8.AddItem "Month"
Combo8.AddItem "Year"
Combo6.Clear
Combo7.Clear
Combo9.Clear
ComboYr.Clear
Combo12.Clear
Combo16.Clear
For i = 2016 To Year(Date)
 Combo9.AddItem i
 Combo7.AddItem i
 Combo6.AddItem i
 ComboYr.AddItem i
 Combo12.AddItem i
 Combo16.AddItem i
Next i
Combo4.Clear
Combo4.AddItem "January"
Combo4.AddItem "February"
Combo4.AddItem "March"
Combo4.AddItem "April"
Combo4.AddItem "May"
Combo4.AddItem "June"
Combo4.AddItem "July"
Combo4.AddItem "August"
Combo4.AddItem "September"
Combo4.AddItem "October"
Combo4.AddItem "November"
Combo4.AddItem "December"
Combo10.Clear
Combo10.AddItem "January"
Combo10.AddItem "February"
Combo10.AddItem "March"
Combo10.AddItem "April"
Combo10.AddItem "May"
Combo10.AddItem "June"
Combo10.AddItem "July"
Combo10.AddItem "August"
Combo10.AddItem "September"
Combo10.AddItem "October"
Combo10.AddItem "November"
Combo10.AddItem "December"
Combo11.Clear
Combo11.AddItem "January"
Combo11.AddItem "February"
Combo11.AddItem "March"
Combo11.AddItem "April"
Combo11.AddItem "May"
Combo11.AddItem "June"
Combo11.AddItem "July"
Combo11.AddItem "August"
Combo11.AddItem "September"
Combo11.AddItem "October"
Combo11.AddItem "November"
Combo11.AddItem "December"
Combo17.Clear
Combo17.AddItem "January"
Combo17.AddItem "February"
Combo17.AddItem "March"
Combo17.AddItem "April"
Combo17.AddItem "May"
Combo17.AddItem "June"
Combo17.AddItem "July"
Combo17.AddItem "August"
Combo17.AddItem "September"
Combo17.AddItem "October"
Combo17.AddItem "November"
Combo17.AddItem "December"

CrentScene = 0
Calendar1.Visible = False
Calendar2.Visible = False
Text3.Text = ""
Text2.Text = ""
Combo1.Clear
Combo2.Clear
Combo1.AddItem "Reg. ID"
Combo1.AddItem "Name"
Combo1.AddItem "Course"
Combo1.AddItem "Enroll Date"
For i = 0 To 3
 Combo(i).Clear
Next i
rs_course.MoveFirst
While rs_course.EOF = False
 Combo(0).AddItem rs_course.Fields(0)
 Combo(1).AddItem rs_course.Fields(0)
 Combo(2).AddItem rs_course.Fields(0)
 Combo(3).AddItem rs_course.Fields(0)
rs_course.MoveNext
Wend
Me.Top = 0
Me.Left = 0
Timer1.Enabled = False
navigation_counter = 0
Frame1.Top = 660
For i = 1 To 4
 frm(i).Visible = False
Next i
setstudrecord
rptbtn_Click (0)
End Sub

Public Sub setstudrecord()
'Setting All Student Records
Set r = c.Execute("select count(*) from rstud where upper(RSTUD_STATUS)='REGISTERED'")
If IsNull(r.Fields(0)) = False Then
 a1(0).Caption = r.Fields(0)
End If
Set r = c.Execute("select count(*) from rstud where upper(RSTUD_STATUS)='REGISTERED' and upper(rstud_gndr)='MALE' ")
If IsNull(r.Fields(0)) = False Then
 a1(1).Caption = r.Fields(0)
End If
Set r = c.Execute("select count(*) from rstud where upper(RSTUD_STATUS)='REGISTERED' and upper(rstud_gndr)='FEMALE'")
If IsNull(r.Fields(0)) = False Then
 a1(2).Caption = r.Fields(0)
End If
Set r = c.Execute("select count(*) from rstud where upper(RSTUD_STATUS)='REGISTERED'")
If IsNull(r.Fields(0)) = False Then
 a1(3).Caption = r.Fields(0)
End If
Set r = c.Execute("select count(*) from rstud where upper(RSTUD_STATUS)='UNREGISTERED'")
If IsNull(r.Fields(0)) = False Then
 a1(4).Caption = r.Fields(0)
End If
Set r = c.Execute("select count(*) from rstud where upper(RSTUD_STATUS)='REGISTERED' and rstud_doj='" & Format(Date, "dd-MMM-yyyy") & "' ")
If IsNull(r.Fields(0)) = False Then
 a1(5).Caption = r.Fields(0)
End If
Set r = c.Execute("select count(*) from rstud where upper(RSTUD_STATUS)='REGISTERED' and rstud_doe='" & Format(Date, "dd-MMM-yyyy") & "'")
If IsNull(r.Fields(0)) = False Then
 a1(6).Caption = r.Fields(0)
End If
Set r = c.Execute("select count(*) from PKG_RENEW ")
If IsNull(r.Fields(0)) = False Then
 a1(7).Caption = r.Fields(0)
End If
End Sub

Private Sub Frame6_Click()
Frame9.Visible = False
End Sub

Private Sub Frame7_Click()
Frame10.Visible = False
Frame8.Visible = False
End Sub

Private Sub frm_Click(Index As Integer)
If Index = 0 Then
cmd_Click
End If
End Sub

Private Sub ovrlbtn_Click()
On Error Resume Next
 If Combo13.Text = "" Then
   MsgBox "Select Creiteria, Either Month Wise or Year Wise.", vbInformation + vbOKOnly, ""
   Exit Sub
 End If
Set r = New ADODB.Recordset
If Combo13.ListIndex = 0 Then 'Month Wise
 If Combo12.Text = "" Or Combo11.Text = "" Then
  MsgBox "Select Year and Month from DropDown List", vbInformation + vbOKOnly, ""
  Combo12.SetFocus
  Exit Sub
 End If
Set r = c.Execute("select sum(INC_AMT) from incm where inc_date between '" & Text4a.Text & "' and '" & Text5a.Text & "' ")
If IsNull(r.Fields(0)) = False Then
 Text4.Text = r.Fields(0)
Else
 Text4.Text = 0
End If
Set r = c.Execute("select sum(EX_AMT) from EXP where EX_date between '" & Text4a.Text & "' and '" & Text5a.Text & "' ")
If IsNull(r.Fields(0)) = False Then
 Text5.Text = r.Fields(0)
 Else
 Text5.Text = 0
End If

DV.CmdIncome "", ""
overall.Sections("section2").Controls("AMTIncm").Caption = Text4.Text
overall.Sections("section2").Controls("AMTExpn").Caption = Text5.Text
overall.Sections("section2").Controls("AMTFinal").Caption = Val(Text4.Text) - Val(Text5.Text)
Text7.Text = "( " & Combo11.Text & " - " & Combo12.Text & " [Monthly Report] )"
overall.Sections("section4").Controls("Label6").Caption = Text7.Text
If (Val(Text4.Text) - Val(Text5.Text)) > 0 Then
 overall.Sections("section2").Controls("Status").Caption = Abs(Val(Text4.Text) - Val(Text5.Text)) & " In Profit (+)"
ElseIf (Val(Text4.Text) - Val(Text5.Text)) < 0 Then
 overall.Sections("section2").Controls("Status").Caption = Abs(Val(Text4.Text) - Val(Text5.Text)) & " In Loss (-)"
Else
  overall.Sections("section2").Controls("Status").Caption = "O    [Zero balance]"
End If
overall.Show 1, MDI
DV.rsCmdIncome.Close
Else 'Year Wise
 If Combo12.Text = "" Then
  MsgBox "Select Valid Year and Month from DropDown List", vbInformation + vbOKOnly, ""
  Combo12.SetFocus
 Exit Sub
 End If
Set r = c.Execute("select sum(INC_AMT) from incm where inc_date between '" & Text4a.Text & "' and '" & Text5a.Text & "' ")
If IsNull(r.Fields(0)) = False Then
 Text4.Text = r.Fields(0)
 Else
 Text4.Text = 0
End If
Set r = c.Execute("select sum(EX_AMT) from EXP where EX_date between '" & Text4a.Text & "' and '" & Text5a.Text & "' ")
If IsNull(r.Fields(0)) = False Then
 Text5.Text = r.Fields(0)
Else
 Text5.Text = 0
End If
DV.CmdIncome "", ""
overall.Sections("section2").Controls("AMTIncm").Caption = Text4.Text
overall.Sections("section2").Controls("AMTExpn").Caption = Text5.Text
overall.Sections("section2").Controls("AMTFinal").Caption = Val(Text4.Text) - Val(Text5.Text)
Text7.Text = "( " & Combo12.Text & " [Annual Report] )"
overall.Sections("section4").Controls("Label6").Caption = Text7.Text
If (Val(Text4.Text) - Val(Text5.Text)) > 0 Then
 overall.Sections("section2").Controls("Status").Caption = Abs(Val(Text4.Text) - Val(Text5.Text)) & " In Profit (+)"
ElseIf (Val(Text4.Text) - Val(Text5.Text)) < 0 Then
 overall.Sections("section2").Controls("Status").Caption = Abs(Val(Text4.Text) - Val(Text5.Text)) & " In Loss (-)"
Else
  overall.Sections("section2").Controls("Status").Caption = "O     [Zero Balance]"
End If
overall.Show 1, MDI
DV.rsCmdIncome.Close
End If
End Sub

Private Sub Picture1_Click(Index As Integer)
studnt_Click (Index)
End Sub

Private Sub rptbtn_Click(Index As Integer)
For i = 0 To 5
 If i = Index Then
  rptbtn(Index).BackColor1 = &H8000000D
  rptbtn(Index).BackColor1 = &H8000000D
  rptbtn(Index).Width = 2795
 Else
  rptbtn(i).BackColor1 = &H6D6D6D
  rptbtn(i).BackColor1 = &H6D6D6D
  rptbtn(i).Width = 2450
 End If
Next i

If Index = 0 Then
 Frame7.Visible = False
 AccountFrame.Visible = False
 Frame3.Visible = False
 frm(0).Visible = True
 Frame6.Visible = False
ElseIf Index = 1 Then
 AccountFrame.Visible = False
 frm(0).Visible = False
 Frame3.Visible = True
 Frame7.Visible = False
  Frame6.Visible = False
ElseIf Index = 2 Then
 AccountFrame.Visible = True
 frm(0).Visible = False
 Frame3.Visible = False
 Frame7.Visible = False
  Frame6.Visible = False
ElseIf Index = 3 Then
 AccountFrame.Visible = False
 frm(0).Visible = False
 Frame3.Visible = False
 Frame7.Visible = False
  Frame6.Visible = True
ElseIf Index = 4 Then
 AccountFrame.Visible = False
 frm(0).Visible = False
 Frame3.Visible = False
 Frame7.Visible = True
  Frame6.Visible = False
ElseIf Index = 5 Then
If MsgBox("Are You Sure to Return To Main Menu ?", vbInformation + vbYesNo, "Main Menu") = vbYes Then
 Unload Me
Else
 rptbtn_Click (0)
End If
End If
End Sub

Private Sub studnt_Click(Index As Integer)
  CrentScene = Index + 1
  Set r = New ADODB.Recordset
 If Index = 0 Then 'All
  Adodc1.RecordSource = "select R.rstud_reg_no, R.rstud_nm, R.rstud_father_nm ,R.RSTUD_STATUS,  R.rstud_mob, R.RSTUD_DOJ, C.c_nm, P.Pkg_nm, S.SCH_TIMING  from rstud R, pkg P, Course C, schdl S where R.c_id=C.c_id and R.pkg_id=P.pkg_id and  R.sch_id=S.sch_id order by R.rstud_reg_no"
  Adodc1.Refresh
  Set r = c.Execute("select count(*) from rstud")
  If r.EOF = False Then
   Txtstud.Text = r.Fields(0) + 1
  End If
 ElseIf Index = 1 Then 'Boys
  Adodc1.RecordSource = "select R.rstud_reg_no, R.rstud_nm, R.rstud_father_nm ,R.RSTUD_STATUS,  R.rstud_mob, R.RSTUD_DOJ, C.c_nm, P.Pkg_nm, S.SCH_TIMING  from rstud R, pkg P, Course C, schdl S where R.c_id=C.c_id and R.pkg_id=P.pkg_id and  R.sch_id=S.sch_id and upper(R.RSTUD_GNDR)='MALE' order by R.rstud_reg_no"
  Adodc1.Refresh
  Txtstud.Text = "MALE"
  ElseIf Index = 2 Then 'Girls
  Adodc1.RecordSource = "select R.rstud_reg_no, R.rstud_nm, R.rstud_father_nm ,R.RSTUD_STATUS,  R.rstud_mob, R.RSTUD_DOJ, C.c_nm, P.Pkg_nm, S.SCH_TIMING  from rstud R, pkg P, Course C, schdl S where R.c_id=C.c_id and R.pkg_id=P.pkg_id and  R.sch_id=S.sch_id and upper(R.RSTUD_GNDR)='FEMALE' order by R.rstud_reg_no"
  Adodc1.Refresh
  Txtstud.Text = "FEMALE"
 ElseIf Index = 3 Then 'package Student
  Adodc1.RecordSource = "select R.rstud_reg_no, R.rstud_nm, R.rstud_father_nm ,R.RSTUD_STATUS,  R.rstud_mob, R.RSTUD_DOJ, C.c_nm, P.Pkg_nm, S.SCH_TIMING  from rstud R, pkg P, Course C, schdl S where R.c_id=C.c_id and R.pkg_id=P.pkg_id and  R.sch_id=S.sch_id and upper(R.RSTUD_STATUS)='REGISTERED' order by R.rstud_reg_no"
  Adodc1.Refresh
  Txtstud.Text = "REGISTERED"
 ElseIf Index = 4 Then 'Non package
 ' Adodc1.RecordSource = "select R.rstud_reg_no, R.rstud_nm, R.rstud_father_nm ,R.RSTUD_STATUS,  R.rstud_mob, R.RSTUD_DOJ, C.c_nm, P.Pkg_nm, S.SCH_TIMING  from rstud R, pkg P, Course C, schdl S where R.c_id=C.c_id and R.pkg_id=P.pkg_id and  R.sch_id=S.sch_id and upper(R.RSTUD_STATUS)<>'REGISTERED' order by R.rstud_reg_no"
 ' Adodc1.Refresh
 ' Txtstud.Text = "UNREGISTERED"
  FrmNonpkg.Show 1, MDI
 ElseIf Index = 5 Then 'Today Enrolled
  Adodc1.RecordSource = "select R.rstud_reg_no, R.rstud_nm, R.rstud_father_nm ,R.RSTUD_STATUS,  R.rstud_mob, R.RSTUD_DOJ, C.c_nm, P.Pkg_nm, S.SCH_TIMING  from rstud R, pkg P, Course C, schdl S where R.c_id=C.c_id and R.pkg_id=P.pkg_id and  R.sch_id=S.sch_id and R.RSTUD_DOJ ='" & Format(Date, "dd-mmm-yyyy") & "' order by R.rstud_reg_no"
  Adodc1.Refresh
  Txtstud.Text = Format(Date, "dd-mmm-yyyy")
  txtStud2.Text = Format(Date, "dd-mmm-yyyy")
 ElseIf Index = 6 Then 'Today expired Pkg
  Adodc1.RecordSource = "select R.rstud_reg_no, R.rstud_nm, R.rstud_father_nm ,R.RSTUD_STATUS,  R.rstud_mob, R.RSTUD_DOJ, C.c_nm, P.Pkg_nm, S.SCH_TIMING  from rstud R, pkg P, Course C, schdl S where R.c_id=C.c_id and R.pkg_id=P.pkg_id and  R.sch_id=S.sch_id and R.RSTUD_DOE ='" & Format(Date, "dd-mmm-yyyy") & "'order by R.rstud_reg_no"
  Adodc1.Refresh
  Txtstud.Text = Format(Date, "dd-mmm-yyyy")
 ElseIf Index = 7 Then 'Pending request
 Adodc1.RecordSource = "select R.rstud_reg_no, R.rstud_nm, R.rstud_father_nm ,R.RSTUD_STATUS,  R.rstud_mob, R.RSTUD_DOJ, C.c_nm, P.Pkg_nm, S.SCH_TIMING  from rstud R, pkg P, Course C, schdl S, PKG_RENEW X where R.c_id=C.c_id and R.pkg_id=P.pkg_id and  R.sch_id=S.sch_id and R.RSTUD_REG_NO=X.RSTUD_REG_NO order by R.rstud_reg_no"
 Adodc1.Refresh
 End If
End Sub

Private Sub Text1_GotFocus()
Text1.Text = ""
Calendar1.Visible = True
End Sub

Private Sub Text1_LostFocus()
 Calendar1.Visible = False
End Sub
Private Sub Text4_GotFocus()
Text4.Text = ""
cld.Visible = True
End Sub

Private Sub Text4_LostFocus()
cld.Visible = False
End Sub

Private Sub Text5_GotFocus()
 Text5.Text = ""
 cld1.Visible = True
End Sub

Private Sub Text5_LostFocus()
cld1.Visible = False
End Sub

Private Sub Text6_GotFocus()
If Text1.Text <> "" Then
 Text6.Text = ""
 Calendar2.Visible = True
 Else
 MsgBox "Select Start Date First", vbInformation + vbOKOnly, " "
 Text1.SetFocus
End If
End Sub

Private Sub Text6_LostFocus()
 Calendar2.Visible = False
End Sub

Private Sub Text7_GotFocus()
 Text7.Text = ""
 cld3.Visible = True
End Sub

Private Sub Text7_LostFocus()
cld3.Visible = False
End Sub

Private Sub Text8_GotFocus()
Text8.Text = ""
cld2.Visible = True
End Sub

Private Sub Text8_LostFocus()
cld2.Visible = False
End Sub


Private Sub Timer1_Timer()
On Error Resume Next
Frame1.Top = 660
If navigation_counter = 0 Then
Frame1.Height = Frame1.Height - 700
If Frame1.Height <= 615 Then
navigation_counter = 1
 Timer1.Enabled = False
End If
ElseIf navigation_counter = 1 Then
Frame1.Height = Frame1.Height + 700
If Frame1.Height >= 9855 Then
navigation_counter = 0
 Timer1.Enabled = False
End If
End If
End Sub


