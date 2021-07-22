VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form Search_registered 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Search Student"
   ClientHeight    =   10785
   ClientLeft      =   -60
   ClientTop       =   -45
   ClientWidth     =   20370
   Icon            =   "Search_regis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10785
   ScaleWidth      =   20370
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command13 
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
      Height          =   375
      Left            =   18960
      MouseIcon       =   "Search_regis.frx":0EE2
      MousePointer    =   99  'Custom
      Picture         =   "Search_regis.frx":1034
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   10120
      Width           =   1315
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   9720
      Top             =   4680
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      RecordSource    =   $"Search_regis.frx":1831
      Caption         =   "Adodc2"
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
   Begin MSACAL.Calendar Calendar1 
      Height          =   3015
      Left            =   360
      TabIndex        =   27
      Top             =   1290
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
   Begin MSACAL.Calendar Calendar2 
      Height          =   3015
      Left            =   2760
      TabIndex        =   26
      Top             =   1290
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Search_regis.frx":1950
      Height          =   7995
      Left            =   45
      TabIndex        =   51
      Top             =   2055
      Width           =   20325
      _ExtentX        =   35851
      _ExtentY        =   14102
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   21
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "RSTUD_REG_NO"
         Caption         =   "   Reg. No"
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
         Caption         =   "Student Name"
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
         Caption         =   "Father's Name"
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
         DataField       =   "RSTUD_MOB"
         Caption         =   "        Mobile No"
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
         DataField       =   "RSTUD_GNDR"
         Caption         =   "     Gender"
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
         DataField       =   "RSTUD_ADHR"
         Caption         =   "             Adhar No"
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
         DataField       =   "RSTUD_DOJ"
         Caption         =   "     Joining date"
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
         DataField       =   "C_NM"
         Caption         =   "           Course"
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
         DataField       =   "PKG_NM"
         Caption         =   "       Package"
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
      BeginProperty Column09 
         DataField       =   "SCH_TIMING"
         Caption         =   "          Schedule"
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
      BeginProperty Column10 
         DataField       =   "RSTUD_AMNT"
         Caption         =   "      Amount"
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
         MarqueeStyle    =   3
         ScrollBars      =   2
         Locked          =   -1  'True
         RecordSelectors =   0   'False
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   1170.142
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2550.047
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2550.047
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1830.047
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1335.118
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   2204.788
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   1665.071
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   1785.26
         EndProperty
         BeginProperty Column08 
            Alignment       =   2
            ColumnWidth     =   1649.764
         EndProperty
         BeginProperty Column09 
            Alignment       =   2
            ColumnWidth     =   1844.787
         EndProperty
         BeginProperty Column10 
            Alignment       =   2
            ColumnWidth     =   1544.882
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   1600
      Left            =   20300
      TabIndex        =   43
      Top             =   380
      Width           =   80
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   1600
      Left            =   20400
      TabIndex        =   42
      Top             =   360
      Width           =   80
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   1600
      Left            =   0
      TabIndex        =   41
      Top             =   380
      Width           =   80
   End
   Begin VB.Frame Frame8 
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   11000
      TabIndex        =   40
      Top             =   395
      Width           =   5750
      Begin VB.ComboBox Combo3 
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
         ItemData        =   "Search_regis.frx":1965
         Left            =   1680
         List            =   "Search_regis.frx":196F
         MouseIcon       =   "Search_regis.frx":198A
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton Command12 
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
         Height          =   375
         Left            =   3900
         MouseIcon       =   "Search_regis.frx":1ADC
         MousePointer    =   99  'Custom
         Picture         =   "Search_regis.frx":1C2E
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   350
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
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
         Left            =   1680
         MouseIcon       =   "Search_regis.frx":258D
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   365
         Width           =   1935
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order     :"
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
         Left            =   500
         TabIndex        =   50
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sort by   : "
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
         Left            =   480
         TabIndex        =   44
         Top             =   360
         Width           =   1065
      End
      Begin VB.Shape Shape24 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         Height          =   1320
         Left            =   180
         Top             =   120
         Width           =   5295
      End
      Begin VB.Shape Shape23 
         BackColor       =   &H80000009&
         BackStyle       =   1  'Opaque
         Height          =   1500
         Left            =   120
         Top             =   0
         Width           =   5415
      End
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Print"
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
      Left            =   9120
      MouseIcon       =   "Search_regis.frx":26DF
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Click To Print"
      Top             =   840
      Width           =   1335
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   9180
      TabIndex        =   33
      Top             =   905
      Width           =   1335
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H80000009&
      Height          =   1525
      Left            =   16700
      TabIndex        =   28
      Top             =   390
      Width           =   3495
      Begin VB.Line Line7 
         BorderWidth     =   2
         X1              =   35
         X2              =   3460
         Y1              =   765
         Y2              =   765
      End
      Begin VB.Label Label10 
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2325
         TabIndex        =   32
         Top             =   900
         Width           =   855
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Today Enrolled  :"
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
         Left            =   480
         TabIndex        =   31
         Top             =   900
         Width           =   1740
      End
      Begin VB.Label Label9 
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2325
         TabIndex        =   30
         Top             =   330
         Width           =   855
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Students   : "
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
         Left            =   480
         TabIndex        =   29
         Top             =   330
         Width           =   1830
      End
      Begin VB.Shape Shape11 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         Height          =   1300
         Left            =   80
         Top             =   150
         Width           =   3355
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10080
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20490
      _ExtentX        =   36142
      _ExtentY        =   17780
      _Version        =   393216
      MousePointer    =   99
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   670
      BackColor       =   0
      MouseIcon       =   "Search_regis.frx":2831
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Search By ID"
      TabPicture(0)   =   "Search_regis.frx":2993
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape15"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Shape16"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Search By Name   "
      TabPicture(1)   =   "Search_regis.frx":29AF
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "Shape13"
      Tab(1).Control(2)=   "Line2"
      Tab(1).Control(3)=   "Shape14"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Search By Course   "
      TabPicture(2)   =   "Search_regis.frx":29CB
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(1)=   "Line4"
      Tab(2).Control(2)=   "lb"
      Tab(2).Control(3)=   "Shape18"
      Tab(2).Control(4)=   "Shape17"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Search By Package"
      TabPicture(3)   =   "Search_regis.frx":29E7
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Text5"
      Tab(3).Control(1)=   "Frame4"
      Tab(3).Control(2)=   "lb2"
      Tab(3).Control(3)=   "Line5"
      Tab(3).Control(4)=   "Shape20"
      Tab(3).Control(5)=   "Shape19"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "Search  By Joining Date "
      TabPicture(4)   =   "Search_regis.frx":2A03
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Shape21"
      Tab(4).Control(1)=   "Shape22"
      Tab(4).Control(2)=   "Line6"
      Tab(4).Control(3)=   "Frame5"
      Tab(4).ControlCount=   4
      Begin VB.ComboBox Text5 
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
         Left            =   -72840
         TabIndex        =   36
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   975
         Left            =   -74520
         TabIndex        =   15
         Top             =   720
         Width           =   9255
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   0
            Width           =   3015
         End
         Begin VB.CommandButton Command10 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Search"
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
            Left            =   5160
            MouseIcon       =   "Search_regis.frx":2A1F
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton Command9 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Refresh"
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
            Left            =   6960
            MouseIcon       =   "Search_regis.frx":2B71
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Course  :"
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
            Left            =   0
            TabIndex        =   47
            Top             =   0
            Width           =   1590
         End
         Begin VB.Label lb3 
            BackColor       =   &H8000000D&
            Height          =   375
            Left            =   4920
            TabIndex        =   39
            Top             =   480
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Enter Package :"
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
            Left            =   0
            TabIndex        =   18
            Top             =   525
            Width           =   1590
         End
         Begin VB.Shape Shape5 
            BackColor       =   &H80000007&
            BackStyle       =   1  'Opaque
            Height          =   375
            Left            =   5205
            Top             =   180
            Width           =   1455
         End
         Begin VB.Shape Shape3 
            BackColor       =   &H80000007&
            BackStyle       =   1  'Opaque
            Height          =   375
            Left            =   7005
            Top             =   180
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   855
         Left            =   -74520
         TabIndex        =   11
         Top             =   720
         Width           =   9255
         Begin VB.ComboBox Text4 
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
            Left            =   1680
            TabIndex        =   35
            Top             =   120
            Width           =   3015
         End
         Begin VB.CommandButton Command8 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Search"
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
            Left            =   5160
            MouseIcon       =   "Search_regis.frx":2CC3
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton Command7 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Refresh"
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
            Left            =   6960
            MouseIcon       =   "Search_regis.frx":2E15
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Enter Course :"
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
            Left            =   0
            TabIndex        =   14
            Top             =   165
            Width           =   1470
         End
         Begin VB.Shape Shape10 
            BackColor       =   &H80000007&
            BackStyle       =   1  'Opaque
            Height          =   375
            Left            =   5205
            Top             =   180
            Width           =   1455
         End
         Begin VB.Shape Shape9 
            BackColor       =   &H80000007&
            BackStyle       =   1  'Opaque
            Height          =   375
            Left            =   7005
            Top             =   180
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   735
         Left            =   -74520
         TabIndex        =   6
         Top             =   720
         Width           =   9255
         Begin VB.TextBox Text3 
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
            Left            =   1680
            TabIndex        =   9
            Top             =   150
            Width           =   3015
         End
         Begin VB.CommandButton Command6 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Search"
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
            Left            =   5160
            MouseIcon       =   "Search_regis.frx":2F67
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Refresh"
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
            Left            =   6960
            MouseIcon       =   "Search_regis.frx":30B9
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Enter Name :"
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
            Left            =   120
            TabIndex        =   10
            Top             =   165
            Width           =   1365
         End
         Begin VB.Shape Shape8 
            BackColor       =   &H80000007&
            BackStyle       =   1  'Opaque
            Height          =   375
            Left            =   5205
            Top             =   180
            Width           =   1455
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H80000007&
            BackStyle       =   1  'Opaque
            Height          =   375
            Left            =   7005
            Top             =   180
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   735
         Left            =   480
         TabIndex        =   1
         Top             =   720
         Width           =   9255
         Begin VB.CommandButton Command5 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Refresh"
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
            Left            =   6960
            MouseIcon       =   "Search_regis.frx":320B
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Search"
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
            Left            =   5160
            MouseIcon       =   "Search_regis.frx":335D
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   120
            Width           =   1455
         End
         Begin VB.TextBox Text2 
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
            Left            =   1680
            TabIndex        =   3
            ToolTipText     =   "Enter Text To Search"
            Top             =   150
            Width           =   3015
         End
         Begin VB.Shape Shape7 
            BackColor       =   &H80000007&
            BackStyle       =   1  'Opaque
            Height          =   375
            Left            =   7005
            Top             =   180
            Width           =   1335
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H80000007&
            BackStyle       =   1  'Opaque
            Height          =   375
            Left            =   5205
            Top             =   180
            Width           =   1455
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Enter ID :"
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
            Left            =   240
            TabIndex        =   2
            Top             =   165
            Width           =   960
         End
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   735
         Left            =   -74760
         TabIndex        =   19
         Top             =   720
         Width           =   10575
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
            Left            =   3240
            TabIndex        =   24
            Top             =   150
            Width           =   1575
         End
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
            TabIndex        =   22
            Top             =   150
            Width           =   1455
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Search"
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
            Left            =   5400
            MouseIcon       =   "Search_regis.frx":34AF
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Refresh"
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
            Left            =   7200
            MouseIcon       =   "Search_regis.frx":3601
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   120
            Width           =   1335
         End
         Begin VB.Line Line1 
            X1              =   -5
            X2              =   -5
            Y1              =   -15
            Y2              =   735
         End
         Begin VB.Label Label6 
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
            Left            =   2640
            TabIndex        =   25
            Top             =   165
            Width           =   390
         End
         Begin VB.Label Label5 
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
            Left            =   120
            TabIndex        =   23
            Top             =   165
            Width           =   690
         End
         Begin VB.Shape Shape6 
            BackColor       =   &H80000007&
            BackStyle       =   1  'Opaque
            Height          =   375
            Left            =   5445
            Top             =   180
            Width           =   1455
         End
         Begin VB.Shape Shape4 
            BackColor       =   &H80000007&
            BackStyle       =   1  'Opaque
            Height          =   375
            Left            =   7245
            Top             =   180
            Width           =   1335
         End
      End
      Begin VB.Shape Shape13 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         Height          =   1350
         Left            =   -74760
         Top             =   495
         Width           =   10695
      End
      Begin VB.Label lb2 
         BackColor       =   &H0000FF00&
         Height          =   375
         Left            =   -69600
         TabIndex        =   38
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Line Line6 
         BorderWidth     =   3
         X1              =   -75000
         X2              =   -54600
         Y1              =   2000
         Y2              =   2000
      End
      Begin VB.Line Line5 
         BorderWidth     =   3
         X1              =   -75000
         X2              =   -54600
         Y1              =   2000
         Y2              =   2000
      End
      Begin VB.Line Line4 
         BorderWidth     =   3
         X1              =   -75000
         X2              =   -54600
         Y1              =   2000
         Y2              =   2000
      End
      Begin VB.Line Line3 
         BorderWidth     =   3
         X1              =   0
         X2              =   20400
         Y1              =   2000
         Y2              =   2000
      End
      Begin VB.Line Line2 
         BorderWidth     =   3
         X1              =   -75000
         X2              =   -54600
         Y1              =   2000
         Y2              =   2000
      End
      Begin VB.Label lb 
         BackColor       =   &H000000FF&
         Height          =   255
         Left            =   -73680
         TabIndex        =   37
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Shape Shape22 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         Height          =   1350
         Left            =   -74765
         Top             =   500
         Width           =   10695
      End
      Begin VB.Shape Shape21 
         BackColor       =   &H80000009&
         BackStyle       =   1  'Opaque
         Height          =   1520
         Left            =   -74817
         Top             =   395
         Width           =   10815
      End
      Begin VB.Shape Shape20 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         Height          =   1350
         Left            =   -74765
         Top             =   500
         Width           =   10695
      End
      Begin VB.Shape Shape19 
         BackColor       =   &H80000009&
         BackStyle       =   1  'Opaque
         Height          =   1520
         Left            =   -74820
         Top             =   395
         Width           =   10815
      End
      Begin VB.Shape Shape18 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         Height          =   1350
         Left            =   -74765
         Top             =   500
         Width           =   10695
      End
      Begin VB.Shape Shape17 
         BackColor       =   &H80000009&
         BackStyle       =   1  'Opaque
         Height          =   1520
         Left            =   -74817
         Top             =   395
         Width           =   10815
      End
      Begin VB.Shape Shape16 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         Height          =   1350
         Left            =   235
         Top             =   500
         Width           =   10695
      End
      Begin VB.Shape Shape15 
         BackColor       =   &H80000009&
         BackStyle       =   1  'Opaque
         Height          =   1520
         Left            =   183
         Top             =   395
         Width           =   10815
      End
      Begin VB.Shape Shape14 
         BackColor       =   &H80000009&
         BackStyle       =   1  'Opaque
         Height          =   1520
         Left            =   -74817
         Top             =   395
         Width           =   10822
      End
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click To see Non registered Student Details. ( Above Records display only Registered Students )"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   150
      MouseIcon       =   "Search_regis.frx":3753
      MousePointer    =   99  'Custom
      TabIndex        =   52
      Top             =   10125
      Width           =   8790
   End
   Begin VB.Shape Shape25 
      BackColor       =   &H80000013&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   0
      Top             =   10080
      Width           =   20475
   End
   Begin VB.Shape Shape12 
      BackColor       =   &H80000007&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   9165
      Top             =   900
      Width           =   1335
   End
End
Attribute VB_Name = "Search_registered"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub combo2_Click()
Text5.Clear
Set r = c.Execute("select c_id from course where upper(c_nm)='" & UCase(Combo2.Text) & "' ")
If r.EOF = False Then
 lb3.Caption = r.Fields(0)
End If
Set r1 = c.Execute("select initcap(Pkg_nm) from pkg where c_id='" & lb3.Caption & "' ")
While r1.EOF = False
 Text5.AddItem r1.Fields(0)
r1.MoveNext
Wend
End Sub

Private Sub Command1_Click()
Command2_Click
End Sub

Private Sub Command10_Click()
If Combo2.Text = "" Then
  MsgBox "Select Course option !!", vbInformation + vbOKOnly, "Select Course"
 Combo2.SetFocus
 Exit Sub
ElseIf Trim(Text5.Text) = "" Then
  MsgBox "Select Or Enter Package !!", vbInformation + vbOKOnly, "Select package"
 Text5.SetFocus
 Exit Sub
End If
If lb2.Caption = "" Then
Set r = New ADODB.Recordset
Set r = c.Execute("select pkg_id from pkg where c_id='" & lb3.Caption & "' and upper(pkg_id)='" & UCase(Trim(Text5.Text)) & "' or upper(pkg_nm)= '" & UCase(Trim(Text5.Text)) & "' ")
If r.EOF = False Then
 lb2.Caption = r.Fields(0)
End If
End If
Adodc1.RecordSource = " select R.rstud_reg_no, R.rstud_nm, R.rstud_father_nm , R.rstud_mob, R.rstud_gndr, R.RSTUD_ADHR,  R.RSTUD_DOJ, C.c_nm, P.Pkg_nm, S.SCH_TIMING, R.RSTUD_AMNT  from rstud R, pkg P, Course C, schdl S where R.c_id=c.c_id and R.pkg_id=P.pkg_id and R.sch_id=S.sch_id and upper(R.Pkg_id)='" & UCase(lb2.Caption) & "' and R.c_id='" & lb3.Caption & "' "
Adodc1.Refresh
End Sub

Private Sub Command11_Click()
c.Execute ("delete from RPTStudentS")
Set r = New ADODB.Recordset
Set r = c.Execute(Adodc1.RecordSource)
If r.EOF = False Then
 Dim ggg As Integer
 ggg = 0
 While r.EOF = False
 ggg = ggg + 1
 r.MoveNext
 Wend
sql = "Insert into RPTStudentS " & Adodc1.RecordSource
 c.Execute (sql)
RptStudentSearch.Sections("section4").Controls("totstu").Caption = Val(ggg)
RptStudentSearch.Show
Else
MsgBox "No Record To Print..", vbInformation + vbOKOnly, ""
End If
End Sub

Private Sub Command12_Click()
If Combo1.Text = "" Then
 MsgBox "Select Sort By option !!", vbInformation + vbOKOnly, "Sorting"
 Combo1.SetFocus
 Exit Sub
End If
If Combo3.Text = "" Then
 MsgBox "Select Order of sorting !!", vbInformation + vbOKOnly, "Sorting"
 Combo3.SetFocus
 Exit Sub
End If
If Combo1.Text = "Reg. No" Then
If Combo3.ListIndex = 0 Then
 Adodc1.RecordSource = "select R.rstud_reg_no, R.rstud_nm, R.rstud_father_nm , R.rstud_mob, R.rstud_gndr, R.RSTUD_ADHR,  R.RSTUD_DOJ, C.c_nm, P.Pkg_nm, S.SCH_TIMING, R.RSTUD_AMNT  from rstud R, pkg P, Course C, schdl S where R.c_id=c.c_id and R.pkg_id=P.pkg_id  and R.sch_id=S.sch_id order by R.rstud_reg_no"
 Adodc1.Refresh
Else
 Adodc1.RecordSource = "select R.rstud_reg_no, R.rstud_nm, R.rstud_father_nm , R.rstud_mob, R.rstud_gndr, R.RSTUD_ADHR,  R.RSTUD_DOJ, C.c_nm, P.Pkg_nm, S.SCH_TIMING, R.RSTUD_AMNT  from rstud R, pkg P, Course C, schdl S where R.c_id=c.c_id and R.pkg_id=P.pkg_id  and R.sch_id=S.sch_id order by R.rstud_reg_no desc"
 Adodc1.Refresh
End If
ElseIf Combo1.Text = "Name" Then
If Combo3.ListIndex = 0 Then
 Adodc1.RecordSource = "select R.rstud_reg_no, R.rstud_nm, R.rstud_father_nm , R.rstud_mob, R.rstud_gndr, R.RSTUD_ADHR,  R.RSTUD_DOJ, C.c_nm, P.Pkg_nm, S.SCH_TIMING, R.RSTUD_AMNT  from rstud R, pkg P, Course C, schdl S where R.c_id=c.c_id and R.pkg_id=P.pkg_id  and R.sch_id=S.sch_id order by R.rstud_nm"
 Adodc1.Refresh
Else
 Adodc1.RecordSource = "select R.rstud_reg_no, R.rstud_nm, R.rstud_father_nm , R.rstud_mob, R.rstud_gndr, R.RSTUD_ADHR,  R.RSTUD_DOJ, C.c_nm, P.Pkg_nm, S.SCH_TIMING, R.RSTUD_AMNT  from rstud R, pkg P, Course C, schdl S where R.c_id=c.c_id and R.pkg_id=P.pkg_id  and R.sch_id=S.sch_id order by R.rstud_nm desc"
 Adodc1.Refresh
End If
ElseIf Combo1.Text = "Gender" Then
If Combo3.ListIndex = 0 Then
 Adodc1.RecordSource = "select R.rstud_reg_no, R.rstud_nm, R.rstud_father_nm , R.rstud_mob, R.rstud_gndr, R.RSTUD_ADHR,  R.RSTUD_DOJ, C.c_nm, P.Pkg_nm, S.SCH_TIMING, R.RSTUD_AMNT  from rstud R, pkg P, Course C, schdl S where R.c_id=c.c_id and R.pkg_id=P.pkg_id  and R.sch_id=S.sch_id order by R.rstud_gndr"
 Adodc1.Refresh
Else
 Adodc1.RecordSource = "select R.rstud_reg_no, R.rstud_nm, R.rstud_father_nm , R.rstud_mob, R.rstud_gndr, R.RSTUD_ADHR,  R.RSTUD_DOJ, C.c_nm, P.Pkg_nm, S.SCH_TIMING, R.RSTUD_AMNT  from rstud R, pkg P, Course C, schdl S where R.c_id=c.c_id and R.pkg_id=P.pkg_id  and R.sch_id=S.sch_id order by R.rstud_gndr desc"
 Adodc1.Refresh
End If
ElseIf Combo1.Text = "Package" Then
If Combo3.ListIndex = 0 Then
 Adodc1.RecordSource = "select R.rstud_reg_no, R.rstud_nm, R.rstud_father_nm , R.rstud_mob, R.rstud_gndr, R.RSTUD_ADHR,  R.RSTUD_DOJ, C.c_nm, P.Pkg_nm, S.SCH_TIMING, R.RSTUD_AMNT  from rstud R, pkg P, Course C, schdl S where R.c_id=c.c_id and R.pkg_id=P.pkg_id  and R.sch_id=S.sch_id order by P.pkg_nm"
 Adodc1.Refresh
Else
 Adodc1.RecordSource = "select R.rstud_reg_no, R.rstud_nm, R.rstud_father_nm , R.rstud_mob, R.rstud_gndr, R.RSTUD_ADHR,  R.RSTUD_DOJ, C.c_nm, P.Pkg_nm, S.SCH_TIMING, R.RSTUD_AMNT  from rstud R, pkg P, Course C, schdl S where R.c_id=c.c_id and R.pkg_id=P.pkg_id  and R.sch_id=S.sch_id order by P.pkg_nm desc"
 Adodc1.Refresh
End If
ElseIf Combo1.Text = "Course" Then
If Combo3.ListIndex = 0 Then
 Adodc1.RecordSource = "select R.rstud_reg_no, R.rstud_nm, R.rstud_father_nm , R.rstud_mob, R.rstud_gndr, R.RSTUD_ADHR,  R.RSTUD_DOJ, C.c_nm, P.Pkg_nm, S.SCH_TIMING, R.RSTUD_AMNT  from rstud R, pkg P, Course C, schdl S where R.c_id=c.c_id and R.pkg_id=P.pkg_id  and R.sch_id=S.sch_id order by C.c_nm"
 Adodc1.Refresh
Else
 Adodc1.RecordSource = "select R.rstud_reg_no, R.rstud_nm, R.rstud_father_nm , R.rstud_mob, R.rstud_gndr, R.RSTUD_ADHR,  R.RSTUD_DOJ, C.c_nm, P.Pkg_nm, S.SCH_TIMING, R.RSTUD_AMNT  from rstud R, pkg P, Course C, schdl S where R.c_id=c.c_id and R.pkg_id=P.pkg_id  and R.sch_id=S.sch_id order by C.c_nm desc"
 Adodc1.Refresh
End If
ElseIf Combo1.Text = "Join date" Then
If Combo3.ListIndex = 0 Then
 Adodc1.RecordSource = "select R.rstud_reg_no, R.rstud_nm, R.rstud_father_nm , R.rstud_mob, R.rstud_gndr, R.RSTUD_ADHR,  R.RSTUD_DOJ, C.c_nm, P.Pkg_nm, S.SCH_TIMING, R.RSTUD_AMNT  from rstud R, pkg P, Course C, schdl S where R.c_id=c.c_id and R.pkg_id=P.pkg_id  and R.sch_id=S.sch_id order by R.rstud_doj"
 Adodc1.Refresh
Else
 Adodc1.RecordSource = "select R.rstud_reg_no, R.rstud_nm, R.rstud_father_nm , R.rstud_mob, R.rstud_gndr, R.RSTUD_ADHR,  R.RSTUD_DOJ, C.c_nm, P.Pkg_nm, S.SCH_TIMING, R.RSTUD_AMNT  from rstud R, pkg P, Course C, schdl S where R.c_id=c.c_id and R.pkg_id=P.pkg_id  and R.sch_id=S.sch_id order by R.rstud_doj desc"
 Adodc1.Refresh
End If
End If
End Sub

Private Sub Command13_Click()
Unload Me
End Sub


Private Sub Command2_Click()
Adodc1.RecordSource = "select R.rstud_reg_no, R.rstud_nm, R.rstud_father_nm , R.rstud_mob, R.rstud_gndr, R.RSTUD_ADHR,  R.RSTUD_DOJ, C.c_nm, P.Pkg_nm, S.SCH_TIMING, R.RSTUD_AMNT  from rstud R, pkg P, Course C, schdl S where R.c_id=c.c_id and R.pkg_id=P.pkg_id and  R.sch_id=S.sch_id order by R.rstud_reg_no"
Adodc1.Refresh
End Sub

Private Sub Command3_Click() 'By Date
Adodc1.RecordSource = "select R.rstud_reg_no, R.rstud_nm, R.rstud_father_nm , R.rstud_mob, R.rstud_gndr, R.RSTUD_ADHR,  R.RSTUD_DOJ, C.c_nm, P.Pkg_nm, S.SCH_TIMING, R.RSTUD_AMNT  from rstud R, pkg P, Course C, schdl S where R.c_id=c.c_id and R.pkg_id=P.pkg_id  and R.sch_id=S.sch_id and R.rstud_doj between '" & Format(Text1.Text, "dd-mmm-yyyy") & "' and '" & Format(Text6.Text, "dd-mmm-yyyy") & "' order by R.rstud_reg_no "
Adodc1.Refresh
End Sub

Private Sub Command4_Click()
Adodc1.RecordSource = "select R.rstud_reg_no, R.rstud_nm, R.rstud_father_nm , R.rstud_mob, R.rstud_gndr, R.RSTUD_ADHR, R.RSTUD_DOJ, C.c_nm, P.Pkg_nm , S.SCH_TIMING, R.RSTUD_AMNT  from rstud R, pkg P, Course C, schdl S where R.c_id=c.c_id and R.pkg_id=P.pkg_id  and R.sch_id=S.sch_id and upper(R.rstud_reg_no) like '" & UCase(Text2.Text) & "%' order by R.rstud_reg_no"
Adodc1.Refresh
End Sub

Private Sub Command5_Click()
Command2_Click
End Sub

Private Sub Command6_Click()
Adodc1.RecordSource = "select R.rstud_reg_no, R.rstud_nm, R.rstud_father_nm , R.rstud_mob, R.rstud_gndr, R.RSTUD_ADHR,  R.RSTUD_DOJ, C.c_nm, P.Pkg_nm, S.SCH_TIMING, R.RSTUD_AMNT  from rstud R, pkg P, Course C, schdl S where R.c_id=c.c_id and R.pkg_id=P.pkg_id and R.sch_id=S.sch_id and upper(R.rstud_nm) like '" & UCase(Text3.Text) & "%' order by R.rstud_reg_no"
Adodc1.Refresh
End Sub

Private Sub Command7_Click()
Command2_Click
End Sub

Private Sub Command8_Click()
If Trim(Text4.Text) = "" Then
 MsgBox "Enter Value first to Search ", vbQuestion + vbOKOnly, "Empty"
 Text4.SetFocus
 Exit Sub
 End If
Set r = New ADODB.Recordset
Set r = c.Execute("select c_id from course where upper(c_id)='" & UCase(Trim(Text4.Text)) & "' or upper(c_nm)like '" & UCase(Trim(Text4.Text)) & "%' ")
If r.EOF = False Then
 lb.Caption = r.Fields(0)
 Else
  lb.Caption = ""
 End If
Adodc1.RecordSource = " select R.rstud_reg_no, R.rstud_nm, R.rstud_father_nm , R.rstud_mob, R.rstud_gndr, R.RSTUD_ADHR,  R.RSTUD_DOJ, C.c_nm, P.Pkg_nm, S.SCH_TIMING, R.RSTUD_AMNT  from rstud R, pkg P, Course C, schdl S where R.c_id=c.c_id and R.pkg_id=P.pkg_id and R.sch_id=S.sch_id and  upper(R.c_id)='" & UCase(lb.Caption) & "' "
Adodc1.Refresh
End Sub

Private Sub Command9_Click()
Command2_Click
End Sub

Private Sub Form_Load()
conn
Me.Width = MDI.Width
SSTab1.Width = Me.Width
DataGrid2.Width = Me.Width - 80
SSTab1.Tab = 1
Me.Top = 0
Me.Left = 0
lb.Caption = ""
lb2.Caption = ""
lb3.Caption = ""
Calendar2.Visible = False
Calendar1.Visible = False
Label9.Caption = 0
Label10.Caption = 0
Combo1.Clear
Combo3.Clear
Combo3.AddItem "Ascending"
Combo3.AddItem "Descending"
Calendar1.Value = Format(Date, "DD-MMM-YY")
Calendar2.Value = Format(Date, "DD-MMM-YY")
Combo1.AddItem "Reg. No"
Combo1.AddItem "Name"
Combo1.AddItem "Gender"
Combo1.AddItem "Package"
Combo1.AddItem "Course"
Combo1.AddItem "Join date"

Set r = New ADODB.Recordset
Set r = c.Execute("select count(*) from rstud")
If r.EOF = False Then
Label9.Caption = r.Fields(0)
End If
Set r = c.Execute("select count(*) from rstud where rstud_doj='" & Format(Date, "dd-mmm-yyyy") & "' ")
If r.EOF = False Then
Label10.Caption = r.Fields(0)
End If
Text4.Clear
Combo2.Clear
Set r = New ADODB.Recordset
Set r = c.Execute("select initcap(c_nm) from course")
While r.EOF = False
 Text4.AddItem r.Fields(0)
 Combo2.AddItem r.Fields(0)
 r.MoveNext
Wend
End Sub


Private Sub Label14_Click()
FrmNonpkg.Show 1, MDI
End Sub

Private Sub Text1_GotFocus()
Text1.Text = ""
Calendar1.Visible = True
End Sub

Private Sub Text1_LostFocus()
 Calendar1.Visible = False
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Then
Text3.SetFocus
Else
KeyAscii = 0
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
Text2.SetFocus
Else
KeyAscii = 0
End If
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then

Else
KeyAscii = 0
End If
End Sub

Private Sub Text5_Click()
lb2.Caption = ""
Set r = c.Execute("select Pkg_id from pkg where upper(Pkg_nm)='" & UCase(Trim(Text5.Text)) & "' and c_id='" & lb3.Caption & "' ")
If r.EOF = False Then
 lb2.Caption = r.Fields(0)
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
Else
KeyAscii = 0
End If
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


Private Sub Text2_Change()
Adodc1.RecordSource = "select R.rstud_reg_no, R.rstud_nm, R.rstud_father_nm, R.rstud_mob, R.rstud_gndr, R.RSTUD_ADHR, R.RSTUD_DOJ, C.c_nm, P.Pkg_nm, S.SCH_TIMING, R.RSTUD_AMNT  from rstud R, pkg P, Course C, schdl S where R.c_id=c.c_id and R.pkg_id=P.pkg_id and R.sch_id=S.sch_id and upper(R.rstud_reg_no) like '" & UCase(Text2.Text) & "%' order by R.rstud_reg_no"
Adodc1.Refresh
End Sub

Private Sub Text3_Change()
Adodc1.RecordSource = "select R.rstud_reg_no, R.rstud_nm, R.rstud_father_nm , R.rstud_mob, R.rstud_gndr, R.RSTUD_ADHR,  R.RSTUD_DOJ, C.c_nm, P.Pkg_nm, S.SCH_TIMING, R.RSTUD_AMNT  from rstud R, pkg P, Course C, schdl S where R.c_id=c.c_id and R.pkg_id=P.pkg_id and R.sch_id=S.sch_id and upper(R.rstud_nm) like '" & UCase(Text3.Text) & "%' order by R.rstud_reg_no"
Adodc1.Refresh
End Sub
