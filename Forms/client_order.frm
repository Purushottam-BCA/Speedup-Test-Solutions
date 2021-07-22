VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form FrmClient2 
   BorderStyle     =   0  'None
   Caption         =   "Client Order"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
   Icon            =   "client_order.frx":0000
   LinkTopic       =   "Form9"
   MDIChild        =   -1  'True
   ScaleHeight     =   11520
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "See All"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   460
      Left            =   13560
      MouseIcon       =   "client_order.frx":08CA
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   55
      ToolTipText     =   "See All Client & Order Lists."
      Top             =   9405
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Order List"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10280
      Left            =   16440
      TabIndex        =   76
      Top             =   0
      Visible         =   0   'False
      Width           =   3975
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   600
         Top             =   1020
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
         Connect         =   "Provider=MSDAORA.1;Password=STS;User ID=STS;Persist Security Info=True"
         OLEDBString     =   "Provider=MSDAORA.1;Password=STS;User ID=STS;Persist Security Info=True"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select O.Ord_no,C. CLNT_NM,O. CSTATUS from client C, CLNT_ORDR_CHLN O where C. CLNT_ID=O. CLNT_ID"
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "client_order.frx":0A1C
         Height          =   9900
         Left            =   0
         TabIndex        =   77
         Top             =   360
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   17463
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   22
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "ORD_NO"
            Caption         =   "Order No"
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
            DataField       =   "CLNT_NM"
            Caption         =   "Client"
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
            DataField       =   "CSTATUS"
            Caption         =   "  Status"
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
            ScrollBars      =   2
            Locked          =   -1  'True
            RecordSelectors =   0   'False
            BeginProperty Column00 
               Alignment       =   2
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1695.118
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   1049.953
            EndProperty
         EndProperty
      End
   End
   Begin VB.TextBox n18 
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
      Height          =   330
      Left            =   14880
      Locked          =   -1  'True
      TabIndex        =   74
      Top             =   3960
      Width           =   1080
   End
   Begin VB.TextBox n17 
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
      Height          =   330
      Left            =   14880
      TabIndex        =   71
      Top             =   3360
      Width           =   1080
   End
   Begin VB.TextBox n16 
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
      Height          =   330
      Left            =   14865
      TabIndex        =   67
      Top             =   2760
      Width           =   1080
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2895
      Left            =   2400
      TabIndex        =   65
      Top             =   3600
      Visible         =   0   'False
      Width           =   3375
      _Version        =   524288
      _ExtentX        =   5953
      _ExtentY        =   5106
      _StockProps     =   1
      BackColor       =   16777215
      Year            =   2019
      Month           =   6
      Day             =   23
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
   Begin VB.TextBox n19 
      Height          =   375
      Left            =   840
      TabIndex        =   64
      Text            =   "Text1"
      Top             =   3120
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   5880
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox n1 
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
      Left            =   3120
      MouseIcon       =   "client_order.frx":0A31
      MousePointer    =   99  'Custom
      Style           =   2  'Dropdown List
      TabIndex        =   63
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox n15 
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
      Height          =   330
      Left            =   6600
      TabIndex        =   62
      Top             =   7200
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   11040
      MouseIcon       =   "client_order.frx":0B83
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   6840
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Default"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   11040
      MouseIcon       =   "client_order.frx":0CD5
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Browse"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   11040
      MouseIcon       =   "client_order.frx":0E27
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      DisabledPicture =   "client_order.frx":0F79
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   460
      Left            =   10455
      MouseIcon       =   "client_order.frx":1803
      MousePointer    =   99  'Custom
      Picture         =   "client_order.frx":1955
      Style           =   1  'Graphical
      TabIndex        =   56
      ToolTipText     =   "Add New Order."
      Top             =   9405
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      DisabledPicture =   "client_order.frx":21DF
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   460
      Left            =   11880
      MouseIcon       =   "client_order.frx":2C3A
      MousePointer    =   99  'Custom
      Picture         =   "client_order.frx":2D8C
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   "Save New Order."
      Top             =   9405
      Width           =   1695
   End
   Begin VB.TextBox n14 
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
      Height          =   330
      Left            =   9120
      TabIndex        =   47
      Top             =   5835
      Width           =   720
   End
   Begin VB.TextBox n8 
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
      Height          =   330
      Left            =   9120
      TabIndex        =   40
      Top             =   3240
      Width           =   960
   End
   Begin VB.TextBox n7 
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
      Height          =   330
      Left            =   9120
      TabIndex        =   39
      Top             =   2760
      Width           =   960
   End
   Begin VB.TextBox n5 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3120
      TabIndex        =   26
      Top             =   6360
      Width           =   1935
   End
   Begin VB.TextBox n3 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3120
      TabIndex        =   24
      Top             =   3840
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      MouseIcon       =   "client_order.frx":37E7
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Top             =   885
      Width           =   1695
   End
   Begin VB.TextBox Order 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   21
      Top             =   915
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   460
      Left            =   14745
      MouseIcon       =   "client_order.frx":3939
      MousePointer    =   99  'Custom
      Picture         =   "client_order.frx":3A8B
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Exit"
      Top             =   9405
      Width           =   1315
   End
   Begin VB.TextBox n4 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   3120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   4440
      Width           =   2655
   End
   Begin VB.TextBox n2 
      Height          =   405
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox n6 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3120
      TabIndex        =   1
      Top             =   7080
      Width           =   1935
   End
   Begin VB.CommandButton cmd_Bk 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Back"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MouseIcon       =   "client_order.frx":4288
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9405
      Width           =   1215
   End
   Begin VB.TextBox n9 
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
      Height          =   330
      Left            =   9120
      MaxLength       =   2
      TabIndex        =   27
      Top             =   3720
      Width           =   480
   End
   Begin VB.TextBox n10 
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
      Height          =   330
      Left            =   10005
      MaxLength       =   2
      TabIndex        =   28
      Top             =   3720
      Width           =   480
   End
   Begin VB.TextBox n13 
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
      Height          =   330
      Left            =   9120
      TabIndex        =   29
      Top             =   5280
      Width           =   720
   End
   Begin VB.TextBox n12 
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
      Height          =   330
      Left            =   9120
      TabIndex        =   30
      Top             =   4680
      Width           =   720
   End
   Begin VB.TextBox n11 
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
      Height          =   330
      Left            =   9120
      TabIndex        =   31
      Top             =   4200
      Width           =   2880
   End
   Begin VB.Label Label20 
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
      Left            =   14640
      TabIndex        =   75
      Top             =   3960
      Width           =   105
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Due  Amount   : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   12840
      TabIndex        =   73
      Top             =   3960
      Width           =   1290
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
      Left            =   14640
      TabIndex        =   72
      Top             =   3360
      Width           =   105
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Paid  Amount   : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   12840
      TabIndex        =   70
      Top             =   3360
      Width           =   1320
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount  : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   12840
      TabIndex        =   69
      Top             =   2880
      Width           =   1290
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
      Left            =   14625
      TabIndex        =   68
      Top             =   2760
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Payment info ( Required )"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   12720
      TabIndex        =   66
      Top             =   2160
      Width           =   2865
   End
   Begin VB.Label Label2 
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
      Left            =   8880
      TabIndex        =   58
      Top             =   6480
      Width           =   105
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "School / institute Logo  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   6600
      TabIndex        =   57
      Top             =   6600
      Width           =   1980
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1395
      Left            =   9120
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label Label21 
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
      Left            =   2880
      TabIndex        =   53
      Top             =   4560
      Width           =   105
   End
   Begin VB.Label Label19 
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
      Left            =   2880
      TabIndex        =   52
      Top             =   3840
      Width           =   105
   End
   Begin VB.Label Label10 
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
      Left            =   2880
      TabIndex        =   51
      Top             =   3240
      Width           =   105
   End
   Begin VB.Label Label5 
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
      Left            =   2880
      TabIndex        =   50
      Top             =   2640
      Width           =   105
   End
   Begin VB.Label Label71 
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
      Left            =   8880
      TabIndex        =   49
      Top             =   5760
      Width           =   105
   End
   Begin VB.Label Label70 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Papers :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   7440
      TabIndex        =   48
      Top             =   5790
      Width           =   1125
   End
   Begin VB.Label Label69 
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
      Left            =   8880
      TabIndex        =   46
      Top             =   5280
      Width           =   105
   End
   Begin VB.Label Label68 
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
      Left            =   8880
      TabIndex        =   45
      Top             =   4800
      Width           =   105
   End
   Begin VB.Label Label67 
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
      Left            =   8880
      TabIndex        =   44
      Top             =   4200
      Width           =   105
   End
   Begin VB.Label Label66 
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
      Left            =   8880
      TabIndex        =   43
      Top             =   3720
      Width           =   105
   End
   Begin VB.Label Label43 
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
      Left            =   8880
      TabIndex        =   42
      Top             =   3240
      Width           =   105
   End
   Begin VB.Label Label64 
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
      Left            =   8880
      TabIndex        =   41
      Top             =   2760
      Width           =   105
   End
   Begin VB.Label Label46 
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
      Left            =   10530
      TabIndex        =   33
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label44 
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
      Left            =   9690
      TabIndex        =   32
      Top             =   3720
      Width           =   225
   End
   Begin VB.Label Label53 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "School / Institute Name :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   750
      TabIndex        =   25
      Top             =   3960
      Width           =   1980
   End
   Begin VB.Label Label62 
      BackStyle       =   0  'Transparent
      Caption         =   "Order No :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label61 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Order Date :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   1755
      TabIndex        =   20
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label59 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Client Name :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   1780
      TabIndex        =   19
      Top             =   2760
      Width           =   1080
   End
   Begin VB.Label Label56 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "School / Institute Address :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   495
      TabIndex        =   18
      Top             =   4560
      Width           =   2235
   End
   Begin VB.Label Label55 
      BackStyle       =   0  'Transparent
      Caption         =   "Client Order Entry"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label54 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Order Information ( Required )"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   600
      TabIndex        =   16
      Top             =   2160
      Width           =   3375
   End
   Begin VB.Label Label52 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Exam Details ( Required )"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   600
      TabIndex        =   15
      Top             =   5880
      Width           =   2670
   End
   Begin VB.Label Label51 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Subject  : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   2040
      TabIndex        =   14
      Top             =   7080
      Width           =   825
   End
   Begin VB.Label Label42 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Name of exam  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   7215
      TabIndex        =   13
      Top             =   4320
      Width           =   1305
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Question Paper Details ( Required )"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   6360
      TabIndex        =   12
      Top             =   2160
      Width           =   3960
   End
   Begin VB.Label Label40 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Marks / Wrong  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   7200
      TabIndex        =   11
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label Label39 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Marks / Correct  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   7095
      TabIndex        =   10
      Top             =   4800
      Width           =   1425
   End
   Begin VB.Label Label38 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Time  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   7515
      TabIndex        =   9
      Top             =   3840
      Width           =   1005
   End
   Begin VB.Label Label37 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Marks  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   7395
      TabIndex        =   8
      Top             =   3360
      Width           =   1110
   End
   Begin VB.Label Label36 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Questions in Paper  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   6375
      TabIndex        =   7
      Top             =   2880
      Width           =   2160
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Class  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   2175
      TabIndex        =   6
      Top             =   6480
      Width           =   615
   End
   Begin VB.Label Label31 
      Caption         =   $"client_order.frx":43DA
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   11520
      TabIndex        =   5
      Top             =   840
      Width           =   4695
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   8505
      Left            =   0
      Top             =   1680
      Width           =   16335
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   8400
      Left            =   120
      Top             =   1845
      Width           =   16260
   End
   Begin VB.Label Label65 
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
      Left            =   8520
      TabIndex        =   38
      Top             =   2880
      Width           =   105
   End
   Begin VB.Label Label63 
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
      Left            =   8520
      TabIndex        =   37
      Top             =   5160
      Width           =   105
   End
   Begin VB.Label Label60 
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
      Left            =   8520
      TabIndex        =   36
      Top             =   3480
      Width           =   105
   End
   Begin VB.Label Label58 
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
      Left            =   8520
      TabIndex        =   35
      Top             =   4080
      Width           =   105
   End
   Begin VB.Label Label50 
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
      Left            =   8520
      TabIndex        =   34
      Top             =   4680
      Width           =   105
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H80000013&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   16380
   End
End
Attribute VB_Name = "FrmClient2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pic_name As String, pic_ext As String, pic_changed As Boolean
Dim NewPmt As Integer

Private Sub Calendar1_Click()
If Calendar1.Value < Date Then
 MsgBox "You cannot Select the date that " & vbCrLf & "is already passed away." & vbCrLf & "Please select valid date", vbExclamation + vbOKOnly, "Invalid date"
 n2.SetFocus
Else
 n2.Text = Calendar1.Day & "-" & Calendar1.Month & "-" & Calendar1.Year
 Calendar1.Visible = False
 End If
End Sub

Private Sub cmd_Bk_Click()
Unload Me
FrmClient1.Show
End Sub

Private Sub Command1_Click()
Unload Me
Unload FrmClient1
End Sub

Private Sub Command3_Click()
If Trim(Order.Text) = "" Then
 Exit Sub
End If
Set r = c.Execute("select * from clnt_ordr_chln where upper(ORD_NO)='" & Trim(UCase(Order.Text)) & "' ")
If r.EOF = False Then
Set r3 = c.Execute("select * from CLIENT_PMT where upper(ORD_NO)='" & Trim(UCase(Order.Text)) & "' ")
If r3.EOF = False Then
 n16.Text = r3.Fields(3)
 n17.Text = r3.Fields(4)
 n18.Text = r3.Fields(6)
End If
 n19.Text = r.Fields(1)
 Set r1 = New ADODB.Recordset
 Set r1 = c.Execute("select CLNT_NM from  client where clnt_id='" & n19.Text & "' ")
 n1.Text = r1.Fields(0)
 n2.Text = r.Fields(2)
 n3.Text = r.Fields(3)
 n4.Text = r.Fields(4)
 n5.Text = r.Fields(5)
 n6.Text = r.Fields(6)
 n7.Text = r.Fields(7)
 n8.Text = r.Fields(8)
 n9.Text = r.Fields(9)
 n10.Text = r.Fields(10)
 n11.Text = r.Fields(11)
 n12.Text = r.Fields(12)
 n13.Text = r.Fields(13)
 n14.Text = r.Fields(14)
 If IsNull(r.Fields(15)) = False Then
  n15.Text = r.Fields(15)
  Image1.Picture = LoadPicture(n15.Text)
 Else
  Image1.Picture = Nothing
 End If
 Command4.Enabled = False
End If
End Sub

Private Sub Command4_Click() 'Save Button
If n1.Text = "" Or n19.Text = "" Then
 MsgBox "Select Client From Drop Down list..", vbCritical + vbOKOnly, "Not Client"
 n1.SetFocus
 Exit Sub
ElseIf Trim(n2.Text) = "" Or Trim(n3.Text) = "" Or Trim(n4.Text) = "" Then
 MsgBox "All Details in Order Information Field is necessary..", vbCritical + vbOKOnly, "Order Information"
 n2.SetFocus
 Exit Sub
ElseIf Trim(n5.Text) = "" Or Trim(n6.Text) = "" Then
 MsgBox "Class and Subject Field Cannot be Blank..", vbCritical + vbOKOnly, "Exam description"
 n5.SetFocus
 Exit Sub
ElseIf Trim(n7.Text) = "" Then
 MsgBox "Enter Total number of questions in the paper..", vbCritical + vbOKOnly, "Total Questions"
 n7.SetFocus
 Exit Sub
ElseIf Trim(n8.Text) = "" Then
 MsgBox "Enter Total Marks For Question Paper..", vbCritical + vbOKOnly, "Total Marks"
 n8.SetFocus
 Exit Sub
ElseIf Trim(n9.Text) = "" Or Trim(n10.Text) = "" Then
 MsgBox "Enter Total Time for exam in Question Paper..", vbCritical + vbOKOnly, "Invalid Time"
 n9.SetFocus
 Exit Sub
ElseIf Trim(n11.Text) = "" Then
 MsgBox "Name Of Exam cannot be Blank ... Eg:-( Unit Test 1, 2nd Semester Exam...)", vbCritical + vbOKOnly, "Exam Name"
 n11.SetFocus
 Exit Sub
ElseIf Trim(n12.Text) = "" Then
 MsgBox "Enter Marks For each Correct Answer..", vbCritical + vbOKOnly, "Correct Marks"
 n12.SetFocus
 Exit Sub
ElseIf Trim(n13.Text) = "" Then
 MsgBox "Enter Marks For each Wrong Answer.. (Enter 0 if No Negative marks).", vbCritical + vbOKOnly, "Wrong Marks"
 n13.SetFocus
 Exit Sub
ElseIf Trim(n14.Text) = "" Then
 MsgBox "Enter Total Number of Papers for Order..", vbCritical + vbOKOnly, "Total Paper"
 n14.SetFocus
 Exit Sub
ElseIf Trim(n15.Text) = "" Then
 MsgBox "Enter Logo For School or Institute..", vbCritical + vbOKOnly, "Logo"
 Exit Sub
ElseIf Trim(n16.Text) = "" Or Val(n16.Text) = 0 Then
 MsgBox "Enter Total Amount for The Order..", vbCritical + vbOKOnly, "Logo"
 n16.SetFocus
 Exit Sub
ElseIf Trim(n17.Text) = "" Or Val(n17.Text) = 0 Then
 MsgBox "Enter Paid Amount ( at least 100) ..", vbCritical + vbOKOnly, "Logo"
 n17.SetFocus
 Exit Sub
End If
Dim jjjjjj As String
If Val(n18.Text) = 0 Or Trim(n18.Text) = "" Then
 jjjjjj = "Completed"
Else
 jjjjjj = "Pending"
End If
c.Execute ("insert into CLNT_ORDR_CHLN values('" & Order.Text & "','" & n19.Text & "','" & Format(n2.Text, "dd-mmm-yyyy") & "','" & n3.Text & "','" & n4.Text & "','" & n5.Text & "','" & n6.Text & "'," & n7.Text & "," & n8.Text & "," & n9.Text & "," & n10.Text & ",'" & n11.Text & "'," & n12.Text & "," & n13.Text & "," & n14.Text & ",'" & n15.Text & "','" & jjjjjj & "' )")
c.Execute ("insert into  client_pmt values(" & NewPmt & ",'" & Order.Text & "','" & n19.Text & "'," & n16.Text & "," & n17.Text & ",'" & Format(n2.Text, "dd-mmm-yyyy") & "'," & n18.Text & ")")
'Inserting into Account
Dim statement As String
Set r = New ADODB.Recordset
statement = " New Order Recieved From " & n1.Text
Set r = c.Execute("select count(*) from incm")
c.Execute ("insert into incm values (" & r.Fields(0) + 1 & ",'" & n1.Text & "','" & statement & "'," & Val(n17.Text) & ",'" & Format(Date, "dd-mmm-yyyy") & "' )")

MsgBox "New Order Added...", vbInformation + vbOKOnly, "New Order Saved"
Cler
Image1.Picture = Nothing
Adodc1.Refresh
End Sub

Private Sub Command5_Click()
If Frame1.Visible = True Then
 Frame1.Visible = False
Else
 Frame1.Visible = True
End If
End Sub

Private Sub Command6_Click()
Cler
idgen
Order.Locked = True
n1.SetFocus
Command4.Enabled = True
Command3.Visible = False
End Sub

Public Function idgen()
Dim t As Integer
Set r1 = New ADODB.Recordset
sql = "select MAX(to_number(substr(ORD_NO,4,length(ORD_NO))))from CLNT_ORDR_CHLN "
Set r1 = c1.Execute(sql)
If IsNull(r1.Fields(0)) Then
 Order.Text = "ORD001"
 NewPmt = 1
Else
 t = r1.Fields(0)
 NewPmt = t + 1
 If t > 0 And t < 9 Then
  Order.Text = "ORD00" & (t + 1)
 ElseIf t < 99 Then
  Order.Text = "ORD0" & (t + 1)
 Else
  Order.Text = "ORD" & (t + 1)
 End If
End If
End Function

Private Sub txt1_KeyPress(KeyAscii As Integer)
   If (((KeyAscii >= 65) And (KeyAscii <= 90)) Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or (KeyAscii = 8) Or (KeyAscii = 32)) Then
        txt1.SetFocus
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txt2_KeyPress(KeyAscii As Integer)
   If (((KeyAscii >= 65) And (KeyAscii <= 90)) Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or (KeyAscii = 8) Or (KeyAscii = 32) Or ((KeyAscii >= 48) And (KeyAscii <= 57))) Then
        txt2.SetFocus
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txt3_KeyPress(KeyAscii As Integer)
 If (((KeyAscii >= 65) And (KeyAscii <= 90)) Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or (KeyAscii = 8) Or (KeyAscii = 32) Or ((KeyAscii >= 48) And (KeyAscii <= 57))) Then
        txt3.SetFocus
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Command7_Click()
Dim filefilter As String
filefilter = "JPEG Image (*.jpg)|*.jpg|All Files (*.*)|*.*"
On Error Resume Next
cd1.Filter = filefilter
cd1.ShowOpen
If cd1.FileName <> "" Then
pic_ext = Right(cd1.FileName, 3)
If UCase(Trim(pic_ext)) = "GIF" Or UCase(Trim(pic_ext)) = "JPG" Then
pic_changed = True
Else
 MsgBox "Invalid Image !! Please Select JPG Image Only", vbCritical + vbOKOnly, "Image"
 pic_name = ""
 Exit Sub
End If
Image1.Picture = LoadPicture(cd1.FileName)
pic_name = cd1.FileName
n15.Text = pic_name
Else
Exit Sub
End If
End Sub

Private Sub Command8_Click()
n15.Text = App.Path & "\Graphics\Gifs\mmf.gif"
Image1.Picture = LoadPicture(n15.Text)
End Sub

Private Sub Command9_Click()
n15.Text = App.Path & "\Graphics\Gifs\mmf.gif"
Image1.Picture = Nothing
End Sub

Private Sub Form_Load()
conn
Me.Top = 0
Me.Left = 0
Cler
Command3.Visible = True
Calendar1.Value = Format(Date, "DD-MMM-YY")
End Sub

Public Sub Cler()
Order.Text = ""
n1.Clear
Calendar1.Visible = False
n19.Text = ""
NewPmt = 0
n17.Text = ""
n16.Text = ""
n18.Text = ""
n2.Text = ""
n3.Text = ""
n4.Text = ""
n5.Text = ""
n6.Text = ""
n7.Text = ""
n8.Text = ""
n9.Text = ""
n10.Text = ""
n11.Text = ""
n12.Text = ""
n13.Text = ""
n14.Text = ""
n15.Text = App.Path & "\Graphics\Gifs\mmf.gif"
Set r = c.Execute("select * from client")
While r.EOF = False
 n1.AddItem r.Fields(1)
r.MoveNext
Wend
If CurrentClient = "" Then
Else
Command6_Click
Order.Text = CurrentClient
Command3.Enabled = False
End If
Command4.Enabled = False
End Sub

Private Sub n1_Click()
Set r1 = New ADODB.Recordset
Set r1 = c.Execute("select CLNT_ID from client where CLNT_NM ='" & n1.Text & "' ")
If r1.EOF = False Then
 n19.Text = r1.Fields(0)
End If
End Sub

Private Sub n16_Change()
If Val(n16.Text) >= Val(n17.Text) Then
 n18.Text = Val(n16.Text) - Val(n17.Text)
End If
End Sub

Private Sub n16_KeyPress(KeyAscii As Integer)
If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Then
       n16.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   n17.SetFocus
  Else
   KeyAscii = 0
  End If
End Sub

Private Sub n17_Change()
If Trim(n16.Text) = "" Then
KeyAscii = 0
MsgBox "First Enter Total Amount Then Paid Amount...", vbInformation + vbOKOnly, " "
n16.SetFocus
Exit Sub
End If
If Val(n16.Text) >= Val(n17.Text) Then
 n18.Text = Val(n16.Text) - Val(n17.Text)
Else
 MsgBox "Paid Amount cannot be more than Total Amount..", vbCritical + vbOKOnly, "Error"
 n17.Text = Int(Val(n17.Text) / 10)
 n18.Text = Val(n16.Text) - Val(n17.Text)
End If
End Sub

Private Sub n17_KeyPress(KeyAscii As Integer)
If Trim(n16.Text) = "" Then
KeyAscii = 0
MsgBox "First Enter Total Amount Then Paid Amount...", vbInformation + vbOKOnly, " "
n16.SetFocus
Exit Sub
End If
If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Then
       n17.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   n18.SetFocus
  Else
   KeyAscii = 0
  End If
If Val(n16.Text) >= Val(n17.Text) Then
n17.SetFocus
Else
KeyAscii = 0
End If
End Sub

Private Sub n2_GotFocus()
n2.Text = ""
Calendar1.Visible = True
End Sub

Private Sub n2_LostFocus()
Calendar1.Visible = False
End Sub

Private Sub n7_KeyPress(KeyAscii As Integer)
 If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Then
        n7.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   n8.SetFocus
  Else
   KeyAscii = 0
   MsgBox "Only numeric Vlue Allowed..", vbInformation + vbOKOnly, ""
  End If
End Sub

Private Sub n8_KeyPress(KeyAscii As Integer)
 If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Then
        n8.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   n9.SetFocus
  Else
   KeyAscii = 0
   MsgBox "Only numeric Vlue Allowed..", vbInformation + vbOKOnly, ""
  End If
End Sub

Private Sub n9_KeyPress(KeyAscii As Integer)
 If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Then
        n9.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   n10.SetFocus
  Else
   KeyAscii = 0
   MsgBox "Only numeric Vlue Allowed..", vbInformation + vbOKOnly, ""
  End If
End Sub

Private Sub n10_KeyPress(KeyAscii As Integer)
 If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Then
        n10.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   n11.SetFocus
  Else
   KeyAscii = 0
   MsgBox "Only numeric Vlue Allowed..", vbInformation + vbOKOnly, ""
  End If
End Sub

Private Sub n12_KeyPress(KeyAscii As Integer)
 If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Then
        n12.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   n13.SetFocus
  Else
   KeyAscii = 0
   MsgBox "Only numeric Vlue Allowed..", vbInformation + vbOKOnly, ""
  End If
End Sub

Private Sub n13_KeyPress(KeyAscii As Integer)
 If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Then
        n13.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   n14.SetFocus
  Else
   KeyAscii = 0
    MsgBox "Only numeric Vlue Allowed..", vbInformation + vbOKOnly, ""
  End If
End Sub

Private Sub n14_KeyPress(KeyAscii As Integer)
 If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Then
        n14.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
  Else
   KeyAscii = 0
   MsgBox "Only numeric Vlue Allowed..", vbInformation + vbOKOnly, ""
  End If
End Sub
