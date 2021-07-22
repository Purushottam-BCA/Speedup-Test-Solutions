VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form QuestionPPRdashboard 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Question Paper Dashboard"
   ClientHeight    =   10500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20400
   LinkTopic       =   "Form8"
   MDIChild        =   -1  'True
   ScaleHeight     =   10500
   ScaleWidth      =   20400
   ShowInTaskbar   =   0   'False
   Begin MSACAL.Calendar cldr2 
      Height          =   2655
      Left            =   17280
      TabIndex        =   30
      Top             =   1080
      Width           =   3015
      _Version        =   524288
      _ExtentX        =   5318
      _ExtentY        =   4683
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2019
      Month           =   5
      Day             =   26
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   0   'False
      ShowTitle       =   0   'False
      ShowVerticalGrid=   0   'False
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
   Begin MSACAL.Calendar cldr1 
      Height          =   2655
      Left            =   13680
      TabIndex        =   29
      Top             =   1080
      Width           =   3015
      _Version        =   524288
      _ExtentX        =   5318
      _ExtentY        =   4683
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2019
      Month           =   5
      Day             =   26
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   0   'False
      ShowTitle       =   0   'False
      ShowVerticalGrid=   0   'False
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   7800
      Top             =   2280
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      RecordSource    =   "select * from qpaprdash order by sno"
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
   Begin VB.Frame Frame2 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   9255
      Left            =   0
      TabIndex        =   1
      Top             =   1270
      Width           =   3015
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clear All Records"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   360
         MouseIcon       =   "PaperGenerateDashboard.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   31
         ToolTipText     =   "Clear All Previous Records List."
         Top             =   2640
         Width           =   1665
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Main Menu"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   360
         MouseIcon       =   "PaperGenerateDashboard.frx":0152
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   3360
         Width           =   1110
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Display All (Refresh)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   360
         MouseIcon       =   "PaperGenerateDashboard.frx":02A4
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   1920
         Width           =   1995
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client Details"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   360
         MouseIcon       =   "PaperGenerateDashboard.frx":03F6
         MousePointer    =   99  'Custom
         TabIndex        =   4
         ToolTipText     =   "Show All Client Details"
         Top             =   1200
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Create Question Paper"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   360
         MouseIcon       =   "PaperGenerateDashboard.frx":0548
         MousePointer    =   99  'Custom
         TabIndex        =   3
         ToolTipText     =   "Click To Create Quesstion Paper."
         Top             =   480
         Width           =   2205
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Caption         =   "Display All"
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20415
      Begin VB.Frame Fram1 
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   14520
         TabIndex        =   26
         Top             =   720
         Width           =   4335
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1440
            TabIndex        =   27
            Top             =   0
            Width           =   2655
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Enter Here : "
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   0
            TabIndex        =   28
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.Frame Fram2 
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   14520
         TabIndex        =   21
         Top             =   720
         Width           =   4095
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2760
            TabIndex        =   23
            Top             =   0
            Width           =   1335
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   720
            TabIndex        =   22
            Top             =   0
            Width           =   1335
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "To :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2280
            TabIndex        =   25
            Top             =   0
            Width           =   375
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "From :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   0
            TabIndex        =   24
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.CommandButton Command1 
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
         Height          =   330
         Left            =   18960
         MouseIcon       =   "PaperGenerateDashboard.frx":069A
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   720
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   15960
         MouseIcon       =   "PaperGenerateDashboard.frx":07EC
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   135
         Width           =   2655
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Question Paper Generator System"
         BeginProperty Font 
            Name            =   "Lucida Blackletter"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   1320
         TabIndex        =   20
         Top             =   315
         Width           =   8055
      End
      Begin VB.Image Image1 
         Height          =   1180
         Left            =   15
         Picture         =   "PaperGenerateDashboard.frx":093E
         Stretch         =   -1  'True
         Top             =   15
         Width           =   930
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Search By :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   14520
         TabIndex        =   18
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3075
      TabIndex        =   7
      Top             =   1275
      Width           =   17295
      Begin VB.Line Line11 
         BorderColor     =   &H00FFFFFF&
         X1              =   17275
         X2              =   17275
         Y1              =   360
         Y2              =   0
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00FFFFFF&
         X1              =   16140
         X2              =   16140
         Y1              =   360
         Y2              =   0
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00FFFFFF&
         X1              =   15050
         X2              =   15050
         Y1              =   360
         Y2              =   0
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFFFF&
         X1              =   12640
         X2              =   12640
         Y1              =   360
         Y2              =   0
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FFFFFF&
         X1              =   10690
         X2              =   10690
         Y1              =   360
         Y2              =   0
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         X1              =   17280
         X2              =   0
         Y1              =   345
         Y2              =   345
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         X1              =   7690
         X2              =   7690
         Y1              =   360
         Y2              =   0
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   4080
         X2              =   4080
         Y1              =   360
         Y2              =   0
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   2325
         X2              =   2325
         Y1              =   360
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   540
         X2              =   540
         Y1              =   360
         Y2              =   0
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Marks"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   16395
         TabIndex        =   16
         Top             =   30
         Width           =   735
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Class"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   11280
         TabIndex        =   15
         Top             =   30
         Width           =   615
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Generate For"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5200
         TabIndex        =   14
         Top             =   25
         Width           =   1935
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Questions"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   15110
         TabIndex        =   13
         Top             =   25
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Test Type"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   8520
         TabIndex        =   12
         Top             =   25
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   13440
         TabIndex        =   11
         Top             =   30
         Width           =   975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Date "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2520
         TabIndex        =   10
         Top             =   25
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Order Date"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   840
         TabIndex        =   9
         Top             =   25
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "No."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   90
         TabIndex        =   8
         Top             =   25
         Width           =   285
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "PaperGenerateDashboard.frx":6F2C
      Height          =   8865
      Left            =   3075
      TabIndex        =   2
      Top             =   1635
      Width           =   17295
      _ExtentX        =   30506
      _ExtentY        =   15637
      _Version        =   393216
      AllowUpdate     =   0   'False
      ColumnHeaders   =   0   'False
      HeadLines       =   1
      RowHeight       =   22
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
         DataField       =   "SNO"
         Caption         =   "SNO"
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
         DataField       =   "ODATE"
         Caption         =   "ODATE"
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
         DataField       =   "DELIVER"
         Caption         =   "DELIVER"
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
         DataField       =   "ORDRTO"
         Caption         =   "ORDRTO"
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
         DataField       =   "TSTTYPE"
         Caption         =   "TSTTYPE"
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
         DataField       =   "CLASS_NM"
         Caption         =   "CLASS_NM"
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
         DataField       =   "SUB_NM"
         Caption         =   "SUB_NM"
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
         DataField       =   "TOTQUES"
         Caption         =   "TOTQUES"
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
         DataField       =   "TOTMRK"
         Caption         =   "TOTMRK"
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
            ColumnWidth     =   540.284
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1785.26
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1769.953
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   3600
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   1950.236
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   2399.811
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column08 
            Alignment       =   2
            ColumnWidth     =   900.284
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   3600
      X2              =   3600
      Y1              =   1320
      Y2              =   1680
   End
End
Attribute VB_Name = "QuestionPPRdashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cldr1_Click()
If cldr1.Value > Date Then
 MsgBox "Cannot Select Future Date", vbQuestion + vbOKOnly, ""
 cldr1.Value = Date
Else
Text2.Text = cldr1.Day & "-" & cldr1.Month & "-" & cldr1.Year
cldr1.Visible = False
End If
End Sub

Private Sub cldr2_Click()
If cldr2.Value < cldr1.Value Then
 Text3.Text = ""
 cldr2.Visible = True
Else
 Text3.Text = cldr2.Day & "-" & cldr2.Month & "-" & cldr2.Year
 cldr2.Visible = False
End If
End Sub

Private Sub Combo1_Click()
 Text1.Text = ""
 Text2.Text = ""
 Text3.Text = ""
 If Combo1.ListIndex = 1 Then 'Order date
  Fram1.Visible = False
  Fram2.Visible = True
 ElseIf Combo1.ListIndex = 2 Then 'delivery Date
  Fram1.Visible = False
  Fram2.Visible = True
 Else
  Fram1.Visible = True
  Fram2.Visible = False
  Text1.SetFocus
 End If
End Sub

Private Sub Command1_Click()
If Trim(Combo1.Text) = "" Then
 MsgBox "Select Search By Option", vbInformation + vbOKOnly, "Empty Search"
 Combo1.SetFocus
Else  'Now Time For Search
If Combo1.ListIndex = 0 Then 'Serial No
If Trim(Text1.Text) = "" Then
 MsgBox "Enter Data to search", vbExclamation + vbOKOnly, "Enter Data to Search"
 Text1.SetFocus
Exit Sub
End If
Adodc1.RecordSource = "select  * from QPAPRDASH where upper(sno)='" & UCase(Trim(Text1.Text)) & "' order by sno"
Adodc1.Refresh
ElseIf Combo1.ListIndex = 1 Then 'Order Date
If Trim(Text2.Text) = "" Then
 MsgBox "Enter From Date ", vbExclamation + vbOKOnly, "Enter Data to Search"
 Text2.SetFocus
ElseIf Trim(Text3.Text) = "" Then
 MsgBox "Enter Date to search", vbExclamation + vbOKOnly, "Enter Data to Search"
 Text3.SetFocus
Exit Sub
End If
Adodc1.RecordSource = "select * from Qpaprdash where ODATE between '" & Format(Text2.Text, "dd-mmm-yyyy") & "' and '" & Format(Text3.Text, "dd-mmm-yyyy") & "' order by sno "
 Adodc1.Refresh
ElseIf Combo1.ListIndex = 2 Then 'Generate Date
If Trim(Text2.Text) = "" Then
 MsgBox "Enter From Date ", vbExclamation + vbOKOnly, "Empty Date"
 Text2.SetFocus
ElseIf Trim(Text3.Text) = "" Then
 MsgBox "Enter Date to search", vbExclamation + vbOKOnly, "Empty Date"
 Text3.SetFocus
Exit Sub
End If
Adodc1.RecordSource = "select * from Qpaprdash where DELIVER between '" & Format(Text2.Text, "dd-mmm-yyyy") & "' and '" & Format(Text3.Text, "dd-mmm-yyyy") & "' order by sno "
 Adodc1.Refresh
ElseIf Combo1.ListIndex = 3 Then 'Class Wise
If Trim(Text1.Text) = "" Then
 MsgBox "Enter Data to search", vbExclamation + vbOKOnly, "Enter Data to Search"
 Text1.SetFocus
Exit Sub
End If
Adodc1.RecordSource = "select  * from QPAPRDASH where upper(CLASS_NM)like '" & UCase(Trim(Text1.Text)) & "%' order by sno"
Adodc1.Refresh
ElseIf Combo1.ListIndex = 4 Then 'Subject Wise
If Trim(Text1.Text) = "" Then
 MsgBox "Enter Data to search", vbExclamation + vbOKOnly, "Enter Data to Search"
 Text1.SetFocus
Exit Sub
End If
Adodc1.RecordSource = "select  * from QPAPRDASH where upper(SUB_NM) like '" & UCase(Trim(Text1.Text)) & "%' order by sno"
Adodc1.Refresh
ElseIf Combo1.ListIndex = 5 Then 'Test Type
If Trim(Text1.Text) = "" Then
 MsgBox "Enter Data to search", vbExclamation + vbOKOnly, "Enter Data to Search"
 Text1.SetFocus
Exit Sub
End If
Adodc1.RecordSource = "select  * from QPAPRDASH where upper(TSTTYPE) like '" & UCase(Trim(Text1.Text)) & "%' order by sno"
Adodc1.Refresh
Else
 MsgBox "Oops !! Record Doesn't Exist", vbInformation + vbOKOnly, "Invalid Search"
End If
End If
End Sub

Private Sub Form_Load()
conn
Adodc1.RecordSource = "select * from qpaprdash order by sno"
Adodc1.Refresh
Me.Top = 0
Me.Left = 0
Me.Width = 20395
Me.Height = 10495
Combo1.Clear
Combo1.AddItem "Serial No"
Combo1.AddItem "Order Date"
Combo1.AddItem "Generate Date"
Combo1.AddItem "Class"
Combo1.AddItem "Subject"
Combo1.AddItem "Paper"
cldr1.Visible = False
cldr2.Visible = False
Fram2.Visible = False
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.FontBold = False
Label2.FontBold = False
Label3.FontBold = False
Label4.FontBold = False
End Sub

Private Sub Label1_Click()
If EMP_login_reg_no = "" Then
 Unload admin_dash
ElseIf admin_login_reg_no = "" Then
 Unload emp_dash
End If
Me.Hide
QpaprSetup.Show
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.FontBold = True
End Sub

Private Sub Label17_Click()
c.Execute ("delete from qpaprdash")
Adodc1.Refresh
End Sub

Private Sub Label17_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label17.FontBold = True
End Sub

Private Sub Label2_Click()
FrmClient1.Show
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.FontBold = True
End Sub

Private Sub Label3_Click()
Adodc1.RecordSource = "select * from qpaprdash order by sno"
Adodc1.Refresh
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.FontBold = True
End Sub

Private Sub Label4_Click()
Unload Me
If EMP_login_reg_no = "" Then
 admin_dash.Show
ElseIf admin_login_reg_no = "" Then
 emp_dash.Show
End If
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.FontBold = True
End Sub

Private Sub Text2_GotFocus()
Text2.Text = ""
cldr1.Visible = True
End Sub

Private Sub Text2_LostFocus()
cldr1.Visible = False
End Sub

Private Sub Text3_GotFocus()
If Text2.Text <> "" Then
 Text3.Text = ""
 cldr2.Visible = True
 Else
 Text2.SetFocus
End If
End Sub

Private Sub Text3_LostFocus()
cldr2.Visible = False
End Sub
