VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form FrmIncmExpense 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Acounts"
   ClientHeight    =   10695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10695
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc Ado2 
      Height          =   375
      Left            =   4440
      Top             =   2400
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
      Connect         =   "Provider=MSDAORA.1;Password=Sts;User ID=Sts;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;Password=Sts;User ID=Sts;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from exp order by ex_no"
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
   Begin MSACAL.Calendar cldr1 
      Height          =   2655
      Left            =   14400
      TabIndex        =   24
      Top             =   675
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
   Begin MSACAL.Calendar cldr2 
      Height          =   2655
      Left            =   16440
      TabIndex        =   23
      Top             =   675
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
   Begin MSAdodcLib.Adodc ado1 
      Height          =   495
      Left            =   11640
      Top             =   1680
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
      RecordSource    =   "select * from incm order by  S_NO"
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
      Height          =   975
      Left            =   2400
      TabIndex        =   3
      Top             =   0
      Width           =   18135
      Begin VB.Frame Fram1 
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   12120
         TabIndex        =   15
         Top             =   240
         Width           =   4215
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   1440
            TabIndex        =   16
            Top             =   35
            Width           =   2655
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Enter Here : "
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   0
            TabIndex        =   17
            Top             =   70
            Width           =   1335
         End
      End
      Begin VB.Frame Fram2 
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   425
         Left            =   12120
         TabIndex        =   18
         Top             =   260
         Width           =   4095
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
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
            Left            =   720
            TabIndex        =   20
            Top             =   0
            Width           =   1335
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
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
            Left            =   2760
            TabIndex        =   19
            Top             =   0
            Width           =   1335
         End
         Begin VB.Label Label19 
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
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   0
            TabIndex        =   22
            Top             =   0
            Width           =   735
         End
         Begin VB.Label Label20 
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
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2280
            TabIndex        =   21
            Top             =   0
            Width           =   495
         End
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
         Left            =   9000
         MouseIcon       =   "Accounts.frx":0000
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   255
         Width           =   2655
      End
      Begin VB.CommandButton Command4 
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
         Height          =   415
         Left            =   16560
         MouseIcon       =   "Accounts.frx":0152
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Click To Search"
         Top             =   280
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         X1              =   120
         X2              =   4815
         Y1              =   700
         Y2              =   700
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Search By :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7560
         TabIndex        =   14
         Top             =   285
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Income && Expense Details  :-"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   190
         Width           =   4695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   10575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Main menu"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   515
         Left            =   30
         MouseIcon       =   "Accounts.frx":02A4
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Return to Main Menu"
         Top             =   2760
         Width           =   2280
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H008080FF&
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   515
         Left            =   30
         MouseIcon       =   "Accounts.frx":03F6
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Exit "
         Top             =   3345
         Width           =   2280
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
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
         Height          =   515
         Left            =   30
         MouseIcon       =   "Accounts.frx":0548
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Refresh All"
         Top             =   2160
         Width           =   2280
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Expense"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   515
         Left            =   30
         MouseIcon       =   "Accounts.frx":069A
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Display Only Expense Info"
         Top             =   1575
         Width           =   2280
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H8000000E&
         Caption         =   "Income"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   515
         Left            =   30
         MouseIcon       =   "Accounts.frx":07EC
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Display Only Income Info"
         Top             =   975
         Width           =   2280
      End
      Begin VB.Image Image1 
         Height          =   975
         Left            =   360
         Picture         =   "Accounts.frx":093E
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1515
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Accounts.frx":53FE
      Height          =   9030
      Left            =   2475
      TabIndex        =   27
      Top             =   1500
      Width           =   17955
      _ExtentX        =   31671
      _ExtentY        =   15928
      _Version        =   393216
      AllowUpdate     =   -1  'True
      ColumnHeaders   =   0   'False
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "EX_NO"
         Caption         =   "EX_NO"
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
         DataField       =   "EX_WHERE"
         Caption         =   "EX_WHERE"
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
         DataField       =   "EX_REASON"
         Caption         =   "EX_REASON"
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
         DataField       =   "EX_AMT"
         Caption         =   "EX_AMT"
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
         DataField       =   "EX_DATE"
         Caption         =   "EX_DATE"
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
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3314.835
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   9929.765
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1934.929
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1920.189
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Accounts.frx":5411
      Height          =   9030
      Left            =   2475
      TabIndex        =   4
      Top             =   1500
      Width           =   17955
      _ExtentX        =   31671
      _ExtentY        =   15928
      _Version        =   393216
      AllowUpdate     =   -1  'True
      ColumnHeaders   =   0   'False
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "S_NO"
         Caption         =   "S_NO"
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
         DataField       =   "INC_FROM"
         Caption         =   "INC_FROM"
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
         DataField       =   "INC_REASON"
         Caption         =   "INC_REASON"
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
         DataField       =   "INC_AMT"
         Caption         =   "INC_AMT"
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
         DataField       =   "INC_DATE"
         Caption         =   "INC_DATE"
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
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3314.835
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   9929.765
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1890.142
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1964.976
         EndProperty
      EndProperty
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date"
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
      Left            =   18240
      TabIndex        =   9
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Amount"
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
      Left            =   16320
      TabIndex        =   8
      Top             =   1080
      Width           =   1905
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Reason"
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
      Left            =   6375
      TabIndex        =   7
      Top             =   1080
      Width           =   9945
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Income From"
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
      Left            =   3065
      TabIndex        =   6
      Top             =   1080
      Width           =   3345
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No"
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
      Left            =   2470
      TabIndex        =   5
      Top             =   1080
      Width           =   615
   End
End
Attribute VB_Name = "FrmIncmExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim currentOne As Byte
Private Sub Combo1_Click()
 If Combo1.ListIndex = 0 Then
  Fram1.Visible = True
  Fram2.Visible = False
  Text1.SetFocus
 Else
  Fram2.Visible = True
  Fram1.Visible = False
  Text2.SetFocus
 End If
End Sub

Private Sub Command1_Click() 'Refresh
currentOne = 0
cldr1.Visible = False
cldr2.Visible = False
Fram1.Visible = True
Fram2.Visible = False
Command2.Width = 2280
Command3.Width = 2280
Command6.Width = 2280
Command5.Width = 2280
Command1.Width = 2405
Command2.BackColor = &HC0C0C0
Command3.BackColor = &HC0C0C0
Command5.BackColor = &HC0C0C0
Command1.BackColor = &HE0E0E0
Label2.Caption = " Income From"
Command2_Click
End Sub

Private Sub Command2_Click()
DataGrid1.Visible = True
DataGrid2.Visible = False
currentOne = 1
Label2.Caption = " Income From"
ado1.RecordSource = "select * from incm order by S_NO"
ado1.Refresh
Command1.Width = 2280
Command3.Width = 2280
Command6.Width = 2280
Command5.Width = 2280
Command2.Width = 2405
Command1.BackColor = &HC0C0C0
Command3.BackColor = &HC0C0C0
Command5.BackColor = &HC0C0C0
Command2.BackColor = &HE0E0E0
End Sub

Private Sub Command3_Click()
DataGrid1.Visible = False
DataGrid2.Visible = True
currentOne = 0
Label2.Caption = " Paid To"
Ado2.RecordSource = "select * from exp order by EX_no "
Ado2.Refresh
Command1.Width = 2280
Command2.Width = 2280
Command6.Width = 2280
Command5.Width = 2280
Command3.Width = 2405
Command1.BackColor = &HC0C0C0
Command2.BackColor = &HC0C0C0
Command5.BackColor = &HC0C0C0
Command3.BackColor = &HE0E0E0
End Sub

Private Sub Command4_Click()
If Trim(Combo1.Text) = "" Then
 MsgBox "Select Search By Option", vbInformation + vbOKOnly, "Empty Search"
 Combo1.SetFocus
 Exit Sub
End If
'Now Time For Search
If Combo1.ListIndex = 0 Then 'Name
 If Trim(Text1.Text) = "" Then
  MsgBox "Enter Data to search", vbExclamation + vbOKOnly, "Enter Data to Search"
  Text1.SetFocus
 Exit Sub
 End If
 If currentOne = 0 Then 'Expense
 DataGrid2.Visible = True
 DataGrid1.Visible = False
 Ado2.RecordSource = "select * from exp where upper(EX_WHERE) like '" & Trim(UCase(Text1.Text)) & "%' order by EX_NO "
 Ado2.Refresh
Else 'income
 DataGrid2.Visible = False
 DataGrid1.Visible = True
 ado1.RecordSource = "select * from incm where upper(INC_FROM) like '" & Trim(UCase(Text1.Text)) & "%' order by  S_NO "
 ado1.Refresh
End If
ElseIf Combo1.ListIndex = 1 Then 'Order Date
If Trim(Text2.Text) = "" Then
 MsgBox "Enter From Date ", vbExclamation + vbOKOnly, "Enter Data to Search"
 Text2.SetFocus
 Exit Sub
ElseIf Trim(Text3.Text) = "" Then
 MsgBox "Enter Date to search", vbExclamation + vbOKOnly, "Enter Data to Search"
 Text3.SetFocus
Exit Sub
End If
If currentOne = 0 Then 'Expense
 DataGrid2.Visible = True
 DataGrid1.Visible = False
 Ado2.RecordSource = "select * from exp where ex_DATE between '" & Format(Text2.Text, "dd-mmm-yyyy") & "' and '" & Format(Text3.Text, "dd-mmm-yyyy") & "' order by EX_NO "
 Ado2.Refresh
Else 'income
 DataGrid2.Visible = False
 DataGrid1.Visible = True
 ado1.RecordSource = "select * from incm where INC_DATE between '" & Format(Text2.Text, "dd-mmm-yyyy") & "' and '" & Format(Text3.Text, "dd-mmm-yyyy") & "' order by  S_NO "
 ado1.Refresh
End If
End If
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Form_Load()
conn
Me.Top = 0
Me.Left = 0
currentOne = 0
cldr1.Visible = False
cldr2.Visible = False
cldr2.Value = Format(Date, "DD-MMM-YY")
cldr1.Value = Format(Date, "DD-MMM-YY")
Fram1.Visible = True
Fram2.Visible = False
Combo1.Clear
Combo1.AddItem "Customer Name"
Combo1.AddItem "Date"
Command2_Click
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

