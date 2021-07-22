VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form stu_prev_record 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Previous Record"
   ClientHeight    =   10620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20475
   Icon            =   "stu_prev_record.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10620
   ScaleWidth      =   20475
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   720
      Top             =   2640
   End
   Begin MSACAL.Calendar Calendar2 
      Height          =   2895
      Left            =   6480
      TabIndex        =   10
      Top             =   7050
      Width           =   3735
      _Version        =   524288
      _ExtentX        =   6588
      _ExtentY        =   5106
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2019
      Month           =   5
      Day             =   3
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
      ShowTitle       =   -1  'True
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
      Height          =   2895
      Left            =   4200
      TabIndex        =   9
      Top             =   7050
      Width           =   3735
      _Version        =   524288
      _ExtentX        =   6588
      _ExtentY        =   5106
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2019
      Month           =   5
      Day             =   3
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
      ShowTitle       =   -1  'True
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   11760
      Top             =   6840
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
      Connect         =   "Provider=MSDAORA.1;Password=sts;User ID=sts;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;Password=sts;User ID=sts;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"stu_prev_record.frx":0EE2
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
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   765
      Left            =   0
      TabIndex        =   1
      Top             =   9915
      Width           =   20535
      Begin VB.CommandButton BtnSearch 
         Height          =   400
         Left            =   9720
         MouseIcon       =   "stu_prev_record.frx":0FB7
         MousePointer    =   99  'Custom
         Picture         =   "stu_prev_record.frx":1109
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton btnrefresh 
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   385
         Left            =   18840
         MouseIcon       =   "stu_prev_record.frx":1884
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Refresh All Record"
         Top             =   120
         Width           =   1455
      End
      Begin VB.Frame dateframe 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   495
         Left            =   4800
         TabIndex        =   4
         Top             =   80
         Visible         =   0   'False
         Width           =   4575
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd-MM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   8
            ToolTipText     =   "Choose End date"
            Top             =   50
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd-MM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   820
            TabIndex        =   6
            ToolTipText     =   "Choose Start Date"
            Top             =   50
            Width           =   1455
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To :"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   2520
            TabIndex        =   7
            Top             =   75
            Width           =   390
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "From : "
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   120
            TabIndex        =   5
            Top             =   75
            Width           =   720
         End
      End
      Begin VB.ComboBox Combo1 
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
         Left            =   1920
         MouseIcon       =   "stu_prev_record.frx":19D6
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Select Search Criteria"
         Top             =   120
         Width           =   2295
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   350
         Left            =   4800
         TabIndex        =   19
         Top             =   150
         Visible         =   0   'False
         Width           =   4320
         Begin VB.ComboBox Combo2 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1680
            TabIndex        =   21
            Text            =   "Combo2"
            Top             =   0
            Width           =   2535
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Enter here : "
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   1590
         End
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search By :"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   240
         TabIndex        =   2
         Top             =   120
         Width           =   1485
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   " "
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20535
      Begin VB.CommandButton Command1 
         Height          =   400
         Left            =   19080
         MouseIcon       =   "stu_prev_record.frx":1B28
         MousePointer    =   99  'Custom
         Picture         =   "stu_prev_record.frx":1C7A
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   200
         Width           =   1215
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Student's  Previous  Test  Records"
         BeginProperty Font 
            Name            =   "Lucida Blackletter"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   5640
         TabIndex        =   31
         Top             =   120
         Width           =   8055
      End
      Begin VB.Image Image1 
         Height          =   735
         Left            =   0
         Picture         =   "stu_prev_record.frx":23AD
         Stretch         =   -1  'True
         Top             =   0
         Width           =   735
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "stu_prev_record.frx":74DC
      Height          =   8865
      Left            =   0
      TabIndex        =   28
      Top             =   1125
      Width           =   20460
      _ExtentX        =   36089
      _ExtentY        =   15637
      _Version        =   393216
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
      ColumnCount     =   13
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
         DataField       =   "SDATE"
         Caption         =   "SDATE"
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
         DataField       =   "TST_TYP"
         Caption         =   "TST_TYP"
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
         DataField       =   "TOTQS"
         Caption         =   "TOTQS"
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
         DataField       =   "TOTCORR"
         Caption         =   "TOTCORR"
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
         DataField       =   "TOTINCORR"
         Caption         =   "TOTINCORR"
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
         DataField       =   "TOTUNATAMPT"
         Caption         =   "TOTUNATAMPT"
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
         DataField       =   "TOT_MRK"
         Caption         =   "TOT_MRK"
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
         DataField       =   "OBT_MRK"
         Caption         =   "OBT_MRK"
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
         DataField       =   "INITCAP(DIF_LVL)"
         Caption         =   "INITCAP(DIF_LVL)"
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
         DataField       =   "Q_STATUS"
         Caption         =   "Q_STATUS"
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
      BeginProperty Column11 
         DataField       =   "TOTTIME"
         Caption         =   "TOTTIME"
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
      BeginProperty Column12 
         DataField       =   "ELAPSEDTIME"
         Caption         =   "ELAPSEDTIME"
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
            ColumnWidth     =   585.071
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   2204.788
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1170.142
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   824.882
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   1440
         EndProperty
         BeginProperty Column08 
            Alignment       =   2
            ColumnWidth     =   1379.906
         EndProperty
         BeginProperty Column09 
            Alignment       =   2
            ColumnWidth     =   2204.788
         EndProperty
         BeginProperty Column10 
            Alignment       =   2
            ColumnWidth     =   2174.74
         EndProperty
         BeginProperty Column11 
            Alignment       =   2
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column12 
            Alignment       =   2
            ColumnWidth     =   2294.929
         EndProperty
      EndProperty
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total time (Minute)"
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
      Left            =   15480
      TabIndex        =   27
      Top             =   735
      Width           =   2415
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Remaining Time"
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
      Left            =   17925
      TabIndex        =   26
      Top             =   735
      Width           =   2535
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Correct"
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
      Left            =   5400
      TabIndex        =   12
      Top             =   735
      Width           =   855
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total ques"
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
      Left            =   4230
      TabIndex        =   22
      Top             =   735
      Width           =   1215
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date"
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
      Left            =   600
      TabIndex        =   24
      Top             =   735
      Width           =   1455
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Test Type"
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
      Left            =   2040
      TabIndex        =   23
      Top             =   735
      Width           =   2415
   End
   Begin VB.Label Label14 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "              Result"
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
      Left            =   13295
      TabIndex        =   18
      Top             =   735
      Width           =   2175
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Diff.  Level"
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
      Left            =   11100
      TabIndex        =   17
      Top             =   735
      Width           =   2175
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Obt. mark"
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
      Left            =   9755
      TabIndex        =   16
      Top             =   735
      Width           =   1335
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Marks"
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
      Left            =   8280
      TabIndex        =   15
      Top             =   735
      Width           =   1455
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Unattempt"
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
      Left            =   7080
      TabIndex        =   14
      Top             =   735
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Wrong"
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
      Left            =   6240
      TabIndex        =   13
      Top             =   735
      Width           =   855
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No"
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
      Left            =   0
      TabIndex        =   11
      Top             =   735
      Width           =   615
   End
End
Attribute VB_Name = "stu_prev_record"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnrefresh_Click()
Adodc1.RecordSource = "select SNO,SDATE,TST_TYP, TOTQS, TOTCORR, TOTINCORR, TOTUNATAMPT, TOT_MRK,OBT_MRK,  initcap(DIF_LVL),Q_STATUS,TOTTIME, ElapsedTime  from STUD_PREV_REC where rstud_reg_no='" & Stu_login_reg_no & "' order by sno "
Adodc1.Refresh
End Sub

Private Sub btnSearch_Click()
If Combo1.Text = "" Then
 MsgBox "Select search criteria ", vbInformation + vbOKOnly, "Invalid Criteria "
 Exit Sub
ElseIf Combo1.ListIndex <> 1 Then 'not the date
 If Combo2.Text = "" Then
  MsgBox "Cannot Be Blank" & vbCrLf & "Either choose or enter data ", vbInformation + vbOKOnly, "Invalid value "
  Exit Sub
 End If
ElseIf Combo1.ListIndex = 1 Then
 If Text1.Text = "" Then
  MsgBox "Please Enter the Start date", vbInformation + vbOKOnly, "Empty Date"
  Text1.SetFocus
  Exit Sub
 ElseIf Text2.Text = "" Then
  MsgBox "Please Enter the End date", vbInformation + vbOKOnly, "Empty Date"
  Text2.SetFocus
  Exit Sub
 End If
End If
If Combo1.ListIndex = 0 Then 'By Serial No
 Adodc1.RecordSource = "select SNO,SDATE,TST_TYP, TOTQS, TOTCORR, TOTINCORR, TOTUNATAMPT, TOT_MRK,OBT_MRK,  initcap(DIF_LVL),Q_STATUS,TOTTIME, ElapsedTime  from STUD_PREV_REC where rstud_reg_no='" & Stu_login_reg_no & "' and upper(sno)=" & UCase(Val(Trim(Combo2.Text))) & " order by sno  "
 Adodc1.Refresh
ElseIf Combo1.ListIndex = 1 Then 'by Date
 Adodc1.RecordSource = "select SNO,SDATE,TST_TYP, TOTQS, TOTCORR, TOTINCORR, TOTUNATAMPT, TOT_MRK,OBT_MRK,  initcap(DIF_LVL),Q_STATUS,TOTTIME, ElapsedTime  from STUD_PREV_REC where rstud_reg_no='" & Stu_login_reg_no & "' and SDATE between '" & Format(Text1.Text, "dd-mmm-yyyy") & "' and '" & Format(Text2.Text, "dd-mmm-yyyy") & "' order by sno "
 Adodc1.Refresh
ElseIf Combo1.ListIndex = 2 Then 'Diff Level
 Adodc1.RecordSource = "select SNO,SDATE,TST_TYP, TOTQS, TOTCORR, TOTINCORR, TOTUNATAMPT, TOT_MRK,OBT_MRK,  initcap(DIF_LVL),Q_STATUS,TOTTIME, ElapsedTime  from STUD_PREV_REC where rstud_reg_no='" & Stu_login_reg_no & "' and upper(DIF_LVL)='" & UCase(Trim(Combo2.Text)) & "' order by sno "
 Adodc1.Refresh
ElseIf Combo1.ListIndex = 3 Then 'Test Type
 Adodc1.RecordSource = "select SNO,SDATE,TST_TYP, TOTQS, TOTCORR, TOTINCORR, TOTUNATAMPT, TOT_MRK,OBT_MRK,  initcap(DIF_LVL),Q_STATUS ,TOTTIME, ElapsedTime from STUD_PREV_REC where rstud_reg_no='" & Stu_login_reg_no & "' and upper(TST_TYP)='" & UCase(Trim(Combo2.Text)) & "' order by sno "
 Adodc1.Refresh
ElseIf Combo1.ListIndex = 4 Then 'Pass Fail
 Adodc1.RecordSource = "select SNO,SDATE,TST_TYP, TOTQS, TOTCORR, TOTINCORR, TOTUNATAMPT, TOT_MRK,OBT_MRK,  initcap(DIF_LVL),Q_STATUS,TOTTIME, ElapsedTime  from STUD_PREV_REC where rstud_reg_no='" & Stu_login_reg_no & "' and upper(Q_STATUS)='" & UCase(Trim(Combo2.Text)) & "' order by sno "
 Adodc1.Refresh
End If
End Sub

Private Sub Calendar1_Click()
Text1.Text = Calendar1.Day & "-" & Calendar1.Month & "-" & Calendar1.Year
Calendar1.Visible = False
End Sub

Private Sub Calendar2_Click()
If Calendar2.Value < Calendar1.Value Then
 Text2.Text = ""
 Calendar2.Visible = True
Else
 Text2.Text = Calendar2.Day & "-" & Calendar2.Month & "-" & Calendar2.Year
 Calendar2.Visible = False
End If
End Sub

Private Sub Combo1_Click()
BtnSearch.Visible = True
Frame4.Visible = True
Combo2.Clear
If Combo1.ListIndex = 0 Then 'Test No
Combo2.Visible = True
dateframe.Visible = False
Set r = c.Execute("select sno from stud_prev_rec where RSTUD_REG_NO ='" & Stu_login_reg_no & "' order by sno ")
While r.EOF = False
 Combo2.AddItem r.Fields(0)
 r.MoveNext
Wend
ElseIf Combo1.ListIndex = 1 Then 'Date
 Frame4.Visible = False
 Combo2.Visible = False
 dateframe.Visible = True
ElseIf Combo1.ListIndex = 2 Then 'Diff Level
 Combo2.Visible = True
 dateframe.Visible = False
 Combo2.AddItem "Easy"
 Combo2.AddItem "Medium"
 Combo2.AddItem "Hard"
 Combo2.AddItem "Mix (All)"

ElseIf Combo1.ListIndex = 3 Then 'Test Type
 Combo2.Visible = True
 dateframe.Visible = False
 Combo2.AddItem "Topic Wise Test"
 Combo2.AddItem "Subject Wise Test"
 Combo2.AddItem "Full Length Test"

ElseIf Combo1.ListIndex = 4 Then 'Result pass/Fail
 Combo2.Visible = True
 dateframe.Visible = False
 Combo2.AddItem "Pass"
 Combo2.AddItem "Fail"
End If
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Top = 0 '1080
Me.Left = 0
BtnSearch.Visible = False
Timer1.Enabled = True
conn
Combo1.Clear
Calendar1.Value = Format(Date, "DD-MMM-YY")
Calendar2.Value = Format(Date, "DD-MMM-YY")
Text1.Text = ""
Combo1.AddItem "Test No"
Combo1.AddItem "Date"
Combo1.AddItem "Difficulty Leval"
Combo1.AddItem "Test Type"
Combo1.AddItem "Result Wise"
Adodc1.Refresh
dateframe.Visible = False
Calendar1.Visible = False
Calendar2.Visible = False
btnrefresh_Click
End Sub

Private Sub Text1_GotFocus()
Text1.Text = ""
Calendar1.Visible = True
End Sub

Private Sub Text1_LostFocus()
Calendar1.Visible = False
End Sub

Private Sub Text2_GotFocus()
If Text1.Text <> "" Then
 Text2.Text = ""
 Calendar2.Visible = True
 Else
 Text1.SetFocus
End If
End Sub

Private Sub Text2_LostFocus()
Calendar2.Visible = False
End Sub
