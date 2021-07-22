VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BF38D12B-22A9-4B10-B26E-019F2B5F9C22}#1.0#0"; "AniGif.ocx"
Begin VB.Form FrmImportQues 
   Caption         =   "Import Questions"
   ClientHeight    =   8145
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8145
   ScaleWidth      =   15000
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Height          =   5055
      Left            =   720
      MouseIcon       =   "ImportQuestions.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "ImportQuestions.frx":0152
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5160
      Visible         =   0   'False
      Width           =   19215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Import Questions"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   10095
      Begin VB.CommandButton Command9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "[1/3] Next >>"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   415
         Left            =   7815
         MouseIcon       =   "ImportQuestions.frx":1CC40
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Click to Procceed Step 2"
         Top             =   5160
         Width           =   1695
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "[2/3] Next >>"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   415
         Left            =   7815
         MouseIcon       =   "ImportQuestions.frx":1CD92
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Click to Procceed to Step 3"
         Top             =   5160
         Width           =   1695
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "<< Close"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   415
         Left            =   6000
         MouseIcon       =   "ImportQuestions.frx":1CEE4
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Return Back."
         Top             =   5160
         Width           =   1695
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "[3/3] Next >>"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   415
         Left            =   7815
         MouseIcon       =   "ImportQuestions.frx":1D036
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Click to Procceed Final Step"
         Top             =   5160
         Width           =   1695
      End
      Begin VB.Frame Frame4 
         Caption         =   "Excel File Format"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   600
         TabIndex        =   20
         Top             =   600
         Width           =   8895
         Begin VB.Line Line1 
            BorderColor     =   &H000000FF&
            X1              =   6800
            X2              =   8640
            Y1              =   1705
            Y2              =   1705
         End
         Begin VB.Label Label7 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "See Proper Format"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   2
            Left            =   6840
            MouseIcon       =   "ImportQuestions.frx":1D188
            MousePointer    =   99  'Custom
            TabIndex        =   22
            ToolTipText     =   "Click To See Excel File Format."
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   $"ImportQuestions.frx":1D2DA
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   990
            Left            =   900
            TabIndex        =   21
            Top             =   480
            Width           =   7200
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Select  File"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   600
         TabIndex        =   17
         Top             =   2880
         Width           =   8895
         Begin VB.CommandButton Command2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Browse"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   415
            Left            =   7080
            MouseIcon       =   "ImportQuestions.frx":1D3A4
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Select File."
            Top             =   920
            Width           =   1335
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select The Excel File (*.CSV , *.XLS ,*.TXT ) From Disk."
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
            Left            =   480
            TabIndex        =   19
            Top             =   480
            Width           =   4890
         End
         Begin VB.Label Label 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   480
            TabIndex        =   18
            Top             =   960
            Width           =   6135
         End
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Next >>"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   840
         MouseIcon       =   "ImportQuestions.frx":1D4F6
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Click to Procceed Next"
         Top             =   1560
         Width           =   105
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000006&
         BorderWidth     =   2
         Index           =   4
         X1              =   0
         X2              =   120
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000006&
         BorderWidth     =   2
         Index           =   3
         X1              =   1850
         X2              =   10145
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   720
         TabIndex        =   30
         Top             =   1080
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "ImportQuestions.frx":1D648
      Height          =   2895
      Left            =   960
      TabIndex        =   0
      Top             =   7680
      Visible         =   0   'False
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   5106
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   16
      BeginProperty Column00 
         DataField       =   "HQ_ID"
         Caption         =   "HQ_ID"
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
         DataField       =   "HQ_NO"
         Caption         =   "HQ_NO"
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
         DataField       =   "HC_ID"
         Caption         =   "HC_ID"
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
         DataField       =   "HSUB_ID"
         Caption         =   "HSUB_ID"
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
         DataField       =   "HTP_ID"
         Caption         =   "HTP_ID"
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
         DataField       =   "HQ_TYP_ID"
         Caption         =   "HQ_TYP_ID"
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
         DataField       =   "HQ_TXT"
         Caption         =   "HQ_TXT"
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
         DataField       =   "HOPT1"
         Caption         =   "HOPT1"
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
         DataField       =   "HOPT2"
         Caption         =   "HOPT2"
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
         DataField       =   "HOPT3"
         Caption         =   "HOPT3"
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
         DataField       =   "HOPT4"
         Caption         =   "HOPT4"
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
         DataField       =   "HANS_TXT"
         Caption         =   "HANS_TXT"
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
         DataField       =   "HANS_NO"
         Caption         =   "HANS_NO"
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
      BeginProperty Column13 
         DataField       =   "HQ_DIF_LVL"
         Caption         =   "HQ_DIF_LVL"
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
      BeginProperty Column14 
         DataField       =   "HQ_EXPLN"
         Caption         =   "HQ_EXPLN"
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
      BeginProperty Column15 
         DataField       =   "HQ_PIC"
         Caption         =   "HQ_PIC"
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
         RecordSelectors =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdb 
      Left            =   10680
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   9960
      Top             =   1920
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
      RecordSource    =   "select * from holdimport"
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
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   195
      Left            =   0
      TabIndex        =   32
      Top             =   -120
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Frame Frame2 
      Caption         =   "Import Questions"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   10095
      Begin VB.ComboBox Combo5 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   390
         Left            =   6960
         MouseIcon       =   "ImportQuestions.frx":1D65D
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2400
         Width           =   2415
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   415
         Left            =   7680
         MouseIcon       =   "ImportQuestions.frx":1D7AF
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Return Back"
         Top             =   4800
         Width           =   1695
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Import"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   415
         Left            =   5865
         MouseIcon       =   "ImportQuestions.frx":1D901
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Import Questions."
         Top             =   4800
         Width           =   1695
      End
      Begin VB.ComboBox Combo4 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   390
         Left            =   6960
         MouseIcon       =   "ImportQuestions.frx":1DA53
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1560
         Width           =   2415
      End
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   390
         Left            =   2265
         MouseIcon       =   "ImportQuestions.frx":1DBA5
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   3480
         Width           =   2385
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   390
         Left            =   2265
         MouseIcon       =   "ImportQuestions.frx":1DCF7
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2520
         Width           =   2415
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   390
         Left            =   2265
         MouseIcon       =   "ImportQuestions.frx":1DE49
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Select Following Properties To Import Questions SuccessFully. (All Options Are Mandatory)."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   16
         Top             =   600
         Width           =   8655
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Question Type  :"
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
         Left            =   5310
         TabIndex        =   15
         Top             =   1560
         Width           =   1515
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Difficulty Level  :"
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
         Left            =   5310
         TabIndex        =   13
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Select  Topic      :"
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
         Index           =   0
         Left            =   480
         TabIndex        =   12
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Select  Subject  :"
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
         Left            =   480
         TabIndex        =   11
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Select  Course   :"
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
         Left            =   480
         TabIndex        =   10
         Top             =   1575
         Width           =   1695
      End
   End
   Begin VB.TextBox Text1 
      Height          =   1365
      Left            =   6840
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000006&
      BorderWidth     =   2
      Index           =   2
      X1              =   435
      X2              =   10580
      Y1              =   6505
      Y2              =   6505
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000006&
      BorderWidth     =   2
      Index           =   1
      X1              =   450
      X2              =   450
      Y1              =   600
      Y2              =   6480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000006&
      BorderWidth     =   2
      Index           =   0
      X1              =   10575
      X2              =   10590
      Y1              =   600
      Y2              =   6480
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Step By Step Instructions to Import Questions."
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
      Height          =   315
      Left            =   12120
      TabIndex        =   34
      Top             =   780
      Width           =   4845
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   5295
      Left            =   11400
      Top             =   1200
      Width           =   8655
   End
   Begin Project1.PictureG PictureG1 
      Height          =   450
      Left            =   11400
      Top             =   740
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   794
      GIF             =   "ImportQuestions.frx":1DF9B
      Stretch         =   2
   End
   Begin VB.Label Label11 
      Caption         =   $"ImportQuestions.frx":23A4D
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   5055
      Left            =   11520
      TabIndex        =   33
      Top             =   1320
      Width           =   8415
   End
   Begin VB.Label m3 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2000
      TabIndex        =   28
      Top             =   3840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label m4 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2000
      TabIndex        =   27
      Top             =   4200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label m2 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2000
      TabIndex        =   26
      Top             =   3480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label m1 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2000
      TabIndex        =   25
      Top             =   3120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   525
      Left            =   11400
      Top             =   690
      Width           =   8655
   End
End
Attribute VB_Name = "FrmImportQues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public file_name As String, file_ext As String
Public Sub Refresh2()
Command4_Click
Command9.Visible = True
Command10.Visible = False
Command11.Visible = False
Command7.Visible = False
Frame1.Visible = True
Frame2.Visible = False
Label.Caption = ""
Label1.Caption = ""
m1.Caption = ""
m2.Caption = ""
m3.Caption = ""
m4.Caption = ""
Text1.Text = ""
Combo1.Clear
If rs_course.EOF = False Then
rs_course.MoveFirst
While rs_course.EOF = False
 Combo1.AddItem rs_course(0)
 rs_course.MoveNext
Wend
End If
Combo4.Clear
If rs_qtyp.EOF = False Then
rs_qtyp.MoveFirst
While rs_qtyp.EOF = False
 Combo4.AddItem rs_qtyp.Fields(1)
 rs_qtyp.MoveNext
Wend
End If
Combo2.Clear
Combo3.Clear
Combo5.Clear
Combo5.AddItem "EASY"
Combo5.AddItem "MEDIUM"
Combo5.AddItem "HARD"

End Sub

Private Sub Combo1_Click()
Combo3.Clear
Combo2.Clear
Set r1 = New ADODB.Recordset
Set r1 = c1.Execute("select sub_nm from sub where c_id =(select c_id from course where upper(c_nm)='" & UCase(Combo1.Text) & "') ")
If IsNull(r1.Fields(0)) = False Then
While r1.EOF = False
 Combo2.AddItem r1.Fields(0)
 r1.MoveNext
Wend
End If
End Sub

Private Sub combo2_Click()
Combo3.Clear
Set r1 = New ADODB.Recordset
Set r1 = c1.Execute("select TP_NM from topic where sub_id =(select sub_id from sub where sub_nm='" & Combo2.Text & "' and c_id=(select c_id from course where c_nm='" & Combo1.Text & "'))and c_id =(select c_id from course where c_nm='" & Combo1.Text & "') ")
If IsNull(r1.Fields(0)) = False Then
 While r1.EOF = False
  Combo3.AddItem r1.Fields(0)
  r1.MoveNext
 Wend
End If
End Sub

Private Sub Command1_Click()
Command1.Visible = False
End Sub

Private Sub Command10_Click()
Copy_Rename
Command10.Visible = False
Command11.Visible = True
End Sub

Private Sub Command11_Click()
Shell "cmd.exe /c " & "sqlldr userid=sts/sts control='" & App.Path & "\Ocx\ControlFile.ctl' skip=1"
Adodc1.Refresh
Command11.Visible = False
Command7.Visible = True
MsgBox "Click 1 more Times...", vbInformation + vbOKOnly, "70% Completed"
End Sub

Private Sub Command2_Click() 'Browse
'On Error Resume Next
' Shell "cmd.exe /c del " & App.Path & "\ocx\mm.csv" 'Deleting if exist
 Dim filefilter As String
 filefilter = "Excel 97-2003 Workbook(*.xls)|*.xls|CSV Files (*.csv)|*.csv|All Files(*.*)|*.*"
 cdb.Filter = filefilter
 cdb.ShowOpen
 If cdb.FileName <> "" And cdb.FileTitle <> "" Then
  file_name = cdb.FileName
  file_ext = Right(cdb.FileTitle, 4)
  If UCase(file_ext) = UCase(".csv") Or UCase(file_ext) = UCase(".txt") Then
   Label1.Caption = file_name
   Label.Caption = file_name
   
   Shell "cmd.exe /c copy /y " & file_name & " " & App.Path & "\ocx\mm.csv"

   MsgBox "CSV File Loaded. Click on Next Button..", vbInformation + vbOKOnly, ""
   Exit Sub
  ElseIf UCase(Trim(file_ext)) = UCase(".xls") Then
   Label.Caption = file_name
   Label1.Caption = App.Path & "\ocx\mm.csv"
   Shell "cmd.exe /c " & App.Path & "\ocx\XlsToCsv.vbs " & file_name & " " & App.Path & "\ocx\mm.csv"
   MsgBox "XLS File Loaded. Click on Next Button..", vbInformation + vbOKOnly, ""
   Exit Sub
  Else
   MsgBox "Choose Correct File.(Only *.Xls or *.Csv) Files..", vbCritical + vbOKOnly, ""
   file_name = ""
   file_ext = ""
   Exit Sub
  End If
 Else
End If
End Sub

Private Sub Command3_Click()
Refresh2
End Sub

Public Function PasteInControlFile() 'Finally writting into control file
 Open App.Path + "\ocx\ControlFile.ctl" For Output As #1
 Print #1, Text1.Text
 Close #1
End Function

Public Function Copy_Rename() 'Copy and renaming
Dim s1 As String, s2 As String
s1 = Label1.Caption
s2 = App.Path & "\ocx\mm.csv"
Shell "cmd.exe /c copy /b/v/y " & s1 & " " & s2
End Function

Private Sub Command4_Click()
' Shell "cmd.exe /c taskkill /IM excel.exe" 'Close EXcel If Open
' Shell "cmd.exe /c taskkill /IM cmd.exe" 'Close EXcel If Open
End Sub

Private Sub Command9_Click()
If Label.Caption = "" Or Trim(Label1.Caption) = "" Then
 MsgBox "Select File First !! Click on Browse Button..", vbExclamation + vbOKOnly, "No File Selected"
 Exit Sub
End If
c.Execute ("delete from holdImport")
Adodc1.Refresh
Text1.Text = "options(skip=1)" & vbCrLf & "LOAD DATA " & vbCrLf & "INFILE '" & App.Path & "\Ocx\mm.csv'" & vbCrLf & "TRUNCATE" & vbCrLf & "INTO TABLE Holdimport" & vbCrLf & "fields terminated by "",""" & vbCrLf & "(" & vbCrLf & " hq_no," & vbCrLf & " hq_txt," & vbCrLf & " hopt1," & vbCrLf & " hopt2," & vbCrLf & " hopt3," & vbCrLf & " hopt4," & vbCrLf & " hans_txt," & vbCrLf & " hans_no," & vbCrLf & " hq_expln" & vbCrLf & ")"
PasteInControlFile
Command9.Visible = False
Command10_Click
MsgBox "Click 2 more Times...", vbInformation + vbOKOnly, "33% Completed"
End Sub

Private Sub Command6_Click() 'Actual Import
If Combo1.Text = "" Then
     MsgBox "Select Course First ", vbCritical + vbOKOnly, "Course Missing"
     Combo1.SetFocus
     Exit Sub
End If
If Combo2.Text = "" Then
     MsgBox "Select Subject Correspondent to the Course", vbCritical + vbOKOnly, "Subject Missing"
     Combo2.SetFocus
     Exit Sub
 End If
If Combo3.Text = "" Then
     MsgBox "Select Topic Correspondent to the Subject of the course", vbCritical + vbOKOnly, "Topic Missing"
     Combo3.SetFocus
     Exit Sub
 End If
If Combo4.Text = "" Then
     MsgBox "Select Question Type", vbCritical + vbOKOnly, "Question type Missing"
     Combo4.SetFocus
     Exit Sub
 End If
If Combo5.Text = "" Then
      MsgBox "Choose Difficulty Level of Question", vbCritical + vbOKOnly, "Select Difficulty"
      Combo5.SetFocus
      Exit Sub
  End If
Dim rt As Integer
updtDA 'Find ID of Course , Subject , Topic Etc
   rt = InputBox("Some Questions May Be Already Available in Database." & vbCrLf & vbCrLf & "1. Keep Old Questions and Import New Questions" & vbCrLf & "2. Remove All Old Questions and Import New Questions" & vbCrLf & "3. Cancel " & vbCrLf & vbCrLf & "Enter Your Choice :-   ")
 If rt = 1 Then
 ElseIf rt = 2 Then
    c.Execute ("delete from quesms where tp_id='" & m3.Caption & "' and sub_id='" & m2.Caption & "' and c_id='" & m1.Caption & "' ")
 Else
    Exit Sub
 End If
Set r = New ADODB.Recordset
Set r = c.Execute("select count(*) from holdimport")
 If (r.Fields(0) = 0) Then
  MsgBox "Some Error Occured. Try Again. click on Close Button To try Again..", vbCritical + vbOKOnly, "Import Error"
 Exit Sub
 End If
SettingForUpdating 'Most important update setting
c.Execute ("update holdimport set hq_id= QNOGENERATOR1.nextval , hc_id='" & m1.Caption & "', hsub_id='" & m2.Caption & "', htp_id='" & m3.Caption & "', hq_typ_id='" & m4.Caption & "', hq_dif_lvl='" & Combo5.Text & "', HQ_PIC=NULL ")
Adodc1.Refresh
Set r = New ADODB.Recordset
Set r = c.Execute("select hq_id from holdimport")
Dim divya As String
While r.EOF = False
       divya = "MS" & Format(r.Fields(0), "0000")
        c.Execute ("update holdimport set hq_id = '" & divya & "' where hq_id='" & r.Fields(0) & "' ")
     r.MoveNext
 Wend
 Adodc1.Refresh
  Shell "cmd.exe /c del " & App.Path & "\ocx\mm.bad"
  Shell "cmd.exe /c wmic process where name='excel.exe' delete" ' Closing Excel
 c.Execute ("insert into quesms select * from holdimport")
 Set r = New ADODB.Recordset
 Set r = c.Execute("select count(*) from holdimport ")
  MsgBox r.Fields(0) & " Questions Imported SuccessFullly..", vbInformation + vbOKOnly, "Import Questions"
 c.Execute ("delete from holdimport")
Command4_Click
Refresh2
End Sub

Private Sub Command7_Click()
On Error GoTo E:
Set r = c.Execute("select count(*) from holdimport")
 If (r.Fields(0) = 0) Then
   MsgBox "Some Internal Error Occured. Try Again...", vbCritical + vbOKOnly, "Import Error"
   Command9.Visible = True
   Command10.Visible = False
   Command11.Visible = False
   Command7.Visible = False
  Exit Sub
End If
 'Shell "cmd.exe /c taskkill /IM excel.exe" 'Close EXcel If Open
 Shell "cmd.exe /c taskkill /IM cmd.exe" 'Close CMD If Open
 'Shell "cmd.exe /c del " & App.Path & "\ocx\mm.csv"
Frame1.Visible = False
Frame2.Visible = True
Adodc1.Refresh
MsgBox "Now Select Questions Property before Importing Them...", vbInformation + vbOKOnly, "99 % Completed"
Exit Sub
E:
  MsgBox "Some Error Occured,Either Imported File is not in proper format or some technical issue.. ", vbCritical + vbOKOnly, "Import Error"
End Sub

Private Sub Command8_Click()
Unload Me
End Sub

Private Sub Form_Load()
conn
Refresh2
End Sub

Public Sub updtDA()
  Set r = New ADODB.Recordset
  Set r1 = New ADODB.Recordset
  Set r2 = New ADODB.Recordset
  Set r3 = New ADODB.Recordset
 Set r = c.Execute("SELECT c_id FROM COURSE WHERE C_NM='" & Combo1.Text & "' ")
If r.EOF = False Then
  m1.Caption = r.Fields(0)
End If
Set r1 = c.Execute("SELECT sub_id from sub where sub_nm='" & Combo2.Text & "' AND C_ID='" & m1.Caption & "' ")
If r1.EOF = False Then
   m2.Caption = r1.Fields(0)
End If
Set r2 = c.Execute("SELECT tp_id from TOPIC where TP_NM ='" & Combo3.Text & "' AND SUB_ID='" & m2.Caption & "' AND C_ID='" & m1.Caption & "' ")
If r2.EOF = False Then
   m3.Caption = r2.Fields(0)
End If
Set r3 = c.Execute("SELECT * FROM Q_TYP WHERE  Q_TYP_NM='" & Combo4.Text & "' ")
If r3.EOF = False Then
  m4.Caption = r3.Fields(0)
End If
End Sub

Public Function FindStartingQID() As Integer
Dim sd As Integer
Set r = New ADODB.Recordset
Set r = c.Execute("select MAX(to_number(substr(q_id,3,length(q_id))))from quesms")
 If IsNull(r.Fields(0)) = True Then
  sd = 1
 Else
  sd = r.Fields(0) + 1
 End If
FindStartingQID = sd
End Function

Public Function SettingForUpdating()
Dim ss As Integer
ss = FindStartingQID()
Set r = New ADODB.Recordset
Set r = c.Execute("select count(*)from user_sequences where sequence_name='QNOGENERATOR1' ")
  If (r.Fields(0) > 0) Then
  c.Execute ("drop sequence QNOGENERATOR1")
  c.Execute (" create sequence QNOGENERATOR1 increment by 1 start with " & ss & " cache 100")
 Else
  c.Execute (" create sequence QNOGENERATOR1 increment by 1 start with " & ss & " cache 100")
 End If
End Function

Private Sub Form_Unload(cancel As Integer)
c.Execute ("delete from holdimport")
Shell "cmd.exe /c del " & App.Path & "\ocx\on.LST" 'Deleting if exist
Shell "cmd.exe /c del " & App.Path & "\ocx\mm.bad" 'Deleting if exist
'Shell "cmd.exe /c del " & App.Path & "\ocx\mm.csv"
'Shell "cmd.exe /c del C:\STS\on.LST"
Command4_Click
End Sub

Private Sub Label7_Click(Index As Integer)
If Command1.Visible = True Then
 Command1.Visible = False
Else
 Command1.Visible = True
 End If
End Sub
