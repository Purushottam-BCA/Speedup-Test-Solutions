VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BF38D12B-22A9-4B10-B26E-019F2B5F9C22}#1.0#0"; "AniGif.ocx"
Begin VB.Form FrmExportQues 
   Caption         =   "Export"
   ClientHeight    =   8265
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13470
   BeginProperty Font 
      Name            =   "Cambria"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "EXportQuestion.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8265
   ScaleWidth      =   13470
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   11280
      Top             =   8280
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      RecordSource    =   "select Q_no,Q_txt,Opt1,opt2,opt3,opt4,Ans_txt,Ans_no,Q_expln from quesms"
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
      Bindings        =   "EXportQuestion.frx":0ECA
      Height          =   3135
      Left            =   480
      TabIndex        =   4
      Top             =   7080
      Width           =   19455
      _ExtentX        =   34316
      _ExtentY        =   5530
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   22
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
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
         DataField       =   "Q_NO"
         Caption         =   "Q No"
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
         DataField       =   "Q_TXT"
         Caption         =   "Questions"
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
         DataField       =   "OPT1"
         Caption         =   "Option 1"
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
         DataField       =   "OPT2"
         Caption         =   "Option 2"
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
         DataField       =   "OPT3"
         Caption         =   "Option 3"
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
         DataField       =   "OPT4"
         Caption         =   "Option 4"
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
         DataField       =   "ANS_TXT"
         Caption         =   "Answer"
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
         DataField       =   "ANS_NO"
         Caption         =   "Ans No"
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
         DataField       =   "Q_EXPLN"
         Caption         =   "Explanation"
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
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   5084.788
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1590.236
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1425.26
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1695.118
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   4860.284
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Preparing Questions"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   10215
      Begin VB.CommandButton Option3 
         Caption         =   "Topic Wise Question"
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
         Left            =   6840
         MouseIcon       =   "EXportQuestion.frx":0EDF
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Select Topic Wise Questions"
         Top             =   1320
         Width           =   3255
      End
      Begin VB.CommandButton Option2 
         Caption         =   "Subject Wise Question"
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
         Left            =   3480
         MouseIcon       =   "EXportQuestion.frx":1031
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Select Subject Wise Questions"
         Top             =   1320
         Width           =   3255
      End
      Begin VB.CommandButton Option1 
         Caption         =   "Course Wise Question"
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
         Left            =   120
         MouseIcon       =   "EXportQuestion.frx":1183
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Select Course Wise Questions"
         Top             =   1320
         Width           =   3255
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   6720
         Top             =   5160
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Export"
         DisabledPicture =   "EXportQuestion.frx":12D5
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   9090
         MouseIcon       =   "EXportQuestion.frx":2277
         MousePointer    =   99  'Custom
         Picture         =   "EXportQuestion.frx":23C9
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Click To Export"
         Top             =   5040
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0FF&
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
         Left            =   4200
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   120
         TabIndex        =   10
         Top             =   1920
         Width           =   10000
         Begin VB.CommandButton Command3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Close"
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
            Left            =   8145
            MouseIcon       =   "EXportQuestion.frx":336B
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Return Back"
            Top             =   2160
            Width           =   1695
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Check"
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
            Left            =   6345
            MouseIcon       =   "EXportQuestion.frx":34BD
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   2160
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
            Left            =   2265
            MouseIcon       =   "EXportQuestion.frx":360F
            MousePointer    =   99  'Custom
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1200
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
            Left            =   7065
            MouseIcon       =   "EXportQuestion.frx":3761
            MousePointer    =   99  'Custom
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1200
            Visible         =   0   'False
            Width           =   2990
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
            Left            =   7065
            MouseIcon       =   "EXportQuestion.frx":38B3
            MousePointer    =   99  'Custom
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   480
            Visible         =   0   'False
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
            MouseIcon       =   "EXportQuestion.frx":3A05
            MousePointer    =   99  'Custom
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   480
            Width           =   2415
         End
         Begin Project1.PictureG PictureG2 
            Height          =   300
            Left            =   3480
            Top             =   2220
            Visible         =   0   'False
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   529
            GIF             =   "EXportQuestion.frx":3B57
            Stretch         =   2
         End
         Begin VB.Label Label5 
            BackColor       =   &H8000000D&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   8760
            TabIndex        =   24
            Top             =   1920
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label m3 
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8745
            TabIndex        =   23
            Top             =   2640
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label m2 
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8745
            TabIndex        =   22
            Top             =   2400
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label m1 
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8745
            TabIndex        =   21
            Top             =   2160
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Level  :"
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
            Left            =   1395
            TabIndex        =   20
            Top             =   1230
            Width           =   675
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Select  Topic  :"
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
            Left            =   5475
            TabIndex        =   19
            Top             =   1200
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
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
            Left            =   5265
            TabIndex        =   18
            Top             =   480
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Select  Course  :"
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
            TabIndex        =   17
            Top             =   495
            Width           =   1695
         End
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Total No. Of Questions (For Export )   :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   5640
         Width           =   3615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
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
         Left            =   4200
         TabIndex        =   6
         Top             =   5025
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Available Questions  ( In DataBase )  : "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   5040
         Width           =   3975
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   $"EXportQuestion.frx":53B5
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   990
         Left            =   405
         TabIndex        =   3
         Top             =   480
         Width           =   9345
      End
   End
   Begin MSComDlg.CommonDialog cdb 
      Left            =   0
      Top             =   -240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      Height          =   2295
      Left            =   2520
      MultiLine       =   -1  'True
      TabIndex        =   25
      Top             =   4200
      Width           =   7695
   End
   Begin VB.TextBox Text3 
      Height          =   1215
      Left            =   2520
      TabIndex        =   26
      Top             =   2880
      Width           =   7695
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All Questions ( In Database )"
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
      Left            =   645
      TabIndex        =   27
      Top             =   6645
      Width           =   2805
   End
   Begin VB.Label Label11 
      Caption         =   $"EXportQuestion.frx":54BA
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
      Height          =   5415
      Left            =   11280
      TabIndex        =   2
      Top             =   960
      Width           =   8415
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Step By Step Instructions to Export Questions."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   12000
      TabIndex        =   1
      Top             =   420
      Width           =   4800
   End
   Begin Project1.PictureG PictureG1 
      Height          =   450
      Left            =   11160
      Top             =   405
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   794
      GIF             =   "EXportQuestion.frx":5915
      Stretch         =   2
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   5655
      Left            =   11160
      Top             =   840
      Width           =   8655
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   495
      Left            =   11160
      Top             =   360
      Width           =   8655
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   480
      Top             =   6600
      Width           =   19455
   End
End
Attribute VB_Name = "FrmExportQues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ak As Integer
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

Private Sub Command1_Click() 'EXPORT Button
c.Execute ("delete from holdimport")
If Val(Trim(Text1.Text)) = 0 Or Trim(Text1.Text) = "" Then
 MsgBox "Enter Total Number of Questions to be Exported", vbQuestion + vbOKOnly, "Invalid Number"
 Text1.SetFocus
 Exit Sub
Else
End If
If Label5.Caption = 1 Then
 If Combo4.ListIndex = 3 Then
  c.Execute ("insert into holdimport select * from quesms where c_id='" & m1.Caption & "' ")
 ElseIf Combo4.ListIndex >= 0 Then
  c.Execute ("insert into holdimport select * from quesms where c_id='" & m1.Caption & "' and q_dif_lvl='" & Combo4.Text & "' ")
 End If
ElseIf Label5.Caption = 2 Then 'Subject Wise
 If Combo4.ListIndex = 3 Then
  c.Execute ("insert into holdimport select * from quesms where c_id='" & m1.Caption & "' and sub_id='" & m2.Caption & "' ")
 ElseIf Combo4.ListIndex >= 0 Then
  c.Execute ("insert into holdimport select * from quesms where c_id='" & m1.Caption & "' and q_dif_lvl='" & Combo4.Text & "' and sub_id='" & m2.Caption & "' ")
 End If
ElseIf Label5.Caption = 3 Then ' Topic Wise
 If Combo4.ListIndex = 3 Then
  c.Execute ("insert into holdimport select * from quesms where c_id='" & m1.Caption & "' and sub_id='" & m2.Caption & "' and tp_id='" & m3.Caption & "' ")
 ElseIf Combo4.ListIndex >= 0 Then
  c.Execute ("insert into holdimport select * from quesms where c_id='" & m1.Caption & "' and q_dif_lvl='" & Combo4.Text & "' and sub_id='" & m2.Caption & "' and tp_id='" & m3.Caption & "' ")
 End If
End If
Text2.Text = ""
Dim tempVar As String
Dim std1 As String, std2 As String
cdb.InitDir = "C:\STS"
cdb.Filter = "Excel 97-2003 Workbook(*.xls)|*.xls|CSV Files (*.csv)|*.csv|All Files(*.*)|*.*"
cdb.ShowSave
std1 = cdb.FileName
If std1 = "" Then
Else
If UCase(Right(std1, 3)) = "XLS" Then
Text2.Text = "set feed off" & vbCrLf & "set markup html on" & vbCrLf & "spool on" & vbCrLf & "spool " & std1 & vbCrLf & "select hq_no,hq_txt,hopt1,hopt2,hopt3,hopt4,hans_txt,hans_no,hq_expln from holdimport;" & vbCrLf & "spool off" & vbCrLf & "set markup html off" & vbCrLf & "commit;" & vbCrLf & "exit;"
Open App.Path + "\ocx\ExportToXls.sql" For Output As #1
 Print #1, Text2.Text
Close #1
 Shell "cmd.exe /c sqlplus sts/sts @" & App.Path & "\Ocx\ExportToXls.sql"
 MsgBox "File Successfully Exported", vbInformation + vbOKOnly, "Export"
 Form_Load
Exit Sub
ElseIf UCase(Right(std1, 3)) = "CSV" Then
 std2 = std1
 std1 = Mid$(std1, 1, Len(std1) - 4) & ".xls"
 Text2.Text = "set feed off" & vbCrLf & "set markup html on" & vbCrLf & "spool on" & vbCrLf & "spool " & std1 & vbCrLf & "select hq_no,hq_txt,hopt1,hopt2,hopt3,hopt4,hans_txt,hans_no,hq_expln from holdimport;" & vbCrLf & "spool off" & vbCrLf & "set markup html off" & vbCrLf & "commit;" & vbCrLf & "exit;"
 Open App.Path + "\ocx\ExportToXls.sql" For Output As #1
 Print #1, Text2.Text
 Close #1
 Shell "cmd.exe /c sqlplus sts/sts @" & App.Path & "\Ocx\ExportToXls.sql"
'---------Now Convert it into csv and delete the Xls File
 Shell "cmd.exe /c " & App.Path & "\ocx\XlsToCsv.vbs " & std1 & " " & std2
 Shell "cmd.exe /c taskkill /IM excel.exe" 'Close EXcel If Open
 MsgBox "CSV File Successfully Generated", vbInformation + vbOKOnly, "Done"
 Shell "cmd.exe /c del " & std1 'Deleting XLS File which is not required
 Shell "cmd.exe /c taskkill /IM excel.exe" 'Close EXcel If Open
 Form_Load
Else
 MsgBox "Invalid Format !! Please Choose either CSV or XLS Type", vbQuestion, ""
 Exit Sub
End If
End If
End Sub

Private Sub Command2_Click()
If Label5.Caption = 1 Then
 If Combo1.Text = "" Then
  MsgBox "Select Course First ", vbInformation + vbOKOnly, "Course Missing"
  Combo1.SetFocus
 Exit Sub
 End If
ElseIf Label5.Caption = 2 Then
 If Combo1.Text = "" Then
  MsgBox "Select Course First ", vbInformation + vbOKOnly, "Course Missing"
  Combo1.SetFocus
  Exit Sub
 ElseIf Combo2.Text = "" Then
  MsgBox "Select Subject Correspondent to the Course", vbInformation + vbOKOnly, "Subject Missing"
  Combo2.SetFocus
  Exit Sub
 End If
ElseIf Label5.Caption = 3 Then
 If Combo1.Text = "" Then
  MsgBox "Select Course First ", vbInformation + vbOKOnly, "Course Missing"
  Combo1.SetFocus
  Exit Sub
 ElseIf Combo2.Text = "" Then
  MsgBox "Select Subject Correspondent to the Course", vbInformation + vbOKOnly, "Subject Missing"
  Combo2.SetFocus
  Exit Sub
 ElseIf Combo3.Text = "" Then
  MsgBox "Select Topic Correspondent to the Subject of the course", vbInformation + vbOKOnly, "Topic Missing"
  Combo3.SetFocus
 Exit Sub
 End If
End If
If Combo4.Text = "" Then
MsgBox "Choose Difficulty Level of Question", vbInformation + vbOKOnly, "Select Difficulty"
Combo4.SetFocus
Exit Sub
End If
Text1.Locked = False
updtDA 'Calling Equivalent Cid , Sub ID and Tp ID
Set r = New ADODB.Recordset
Timer1.Enabled = True
If Label5.Caption = 2 Then
 If Combo4.ListIndex <> 3 Then
 Set r = c.Execute("select count(*) from quesms where c_id='" & m1.Caption & "' and Q_dif_lvl='" & Combo4.Text & "' and sub_id='" & m2.Caption & "' ")
 Else
 Set r = c.Execute("select count(*) from quesms where c_id='" & m1.Caption & "' and sub_id='" & m2.Caption & "' ")
 End If
 Label2.Caption = r.Fields(0)
ElseIf Label5.Caption = 3 Then
  If Combo4.ListIndex <> 3 Then
   Set r = c.Execute("select count(*) from quesms where c_id='" & m1.Caption & "' and Q_dif_lvl='" & Combo4.Text & "' and sub_id='" & m2.Caption & "' and tp_id='" & m3.Caption & "' ")
  Else
   Set r = c.Execute("select count(*) from quesms where c_id='" & m1.Caption & "' and sub_id='" & m2.Caption & "' and tp_id='" & m3.Caption & "' ")
  End If
  Label2.Caption = r.Fields(0)
ElseIf Label5.Caption = 1 Then
If Combo4.ListIndex <> 3 Then
 Set r = c.Execute("select count(*) from quesms where c_id='" & m1.Caption & "' and Q_dif_lvl='" & Combo4.Text & "' ")
Else
 Set r = c.Execute("select count(*) from quesms where c_id='" & m1.Caption & "' ")
End If
 Label2.Caption = r.Fields(0)
End If
Command1.Enabled = True
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Text1.Visible = True
Command1.Visible = True
Text1.SetFocus
End Sub
Sub updtDA() 'Important to load before save
Set r1 = New ADODB.Recordset
Set r2 = New ADODB.Recordset
Set r3 = New ADODB.Recordset
Set r1 = c1.Execute("SELECT c_id FROM COURSE WHERE C_NM='" & Combo1.Text & "'")
If IsNull(r1.Fields(0)) = False Then
 m1.Caption = r1.Fields(0)
End If
If Label5.Caption = 2 Then
  Set r2 = c1.Execute("SELECT sub_id from sub where sub_nm='" & Combo2.Text & "' AND C_ID=(SELECT C_ID FROM COURSE WHERE C_NM='" & Combo1.Text & "')")
 If IsNull(r2.Fields(0)) = False Then
  m2.Caption = r2.Fields(0)
 End If
ElseIf Label5.Caption = 3 Then
 Set r2 = c1.Execute("SELECT sub_id from sub where sub_nm='" & Combo2.Text & "' AND C_ID=(SELECT C_ID FROM COURSE WHERE C_NM='" & Combo1.Text & "')")
 Set r3 = c1.Execute("SELECT tp_id from TOPIC where TP_NM ='" & Combo3.Text & "' AND SUB_ID=(SELECT SUB_ID FROM SUB WHERE sub_nm='" & Combo2.Text & "' AND C_ID=(SELECT C_ID FROM COURSE WHERE C_NM='" & Combo1.Text & "')) AND C_ID=(SELECT C_ID FROM COURSE WHERE C_NM='" & Combo1.Text & "') ")
 If IsNull(r2.Fields(0)) = False Then
  m2.Caption = r2.Fields(0)
 End If
 If IsNull(r3.Fields(0)) = False Then
  m3.Caption = r3.Fields(0)
 End If
End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
conn
Command1.Enabled = False
Text1.Text = ""
Text1.Locked = True
Label2.Caption = ""
Label5.Caption = ""
m1.Caption = ""
m2.Caption = ""
m3.Caption = ""
Combo1.Clear
If rs_course.EOF = False Then
rs_course.MoveFirst
While rs_course.EOF = False
 Combo1.AddItem rs_course(0)
 rs_course.MoveNext
Wend
End If
Combo2.Clear
Combo3.Clear
Combo4.Clear
Combo4.AddItem "EASY"
Combo4.AddItem "MEDIUM"
Combo4.AddItem "HARD"
Combo4.AddItem "All"
Frame2.Enabled = False
PictureG2.Visible = False
Label5.Caption = 1
Frame2.Enabled = True
Label6.Visible = False
Label7.Visible = False
Combo2.Visible = False
Combo3.Visible = False
Text1.Locked = True
Command1.Enabled = False

Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Text1.Visible = False
Command1.Visible = False
Option1_Click
End Sub

Private Sub Option1_Click()
Label5.Caption = 1
Frame2.Enabled = True
Label6.Visible = False
Label7.Visible = False
Combo2.Visible = False
Combo3.Visible = False
Text1.Locked = True
Command1.Enabled = False
Option1.BackColor = &HC0C0C0
Option2.BackColor = &H8000000F
Option3.BackColor = &H8000000F
End Sub

Private Sub Option2_Click()
Label5.Caption = 2
Frame2.Enabled = True
Label6.Visible = True
Label7.Visible = False
Combo2.Visible = True
Combo3.Visible = False
Command1.Enabled = False
Text1.Locked = True
Option2.BackColor = &HC0C0C0
Option3.BackColor = &H8000000F
Option1.BackColor = &H8000000F
Combo1.SetFocus
End Sub

Private Sub Option3_Click()
Label5.Caption = 3
Frame2.Enabled = True
Label6.Visible = True
Label7.Visible = True
Combo2.Visible = True
Combo3.Visible = True
Command1.Enabled = False
Text1.Locked = True
Option3.BackColor = &HC0C0C0
Option2.BackColor = &H8000000F
Option1.BackColor = &H8000000F
Combo1.SetFocus
End Sub

Private Sub Text1_Change()
If Val(Text1.Text) > Val(Label2.Caption) Then
 MsgBox "Cannot Enter More Questions Than Exist in database, Enter Less Questions..", vbInformation + vbOKOnly, ""
 Text1.Text = ""
 Text1.SetFocus
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
Text1.SetFocus
ElseIf KeyAscii = 13 Then
 KeyAscii = 0
 If Val(Label2.Caption) >= Val(Text1.Text) Then
  Command1_Click
 End If
Else
KeyAscii = 0
End If
End Sub

Private Sub Text1_LostFocus()
If Label2.Caption <> "" Then
 If Val(Label2.Caption) < Val(Text1.Text) Then
 Text1.Text = ""
 MsgBox "More Than Total Questions..", vbCritical + vbOKOnly, "Invalid"
 Text1.SetFocus
End If
End If
End Sub

Private Sub Timer1_Timer()
ak = ak + 1
If ak > 7 Then
PictureG2.Visible = False
Timer1.Enabled = False
Else
PictureG2.Visible = True
End If
End Sub
