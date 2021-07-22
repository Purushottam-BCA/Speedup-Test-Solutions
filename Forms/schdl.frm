VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{EAFDAFBF-1D88-41DD-B117-60ECBC4B8441}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form FrmSchedule 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Schedule"
   ClientHeight    =   9585
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9870
   Icon            =   "schdl.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9585
   ScaleWidth      =   9870
   Begin VB.PictureBox a 
      Height          =   10335
      Left            =   0
      ScaleHeight     =   10275
      ScaleWidth      =   9915
      TabIndex        =   5
      Top             =   0
      Width           =   9975
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   8400
         Top             =   4920
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
         RecordSource    =   "select s.sch_id,s.sch_STRNTH,s.sch_timing,s.strt_time,s.end_time,c.c_nm from schdl s,course c where c.c_id=s.c_id"
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
         Bindings        =   "schdl.frx":0ECA
         Height          =   5235
         Left            =   0
         TabIndex        =   16
         Top             =   4320
         Width           =   9840
         _ExtentX        =   17357
         _ExtentY        =   9234
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         BackColor       =   14737632
         HeadLines       =   1
         RowHeight       =   22
         TabAction       =   2
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "SCH_ID"
            Caption         =   "Schedule ID"
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
            DataField       =   "SCH_STRNTH"
            Caption         =   "   Strength"
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
            DataField       =   "SCH_TIMING"
            Caption         =   "      Timing"
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
            DataField       =   "STRT_TIME"
            Caption         =   "  Start Time"
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
            DataField       =   "END_TIME"
            Caption         =   " End Time"
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
            DataField       =   "C_NM"
            Caption         =   "       Course"
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
            AllowRowSizing  =   0   'False
            Locked          =   -1  'True
            RecordSelectors =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   1500.095
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   1454.74
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   1755.213
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   1484.787
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnWidth     =   1409.953
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   1964.976
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   650
         Left            =   0
         TabIndex        =   14
         Top             =   3675
         Width           =   10140
         Begin VB.CommandButton btnback 
            BackColor       =   &H00C0C0FF&
            DisabledPicture =   "schdl.frx":0EDF
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   8640
            MouseIcon       =   "schdl.frx":1539
            MousePointer    =   99  'Custom
            Picture         =   "schdl.frx":168B
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Exit From Here"
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton btnUpdate 
            DisabledPicture =   "schdl.frx":1CE5
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   4605
            MouseIcon       =   "schdl.frx":2378
            MousePointer    =   99  'Custom
            Picture         =   "schdl.frx":24CA
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton btnDelete 
            DisabledPicture =   "schdl.frx":2B5D
            Height          =   390
            Left            =   3000
            MouseIcon       =   "schdl.frx":328C
            MousePointer    =   99  'Custom
            Picture         =   "schdl.frx":33DE
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   120
            Width           =   1400
         End
         Begin VB.CommandButton btnsave 
            BackColor       =   &H8000000E&
            DisabledPicture =   "schdl.frx":3B0D
            Height          =   390
            Left            =   1560
            MouseIcon       =   "schdl.frx":41C0
            MousePointer    =   99  'Custom
            Picture         =   "schdl.frx":4312
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   120
            Width           =   1245
         End
         Begin VB.CommandButton btnadd 
            BackColor       =   &H8000000E&
            DisabledPicture =   "schdl.frx":49C5
            Height          =   390
            Left            =   120
            MouseIcon       =   "schdl.frx":505E
            MousePointer    =   99  'Custom
            Picture         =   "schdl.frx":51B0
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   120
            Width           =   1230
         End
         Begin VB.CommandButton btnclear 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Clear"
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
            Left            =   6000
            MouseIcon       =   "schdl.frx":5849
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton btnsearch 
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
            Height          =   390
            Left            =   7320
            MouseIcon       =   "schdl.frx":599B
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   120
            Width           =   1095
         End
         Begin vkUserContolsXP.vkCommand vkCommand9 
            Height          =   495
            Left            =   11475
            TabIndex        =   15
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Caption         =   "&OK"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   2895
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   9855
         Begin VB.TextBox temp 
            BackColor       =   &H00FF80FF&
            Height          =   285
            Left            =   3960
            TabIndex        =   24
            Top             =   0
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Search Here"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000009&
            Height          =   1335
            Left            =   5760
            TabIndex        =   18
            Top             =   1560
            Visible         =   0   'False
            Width           =   3855
            Begin VB.CommandButton btnGO 
               Caption         =   "GO"
               Height          =   425
               Left            =   2880
               MouseIcon       =   "schdl.frx":5AED
               MousePointer    =   99  'Custom
               TabIndex        =   21
               Top             =   700
               Width           =   735
            End
            Begin VB.TextBox Text5 
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   425
               Left            =   240
               TabIndex        =   19
               Top             =   700
               Width           =   2415
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Enter ID or Start Time (Eg: 09:30 )"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   270
               Left            =   240
               TabIndex        =   20
               Top             =   345
               Width           =   3000
            End
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   345
            Left            =   2520
            MouseIcon       =   "schdl.frx":5C3F
            MousePointer    =   99  'Custom
            TabIndex        =   0
            Top             =   360
            Width           =   1935
         End
         Begin VB.ComboBox Combo2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   2520
            MouseIcon       =   "schdl.frx":5D91
            MousePointer    =   99  'Custom
            TabIndex        =   1
            Top             =   960
            Width           =   1935
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7365
            MaxLength       =   2
            TabIndex        =   2
            Top             =   210
            Width           =   1095
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   325
            Left            =   2640
            MaxLength       =   5
            TabIndex        =   3
            Top             =   1695
            Width           =   1695
         End
         Begin VB.TextBox Text4 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2640
            MaxLength       =   5
            TabIndex        =   4
            Top             =   2280
            Width           =   1695
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   7365
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   735
            Width           =   1695
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(Eg: 10:30 )"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   4560
            TabIndex        =   26
            Top             =   2280
            Width           =   990
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(Eg: 09:30 )"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   4560
            TabIndex        =   25
            Top             =   1680
            Width           =   990
         End
         Begin VB.Shape Shape3 
            BackColor       =   &H00C0FFFF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000006&
            Height          =   375
            Left            =   2520
            Shape           =   4  'Rounded Rectangle
            Top             =   1680
            Width           =   1935
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000006&
            Height          =   375
            Left            =   7320
            Shape           =   4  'Rounded Rectangle
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Course "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   400
            TabIndex        =   13
            Top             =   360
            Width           =   870
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Schedule ID"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   400
            TabIndex        =   12
            Top             =   960
            Width           =   1305
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Strength"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5820
            TabIndex        =   11
            Top             =   195
            Width           =   960
         End
         Begin VB.Shape Shape6 
            BackColor       =   &H00C0FFFF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000006&
            Height          =   375
            Left            =   7320
            Shape           =   4  'Rounded Rectangle
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Batch Start Time"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   400
            TabIndex        =   10
            Top             =   1680
            Width           =   1875
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Batch End Time"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   400
            TabIndex        =   9
            Top             =   2280
            Width           =   1725
         End
         Begin VB.Shape Shape7 
            BackColor       =   &H00C0FFFF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000006&
            Height          =   375
            Left            =   2520
            Shape           =   4  'Rounded Rectangle
            Top             =   2265
            Width           =   1935
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Timing"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5820
            TabIndex        =   8
            Top             =   720
            Width           =   720
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   3975
         X2              =   6000
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   3360
         Picture         =   "schdl.frx":5EE3
         Top             =   80
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Batch Time"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   420
         Left            =   4080
         TabIndex        =   17
         Top             =   90
         Width           =   1875
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000013&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   615
         Left            =   0
         Top             =   0
         Width           =   9975
      End
   End
End
Attribute VB_Name = "FrmSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim t As Integer, opt1 As Integer
Dim opt As String

Private Sub btnClear_Click()
btnadd.Enabled = True
btnsave.Enabled = False
btnUpdate.Enabled = False
btnDelete.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = True
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Combo2.Clear
sql = "select distinct (sch_id) from schdl"
Set r1 = c1.Execute(sql)
While r1.EOF = False
Combo2.AddItem r1.Fields(0)
r1.MoveNext
Wend
opt1 = 1
End Sub

Private Sub btnGO_Click()
If Text5.Text = "" Then
 MsgBox "Emplty Field", vbInformation + vbOKOnly, ""
 Text5.SetFocus
 Exit Sub
Else
Set r = c.Execute("select s.sch_id,s.sch_STRNTH,s.sch_timing,s.strt_time,s.end_time,c.c_nm,c.c_id from schdl s,course c where c.c_id=s.c_id and (upper(s.sch_id)='" & UCase(Trim(Text5.Text)) & "' or upper(s.strt_time)='" & UCase(Trim(Text5.Text)) & "') ")
If IsNull(r.Fields(0)) = False Then
 Text1.Text = UCase(r.Fields(1))
 Text2.Text = UCase(r.Fields(2))
 Text3.Text = UCase(r.Fields(3))
 Text4.Text = UCase(r.Fields(4))
 Combo1.Text = UCase(r.Fields(5))
 Combo2.Text = UCase(r.Fields(0))
 temp.Text = r.Fields(6)
 btnsave.Enabled = False
 btnUpdate.Enabled = True
 btnDelete.Enabled = True
Else
MsgBox "Invalid ID or Time, Try Again", vbInformation, ""
Text5.Text = ""
Text5.SetFocus
End If
End If
End Sub

Private Sub btnSearch_Click()
If opt1 = 1 Then
 Frame3.Visible = True
 opt1 = 0
Else
Frame3.Visible = False
opt1 = 1
End If
 btnsave.Enabled = False
End Sub

Private Sub Combo1_Click() 'Course Combo
'sql = "select * from schdl where c_id=(select c_id from course where c_nm='" & Combo1.Text & "')"
'Set r1 = c1.Execute(sql)
'Combo2.Enabled = True
'While (r1.EOF) = False
' Combo2.AddItem r1.Fields(0)
'r1.MoveNext
'Wend
temp.Text = ""
Set r = c.Execute("select c_id from course where upper(c_nm)='" & UCase(Combo1.Text) & "' ")
  temp.Text = r.Fields(0)
  Text1.SetFocus
End Sub

Private Sub combo2_Click()
Set r = c.Execute("select s.sch_STRNTH,s.sch_timing,s.strt_time,s.end_time,c.c_nm,c.c_id from schdl s,course c where c.c_id=s.c_id and upper(s.sch_id)='" & UCase(Trim(Combo2.Text)) & "' ")
If IsNull(r.Fields(0)) = False Then
 Text1.Text = UCase(r.Fields(0))
 Text2.Text = UCase(r.Fields(1))
 Text3.Text = UCase(r.Fields(2))
 Text4.Text = UCase(r.Fields(3))
 Combo1.Text = UCase(r.Fields(4))
 temp.Text = r.Fields(5)
 btnsave.Enabled = False
 btnUpdate.Enabled = True
 btnDelete.Enabled = True
Else
 MsgBox "Record Not Found", vbInformation + vbOKOnly, ""
End If
End Sub

Private Sub DataGrid1_Click()
'btnSave.Enabled = False
'btnUpdate.Enabled = True
'BtnDelete.Enabled = True
'Set Combo1.DataSource = DataGrid1.DataSource
'Set Combo2.DataSource = DataGrid1.DataSource
'Set Text1.DataSource = DataGrid1.DataSource
'Set Text2.DataSource = DataGrid1.DataSource
'Set Text3.DataSource = DataGrid1.DataSource
'Set Text4.DataSource = DataGrid1.DataSource
End Sub

Private Sub Form_Load()
Me.Top = 500
Me.Left = 5500
conn
btnadd.Enabled = True
btnsave.Enabled = False
btnUpdate.Enabled = False
btnDelete.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = True
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Combo1.Clear
Combo2.Clear
Set r1 = New ADODB.Recordset
sql = "select  initcap(c_nm) from course"
Set r1 = c1.Execute(sql)
While r1.EOF = False
Combo1.AddItem r1.Fields(0)
r1.MoveNext
Wend

sql = "select distinct (sch_id) from schdl"
Set r1 = c1.Execute(sql)
While r1.EOF = False
Combo2.AddItem r1.Fields(0)
r1.MoveNext
Wend
opt1 = 1
End Sub

Private Sub btnadd_Click() 'new'
cauto_id
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Combo1.Text = ""
Combo1.Enabled = True
Combo2.Enabled = False
btnsave.Enabled = True
btnadd.Enabled = False
btnDelete.Enabled = False
btnUpdate.Enabled = False
End Sub

Private Sub btnsave_Click() 'save'
If Trim(Combo1.Text) = "" Then
MsgBox "Course Name Blank", vbCritical + vbOKOnly, "warning"
Combo1.SetFocus
ElseIf Trim(Combo2.Text) = "" Then
MsgBox "Schdule Name Blank", vbCritical + vbOKOnly, "Warning"
Combo2.SetFocus
ElseIf Trim(Text1.Text) = "" Then
 MsgBox "Enter Schdule Strength", vbCritical + vbOKOnly, "Warning"
Text1.SetFocus
ElseIf Trim(Text2.Text) = "" Then
 MsgBox "Enter Schdule Timing", vbCritical + vbOKOnly, "Warning"
Text2.SetFocus
ElseIf Trim(Text3.Text) = "" Then
 MsgBox "Enter Start Time", vbCritical + vbOKOnly, "Warning"
Text3.SetFocus
ElseIf Trim(Text4.Text) = "" Then
 MsgBox "Enter End Time", vbCritical + vbOKOnly, "Warning"
Text4.SetFocus
Else
If InStr(Text3.Text, ":") = False Or Len(Text3.Text) < 4 Then
 MsgBox "Enter The Time In Proper Format ( 09:00 )", vbCritical + vbOKOnly, "Invalid Format"
 Text3.SetFocus
 Exit Sub
End If
If InStr(Text4.Text, ":") = False Or Len(Text4.Text) < 4 Then
 MsgBox "Enter The Time In Proper Format ( 09:00 )", vbCritical + vbOKOnly, "Invalid Format"
 Text4.SetFocus
 Exit Sub
End If
Set r1 = New ADODB.Recordset
sql = "insert into schdl values ('" & Combo2.Text & "'," & Text1.Text & ",'" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & temp.Text & "')"
Set r1 = c1.Execute(sql)
MsgBox "Data Saved", vbInformation, ""
Adodc1.Refresh
btnadd.Enabled = True
btnsave.Enabled = False
btnUpdate.Enabled = False
btnDelete.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = True
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Combo2.Clear
sql = "select distinct (sch_id) from schdl"
Set r1 = c1.Execute(sql)
While r1.EOF = False
Combo2.AddItem r1.Fields(0)
r1.MoveNext
Wend
opt1 = 1
End If
End Sub

Private Sub btnDelete_Click() 'delete'
If Trim(Combo1.Text) = "" Or Trim(Combo2.Text) = "" Then
 MsgBox "Select Corrrect Schedule", vbCritical + vbOKOnly, "Delete ERROR"
 Exit Sub
Else
opt = MsgBox("Are You Sure to Delete ?", vbQuestion + vbYesNo, "Delete conformation!")
If opt = vbYes Then
Set r1 = New ADODB.Recordset
sql = " delete from schdl where sch_id='" & Combo2.Text & "'"
Set r1 = c1.Execute(sql)
MsgBox "record deleted", vbInformation, ""
Adodc1.Refresh
Form_Load
End If
End If
End Sub

Private Sub btnUpdate_Click() 'update'
If Trim(Combo1.Text) = "" Or Trim(Combo2.Text) = "" Then
 MsgBox "Select Corrrect Schdule", vbCritical + vbOKOnly, "Update ERROR"
 Else
  opt = MsgBox("Are You Sure to Update ?", vbQuestion + vbYesNo, "UPDATE")
If opt = vbYes Then
 If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Then
  MsgBox "fill all fields", vbInformation + vbOKOnly, ""
  Text1.SetFocus
  Exit Sub
 Else
 temp.Text = ""
  Set r = c.Execute("select c_id from course where upper(c_nm)='" & UCase(Combo1.Text) & "' ")
  temp.Text = r.Fields(0)
  sql = " update schdl set sch_strnth=" & Text1.Text & ",sch_timing='" & Text2.Text & "',strt_time='" & Text3.Text & "',c_id='" & temp.Text & "',end_time='" & Text4.Text & "' where sch_id='" & Combo2.Text & "'"
  Set r1 = c1.Execute(sql)
   MsgBox "record updated"
   Adodc1.Refresh
   Form_Load
 End If
End If
End If
End Sub

Private Sub btnBack_Click()
Unload Me
End Sub
Public Function cauto_id()
Set r1 = New ADODB.Recordset
sql = "select max(to_number(substr(sch_id,4,length(sch_id))))from schdl"
Set r1 = c1.Execute(sql)
If IsNull(r1.Fields(0)) Then
Combo2.Text = "SCH00" & 1
Else
t = r1.Fields(0)
If t > 0 And t < 9 Then
 Combo2.Text = "SCH00" & (t + 1)
ElseIf t < 99 Then
 Combo2.Text = "SCH0" & (t + 1)
End If
End If
End Function

Private Sub Text1_KeyPress(KeyAscii As Integer)
If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Then
        Text1.SetFocus
ElseIf KeyAscii = 13 Then
 KeyAscii = 0
 Text3.SetFocus
Else
KeyAscii = 0
End If
End Sub
Private Sub Text3_Change()
Text2.Text = Text3.Text
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Or KeyAscii = 58 Then
        Text3.SetFocus
ElseIf KeyAscii = 13 Then
       KeyAscii = 0
       Text4.SetFocus
Else
       KeyAscii = 0
End If
End Sub

Private Sub Text4_Change()
If Trim(Text3.Text) <> "" Then
 Text2.Text = Text3.Text & "-" & Text4.Text
Else
 Text2.Text = " - " & Text4.Text
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
Text2.Text = ""
If Text3.Text <> "" And Text4.Text <> "" Then
Text2.Text = Text3.Text & "-" & Text4.Text & Chr(KeyAscii)
End If
End Sub

Private Sub Text5_Change()
On Error Resume Next
Adodc1.RecordSource = "select s.sch_id,s.sch_STRNTH,s.sch_timing,s.strt_time,s.end_time,c.c_nm,c.c_id from schdl s,course c where c.c_id=s.c_id and (upper(s.sch_id) like '" & UCase(Trim(Text5.Text)) & "%' or upper(s.strt_time)='" & UCase(Trim(Text5.Text)) & "%')"
Adodc1.Refresh
Set r = c.Execute("select s.sch_id,s.sch_STRNTH,s.sch_timing,s.strt_time,s.end_time,c.c_nm,c.c_id from schdl s,course c where c.c_id=s.c_id and (upper(s.sch_id)like '" & UCase(Trim(Text5.Text)) & "%' or upper(s.strt_time) like '" & UCase(Trim(Text5.Text)) & "%') ")
If IsNull(r.Fields(0)) = False Then
 Text1.Text = UCase(r.Fields(1))
 Text2.Text = UCase(r.Fields(2))
 Text3.Text = UCase(r.Fields(3))
 Text4.Text = UCase(r.Fields(4))
 Combo1.Text = UCase(r.Fields(5))
 Combo2.Text = UCase(r.Fields(0))
 temp.Text = r.Fields(6)
 btnsave.Enabled = False
 btnUpdate.Enabled = True
 btnDelete.Enabled = True
 Else
 Text1.Text = ""
 Text2.Text = ""
 Text3.Text = ""
 Text4.Text = ""
 Combo1.Text = ""
 Combo2.Text = ""
 temp.Text = ""
 End If
End Sub
