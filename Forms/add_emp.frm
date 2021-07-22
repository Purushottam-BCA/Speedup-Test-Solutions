VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form FrmEmpMaster 
   BackColor       =   &H00E0E0E0&
   Caption         =   "User"
   ClientHeight    =   10290
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   Icon            =   "add_emp.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10290
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid db1 
      Bindings        =   "add_emp.frx":0EE2
      Height          =   8415
      Left            =   10440
      TabIndex        =   53
      Top             =   480
      Width           =   10000
      _ExtentX        =   17648
      _ExtentY        =   14843
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      HeadLines       =   1
      RowHeight       =   22
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri Light"
         Size            =   11.25
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
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "EMP_ID"
         Caption         =   "User ID"
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
         DataField       =   "E_NM"
         Caption         =   "Name"
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
         DataField       =   "E_FATHER"
         Caption         =   "Father Name"
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
         DataField       =   "E_MOTHER"
         Caption         =   "Mother Name"
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
         DataField       =   "E_MOB"
         Caption         =   "Mobile No"
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
         DataField       =   "E_GNDR"
         Caption         =   "Gender"
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
         DataField       =   "E_ADD"
         Caption         =   "Address"
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
         DataField       =   "E_STATE"
         Caption         =   "State"
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
         DataField       =   "E_J_DT"
         Caption         =   "Join Date"
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
         DataField       =   "E_SAL"
         Caption         =   "Salary"
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
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         RecordSelectors =   0   'False
         BeginProperty Column00 
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1769.953
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1544.882
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1665.071
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1635.024
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1124.787
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid Db2 
      Bindings        =   "add_emp.frx":0EF7
      Height          =   8415
      Left            =   10455
      TabIndex        =   54
      Top             =   480
      Width           =   9845
      _ExtentX        =   17357
      _ExtentY        =   14843
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   22
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri Light"
         Size            =   11.25
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
         DataField       =   "A_ID"
         Caption         =   "ID"
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
         DataField       =   "A_NM"
         Caption         =   "Name"
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
         DataField       =   "A_FATHER"
         Caption         =   "Father Name"
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
         DataField       =   "A_MOTHER"
         Caption         =   "Mother Name"
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
         DataField       =   "A_MOB"
         Caption         =   "Mobile"
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
         DataField       =   "A_GNDR"
         Caption         =   "Gender"
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
         DataField       =   "A_ADD"
         Caption         =   "Address"
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
         DataField       =   "A_STATE"
         Caption         =   "State"
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
         DataField       =   "A_J_DT"
         Caption         =   "Join Date"
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
         Locked          =   -1  'True
         RecordSelectors =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1755.213
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1574.929
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1649.764
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1230.236
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1649.764
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   2129.953
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   12360
      Top             =   9120
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      RecordSource    =   "select a_id,a_nm, a_father,a_mother, a_mob, a_gndr, a_add, a_state, a_j_dt from adminTbl"
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
      Height          =   2775
      Left            =   3600
      TabIndex        =   50
      Top             =   6550
      Width           =   3255
      _Version        =   524288
      _ExtentX        =   5741
      _ExtentY        =   4895
      _StockProps     =   1
      BackColor       =   14737632
      Year            =   2019
      Month           =   6
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1095
      Left            =   10440
      TabIndex        =   3
      Top             =   9010
      Width           =   10335
      Begin VB.CommandButton back_btn 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   425
         Left            =   8640
         MouseIcon       =   "add_emp.frx":0F0C
         MousePointer    =   99  'Custom
         TabIndex        =   48
         Top             =   405
         Width           =   1215
      End
      Begin VB.CommandButton Clear 
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   425
         Left            =   6960
         MouseIcon       =   "add_emp.frx":105E
         MousePointer    =   99  'Custom
         TabIndex        =   47
         Top             =   405
         Width           =   1455
      End
      Begin VB.CommandButton dl_btn 
         Caption         =   "Remove"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   425
         Left            =   5280
         MouseIcon       =   "add_emp.frx":11B0
         MousePointer    =   99  'Custom
         TabIndex        =   46
         Top             =   405
         Width           =   1455
      End
      Begin VB.CommandButton update_btn 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   425
         Left            =   3600
         MouseIcon       =   "add_emp.frx":1302
         MousePointer    =   99  'Custom
         TabIndex        =   45
         Top             =   405
         Width           =   1455
      End
      Begin VB.CommandButton sv_btn 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   425
         Left            =   1920
         MouseIcon       =   "add_emp.frx":1454
         MousePointer    =   99  'Custom
         TabIndex        =   44
         Top             =   405
         Width           =   1455
      End
      Begin VB.CommandButton add_btn 
         Caption         =   "Add new"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   425
         Left            =   240
         MouseIcon       =   "add_emp.frx":15A6
         MousePointer    =   99  'Custom
         TabIndex        =   43
         Top             =   405
         Width           =   1455
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   0
         Top             =   120
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
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
         RecordSource    =   "select emp_id,e_nm, e_father,e_mother, e_mob, e_gndr, e_add, e_state, e_j_dt,e_sal  from emp"
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
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   10455
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   10425
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   1095
         Left            =   3720
         TabIndex        =   40
         Top             =   240
         Width           =   3975
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
            Height          =   350
            Left            =   2520
            MouseIcon       =   "add_emp.frx":16F8
            MousePointer    =   99  'Custom
            TabIndex        =   42
            Top             =   480
            Width           =   1215
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00FFFFFF&
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
            Left            =   300
            Style           =   1  'Simple Combo
            TabIndex        =   41
            Text            =   "Mukesh Lal"
            Top             =   480
            Width           =   2055
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "User Type"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   2055
         Left            =   7920
         TabIndex        =   36
         Top             =   3480
         Width           =   2295
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
            ForeColor       =   &H00FFFFFF&
            Height          =   390
            Left            =   240
            MouseIcon       =   "add_emp.frx":184A
            MousePointer    =   99  'Custom
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User Role :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   480
            TabIndex        =   38
            Top             =   520
            Width           =   1125
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Contact && Residential"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   3375
         Left            =   240
         TabIndex        =   11
         Top             =   5640
         Width           =   9975
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
            ForeColor       =   &H00FFFFFF&
            Height          =   390
            Left            =   7320
            MouseIcon       =   "add_emp.frx":199C
            MousePointer    =   99  'Custom
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   1920
            Width           =   2415
         End
         Begin VB.TextBox PinCode 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   7500
            MaxLength       =   6
            TabIndex        =   34
            Top             =   1245
            Width           =   1200
         End
         Begin VB.TextBox mobno2 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   7500
            MaxLength       =   10
            TabIndex        =   32
            Top             =   525
            Width           =   2040
         End
         Begin VB.TextBox address 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1020
            Left            =   2340
            MaxLength       =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   25
            Top             =   1275
            Width           =   2850
         End
         Begin VB.TextBox mobNo 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2460
            MaxLength       =   10
            TabIndex        =   14
            Top             =   525
            Width           =   2040
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
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
            Left            =   2380
            MaxLength       =   12
            TabIndex        =   13
            Top             =   2700
            Width           =   2760
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   7400
            MaxLength       =   39
            TabIndex        =   12
            Top             =   2700
            Width           =   2280
         End
         Begin VB.Shape Shape5 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00808080&
            Height          =   405
            Index           =   5
            Left            =   7320
            Shape           =   4  'Rounded Rectangle
            Top             =   2640
            Width           =   2475
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "State  :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   6360
            TabIndex        =   35
            Top             =   1920
            Width           =   690
         End
         Begin VB.Shape Shape5 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00808080&
            Height          =   405
            Index           =   11
            Left            =   7320
            Shape           =   4  'Rounded Rectangle
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pin Code  :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   5955
            TabIndex        =   33
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Shape Shape5 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00808080&
            Height          =   405
            Index           =   8
            Left            =   2280
            Shape           =   4  'Rounded Rectangle
            Top             =   2640
            Width           =   2955
         End
         Begin VB.Shape Shape5 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00808080&
            Height          =   1125
            Index           =   4
            Left            =   2280
            Shape           =   4  'Rounded Rectangle
            Top             =   1215
            Width           =   2955
         End
         Begin VB.Shape Shape5 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00808080&
            Height          =   405
            Index           =   1
            Left            =   7320
            Shape           =   4  'Rounded Rectangle
            Top             =   480
            Width           =   2415
         End
         Begin VB.Shape Shape5 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00808080&
            Height          =   405
            Index           =   0
            Left            =   2280
            Shape           =   4  'Rounded Rectangle
            Top             =   480
            Width           =   2415
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address  :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1080
            TabIndex        =   26
            Top             =   1200
            Width           =   1020
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact No 2  :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   5535
            TabIndex        =   18
            Top             =   480
            Width           =   1635
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact No 1  :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   600
            TabIndex        =   17
            Top             =   480
            Width           =   1515
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Adhar No.  :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   885
            TabIndex        =   16
            Top             =   2640
            Width           =   1260
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Email ID  :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   5985
            TabIndex        =   15
            Top             =   2670
            Width           =   1140
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Personal information"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   4095
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   7455
         Begin VB.TextBox emp_mother 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2340
            MaxLength       =   50
            TabIndex        =   30
            Top             =   2325
            Width           =   4320
         End
         Begin VB.ComboBox Combo3 
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   390
            Left            =   5160
            MouseIcon       =   "add_emp.frx":1AEE
            MousePointer    =   99  'Custom
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   3150
            Width           =   1575
         End
         Begin VB.TextBox emp_father 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2340
            MaxLength       =   50
            TabIndex        =   27
            Top             =   1485
            Width           =   4320
         End
         Begin VB.TextBox emp_name 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2340
            MaxLength       =   50
            TabIndex        =   7
            Top             =   705
            Width           =   4320
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd-MMM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   375
            Left            =   2280
            TabIndex        =   6
            Top             =   3150
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            MousePointer    =   99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "add_emp.frx":1C40
            Format          =   101056515
            CurrentDate     =   43579
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mother Name :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   550
            TabIndex        =   31
            Top             =   2325
            Width           =   1545
         End
         Begin VB.Shape Shape5 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00808080&
            Height          =   405
            Index           =   6
            Left            =   2280
            Shape           =   4  'Rounded Rectangle
            Top             =   2280
            Width           =   4455
         End
         Begin VB.Shape Shape5 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00808080&
            Height          =   405
            Index           =   10
            Left            =   2280
            Shape           =   4  'Rounded Rectangle
            Top             =   1440
            Width           =   4455
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Father Name :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   630
            TabIndex        =   28
            Top             =   1500
            Width           =   1470
         End
         Begin VB.Shape Shape5 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00808080&
            Height          =   405
            Index           =   3
            Left            =   2280
            Shape           =   4  'Rounded Rectangle
            Top             =   660
            Width           =   4455
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   " Gender  :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3960
            TabIndex        =   10
            Top             =   3150
            Width           =   1020
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name  :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1320
            TabIndex        =   9
            Top             =   720
            Width           =   780
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   " Date of Birth  :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   525
            TabIndex        =   8
            Top             =   3150
            Width           =   1575
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Others"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   1095
         Left            =   240
         TabIndex        =   19
         Top             =   9120
         Width           =   9975
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
            ForeColor       =   &H00FFFFFF&
            Height          =   390
            Left            =   8280
            MouseIcon       =   "add_emp.frx":1DA2
            MousePointer    =   99  'Custom
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox Text4 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   4980
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   405
            Width           =   1440
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1815
            MaxLength       =   5
            TabIndex        =   20
            Top             =   405
            Width           =   1200
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Rs."
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1465
            TabIndex        =   51
            Top             =   400
            Width           =   255
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " Joining Date  :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3360
            TabIndex        =   24
            Top             =   360
            Width           =   1500
         End
         Begin VB.Shape Shape5 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00808080&
            Height          =   405
            Index           =   7
            Left            =   4920
            Shape           =   4  'Rounded Rectangle
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qualification  :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6765
            TabIndex        =   23
            Top             =   360
            Width           =   1470
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Salary  :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   240
            TabIndex        =   21
            Top             =   360
            Width           =   1020
         End
         Begin VB.Shape Shape5 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00808080&
            Height          =   405
            Index           =   2
            Left            =   1395
            Shape           =   4  'Rounded Rectangle
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.CommandButton upload 
         Caption         =   "Browse..."
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
         Left            =   8040
         MouseIcon       =   "add_emp.frx":1EF4
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   2880
         Width           =   2175
      End
      Begin MSComDlg.CommonDialog cd1 
         Left            =   7680
         Top             =   1680
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   2415
         Left            =   8040
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "E001"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   325
         Left            =   2040
         TabIndex        =   2
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee ID  :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   1500
      End
   End
   Begin MSComctlLib.TabStrip Tb1 
      Height          =   8895
      Left            =   10440
      TabIndex        =   52
      Top             =   0
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   15690
      Style           =   1
      TabMinWidth     =   3528
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "     Standard User"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "        Admin User"
            ImageVarType    =   2
         EndProperty
      EndProperty
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "add_emp.frx":2046
   End
End
Attribute VB_Name = "FrmEmpMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pic_name As String, pic_ext As String, pic_changed As Boolean
Option Explicit
Dim t As Integer
Dim opt As String
Private Sub add_btn_Click()
cauto_id
emp_name.Text = ""
emp_father.Text = ""
emp_mother.Text = ""
address.Text = ""
mobNo.Text = ""
mobno2.Text = ""
emp_name.Text = ""
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
PinCode.Text = ""
pic_name = ""
Image1.Picture = Nothing
sv_btn.Enabled = True
End Sub

Private Sub address_KeyPress(KeyAscii As Integer)
If (((KeyAscii >= 65) And (KeyAscii <= 90)) Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or (KeyAscii = 8) Or (KeyAscii = 32)) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 44 Then
   address.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   PinCode.SetFocus
  Else
   KeyAscii = 0
  End If
End Sub

Private Sub back_btn_Click()
Unload Me
End Sub

Private Sub Calendar1_Click()
If Calendar1.Value < Date Then
 MsgBox "You cannot Select the date that " & vbCrLf & "is already passed away." & vbCrLf & "Please select valid date", vbExclamation + vbOKOnly, "Invalid date"
 Text4.SetFocus
Else
 Text4.Text = Calendar1.Day & "-" & Calendar1.Month & "-" & Calendar1.Year
 Calendar1.Visible = False
 End If
End Sub

Private Sub clear_Click()
On Error Resume Next
Image1.Picture = Nothing
pic_name = ""
Form_Load
End Sub
Public Function Search_Rec()
On Error GoTo k:
Set r1 = New ADODB.Recordset
sql = " select * from emp where upper(emp_id)='" & UCase(Trim(Combo1.Text)) & "' or upper(e_nm)= '" & UCase(Trim(Combo1.Text)) & "'  "
Set r1 = c1.Execute(sql)
If r1.EOF = False Then
sv_btn.Enabled = False
dl_btn.Enabled = True
update_btn.Enabled = True
add_btn.Enabled = True
 
 Label10.Caption = r1.Fields(0)
 emp_name.Text = r1.Fields(1)
 emp_father.Text = r1.Fields(2)
  emp_mother.Text = r1.Fields(3)
 address.Text = r1.Fields(4)
 Combo4.Text = r1.Fields(5)
 mobNo.Text = r1.Fields(6)
 If IsNull(r1.Fields(7)) = False Then
 mobno2.Text = r1.Fields(7)
 Else
 mobno2.Text = ""
 End If
 DTPicker1.Value = r1.Fields(8)
 Combo3.Text = r1.Fields(9)
 Text1.Text = r1.Fields(10)
 Text2.Text = r1.Fields(11)
 Text3.Text = r1.Fields(12)
 PinCode.Text = r1.Fields(13)
  Text4.Text = r1.Fields(14)
 Combo5.Text = r1.Fields(15)
 pic_name = r1.Fields(16)
 Combo2.Text = "Employee"
 Image1.Picture = LoadPicture(pic_name)
 Adodc1.RecordSource = " select emp_id,e_nm, e_father,e_mother, e_mob, e_gndr, e_add, e_state, e_j_dt,e_sal from emp where upper(emp_id)='" & UCase(Trim(Combo1.Text)) & "' or upper(e_nm)= '" & UCase(Trim(Combo1.Text)) & "' "
 Adodc1.Refresh
Exit Function
End If
 sql = "Select * from adminTbl where upper(a_id)='" & UCase(Trim(Combo1.Text)) & "' or upper(a_nm)= '" & UCase(Trim(Combo1.Text)) & "'  "
 Set r1 = c1.Execute(sql)
 If r1.EOF = False Then
sv_btn.Enabled = False
dl_btn.Enabled = True
update_btn.Enabled = True
add_btn.Enabled = True
  
  Label10.Caption = r1.Fields(0)
  emp_name.Text = r1.Fields(1)
  emp_father.Text = r1.Fields(2)
  emp_mother.Text = r1.Fields(3)
  address.Text = r1.Fields(4)
  Combo4.Text = r1.Fields(5)
  mobNo.Text = r1.Fields(6)
  If IsNull(r1.Fields(7)) = False Then
   mobno2.Text = r1.Fields(7)
  Else
   mobno2.Text = ""
  End If
 DTPicker1.Value = r1.Fields(8)
 Combo3.Text = r1.Fields(9)
 Text1.Text = r1.Fields(10)
 Text2.Text = r1.Fields(11)
 Text3.Text = r1.Fields(12)
 PinCode.Text = r1.Fields(13)
  Text4.Text = r1.Fields(14)
 Combo5.Text = r1.Fields(15)
 pic_name = r1.Fields(16)
 Image1.Picture = LoadPicture(pic_name)
 Combo2.Text = "Admin"
 Adodc1.RecordSource = " select a_id,a_nm, a_father,a_mother, a_mob, a_gndr, a_add, a_state, a_j_dt from adminTbl where upper(a_id)='" & UCase(Trim(Combo1.Text)) & "' or upper(a_nm)= '" & UCase(Trim(Combo1.Text)) & "'  "
 Adodc1.Refresh
 Exit Function
 End If
 MsgBox "Record Not Found", vbExclamation + vbOKOnly, "No data Available"
 Exit Function
k:
 Image1.Picture = LoadPicture(App.Path & "\Graphics\Main_Screen_Icon\PicNotAvail.jpg")
End Function
Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
Search_Rec
End If
End Sub
Private Sub Combo5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 KeyAscii = 0
 Combo3.SetFocus
End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 KeyAscii = 0
 Combo2.SetFocus
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 KeyAscii = 0
 Combo4.SetFocus
End If
End Sub
Private Sub Combo4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 KeyAscii = 0
 Combo5.SetFocus
End If
End Sub

Private Sub Command1_Click()
If Trim(Combo1.Text) = "" Then
 MsgBox "Enter Some Value For Searching...", vbExclamation + vbOKOnly, ""
 Combo1.SetFocus
 Exit Sub
 End If
On Error GoTo k:
Set r1 = New ADODB.Recordset
sql = " select * from emp where upper(emp_id)='" & UCase(Trim(Combo1.Text)) & "' or upper(e_nm)= '" & UCase(Trim(Combo1.Text)) & "'  "
Set r1 = c1.Execute(sql)
If r1.EOF = False Then
sv_btn.Enabled = False
dl_btn.Enabled = True
update_btn.Enabled = True
add_btn.Enabled = True
 
 Label10.Caption = r1.Fields(0)
 emp_name.Text = r1.Fields(1)
 emp_father.Text = r1.Fields(2)
  emp_mother.Text = r1.Fields(3)
 address.Text = r1.Fields(4)
 Combo4.Text = r1.Fields(5)
 mobNo.Text = r1.Fields(6)
 If IsNull(r1.Fields(7)) = False Then
 mobno2.Text = r1.Fields(7)
 Else
 mobno2.Text = ""
 End If
 DTPicker1.Value = r1.Fields(8)
 Combo3.Text = r1.Fields(9)
 Text1.Text = r1.Fields(10)
 Text2.Text = r1.Fields(11)
 Text3.Text = r1.Fields(12)
 PinCode.Text = r1.Fields(13)
  Text4.Text = r1.Fields(14)
 Combo5.Text = r1.Fields(15)
 pic_name = r1.Fields(16)
 Combo2.Text = "Admin"
 Image1.Picture = LoadPicture(pic_name)
 Adodc1.RecordSource = " select emp_id,e_nm, e_father,e_mother, e_mob, e_gndr, e_add, e_state, e_j_dt,e_sal from emp where upper(emp_id)='" & UCase(Trim(Combo1.Text)) & "' or upper(e_nm)= '" & UCase(Trim(Combo1.Text)) & "' "
 Adodc1.Refresh
Exit Sub
End If
 sql = "Select * from adminTbl where upper(a_id)='" & UCase(Trim(Combo1.Text)) & "' or upper(a_nm)= '" & UCase(Trim(Combo1.Text)) & "'  "
 Set r1 = c1.Execute(sql)
 If r1.EOF = False Then
sv_btn.Enabled = False
dl_btn.Enabled = True
update_btn.Enabled = True
add_btn.Enabled = True
  
  Label10.Caption = r1.Fields(0)
  emp_name.Text = r1.Fields(1)
  emp_father.Text = r1.Fields(2)
  emp_mother.Text = r1.Fields(3)
  address.Text = r1.Fields(4)
  Combo4.Text = r1.Fields(5)
  mobNo.Text = r1.Fields(6)
  If IsNull(r1.Fields(6)) = False Then
   mobno2.Text = r1.Fields(7)
  Else
   mobno2.Text = ""
  End If
 DTPicker1.Value = r1.Fields(8)
 Combo3.Text = r1.Fields(9)
 Text1.Text = r1.Fields(10)
 Text2.Text = r1.Fields(11)
 Text3.Text = r1.Fields(12)
 PinCode.Text = r1.Fields(13)
  Text4.Text = r1.Fields(14)
 Combo5.Text = r1.Fields(15)
 pic_name = r1.Fields(16)
 Image1.Picture = LoadPicture(pic_name)
 Combo2.Text = "Admin"
 Adodc1.RecordSource = " select a_id,a_nm, a_father,a_mother, a_mob, a_gndr, a_add, a_state, a_j_dt from adminTbl where upper(a_id)='" & UCase(Trim(Combo1.Text)) & "' or upper(a_nm)= '" & UCase(Trim(Combo1.Text)) & "'  "
 Adodc1.Refresh
 Exit Sub
 End If
 MsgBox "Record Not Found", vbExclamation + vbOKOnly, "No data Available"
 Exit Sub
k:
 Image1.Picture = LoadPicture(App.Path & "\Graphics\Main_Screen_Icon\PicNotAvail.jpg")
End Sub

Private Sub dl_btn_Click()
conn
If Trim(Combo1.Text) = "" Or Label10.Caption = "" Then
 MsgBox "Select Corrrect User ID or Name..", vbCritical + vbOKOnly, "Delete ERROR"
 Combo1.SetFocus
 Exit Sub
End If
Set r1 = New ADODB.Recordset
opt = MsgBox("Are You Sure to Delete ???" & vbCrLf & "This will delete all record of User", vbQuestion + vbYesNo, "Delete conformation!")
If opt = vbYes Then
 If UCase(Combo2.Text) = UCase("Admin") Then
  sql = " delete from adminTbl where a_id='" & Label10.Caption & "'"
 ElseIf UCase(Trim(Combo2.Text)) <> "" Then
  sql = " delete from emp where emp_id='" & Label10.Caption & "'"
 End If
  c1.Execute (sql)
  MsgBox "Record SuccessFully deleted", vbInformation + vbOKOnly, "Delete"
  Adodc1.Refresh
  Adodc2.Refresh
  Form_Load
End If
End Sub

Private Sub emp_father_KeyPress(KeyAscii As Integer)
If (((KeyAscii >= 65) And (KeyAscii <= 90)) Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or (KeyAscii = 8) Or (KeyAscii = 32)) Then
   emp_father.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   emp_mother.SetFocus
  Else
   KeyAscii = 0
  End If
End Sub
Private Sub emp_mother_KeyPress(KeyAscii As Integer)
If (((KeyAscii >= 65) And (KeyAscii <= 90)) Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or (KeyAscii = 8) Or (KeyAscii = 32)) Then
   emp_mother.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   mobNo.SetFocus
  Else
   KeyAscii = 0
  End If
End Sub

Private Sub emp_name_KeyPress(KeyAscii As Integer)
  If (((KeyAscii >= 65) And (KeyAscii <= 90)) Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or (KeyAscii = 8) Or (KeyAscii = 32)) Then
        emp_name.SetFocus
  ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        emp_father.SetFocus
  Else
   KeyAscii = 0
  End If
End Sub

Private Sub Form_Load()
CenterForm Me
conn
add_btn.Enabled = True
sv_btn.Enabled = False
update_btn.Enabled = False
dl_btn.Enabled = False
Calendar1.Value = Format(Date, "DD-MMM-YY")
Calendar1.Visible = False
DTPicker1.MaxDate = Date - (20 * 365)
DTPicker1.MinDate = Date - (55 * 365)
DTPicker1.Value = DTPicker1.MaxDate
emp_father.Text = ""
emp_mother.Text = ""
Combo2.Locked = False
mobNo.Text = ""
mobno2.Text = ""
pic_name = ""
Label10.Caption = ""
emp_name.Text = ""
address.Text = ""
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
PinCode.Text = ""
Combo2.Clear
 Combo2.AddItem "Admin"
 Combo2.AddItem "Employee"
'or Checking at start
If Frm_Starting = 1 Then
 Combo2.Clear
 Combo2.AddItem "Admin"
 dl_btn.Enabled = False
 update_btn.Enabled = False
 Clear.Enabled = False
 back_btn.Enabled = True
 Command1.Enabled = False
 add_btn_Click
 add_btn.Enabled = False
 Combo2.Text = "Admin"
ElseIf Frm_Starting = 2 Then
Combo2.Clear
 Combo2.AddItem "Admin"
 Combo2.AddItem "Employee"
 add_btn.Enabled = True
 dl_btn.Enabled = False
 update_btn.Enabled = False
 Clear.Enabled = True
 back_btn.Enabled = True
 Command1.Enabled = True
End If
Combo3.Clear
Combo3.AddItem "Male"
Combo3.AddItem "Female"
Combo3.AddItem "Other"

Combo5.Clear
Combo5.AddItem "I.A."
Combo5.AddItem "I.Com"
Combo5.AddItem "I.Sc."
Combo5.AddItem "B.A."
Combo5.AddItem "B.Com"
Combo5.AddItem "B.Sc"
Combo5.AddItem "M.A."
Combo5.AddItem "M.Com"
Combo5.AddItem "M.Sc"
Combo5.AddItem "Ph.D."
Combo5.AddItem "Other"

Combo4.Clear
Combo4.AddItem "Andhra Pradesh"
Combo4.AddItem "Arunachal Pradesh"
Combo4.AddItem "Assam"
Combo4.AddItem "Bihar"
Combo4.AddItem "Chhattisgarh"
Combo4.AddItem "Goa"
Combo4.AddItem "Gujrat"
Combo4.AddItem "Haryana"
Combo4.AddItem "Himachal Pradesh"
Combo4.AddItem "Jammu and Kashmir"
Combo4.AddItem "Jharkhand"
Combo4.AddItem "Karnataka"
Combo4.AddItem "Kerala"
Combo4.AddItem "Madhya Pradesh"
Combo4.AddItem "Maharashtra"
Combo4.AddItem "Manipur"
Combo4.AddItem "Meghalaya"
Combo4.AddItem "Mizoram"
Combo4.AddItem "Nagaland"
Combo4.AddItem "Odisha"
Combo4.AddItem "Punjab"
Combo4.AddItem "Rajsthan"
Combo4.AddItem "Sikkim"
Combo4.AddItem "Tamil Nadu"
Combo4.AddItem "Telangana"
Combo4.AddItem "Tripura"
Combo4.AddItem "Uttar Pradesh"
Combo4.AddItem "uttarakhand"
Combo4.AddItem "West Bengal"

Combo1.Clear
Image1.Picture = Nothing
Set r1 = New ADODB.Recordset
sql = "select emp_id from emp"
Set r1 = c1.Execute(sql)
While r1.EOF = False
Combo1.AddItem r1.Fields(0)
r1.MoveNext
Wend
Set r = New ADODB.Recordset
Set r = c.Execute("select a_id from adminTbl")
While r.EOF = False
Combo1.AddItem r.Fields(0)
r.MoveNext
Wend
End Sub
Public Function cauto_id()
Dim sql2 As String
Set r1 = New ADODB.Recordset
Set r = New ADODB.Recordset
sql = "select max(to_number(substr(emp_id,2,length(emp_id))))from emp"
Set r1 = c1.Execute(sql)
sql2 = "select max(to_number(substr(a_id,2,length(a_id))))from adminTbl"
Set r = c1.Execute(sql2)
If IsNull(r.Fields(0)) And IsNull(r1.Fields(0)) Then
Label10.Caption = "E00" & 1
Else
If IsNull(r1.Fields(0)) = False And IsNull(r.Fields(0)) Then
 t = r1.Fields(0)
ElseIf IsNull(r1.Fields(0)) And IsNull(r.Fields(0)) = False Then
 t = r.Fields(0)
ElseIf IsNull(r1.Fields(0)) = False And IsNull(r.Fields(0)) = False Then
 If r1.Fields(0) > r.Fields(0) Then
  t = r1.Fields(0)
 Else
  t = r.Fields(0)
 End If
End If
If t > 0 And t < 9 Then
 Label10.Caption = "E00" & (t + 1)
Else
 Label10.Caption = "E0" & (t + 1)
End If
End If
End Function

Private Sub Label8_Click()
Text1.SetFocus
End Sub

Private Sub mobNo_KeyPress(KeyAscii As Integer)
If Len(Trim(mobNo.Text)) = 0 Then
If KeyAscii >= 48 And KeyAscii <= 53 Then
 MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
KeyAscii = 0
mobNo.SetFocus
Exit Sub
End If
End If
If Len(Trim(mobNo.Text)) = 1 Then
 If mobNo.Text = 6 Then
  If KeyAscii <> 50 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   mobNo.SetFocus
  Exit Sub
 End If
End If
End If
If Len(Trim(mobNo.Text)) = 6 Then
 If Right(mobNo.Text, 4) = "0000" Then
  If KeyAscii = 48 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   mobNo.SetFocus
  Exit Sub
 End If
End If
End If
If Len(Trim(mobNo.Text)) = 7 Then
 If Right(mobNo.Text, 5) = "00000" Then
  If KeyAscii = 48 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   mobNo.SetFocus
  Exit Sub
 End If
End If
If Right(mobNo.Text, 5) = "11111" Then
  If KeyAscii = 49 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   mobNo.SetFocus
  Exit Sub
 End If
End If
If Right(mobNo.Text, 5) = "22222" Then
  If KeyAscii = 50 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   mobNo.SetFocus
  Exit Sub
 End If
End If
If Right(mobNo.Text, 5) = "55555" Then
  If KeyAscii = 53 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   mobNo.SetFocus
  Exit Sub
 End If
End If
If Right(mobNo.Text, 5) = "66666" Then
  If KeyAscii = 54 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   mobNo.SetFocus
  Exit Sub
 End If
End If
If Right(mobNo.Text, 5) = "77777" Then
  If KeyAscii = 55 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   mobNo.SetFocus
  Exit Sub
 End If
End If
If Right(mobNo.Text, 5) = "88888" Then
  If KeyAscii = 56 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   mobNo.SetFocus
  Exit Sub
 End If
End If
If Right(mobNo.Text, 5) = "99999" Then
  If KeyAscii = 57 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   mobNo.SetFocus
  Exit Sub
 End If
End If

End If
If Len(Trim(mobNo.Text)) = 8 Then
 If Right(mobNo.Text, 6) = "000000" Then
  If KeyAscii = 48 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   mobNo.SetFocus
  Exit Sub
 End If
End If
 If Right(mobNo.Text, 6) = "111111" Then
  If KeyAscii = 48 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   mobNo.SetFocus
  Exit Sub
 End If
End If
End If

If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Then
        mobNo.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   mobno2.SetFocus
  Else
   KeyAscii = 0
  End If
End Sub

Private Sub mobNo_LostFocus()
If (mobNo.Text <> "") Then
        If (Len(mobNo.Text) < 10) Then
            MsgBox "Invalid MOBILE NUMBER", vbExclamation + vbOKOnly, "Invalid  Mobile No"
            mobNo.Text = ""
            mobNo.SetFocus
        End If
End If
End Sub

Private Sub mobNo2_KeyPress(KeyAscii As Integer)
If Len(Trim(mobno2.Text)) = 0 Then
If KeyAscii >= 48 And KeyAscii <= 53 Then
 MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
KeyAscii = 0
mobno2.SetFocus
Exit Sub
End If
End If
If Len(Trim(mobno2.Text)) = 1 Then
 If mobno2.Text = 6 Then
  If KeyAscii >= 48 And KeyAscii <= 49 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   mobno2.SetFocus
  Exit Sub
 End If
End If
End If
If Len(Trim(mobno2.Text)) = 6 Then
 If Right(mobno2.Text, 4) = "0000" Then
  If KeyAscii = 48 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   mobno2.SetFocus
  Exit Sub
 End If
End If
End If
If Len(Trim(mobno2.Text)) = 7 Then
 If Right(mobno2.Text, 5) = "00000" Then
  If KeyAscii = 48 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   mobno2.SetFocus
  Exit Sub
 End If
End If
If Right(mobno2.Text, 5) = "11111" Then
  If KeyAscii = 49 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   mobno2.SetFocus
  Exit Sub
 End If
End If
If Right(mobno2.Text, 5) = "22222" Then
  If KeyAscii = 50 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   mobno2.SetFocus
  Exit Sub
 End If
End If
If Right(mobno2.Text, 5) = "55555" Then
  If KeyAscii = 53 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   mobno2.SetFocus
  Exit Sub
 End If
End If
If Right(mobno2.Text, 5) = "66666" Then
  If KeyAscii = 54 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   mobno2.SetFocus
  Exit Sub
 End If
End If
If Right(mobno2.Text, 5) = "77777" Then
  If KeyAscii = 55 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   mobno2.SetFocus
  Exit Sub
 End If
End If
If Right(mobno2.Text, 5) = "88888" Then
  If KeyAscii = 56 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   mobno2.SetFocus
  Exit Sub
 End If
End If
If Right(mobno2.Text, 5) = "99999" Then
  If KeyAscii = 57 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   mobno2.SetFocus
  Exit Sub
 End If
End If

End If
If Len(Trim(mobno2.Text)) = 8 Then
 If Right(mobno2.Text, 6) = "000000" Then
  If KeyAscii = 48 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   mobno2.SetFocus
  Exit Sub
 End If
End If
 If Right(mobno2.Text, 6) = "111111" Then
  If KeyAscii = 48 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   mobno2.SetFocus
  Exit Sub
 End If
End If
End If
If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Then
        mobno2.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   address.SetFocus
  Else
   KeyAscii = 0
  End If
End Sub

Private Sub mobNo2_LostFocus()
If (mobno2.Text <> "") Then
        If (Len(mobno2.Text) < 10) Then
            MsgBox "Invalid MOBILE NUMBER", vbExclamation + vbOKOnly, "Invalid  Mobile No"
            mobno2.SetFocus
        End If
End If
End Sub
Private Sub PinCode_KeyPress(KeyAscii As Integer)
If Len(Trim(PinCode.Text)) = 0 Then
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 32 Then
Else
 MsgBox "Invalid Pin Code !!", vbInformation + vbOKOnly, "Pin Code"
 KeyAscii = 0
 PinCode.SetFocus
Exit Sub
End If
End If
If Len(Trim(PinCode.Text)) = 0 Then
  If KeyAscii = 56 Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 32 Then
   Else
   MsgBox "Invalid PIN Code", vbInformation + vbOKOnly, "Pin Code"
   KeyAscii = 0
   PinCode.SetFocus
  Exit Sub
 End If
End If
If Len(Trim(PinCode.Text)) = 1 Then
  If KeyAscii <> 48 Then
   MsgBox "Invalid PIN Code", vbInformation + vbOKOnly, "Pin Code"
   KeyAscii = 0
   PinCode.SetFocus
  Exit Sub
 End If
End If

If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Then
        PinCode.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   Text2.SetFocus
  Else
   KeyAscii = 0
  End If
End Sub

Private Sub PinCode_LostFocus()
If (Trim(PinCode.Text) <> "") Then
        If (Len(PinCode.Text) < 6) Then
            MsgBox "Invalid PIN Code !!", vbExclamation + vbOKOnly, "Invalid PIN Code"
            PinCode.SetFocus
        End If
End If
End Sub

Private Sub sv_btn_Click()
If Trim(Label10.Caption) = "" Then
MsgBox "User ID Blank", vbCritical + vbOKOnly, "Warning"
Exit Sub
ElseIf Trim(emp_name.Text) = "" Then
MsgBox "Name Cannot Be Blank", vbCritical + vbOKOnly, "Warning"
emp_name.SetFocus
Exit Sub
ElseIf Trim(emp_father.Text) = "" Then
MsgBox "Father Name is required..", vbCritical + vbOKOnly, "Warning"
emp_father.SetFocus
Exit Sub
ElseIf Trim(emp_mother.Text) = "" Then
MsgBox "Mother Name is Required..", vbCritical + vbOKOnly, "Warning"
emp_mother.SetFocus
Exit Sub
ElseIf Trim(address.Text) = "" Then
 MsgBox "Enter User Address", vbCritical + vbOKOnly, "Warning"
address.SetFocus
Exit Sub
ElseIf Trim(mobNo.Text) = "" Then
 MsgBox "Enter Mobile No", vbCritical + vbOKOnly, "Warning"
mobNo.SetFocus
Exit Sub
ElseIf Trim(Combo3.Text) = "" Then
 MsgBox "Gender field cann't be blank", vbCritical + vbOKOnly, "Warning"
Combo3.SetFocus
Exit Sub
ElseIf Trim(Text1.Text) = "" Then
 MsgBox "Enter Salary", vbCritical + vbOKOnly, "Warning"
Text1.SetFocus
Exit Sub
ElseIf Trim(Text2.Text) = "" Then
 MsgBox "Adhar No is Mandatory", vbCritical + vbOKOnly, "Warning"
Text2.SetFocus
Exit Sub
ElseIf Trim(Text3.Text) = "" Then
 MsgBox "Enter Email ID", vbCritical + vbOKOnly, "Warning"
Text3.SetFocus
Exit Sub
ElseIf Trim(Text4.Text) = "" Then
 MsgBox "Enter Joining Date", vbCritical + vbOKOnly, "Warning"
Text4.SetFocus
Exit Sub
ElseIf Trim(Combo5.Text) = "" Then
MsgBox "Enter Qualification", vbCritical + vbOKOnly, "Warning"
Combo5.SetFocus
Exit Sub
ElseIf Trim(Combo4.Text) = "" Then
MsgBox "Enter State from where User belong To..", vbCritical + vbOKOnly, "Warning"
Combo4.SetFocus
Exit Sub
ElseIf Trim(PinCode.Text) = "" Then
MsgBox "PinCode Field cann't Blank ???", vbCritical + vbOKOnly, "Warning"
PinCode.SetFocus
Exit Sub
ElseIf pic_name = "" Then
MsgBox "Photo is Required, Browse Photo", vbCritical + vbOKOnly, "Warning"
upload.SetFocus
Exit Sub
ElseIf Combo2.Text = "" Then
MsgBox "Select User Role for system", vbCritical + vbOKOnly, "Warning"
Combo2.SetFocus
Exit Sub
End If
If UCase(Combo2.Text) = "ADMIN" Then 'If Admin is hired
 If Trim(mobno2.Text) = "" Then
  sql = "insert into AdminTbl values ('" & Label10.Caption & "','" & UCase(emp_name.Text) & "','" & UCase(emp_father.Text) & "','" & UCase(emp_mother.Text) & "','" & Trim(address.Text) & "','" & Combo4.Text & "'," & mobNo.Text & ",NULL,'" & Format(DTPicker1.Value, "dd-mmm-yyyy") & "','" & Combo3.Text & "'," & Text1.Text & ",'" & Text2.Text & "','" & Text3.Text & "'," & PinCode.Text & ",'" & Format(Text4.Text, "dd-mmm-yyyy") & "','" & Combo5.Text & "','" & pic_name & "')"
 Else
  sql = "insert into AdminTbl values ('" & Label10.Caption & "','" & UCase(emp_name.Text) & "','" & UCase(emp_father.Text) & "','" & UCase(emp_mother.Text) & "','" & Trim(address.Text) & "','" & Combo4.Text & "'," & mobNo.Text & "," & mobno2.Text & ",'" & Format(DTPicker1.Value, "dd-mmm-yyyy") & "','" & Combo3.Text & "'," & Text1.Text & ",'" & Text2.Text & "','" & Text3.Text & "'," & PinCode.Text & ",'" & Format(Text4.Text, "dd-mmm-yyyy") & "','" & Combo5.Text & "','" & pic_name & "')"
 End If
  c1.Execute (sql)
  Adodc2.Refresh
  emp_id_pass.id.Caption = Label10.Caption
  emp_id_pass.log_id.Caption = UCase(Trim(Mid$(emp_name.Text, 1, 4))) & DTPicker1.Year & Label10.Caption
  emp_id_pass.Password.Caption = Label10.Caption & DTPicker1.Day & "Admin@STS"
Else 'Standard User
 If Trim(mobno2.Text) = "" Then
  sql = "insert into emp values ('" & Label10.Caption & "','" & UCase(emp_name.Text) & "','" & UCase(emp_father.Text) & "','" & UCase(emp_mother.Text) & "','" & Trim(address.Text) & "','" & Combo4.Text & "'," & mobNo.Text & ",NULL,'" & Format(DTPicker1.Value, "dd-mmm-yyyy") & "','" & Combo3.Text & "'," & Text1.Text & ",'" & Text2.Text & "','" & Text3.Text & "'," & PinCode.Text & ",'" & Format(Text4.Text, "dd-mmm-yyyy") & "','" & Combo5.Text & "','" & pic_name & "')"
 Else
  sql = "insert into emp values ('" & Label10.Caption & "','" & UCase(emp_name.Text) & "','" & UCase(emp_father.Text) & "','" & UCase(emp_mother.Text) & "','" & Trim(address.Text) & "','" & Combo4.Text & "'," & mobNo.Text & "," & mobno2.Text & ",'" & Format(DTPicker1.Value, "dd-mmm-yyyy") & "','" & Combo3.Text & "'," & Text1.Text & ",'" & Text2.Text & "','" & Text3.Text & "'," & PinCode.Text & ",'" & Format(Text4.Text, "dd-mmm-yyyy") & "','" & Combo5.Text & "','" & pic_name & "')"
 End If
  c1.Execute (sql)
  Adodc1.Refresh
  emp_id_pass.id.Caption = Label10.Caption
  emp_id_pass.log_id.Caption = UCase(Trim(Mid$(emp_name.Text, 1, 4))) & DTPicker1.Year & Label10.Caption
  emp_id_pass.Password.Caption = Label10.Caption & DTPicker1.Day & "@STS"
End If
emp_id_pass.Role.Caption = Combo2.Text
emp_id_pass.Show vbModal
If Frm_Starting = 1 Then
 Unload Me
Exit Sub
End If
Form_Load
End Sub

Private Sub Tb1_Click()
If Trim(Tb1.SelectedItem.Caption) = "Standard User" Then
    db1.Visible = True
    Db2.Visible = False
Else
    db1.Visible = False
    Db2.Visible = True
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Then
        Text1.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   Text4.SetFocus
  Else
   KeyAscii = 0
  End If
End Sub

Private Sub Text1_LostFocus()
If Len(Trim(Text1.Text)) <> 0 Then
 If Val(Text1.Text) <= 1000 Or Val(Text1.Text) >= 50000 Then
  MsgBox "Salary range Can Be Between 1000 to 50000", vbInformation + vbOKOnly, "Salary range"
  Text1.SetFocus
  End If
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Then
     Text2.SetFocus
    ElseIf KeyAscii = 13 Then
     KeyAscii = 0
     Text3.SetFocus
    Else
     KeyAscii = 0
   End If
End Sub

Private Sub Text2_LostFocus()
If (Text2.Text <> "") Then
        If (Len(Text2.Text) < 12) Then
            MsgBox "Invalid AADHAR number", vbCritical + vbOKOnly, "Invalid Adhar"
            Text2.Text = ""
            Text2.SetFocus
        End If
    End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
If Len(Trim(Text3.Text)) = 0 Then
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 32 Then
Else
 MsgBox "Email Id Must start With Character!!", vbInformation + vbOKOnly, "Email"
 KeyAscii = 0
 Text3.SetFocus
Exit Sub
End If
End If
If InStr(Text3.Text, "@") = False Then
 If KeyAscii = 95 Or KeyAscii = 46 Or KeyAscii = 64 Or (((KeyAscii >= 65) And (KeyAscii <= 90)) Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or (KeyAscii = 8) Or (KeyAscii = 32)) Or (KeyAscii >= 48 And KeyAscii <= 57) Then
   Text3.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   Text1.SetFocus
  Else
   KeyAscii = 0
  End If
Else
  If KeyAscii = 95 Or KeyAscii = 46 Or (((KeyAscii >= 65) And (KeyAscii <= 90)) Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or (KeyAscii = 8) Or (KeyAscii = 32)) Or (KeyAscii >= 48 And KeyAscii <= 57) Then
   Text3.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   Text1.SetFocus
  Else
   KeyAscii = 0
  End If
End If
End Sub

Private Sub Text3_LostFocus() 'check It Contains @ or not
Dim domain As String
If Len(Trim(Text3.Text)) <> 0 Then
 If Len(Trim(Text3.Text)) <= 12 Then
 MsgBox "Invalid Email, Too Short Email", vbCritical + vbOKOnly, "Email"
Text3.SetFocus
Exit Sub
End If
If InStr(Text3.Text, "@") = False Then
 MsgBox "Invalid Email, It Must contain @..", vbCritical + vbOKOnly, "Email"
 Text3.SetFocus
Exit Sub
End If
domain = Right(Text3.Text, 4)
If UCase(domain) = UCase(".COM") Or UCase(domain) = UCase(".NET") Then
Exit Sub
Else
 MsgBox "Invalid Email", vbCritical + vbOKOnly, "Email"
 Text3.SetFocus
 Exit Sub
 End If
domain = Right(Text3.Text, 3)
If UCase(domain) = UCase(".TK") Or UCase(domain) = UCase(".IN") Then
Exit Sub
Else
 MsgBox "Invalid Email", vbCritical + vbOKOnly, "Email"
 Text3.SetFocus
 Exit Sub
End If
 End If
End Sub

Private Sub Text4_GotFocus()
Text4.Text = ""
Calendar1.Visible = True
End Sub

Private Sub Text4_LostFocus()
Calendar1.Visible = False
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   Combo5.SetFocus
End If
End Sub

Private Sub update_btn_Click()
conn
If Trim(Combo1.Text) = "" Then
 MsgBox "Select Corrrect User ID Or Name..", vbCritical + vbOKOnly, "Update ERROR"
 Combo1.SetFocus
 Exit Sub
 End If
 If Trim(emp_name.Text) = "" Then
MsgBox "Employee Name Blank", vbCritical + vbOKOnly, "Warning"
emp_name.SetFocus
Exit Sub
ElseIf Trim(emp_father.Text) = "" Then
MsgBox "Employee father Name Blank", vbCritical + vbOKOnly, "Warning"
emp_father.SetFocus
Exit Sub
ElseIf Trim(emp_mother.Text) = "" Then
MsgBox "Employee Mother Name Blank", vbCritical + vbOKOnly, "Warning"
emp_mother.SetFocus
Exit Sub
ElseIf Trim(address.Text) = "" Then
 MsgBox "Enter Employee Address", vbCritical + vbOKOnly, "Warning"
address.SetFocus
Exit Sub
ElseIf Trim(mobNo.Text) = "" Then
 MsgBox "Enter Mobile No", vbCritical + vbOKOnly, "Warning"
mobNo.SetFocus
Exit Sub
ElseIf Trim(Combo3.Text) = "" Then
 MsgBox "Gender field cann't be blank", vbCritical + vbOKOnly, "Warning"
Combo3.SetFocus
Exit Sub
ElseIf Trim(Text1.Text) = "" Then
 MsgBox "Enter Salary", vbCritical + vbOKOnly, "Warning"
Text1.SetFocus
Exit Sub
ElseIf Trim(Text2.Text) = "" Then
 MsgBox "Adhar No is Mandatory", vbCritical + vbOKOnly, "Warning"
Text2.SetFocus
Exit Sub
ElseIf Trim(Text3.Text) = "" Then
 MsgBox "Enter Email ID", vbCritical + vbOKOnly, "Warning"
Text3.SetFocus
Exit Sub
ElseIf Trim(Text4.Text) = "" Then
 MsgBox "Enter Joining Date", vbCritical + vbOKOnly, "Warning"
Text4.SetFocus
Exit Sub
ElseIf Trim(Combo5.Text) = "" Then
MsgBox "Enter Qualification", vbCritical + vbOKOnly, "Warning"
Combo5.SetFocus
Exit Sub
ElseIf Trim(Combo4.Text) = "" Then
MsgBox "Enter State from where User belong To..", vbCritical + vbOKOnly, "Warning"
Combo4.SetFocus
Exit Sub
ElseIf Trim(PinCode.Text) = "" Then
MsgBox "PinCode Field cann't Blank ???", vbCritical + vbOKOnly, "Warning"
PinCode.SetFocus
Exit Sub
ElseIf Combo2.Text = "" Then
MsgBox "Select User Role for system", vbCritical + vbOKOnly, "Warning"
Combo2.SetFocus
Exit Sub
End If
conn
opt = MsgBox("Are You Sure to Update Record ?", vbQuestion + vbYesNo, "UPDATE")
If opt = vbYes Then
  If UCase(Combo2.Text) = UCase("Admin") Then
   If Trim(mobno2.Text) = "" Then
    sql = " Update adminTbl set a_nm='" & emp_name.Text & "',a_father='" & emp_father.Text & "',a_mother='" & emp_mother.Text & "',a_add='" & Trim(address.Text) & "',a_state='" & Combo4.Text & "',a_mob=" & mobNo.Text & ",a_dob='" & Format(DTPicker1.Value, "dd-mmm-yyyy") & "',a_sal=" & Text1.Text & ",a_adhr='" & Text2.Text & "',a_email='" & Text3.Text & "',a_mob2=NULL,a_pincd='" & PinCode.Text & "',a_qualif='" & Combo5.Text & "',a_pic='" & pic_name & "' where a_id='" & Label10.Caption & "'"
   Else
    sql = " Update adminTbl set a_nm='" & emp_name.Text & "',a_father='" & emp_father.Text & "',a_mother='" & emp_mother.Text & "',a_add='" & Trim(address.Text) & "',a_state='" & Combo4.Text & "',a_mob=" & mobNo.Text & ",a_dob='" & Format(DTPicker1.Value, "dd-mmm-yyyy") & "',a_sal=" & Text1.Text & ",a_adhr='" & Text2.Text & "',a_email='" & Text3.Text & "',a_mob2=" & mobno2.Text & ",a_pincd='" & PinCode.Text & "',a_qualif='" & Combo5.Text & "',a_pic='" & pic_name & "' where a_id='" & Label10.Caption & "'"
   End If
  Else
   If Trim(mobno2.Text) = "" Then
    sql = " Update emp set e_nm='" & emp_name.Text & "',e_father='" & emp_father.Text & "',e_mother='" & emp_mother.Text & "',e_add='" & Trim(address.Text) & "',e_state='" & Combo4.Text & "',e_mob=" & mobNo.Text & ",e_dob='" & Format(DTPicker1.Value, "dd-mmm-yyyy") & "',e_sal=" & Text1.Text & ",e_adhr='" & Text2.Text & "',e_email='" & Text3.Text & "',e_mob2=NULL,e_pincd='" & PinCode.Text & "',e_qualif='" & Combo5.Text & "',e_pic='" & pic_name & "' where emp_id='" & Label10.Caption & "'"
   Else
    sql = " Update emp set e_nm='" & emp_name.Text & "',e_father='" & emp_father.Text & "',e_mother='" & emp_mother.Text & "',e_add='" & Trim(address.Text) & "',e_state='" & Combo4.Text & "',e_mob=" & mobNo.Text & ",e_dob='" & Format(DTPicker1.Value, "dd-mmm-yyyy") & "',e_sal=" & Text1.Text & ",e_adhr='" & Text2.Text & "',e_email='" & Text3.Text & "',e_mob2=" & mobno2.Text & ",e_pincd='" & PinCode.Text & "',e_qualif='" & Combo5.Text & "',e_pic='" & pic_name & "' where emp_id='" & Label10.Caption & "'"
   End If
  End If
End If
c1.Execute (sql)
MsgBox "Record SuccessFully Updated !!!", vbInformation + vbOKOnly, "Record Update"
Adodc1.Refresh
Adodc2.Refresh
Form_Load
End Sub

Private Sub upload_Click()
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
Else
Exit Sub
End If
End Sub
