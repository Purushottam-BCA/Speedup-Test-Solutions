VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmTopicMaster 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Topic Master"
   ClientHeight    =   10770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
   Icon            =   "Topic.frx":0000
   LinkTopic       =   "Form8"
   MDIChild        =   -1  'True
   ScaleHeight     =   10770
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Topic.frx":0ECA
      Height          =   9885
      Left            =   8640
      TabIndex        =   13
      Top             =   495
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   17436
      _Version        =   393216
      AllowUpdate     =   -1  'True
      ColumnHeaders   =   0   'False
      HeadLines       =   0
      RowHeight       =   22
      FormatLocked    =   -1  'True
      AllowDelete     =   -1  'True
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "TP_ID"
         Caption         =   "TP_ID"
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
         DataField       =   "TP_NM"
         Caption         =   "TP_NM"
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
         DataField       =   "TP_DUR"
         Caption         =   "TP_DUR"
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
      BeginProperty Column04 
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
      SplitCount      =   1
      BeginProperty Split0 
         ScrollBars      =   2
         Locked          =   -1  'True
         RecordSelectors =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   5490.142
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   945.071
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1860.095
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1980.284
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2520
      Top             =   8040
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
      Connect         =   "Provider=MSDAORA.1;User ID=sts/sts;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=sts/sts;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"Topic.frx":0EDF
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   4335
      Left            =   600
      TabIndex        =   0
      Top             =   2280
      Width           =   7335
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2640
         MouseIcon       =   "Topic.frx":0F73
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Left            =   2645
         MaxLength       =   2
         TabIndex        =   4
         Top             =   3200
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2640
         MouseIcon       =   "Topic.frx":10C5
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   360
         Width           =   2535
      End
      Begin VB.ComboBox Combo3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2640
         MouseIcon       =   "Topic.frx":1217
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   1035
         Width           =   2535
      End
      Begin VB.TextBox Text1 
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
         Height          =   315
         Left            =   2660
         TabIndex        =   1
         Top             =   2440
         Width           =   4095
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   $"Topic.frx":1369
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   450
         Left            =   120
         TabIndex        =   44
         Top             =   3840
         Width           =   7215
      End
      Begin VB.Label lblTopic 
         BackColor       =   &H0000FFFF&
         Height          =   375
         Left            =   5520
         TabIndex        =   24
         Top             =   1320
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblSubject 
         BackColor       =   &H0000FFFF&
         Height          =   375
         Left            =   5520
         TabIndex        =   23
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblCourse 
         BackColor       =   &H0000FFFF&
         Height          =   375
         Left            =   5520
         TabIndex        =   22
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Topic ID"
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
         Left            =   800
         TabIndex        =   12
         Top             =   1725
         Width           =   825
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Classes"
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
         Left            =   3845
         TabIndex        =   9
         Top             =   3225
         Width           =   765
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Duration"
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
         Left            =   800
         TabIndex        =   8
         Top             =   3270
         Width           =   915
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000006&
         Height          =   375
         Left            =   2640
         Shape           =   4  'Rounded Rectangle
         Top             =   3195
         Width           =   2055
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Course "
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
         Left            =   800
         TabIndex        =   7
         Top             =   360
         Width           =   780
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Subject"
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
         Left            =   800
         TabIndex        =   6
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Topic Name"
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
         Left            =   800
         TabIndex        =   5
         Top             =   2475
         Width           =   1230
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000006&
         Height          =   375
         Left            =   2640
         Shape           =   4  'Rounded Rectangle
         Top             =   2400
         Width           =   4215
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   4335
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   8535
      Begin VB.Frame Frame6 
         Caption         =   "Frame6"
         Height          =   2055
         Left            =   2280
         TabIndex        =   25
         Top             =   2280
         Width           =   135
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "( TOPIC )"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3600
         TabIndex        =   21
         Top             =   1320
         Width           =   1635
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
         Height          =   420
         Left            =   1800
         TabIndex        =   20
         Top             =   1305
         Width           =   5370
      End
      Begin VB.Image Image1 
         Height          =   1740
         Left            =   0
         Picture         =   "Topic.frx":1414
         Stretch         =   -1  'True
         Top             =   0
         Width           =   8610
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   5775
      Left            =   0
      TabIndex        =   26
      Top             =   4320
      Width           =   8535
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   960
         Left            =   615
         TabIndex        =   31
         Top             =   3000
         Width           =   7300
         Begin VB.CommandButton addbtn 
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
            Height          =   375
            Left            =   240
            MouseIcon       =   "Topic.frx":A53A
            MousePointer    =   99  'Custom
            TabIndex        =   43
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cancelbtn 
            Caption         =   "Cancel"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            MouseIcon       =   "Topic.frx":A68C
            MousePointer    =   99  'Custom
            TabIndex        =   32
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton backbtn 
            BackColor       =   &H00FFC0FF&
            Caption         =   "Exit"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6000
            MouseIcon       =   "Topic.frx":A7DE
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Click To Exit"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton updatebtn 
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
            Height          =   375
            Left            =   4560
            MouseIcon       =   "Topic.frx":A930
            MousePointer    =   99  'Custom
            TabIndex        =   35
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton savebtn 
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
            Height          =   375
            Left            =   1680
            MouseIcon       =   "Topic.frx":AA82
            MousePointer    =   99  'Custom
            TabIndex        =   34
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton deletebtn 
            Caption         =   "Delete"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3120
            MouseIcon       =   "Topic.frx":ABD4
            MousePointer    =   99  'Custom
            TabIndex        =   33
            Top             =   240
            Width           =   1095
         End
         Begin VB.PictureBox add 
            BackColor       =   &H00404040&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   360
            ScaleHeight     =   315
            ScaleWidth      =   1005
            TabIndex        =   42
            Top             =   315
            Width           =   1060
         End
         Begin VB.PictureBox xpButton4 
            BackColor       =   &H00404040&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   6120
            ScaleHeight     =   315
            ScaleWidth      =   1005
            TabIndex        =   41
            Top             =   315
            Width           =   1065
         End
         Begin VB.PictureBox xpButton3 
            BackColor       =   &H00404040&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   4680
            ScaleHeight     =   315
            ScaleWidth      =   1005
            TabIndex        =   40
            Top             =   315
            Width           =   1065
         End
         Begin VB.PictureBox xpButton2 
            BackColor       =   &H00404040&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   3240
            ScaleHeight     =   315
            ScaleWidth      =   1005
            TabIndex        =   39
            Top             =   315
            Width           =   1065
         End
         Begin VB.PictureBox xpButton1 
            BackColor       =   &H00404040&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   1800
            ScaleHeight     =   315
            ScaleWidth      =   1005
            TabIndex        =   38
            Top             =   315
            Width           =   1065
         End
         Begin VB.PictureBox vkCommand9 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   11475
            ScaleHeight     =   435
            ScaleWidth      =   1395
            TabIndex        =   37
            Top             =   720
            Width           =   1455
         End
      End
      Begin VB.Frame Frame5 
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
         Height          =   1125
         Left            =   600
         TabIndex        =   27
         Top             =   4515
         Width           =   7335
         Begin VB.CommandButton Command4 
            Caption         =   "Search"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3885
            MouseIcon       =   "Topic.frx":AD26
            MousePointer    =   99  'Custom
            TabIndex        =   30
            Top             =   415
            Width           =   1455
         End
         Begin VB.PictureBox Picture1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   11475
            ScaleHeight     =   435
            ScaleWidth      =   1395
            TabIndex        =   29
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   28
            Top             =   415
            Width           =   3375
         End
      End
      Begin VB.Shape Shape1 
         Height          =   1500
         Left            =   600
         Top             =   2760
         Width           =   7335
      End
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   20385
      X2              =   20385
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   8650
      X2              =   8650
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   18200
      X2              =   18200
      Y1              =   0
      Y2              =   520
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   16320
      X2              =   16320
      Y1              =   0
      Y2              =   520
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   15375
      X2              =   15375
      Y1              =   0
      Y2              =   520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   9880
      X2              =   9880
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   18720
      TabIndex        =   19
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Course"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   16920
      TabIndex        =   18
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Session"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   15435
      TabIndex        =   17
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Topic Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   10680
      TabIndex        =   16
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Topic ID"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8880
      TabIndex        =   15
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000011&
      Height          =   495
      Left            =   8640
      TabIndex        =   14
      Top             =   0
      Width           =   11775
   End
End
Attribute VB_Name = "FrmTopicMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim t As Integer
Dim opt As String

Private Sub Combo1_Click()
Set r1 = New ADODB.Recordset
sql = " select * from topic where tp_id='" & Combo1.Text & "'"
Set r1 = c1.Execute(sql)
If IsNull(r1.Fields(0)) = False Then
 lblTopic.Caption = r1.Fields(0)
 Text1.Text = r1.Fields(1)
 Text2.Text = r1.Fields(2)
 lblSubject.Caption = r1.Fields(3)
 lblCourse.Caption = r1.Fields(4)
 Set r = New ADODB.Recordset
 Set r = c.Execute("select c_nm from course where c_id='" & lblCourse.Caption & "' ")
 Combo2.Text = r.Fields(0)
 Set r = c.Execute("select sub_nm from sub where sub_id='" & lblSubject.Caption & "' ")
 Combo3.Text = r.Fields(0)
 End If
 updatebtn.Enabled = True
 deletebtn.Enabled = True
 savebtn.Enabled = False
 addbtn.Enabled = False
 cancelbtn.Enabled = True
 cancelbtn.Visible = True
 addbtn.Visible = False
End Sub

Private Sub combo2_Click()
Set r1 = New ADODB.Recordset
Set r1 = c1.Execute("select c_id from course where c_nm='" & Combo2.Text & "' ")
lblCourse.Caption = r1.Fields(0)

sql = "select * from sub where c_id='" & lblCourse.Caption & "' "
Set r1 = c1.Execute(sql)
 Combo3.Enabled = True
 Combo3.Clear
While (r1.EOF) = False
 Combo3.AddItem r1.Fields(2)
 r1.MoveNext
Wend
 Combo3.SetFocus
End Sub

Private Sub combo3_Click()
Set r1 = New ADODB.Recordset
Set r1 = c1.Execute("select sub_id from sub where sub_nm='" & Combo3.Text & "' and c_id='" & lblCourse.Caption & "' ")
lblSubject.Caption = r1.Fields(0)
Text1.SetFocus
End Sub


Private Sub addbtn_Click()
Tpicauto_id
Combo2.Text = ""
Combo3.Text = ""
Text1.Text = ""
Text2.Text = 0
Text2.Locked = True
Combo1.Enabled = False
 savebtn.Enabled = True
 cancelbtn.Visible = True
 addbtn.Visible = False
 Combo2.Enabled = True
 Combo2.SetFocus
End Sub

Private Sub cancelbtn_Click()
Form_Load
End Sub
'
Private Sub backbtn_Click()
Unload Me
End Sub

'Private Sub combo1_Click()
'Set r1 = New ADODB.Recordset
'sql = "select * from topic where trim(tp_id) = trim('" & Combo1.Text & "') "
'Set r1 = c1.Execute(sql)
'If IsNull(r1.Fields(0)) = False Then
'Text1.Text = r1.Fields(1)
'Text2.Text = r1.Fields(2)
'End If
'End Sub
'
'Private Sub combo3_Click() 'Searching Option for Subject
'Set r1 = New ADODB.Recordset
'sql = "select * from topic where trim(sub_id)=trim('" & Combo3.Text & "')and trim(c_id)=trim('" & Combo2.Text & "') "
'Set r1 = c1.Execute(sql)
'If IsNull(r1.Fields(0)) = False Then
' Combo1.Enabled = True
' 'combo1.clear
' While r1.EOF = False
' Combo1.AddItem Space(25) & r1.Fields(0)    '& Space(10) & r1.Fields(1)
' r1.MoveNext
'Wend
''Else
''MsgBox "Invalid Course ID", vbOKOnly, " "
''Combo1.SetFocus
'End If
'End Sub
'Private Sub combo2_Click() 'Searching Option Course
'Set r1 = New ADODB.Recordset
'sql = "select * from sub where trim(c_id)=trim('" & Combo2.Text & "')"
'Set r1 = c1.Execute(sql)
'Combo3.Enabled = True
'Combo3.Clear
'If IsNull(r1.Fields(0)) = False Then
' While r1.EOF = False
' Combo3.AddItem Space(25) & r1.Fields(0) '& Space(5) & r1.Fields(1)
' r1.MoveNext
' Wend
'' Adodc1.RecordSource = "select * from course where trim(c_id)=trim('" & Combo1.Text & "')"
'' Adodc1.Refresh
'Else
'MsgBox "Invalid Course ID", vbOKOnly, " "
'Combo2.SetFocus
'End If
'End Sub
'
'Private Sub Form_Load()
'CenterForm Me
'conn
'Combo1.Enabled = False
'Combo3.Enabled = False
'cancelbtn.Enabled = False
'save.Enabled = False
'delete.Enabled = False
'update.Enabled = False
'data_in_combo2 'call Course
'End Sub
'Public Function data_in_combo1() 'For Topic ID
'Combo1.Clear
'Set r1 = New ADODB.Recordset
'sql = "select * from topic"
'Set r1 = c1.Execute(sql)
'While r1.EOF = False
''Text1.Text = r1.Fields(1)
''Text2.Text = r1.Fields(2)
' Combo1.AddItem Space(25) & r1.Fields(0) '& Space(5) & r1.Fields(1)
' r1.MoveNext
'Wend
'End Function
'
'Public Function data_in_combo2() 'For Course Id
'Set r1 = New ADODB.Recordset
'Combo2.Clear
'sql = "select * from course"
'Set r1 = c1.Execute(sql)
'While r1.EOF = False
' Combo2.AddItem Space(25) & r1.Fields(0)    '& Space(10) & r1.Fields(1)
' r1.MoveNext
'Wend
'End Function
'Public Function data_in_combo3() 'For Sub ID
'Combo3.Clear
'Set r1 = New ADODB.Recordset
'sql = "select * from sub"
'Set r1 = c1.Execute(sql)
'While r1.EOF = False
' Combo1.AddItem Space(25) & r1.Fields(0) '& Space(5) & r1.Fields(1)
' r1.MoveNext
'Wend
'End Function
'
Public Function Tpicauto_id()
Set r1 = New ADODB.Recordset
sql = "select MAX(to_number(substr(tp_id,2,length(tp_id))))from topic"
Set r1 = c1.Execute(sql)
If IsNull(r1.Fields(0)) Then
 Combo1.Text = "T00" & 1
Else
 t = r1.Fields(0)
 If t > 0 And t < 9 Then
  Combo1.Text = "T00" & (t + 1)
 ElseIf t < 99 Then
  Combo1.Text = "T0" & (t + 1)
 Else
  Combo1.Text = "T" & (t + 1)
End If
End If
Combo1.Locked = True
End Function

Private Sub Command4_Click()
If Trim(Text3.Text) <> "" Then
Set r = New ADODB.Recordset
Set r = c.Execute("select * from topic where upper(tp_nm)='" & UCase(Trim(Text3.Text)) & "' or upper(tp_id)='" & Trim(UCase(Text3.Text)) & "' ")
If r.EOF = False Then
 Combo1.Text = r.Fields(0)
 Text1.Text = r.Fields(1)
 Text2.Text = r.Fields(2)
 lblCourse.Caption = r.Fields(4)
 lblSubject.Caption = r.Fields(3)
 Set r2 = c.Execute("select initcap(c_nm) from course where c_id='" & lblCourse.Caption & "' ")
  Combo2.Text = r2.Fields(0)
 Set r3 = c.Execute("select initcap(sub_nm) from sub where sub_id='" & lblSubject.Caption & "' ")
  Combo3.Text = r3.Fields(0)
  lblCourse.Caption = ""
  lblSubject.Caption = ""
  Combo2.Enabled = True
  Combo3.Enabled = True
  Combo1.Locked = True
  deletebtn.Enabled = True
  updatebtn.Enabled = True
  savebtn.Enabled = False
  addbtn.Visible = False
  cancelbtn.Enabled = True
  cancelbtn.Visible = True
Else
 MsgBox "Topic Not Found..", vbQuestion + vbOKOnly, "Not Found"
 Exit Sub
End If
Else
End If
End Sub

Private Sub savebtn_Click() 'Save
Set r1 = New ADODB.Recordset
 If Trim(Combo2.Text) = "" Then
 MsgBox " Select Course ", vbCritical + vbOKOnly, "Warning"
 Combo2.SetFocus
ElseIf Trim(Combo3.Text) = "" Then
 MsgBox " Select Subject ", vbCritical + vbOKOnly, "Warning"
 Combo3.SetFocus
ElseIf Trim(Text1.Text) = "" Then
 MsgBox "Enter Chapter Name", vbCritical + vbOKOnly, "Warning"
Text1.SetFocus
ElseIf Trim(Text2.Text) = "" Then
 MsgBox "Enter Chapter duration", vbCritical + vbOKOnly, "Warning"
Text2.SetFocus
Else
 Set r = New ADODB.Recordset
 Set r = c.Execute("select * from topic")
 While r.EOF = False
  If UCase(r.Fields(1)) = UCase(Trim(Text1.Text)) And UCase(r.Fields(3)) = UCase(Trim(lblSubject.Caption)) And UCase(r.Fields(4)) = UCase(Trim(lblCourse.Caption)) Then
   MsgBox "Topic Already Exist..", vbCritical + vbOKOnly, "Duplicate Topic"
   Exit Sub
  End If
 r.MoveNext
 Wend
 sql = "insert into topic values ('" & Combo1.Text & "','" & Text1.Text & "'," & Text2.Text & ",'" & lblSubject.Caption & "','" & lblCourse.Caption & "')"
 Set r1 = c1.Execute(sql)
 MsgBox "Topic Successfully added", vbApplicationModal + vbInformation + vbOKOnly, ""
 Adodc1.Refresh
 Form_Load
End If
End Sub

Private Sub deletebtn_Click() 'Delete
If Trim(Combo2.Text) = "" Or Trim(Combo3.Text) = "" Or Trim(Combo1.Text) = "" Then
 MsgBox "Select Corrrect topic", vbCritical + vbOKOnly, "Delete ERROR"
Else
Set r1 = New ADODB.Recordset
opt = MsgBox("Are You Sure to Delete ?", vbQuestion + vbYesNo, "Delete conformation!")
If opt = vbYes Then
 sql = " delete from topic where tp_id='" & lblTopic.Caption & "'"
 c1.Execute (sql)
 MsgBox "Topic Successfully Deleted!!", vbInformation + vbOKOnly, "Delete Topic !"
 Adodc1.Refresh  'DataGrid Updated
 Form_Load
 Else
End If
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 32 Or KeyAscii = 8 Then
 Text2.SetFocus
ElseIf KeyAscii = 13 Then
 KeyAscii = 0
Else
KeyAscii = 0
End If
End Sub

Private Sub updatebtn_Click()
 If Trim(Combo2.Text) = "" Or Trim(Combo3.Text) = "" Or Trim(Combo1.Text) = "" Then
 MsgBox "Select Corrrect topic", vbCritical + vbOKOnly, "Update ERROR"
 Else
  conn
  opt = MsgBox("Are You Sure to Update ?", vbQuestion + vbYesNo, "UPDATE")
   If opt = vbYes Then
    sql = " update topic set tp_nm='" & Text1.Text & "',tp_dur=" & Text2.Text & " where tp_id='" & lblTopic.Caption & "'"
    c1.Execute (sql)
    MsgBox "topic Successfully Updated!!", vbInformation + vbOKOnly, "Update Topic !"
    Adodc1.Refresh
    Combo2.Text = ""
    Combo3.Text = ""
  End If
  End If
  Form_Load
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
conn
Combo2.Enabled = False
Combo3.Enabled = False
Combo1.Enabled = True
Combo1.Locked = False
cancelbtn.Visible = False
cancelbtn.Enabled = True
savebtn.Enabled = False
deletebtn.Enabled = False
updatebtn.Enabled = False
addbtn.Visible = True
addbtn.Enabled = True
Text1.Text = ""
Text2.Text = ""
Combo3.Clear
Combo2.Clear
Set r1 = New ADODB.Recordset
sql = "select * from course"
Set r1 = c1.Execute(sql)
While r1.EOF = False
Combo2.AddItem r1.Fields(1)
r1.MoveNext
Wend
Text2.Text = 0
Combo1.Clear
sql = "select tp_id from topic"
Set r1 = c1.Execute(sql)
While r1.EOF = False
Combo1.AddItem r1.Fields(0)
r1.MoveNext
Wend

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Then
 Text1.SetFocus
ElseIf KeyAscii = 13 Then
 KeyAscii = 0
 Text2.SetFocus
 Else
  KeyAscii = 0
  MsgBox "Can Contain Only Characters nothing Else", vbInformation + vbOKOnly, ""
End If
End Sub
