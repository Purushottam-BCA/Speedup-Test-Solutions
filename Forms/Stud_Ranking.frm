VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Stud_Ranking 
   BorderStyle     =   0  'None
   Caption         =   "Student Ranking"
   ClientHeight    =   10560
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20415
   Icon            =   "Stud_Ranking.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10560
   ScaleWidth      =   20415
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   10555
      Left            =   15720
      ScaleHeight     =   10500
      ScaleWidth      =   4665
      TabIndex        =   17
      Top             =   0
      Width           =   4725
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2640
         MouseIcon       =   "Stud_Ranking.frx":0EE2
         MousePointer    =   99  'Custom
         Picture         =   "Stud_Ranking.frx":1034
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Close It Now"
         Top             =   9240
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   600
         MouseIcon       =   "Stud_Ranking.frx":1C42
         MousePointer    =   99  'Custom
         Picture         =   "Stud_Ranking.frx":1D94
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Print "
         Top             =   9240
         Width           =   1455
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student Info"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   345
         Left            =   1515
         TabIndex        =   32
         Top             =   65
         Width           =   1485
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   4670
      End
      Begin VB.Label l2 
         BackStyle       =   0  'Transparent
         Caption         =   "Not Available"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   31
         Top             =   5760
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact No     :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   630
         TabIndex        =   30
         Top             =   7560
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Course            :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   630
         TabIndex        =   29
         Top             =   6960
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Join date         :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   630
         TabIndex        =   28
         Top             =   6360
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Father Name   :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   630
         TabIndex        =   27
         Top             =   5760
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name              :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   630
         TabIndex        =   26
         Top             =   5160
         Width           =   1575
      End
      Begin VB.Label l5 
         BackStyle       =   0  'Transparent
         Caption         =   "Not Available"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   25
         Top             =   7560
         Width           =   2055
      End
      Begin VB.Label l4 
         BackStyle       =   0  'Transparent
         Caption         =   "Not Available"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   24
         Top             =   6960
         Width           =   2055
      End
      Begin VB.Label l3 
         BackStyle       =   0  'Transparent
         Caption         =   "Not Available"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   23
         Top             =   6360
         Width           =   2055
      End
      Begin VB.Label l1 
         BackStyle       =   0  'Transparent
         Caption         =   "Not Availble"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   22
         Top             =   5160
         Width           =   2175
      End
      Begin VB.Image Image1 
         Height          =   3600
         Left            =   1020
         Picture         =   "Stud_Ranking.frx":2749
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   2760
      End
      Begin VB.Label l6 
         BackStyle       =   0  'Transparent
         Caption         =   "Not Available"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   21
         Top             =   8220
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Student Type  :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   615
         TabIndex        =   20
         Top             =   8220
         Width           =   1815
      End
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   6960
      TabIndex        =   14
      Top             =   5280
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   750
      Left            =   15
      TabIndex        =   7
      Top             =   9865
      Width           =   15685
      Begin VB.CommandButton Command3 
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
         Height          =   395
         Left            =   10440
         MouseIcon       =   "Stud_Ranking.frx":40B6
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Click To Search"
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
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
         Height          =   405
         Left            =   13560
         MouseIcon       =   "Stud_Ranking.frx":4208
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Click to refresh All"
         Top             =   120
         Width           =   1935
      End
      Begin VB.ComboBox Combo2 
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
         Left            =   6960
         MouseIcon       =   "Stud_Ranking.frx":435A
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   120
         Width           =   3135
      End
      Begin VB.ComboBox Combo1 
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
         Left            =   1800
         MouseIcon       =   "Stud_Ranking.frx":44AC
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select or Enter Here : "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4800
         TabIndex        =   11
         Top             =   150
         Width           =   2115
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   480
         TabIndex        =   9
         Top             =   135
         Width           =   1065
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   6960
      Top             =   5280
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      CacheSize       =   80
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDAORA.1;Password=sts;User ID=STS;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;Password=sts;User ID=STS;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"Stud_Ranking.frx":45FE
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
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   6960
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   5280
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   6960
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   5280
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Stud_Ranking.frx":4720
      Height          =   8895
      Left            =   15
      TabIndex        =   6
      Top             =   960
      Width           =   15690
      _ExtentX        =   27675
      _ExtentY        =   15690
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   20
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
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "INITCAP(R.RSTUD_REG_NO)"
         Caption         =   "Registration No"
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
         DataField       =   "INITCAP(R.RSTUD_NM)"
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
         DataField       =   "INITCAP(C.C_NM)"
         Caption         =   "             Course"
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
         DataField       =   "TOT_MRK"
         Caption         =   "  Full Marks"
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
         DataField       =   "OBT_MRK"
         Caption         =   "   Obtained Marks"
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
         DataField       =   "DIF_LVL"
         Caption         =   "     Diff. level"
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
         DataField       =   "RSTUD_MOB"
         Caption         =   "             Mobile No"
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
            WrapText        =   -1  'True
            ColumnWidth     =   1590.236
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3674.835
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   2310.236
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   2025.071
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   2700.284
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton opt6 
      BackColor       =   &H80000016&
      Caption         =   "Full Length Test"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10575
      MaskColor       =   &H00FFFFFF&
      MouseIcon       =   "Stud_Ranking.frx":4735
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   450
      Width           =   5150
   End
   Begin VB.CommandButton opt5 
      BackColor       =   &H80000016&
      Caption         =   "Subject Wise Test"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5290
      MaskColor       =   &H00FFFFFF&
      MouseIcon       =   "Stud_Ranking.frx":4887
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   450
      Width           =   5295
   End
   Begin VB.CommandButton opt4 
      BackColor       =   &H80000016&
      Caption         =   "Topic Wise Test"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15
      MaskColor       =   &H00FFFFFF&
      MouseIcon       =   "Stud_Ranking.frx":49D9
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   450
      Width           =   5295
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9945
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15735
      _ExtentX        =   27755
      _ExtentY        =   17542
      _Version        =   393216
      MousePointer    =   99
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   706
      MouseIcon       =   "Stud_Ranking.frx":4B2B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Today"
      TabPicture(0)   =   "Stud_Ranking.frx":4C8D
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "This Week"
      TabPicture(1)   =   "Stud_Ranking.frx":4CA9
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "This Month"
      TabPicture(2)   =   "Stud_Ranking.frx":4CC5
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "This Year"
      TabPicture(3)   =   "Stud_Ranking.frx":4CE1
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "All Time"
      TabPicture(4)   =   "Stud_Ranking.frx":4CFD
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
   End
   Begin VB.Label Timi 
      Caption         =   "Label10"
      Height          =   1575
      Left            =   12840
      TabIndex        =   16
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   255
      Left            =   12120
      TabIndex        =   15
      Top             =   1920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Menu FFQRERW 
      Caption         =   "mm"
      Visible         =   0   'False
      Begin VB.Menu cdesf 
         Caption         =   "-"
      End
      Begin VB.Menu shdtails 
         Caption         =   "Show details"
      End
      Begin VB.Menu dsvb 
         Caption         =   "-"
      End
      Begin VB.Menu extcvvvv 
         Caption         =   "Exit"
      End
      Begin VB.Menu dsssfg 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "Stud_Ranking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurrentBtn As Byte

Private Sub Combo1_Click()
Combo2.Clear
Set r = New ADODB.Recordset
If Combo1.ListIndex = 0 Then
 Set r = c.Execute("select rstud_reg_no from rstud ")
 While r.EOF = False
  Combo2.AddItem r.Fields(0)
 r.MoveNext
 Wend
ElseIf Combo1.ListIndex = 1 Then
 Set r = c.Execute("select initcap(rstud_nm) from rstud ")
 While r.EOF = False
  Combo2.AddItem r.Fields(0)
 r.MoveNext
 Wend
ElseIf Combo1.ListIndex = 2 Then
 Set r = c.Execute("select initcap(c_nm) from course ")
 While r.EOF = False
  Combo2.AddItem r.Fields(0)
 r.MoveNext
 Wend
ElseIf Combo1.ListIndex = 3 Then
 Set r = c.Execute("select rstud_mob from rstud")
 While r.EOF = False
  Combo2.AddItem r.Fields(0)
 r.MoveNext
 Wend
End If
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
 RefreshIt
 Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Topic Wise Test') and A.sdate between '" & Text1.Text & "' and '" & Text2.Text & "' order by A.obt_mrk desc"
 Adodc1.Refresh
 opt4_Click
End Sub

Private Sub Command3_Click()
If Combo1.Text = "" Then
 MsgBox "Select Search By Option...", vbInformation + vbOKOnly, "Search By Empty"
 Combo1.SetFocus
 Exit Sub
ElseIf Trim(Combo2.Text) = "" Then
 MsgBox "Select or Enter Value to Search...", vbInformation + vbOKOnly, "Search Value"
 Combo2.SetFocus
 Exit Sub
End If
If SSTab1.Tab <> 4 Then ' All Except "All time Criteria"
 If CurrentBtn = 1 Then 'topic Wise Test
  If Combo1.ListIndex = 0 Then     'Reg.No
   Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Topic Wise Test') and A.sdate between '" & Text1.Text & "' and '" & Text2.Text & "' and upper(A.rstud_reg_no)='" & UCase(Trim(Combo2.Text)) & "' order by A.obt_mrk desc"
   Adodc1.Refresh
  ElseIf Combo1.ListIndex = 1 Then 'Name
   Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Topic Wise Test') and A.sdate between '" & Text1.Text & "' and '" & Text2.Text & "' and upper(R.rstud_nm) like '" & UCase(Trim(Combo2.Text)) & "%' order by A.obt_mrk desc"
   Adodc1.Refresh
  ElseIf Combo1.ListIndex = 2 Then 'Course
  Set r = New ADODB.Recordset
  Set r = c.Execute("select c_id from course where upper(c_nm)='" & UCase(Trim(Combo2.Text)) & "' ")
  If r.EOF = False Then
    Text3.Text = r.Fields(0)
  End If
   Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Topic Wise Test') and A.sdate between '" & Text1.Text & "' and '" & Text2.Text & "' and upper(r.c_id)='" & UCase(Trim(Text3.Text)) & "' order by A.obt_mrk desc"
   Adodc1.Refresh
   Text3.Text = ""
  ElseIf Combo1.ListIndex = 3 Then 'Mobile No
   Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Topic Wise Test') and A.sdate between '" & Text1.Text & "' and '" & Text2.Text & "' and upper(r.RSTUD_MOB)='" & UCase(Trim(Combo2.Text)) & "' order by A.obt_mrk desc"
   Adodc1.Refresh
  End If
 ElseIf CurrentBtn = 2 Then 'Subject wise Test
  If Combo1.ListIndex = 0 Then     'Reg.No
   Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Subject Wise Test')and A.sdate between '" & Text1.Text & "' and '" & Text2.Text & "'and upper(A.rstud_reg_no)='" & UCase(Trim(Combo2.Text)) & "' order by A.obt_mrk desc"
   Adodc1.Refresh
  ElseIf Combo1.ListIndex = 1 Then 'Name
   Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Subject Wise Test')and A.sdate between '" & Text1.Text & "' and '" & Text2.Text & "'and upper(R.rstud_nm) like '" & UCase(Trim(Combo2.Text)) & "%' order by A.obt_mrk desc"
   Adodc1.Refresh
  ElseIf Combo1.ListIndex = 2 Then 'Course
   Set r = New ADODB.Recordset
  Set r = c.Execute("select c_id from course where upper(c_nm)='" & UCase(Trim(Combo2.Text)) & "' ")
  If r.EOF = False Then
    Text3.Text = r.Fields(0)
  End If
   Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Subject Wise Test')and A.sdate between '" & Text1.Text & "' and '" & Text2.Text & "' and upper(r.c_id)='" & UCase(Trim(Text3.Text)) & "' order by A.obt_mrk desc"
   Adodc1.Refresh
   Text3.Text = ""
  ElseIf Combo1.ListIndex = 3 Then 'Mobile No
   Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Subject Wise Test')and A.sdate between '" & Text1.Text & "' and '" & Text2.Text & "'and upper(r.RSTUD_MOB)='" & UCase(Trim(Combo2.Text)) & "' order by A.obt_mrk desc"
   Adodc1.Refresh
  End If
 ElseIf CurrentBtn = 3 Then 'Full Length Test
    If Combo1.ListIndex = 0 Then     'Reg.No
   Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Full Length Test')and A.sdate between '" & Text1.Text & "' and '" & Text2.Text & "'and upper(A.rstud_reg_no)='" & UCase(Trim(Combo2.Text)) & "' order by A.obt_mrk desc"
   Adodc1.Refresh
  ElseIf Combo1.ListIndex = 1 Then 'Name
   Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Full Length Test')and A.sdate between '" & Text1.Text & "' and '" & Text2.Text & "'and upper(R.rstud_nm) like '" & UCase(Trim(Combo2.Text)) & "%' order by A.obt_mrk desc"
   Adodc1.Refresh
  ElseIf Combo1.ListIndex = 2 Then 'Course
   Set r = New ADODB.Recordset
  Set r = c.Execute("select c_id from course where upper(c_nm)='" & UCase(Trim(Combo2.Text)) & "' ")
  If r.EOF = False Then
    Text3.Text = r.Fields(0)
  End If
   Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Full Length Test')and A.sdate between '" & Text1.Text & "' and '" & Text2.Text & "'and upper(r.c_id)='" & UCase(Trim(Text3.Text)) & "' order by A.obt_mrk desc"
   Adodc1.Refresh
   Text3.Text = ""
  ElseIf Combo1.ListIndex = 3 Then 'Mobile No
   Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Full Length Test')and A.sdate between '" & Text1.Text & "' and '" & Text2.Text & "'and upper(r.RSTUD_MOB)='" & UCase(Trim(Combo2.Text)) & "' order by A.obt_mrk desc"
   Adodc1.Refresh
  End If
 End If
ElseIf SSTab1.Tab = 4 Then 'All Time
 If CurrentBtn = 1 Then 'topic Wise  Test
  If Combo1.ListIndex = 0 Then     'Reg.No
   Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Topic Wise Test')and upper(A.rstud_reg_no)='" & UCase(Trim(Combo2.Text)) & "' order by A.obt_mrk desc"
   Adodc1.Refresh
  ElseIf Combo1.ListIndex = 1 Then 'Name
   Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Topic Wise Test')and upper(R.rstud_nm) like '" & UCase(Trim(Combo2.Text)) & "%' order by A.obt_mrk desc"
   Adodc1.Refresh
  ElseIf Combo1.ListIndex = 2 Then 'Course
   Set r = New ADODB.Recordset
  Set r = c.Execute("select c_id from course where upper(c_nm)='" & UCase(Trim(Combo2.Text)) & "' ")
  If r.EOF = False Then
    Text3.Text = r.Fields(0)
  End If
   Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Topic Wise Test')and upper(r.c_id)='" & UCase(Trim(Text3.Text)) & "' order by A.obt_mrk desc"
   Adodc1.Refresh
   Text3.Text = ""
  ElseIf Combo1.ListIndex = 3 Then 'Mobile No
   Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Topic Wise Test')and upper(r.RSTUD_MOB)='" & UCase(Trim(Combo2.Text)) & "' order by A.obt_mrk desc"
   Adodc1.Refresh
  End If
 ElseIf CurrentBtn = 2 Then 'Subject wise Test
  If Combo1.ListIndex = 0 Then     'Reg.No
   Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Subject Wise Test')and upper(A.rstud_reg_no)='" & UCase(Trim(Combo2.Text)) & "'order by A.obt_mrk desc"
   Adodc1.Refresh
  ElseIf Combo1.ListIndex = 1 Then 'Name
   Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Subject Wise Test')and upper(R.rstud_nm) like '" & UCase(Trim(Combo2.Text)) & "%' order by A.obt_mrk desc"
   Adodc1.Refresh
  ElseIf Combo1.ListIndex = 2 Then 'Course
   Set r = New ADODB.Recordset
  Set r = c.Execute("select c_id from course where upper(c_nm)='" & UCase(Trim(Combo2.Text)) & "' ")
  If r.EOF = False Then
    Text3.Text = r.Fields(0)
  End If
   Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Subject Wise Test')and upper(r.c_id)='" & UCase(Trim(Text3.Text)) & "' order by A.obt_mrk desc"
   Adodc1.Refresh
 Text3.Text = ""
  ElseIf Combo1.ListIndex = 3 Then 'Mobile No
   Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Subject Wise Test')and upper(r.RSTUD_MOB)='" & UCase(Trim(Combo2.Text)) & "' order by A.obt_mrk desc"
   Adodc1.Refresh
  End If
 ElseIf CurrentBtn = 3 Then 'Full length test
  If Combo1.ListIndex = 0 Then     'Reg.No
   Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Full Length Test')and upper(A.rstud_reg_no)='" & UCase(Trim(Combo2.Text)) & "' order by A.obt_mrk desc"
   Adodc1.Refresh
  ElseIf Combo1.ListIndex = 1 Then 'Name
   Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Full Length Test')and upper(R.rstud_nm) like '" & UCase(Trim(Combo2.Text)) & "%' order by A.obt_mrk desc"
   Adodc1.Refresh
  ElseIf Combo1.ListIndex = 2 Then 'Course
   Set r = New ADODB.Recordset
  Set r = c.Execute("select c_id from course where upper(c_nm)='" & UCase(Trim(Combo2.Text)) & "' ")
  If r.EOF = False Then
    Text3.Text = r.Fields(0)
  End If
   Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Full Length Test')and upper(r.c_id)='" & UCase(Trim(Text3.Text)) & "' order by A.obt_mrk desc"
   Adodc1.Refresh
   Text3.Text = ""
  ElseIf Combo1.ListIndex = 3 Then 'Mobile No
   Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Full Length Test')and upper(r.RSTUD_MOB)='" & UCase(Trim(Combo2.Text)) & "' order by A.obt_mrk desc"
   Adodc1.Refresh
  End If
 End If
End If
End Sub

Private Sub Command4_Click() 'Print Button
 c.Execute ("delete from stuRank")
 sql = "insert into StuRank " & Adodc1.RecordSource
 c.Execute (sql)
DV.rsStudRanking.Open
 RPTStuRank.Sections("section4").Controls("RPTTstType").Caption = Label6.Caption
 RPTStuRank.Sections("section4").Controls("label13").Caption = Timi.Caption
 RPTStuRank.Refresh
 RPTStuRank.Show vbModal, MDI
DV.rsStudRanking.Close
End Sub

Private Sub DataGrid1_Click()
On Error Resume Next
Dim img As String
If Stu_login_reg_no = "" And Current_Logged_ID = "" Then
' Exit Sub
End If
If DataGrid1.Columns(0).Text = "" Then
 MsgBox "Either Select a Record then Double Click or Right Click To see Details ...", vbInformation + vbOKOnly, "Not Selected"
 Exit Sub
Else
 Set r = New ADODB.Recordset
 Set r = c.Execute("select Rn.rstud_pic, initcap(Rn.rstud_nm),initcap(Rn.RSTUD_FATHER_NM),Rn.rstud_doj,initcap(C.c_nm),Rn.RSTUD_MOB,initcap(Rn.RSTUD_STATUS) from rstud Rn, Course C where upper(Rn.rstud_reg_no)='" & UCase(DataGrid1.Columns(0).Text) & "' and Rn.c_id=C.c_id")
 If r.EOF = False Then
  img = r.Fields(0)
  Image1.Picture = LoadPicture(img)
  l1.Caption = r.Fields(1)
  l2.Caption = r.Fields(2)
  l3.Caption = r.Fields(3)
  l4.Caption = r.Fields(4)
  l5.Caption = r.Fields(5)
  l6.Caption = r.Fields(6)
 End If
End If
End Sub

Private Sub DataGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu FFQRERW
End If
End Sub

Private Sub extcvvvv_Click()
Unload Me
End Sub
Public Sub RefreshIt()
Image1.Picture = LoadPicture(App.Path & "\Graphics\#\PicNotAvail.jpg")
l1.Caption = "Not Available"
l2.Caption = "Not Available"
l3.Caption = "Not Available"
l4.Caption = "Not Available"
l5.Caption = "Not Available"
l6.Caption = "Not Available"
Text3.Text = ""
End Sub
Private Sub Form_Load()
conn
RefreshIt
Me.Left = 0
Me.Top = 0
 Combo1.Clear
 Combo1.AddItem "Registration No"
 Combo1.AddItem "Student Name"
 Combo1.AddItem "Course"
 Combo1.AddItem "Mobile No"
If Stu_login_reg_no <> "" Or Current_Logged_ID <> "" Then
 Command4.Enabled = False
 Frame1.Enabled = False
End If
 Text1.Text = Format(Date, "dd-mmm-yyyy")
 Text2.Text = Format(Date, "dd-mmm-yyyy")
 SSTab1.Tab = 0
 Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Topic Wise Test') and A.sdate between '" & Text1.Text & "' and '" & Text2.Text & "' order by A.obt_mrk desc"
 Adodc1.Refresh
 opt4_Click
End Sub

Private Sub opt4_Click() 'Only For Topic Wise Test
CurrentBtn = 1
opt4.BackColor = &H8000000A
opt5.BackColor = &H80000016
opt6.BackColor = &H80000016
Label6.Caption = opt4.Caption
If SSTab1.Tab = 0 Then     'Today
 Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Topic Wise Test') and A.sdate between '" & Text1.Text & "' and '" & Text2.Text & "' order by A.obt_mrk desc"
 Adodc1.Refresh
ElseIf SSTab1.Tab = 1 Then  'This Week
 Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Topic Wise Test') and A.sdate between '" & Text1.Text & "' and '" & Text2.Text & "' order by A.obt_mrk desc"
 Adodc1.Refresh
ElseIf SSTab1.Tab = 2 Then  'This Month
 Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Topic Wise Test') and A.sdate between '" & Text1.Text & "' and '" & Text2.Text & "' order by A.obt_mrk desc"
 Adodc1.Refresh
ElseIf SSTab1.Tab = 3 Then  'This Year
 Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Topic Wise Test') and A.sdate between '" & Text1.Text & "' and '" & Text2.Text & "' order by A.obt_mrk desc"
 Adodc1.Refresh
ElseIf SSTab1.Tab = 4 Then  'All Time
 Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Topic Wise Test') order by A.obt_mrk desc"
 Adodc1.Refresh
End If
End Sub

Private Sub opt5_Click() 'Subject Wise Test
CurrentBtn = 2
opt5.BackColor = &H8000000A
opt4.BackColor = &H80000016
opt6.BackColor = &H80000016
Label6.Caption = opt5.Caption
If SSTab1.Tab = 0 Then     'Today
Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Subject Wise Test')and A.sdate between '" & Text1.Text & "' and '" & Text2.Text & "' order by A.obt_mrk desc"
Adodc1.Refresh
ElseIf SSTab1.Tab = 1 Then  'This Week
Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Subject Wise Test')and A.sdate between '" & Text1.Text & "' and '" & Text2.Text & "' order by A.obt_mrk desc"
Adodc1.Refresh
ElseIf SSTab1.Tab = 2 Then  'This Month
Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Subject Wise Test')and A.sdate between '" & Text1.Text & "' and '" & Text2.Text & "' order by A.obt_mrk desc"
Adodc1.Refresh
ElseIf SSTab1.Tab = 3 Then  'This Year
Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Subject Wise Test')and A.sdate between '" & Text1.Text & "' and '" & Text2.Text & "' order by A.obt_mrk desc"
Adodc1.Refresh
ElseIf SSTab1.Tab = 4 Then  'All Time
Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Subject Wise Test')order by A.obt_mrk desc"
Adodc1.Refresh
End If
End Sub

Private Sub opt6_Click() 'Full Length Test
CurrentBtn = 3
opt6.BackColor = &H8000000A
opt5.BackColor = &H80000016
opt4.BackColor = &H80000016
Label6.Caption = opt6.Caption
If SSTab1.Tab = 0 Then     'Today
 Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Full Length Test')and A.sdate between '" & Text1.Text & "' and '" & Text2.Text & "' order by A.obt_mrk desc"
 Adodc1.Refresh
ElseIf SSTab1.Tab = 1 Then  'This Week
 Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Full Length Test')and A.sdate between '" & Text1.Text & "' and '" & Text2.Text & "' order by A.obt_mrk desc"
 Adodc1.Refresh
ElseIf SSTab1.Tab = 2 Then  'This Month
 Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Full Length Test')and A.sdate between '" & Text1.Text & "' and '" & Text2.Text & "' order by A.obt_mrk desc"
 Adodc1.Refresh
ElseIf SSTab1.Tab = 3 Then  'This Year
 Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Full Length Test')and A.sdate between '" & Text1.Text & "' and '" & Text2.Text & "' order by A.obt_mrk desc"
 Adodc1.Refresh
ElseIf SSTab1.Tab = 4 Then  'All Time
 Adodc1.RecordSource = "select initcap(r.rstud_reg_no), initcap(r.rstud_nm), initcap(C.c_nm), A.tot_mrk, A.obt_mrk, A.Dif_lvl,  r.rstud_mob  from course C, stud_prev_rec A, rstud r where A.rstud_reg_no = r.rstud_reg_no and r.c_id =C.c_id and upper(A.tst_typ)= upper('Full Length Test') order by A.obt_mrk desc"
 Adodc1.Refresh
End If
End Sub

Private Sub shdtails_Click() 'Show Details
On Error Resume Next
Dim img As String
If Stu_login_reg_no = "" And Current_Logged_ID = "" Then
' Exit Sub
End If
If DataGrid1.Columns(0).Text = "" Then
 MsgBox "Either Select a Record then Double Click or Right Click To see Details ...", vbInformation + vbOKOnly, "Not Selected"
 Exit Sub
Else
 Set r = New ADODB.Recordset
 Set r = c.Execute("select Rn.rstud_pic, initcap(Rn.rstud_nm),initcap(Rn.RSTUD_FATHER_NM),Rn.rstud_doj,initcap(C.c_nm),Rn.RSTUD_MOB,initcap(Rn.RSTUD_STATUS) from rstud Rn, Course C where upper(Rn.rstud_reg_no)='" & UCase(DataGrid1.Columns(0).Text) & "' and Rn.c_id=C.c_id")
 If r.EOF = False Then
  img = r.Fields(0)
  Image1.Picture = LoadPicture(img)
  l1.Caption = r.Fields(1)
  l2.Caption = r.Fields(2)
  l3.Caption = r.Fields(3)
  l4.Caption = r.Fields(4)
  l5.Caption = r.Fields(5)
  l6.Caption = r.Fields(6)
 End If
End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Set r = New ADODB.Recordset
If SSTab1.Tab = 0 Then 'Today
 Text1.Text = Format(Date, "dd-mmm-yyyy")
 Text2.Text = Format(Date, "dd-mmm-yyyy")
 Timi.Caption = "Daily (today)"
  opt4_Click
ElseIf SSTab1.Tab = 1 Then 'Week
 Text1.Text = Format(Date - (Weekday(Date, vbSunday) - 1), "DD-MMM-YYYY")
 Text2.Text = Format(Date, "dd-mmm-yyyy")
  Timi.Caption = "Weekly (this week)"
  opt4_Click
ElseIf SSTab1.Tab = 2 Then 'Month
 Text1.Text = "01-" & Format(Date, "mmm-yyyy")
 Text2.Text = Format(Date, "dd-mmm-yyyy")
   Timi.Caption = "Monthly (this month)"
  opt4_Click
ElseIf SSTab1.Tab = 3 Then 'Year
 Text1.Text = "01-Jan-" & Format(Date, "yyyy")
 Text2.Text = Format(Date, "dd-mmm-yyyy")
    Timi.Caption = "Yearly (this Year)"
opt4_Click
ElseIf SSTab1.Tab = 4 Then 'All
 Text1.Text = ""
 Text2.Text = Format(Date, "dd-mmm-yyyy")
   Timi.Caption = "All Time"
 opt4_Click
End If
End Sub
