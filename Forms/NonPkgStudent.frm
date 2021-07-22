VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmNonpkg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unregistered Student"
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16170
   Icon            =   "NonPkgStudent.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   16170
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Txtstud 
      Height          =   405
      Left            =   2640
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   6000
      Top             =   8520
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      RecordSource    =   $"NonPkgStudent.frx":0EE2
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
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Height          =   675
      Left            =   15165
      MouseIcon       =   "NonPkgStudent.frx":0FE7
      MousePointer    =   99  'Custom
      Picture         =   "NonPkgStudent.frx":1139
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Print Report"
      Top             =   8400
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "NonPkgStudent.frx":20DB
      Height          =   7035
      Left            =   0
      TabIndex        =   0
      Top             =   1350
      Width           =   16140
      _ExtentX        =   28469
      _ExtentY        =   12409
      _Version        =   393216
      AllowUpdate     =   -1  'True
      ColumnHeaders   =   0   'False
      HeadLines       =   1
      RowHeight       =   23
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
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "RSTUD_REG_NO"
         Caption         =   "RSTUD_REG_NO"
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
         DataField       =   "RSTUD_NM"
         Caption         =   "RSTUD_NM"
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
         DataField       =   "RSTUD_FATHER_NM"
         Caption         =   "RSTUD_FATHER_NM"
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
         DataField       =   "RSTUD_STATUS"
         Caption         =   "RSTUD_STATUS"
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
         DataField       =   "RSTUD_MOB"
         Caption         =   "RSTUD_MOB"
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
         DataField       =   "RSTUD_DOJ"
         Caption         =   "RSTUD_DOJ"
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
      BeginProperty Column07 
         DataField       =   "SCH_TIMING"
         Caption         =   "SCH_TIMING"
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
            ColumnWidth     =   1769.953
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2894.74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2640.189
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1844.787
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1709.858
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1574.929
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1934.929
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   5280
      X2              =   11055
      Y1              =   630
      Y2              =   645
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Non Package / Unregistered Students"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   5400
      TabIndex        =   10
      Top             =   120
      Width           =   5655
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000013&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000013&
      Height          =   820
      Left            =   0
      Top             =   0
      Width           =   16135
   End
   Begin VB.Line Line10 
      Index           =   7
      X1              =   1785
      X2              =   1785
      Y1              =   840
      Y2              =   1340
   End
   Begin VB.Line Line10 
      Index           =   5
      X1              =   13965
      X2              =   13965
      Y1              =   840
      Y2              =   1320
   End
   Begin VB.Line Line10 
      Index           =   4
      X1              =   12375
      X2              =   12378
      Y1              =   840
      Y2              =   1320
   End
   Begin VB.Line Line10 
      Index           =   3
      X1              =   9150
      X2              =   9150
      Y1              =   840
      Y2              =   1320
   End
   Begin VB.Line Line10 
      Index           =   2
      X1              =   10650
      X2              =   10650
      Y1              =   840
      Y2              =   1320
   End
   Begin VB.Line Line10 
      Index           =   1
      X1              =   4680
      X2              =   4680
      Y1              =   840
      Y2              =   1320
   End
   Begin VB.Line Line10 
      Index           =   0
      X1              =   7305
      X2              =   7305
      Y1              =   840
      Y2              =   1320
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No"
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
      Index           =   12
      Left            =   9390
      TabIndex        =   8
      Top             =   945
      Width           =   975
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Join Date"
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
      Index           =   11
      Left            =   11025
      TabIndex        =   7
      Top             =   945
      Width           =   855
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Course"
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
      Index           =   10
      Left            =   12825
      TabIndex        =   6
      Top             =   945
      Width           =   645
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Schedule "
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
      Index           =   8
      Left            =   14280
      TabIndex        =   5
      Top             =   945
      Width           =   900
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Student Type"
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
      Index           =   7
      Left            =   7605
      TabIndex        =   4
      Top             =   945
      Width           =   1230
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Father's Name"
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
      Index           =   6
      Left            =   4725
      TabIndex        =   3
      Top             =   945
      Width           =   1335
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Student Name"
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
      Index           =   5
      Left            =   1875
      TabIndex        =   2
      Top             =   945
      Width           =   1335
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registration No"
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
      Index           =   4
      Left            =   150
      TabIndex        =   1
      Top             =   945
      Width           =   1440
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H00808080&
      Height          =   495
      Left            =   0
      Top             =   840
      Width           =   16140
   End
End
Attribute VB_Name = "FrmNonpkg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
Txtstud.Text = "UnRegistered"
Set r1 = c.Execute("select count(*) from rstud where upper(RSTUD_STATUS)='UNREGISTERED' ")
    If r1.Fields(0) > 0 Then
     DV.CmdStudRep "", "", Txtstud.Text, "", "", "", "", ""
     StudReport.Sections("section4").Controls("TotStu").Caption = r1.Fields(0)
     StudReport.Show 1, MDI
     DV.rsCmdStudRep.Close
    End If
End Sub

Private Sub Form_Load()
conn
Me.Top = 450
Me.Left = 2000
End Sub
