VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form LoginINFO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login Info"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10560
   Icon            =   "Login_Info.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   10560
   Begin VB.CommandButton Command5 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      MouseIcon       =   "Login_Info.frx":0EE2
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   8430
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      MouseIcon       =   "Login_Info.frx":1034
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   8430
      Width           =   1095
   End
   Begin VB.TextBox Text1 
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
      Left            =   3960
      TabIndex        =   9
      Text            =   "hjbhjbhjbhjbh"
      Top             =   8415
      Width           =   2295
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8415
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   14843
      _Version        =   393216
      MousePointer    =   99
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   794
      BackColor       =   -2147483635
      ForeColor       =   128
      MouseIcon       =   "Login_Info.frx":1186
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Student's Login Info"
      TabPicture(0)   =   "Login_Info.frx":12E8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DataGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Adodc1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "User's Login Info"
      TabPicture(1)   =   "Login_Info.frx":1304
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataGrid2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Adodc2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame2 
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   -75000
         TabIndex        =   6
         Top             =   7890
         Width           =   10575
         Begin VB.CommandButton Command4 
            Caption         =   "<<<  Student Menu   "
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8295
            MouseIcon       =   "Login_Info.frx":1320
            MousePointer    =   99  'Custom
            TabIndex        =   8
            Top             =   65
            Width           =   2175
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Exit"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   80
            MouseIcon       =   "Login_Info.frx":1472
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   65
            Width           =   855
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Search Name :"
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
            Left            =   2400
            TabIndex        =   11
            Top             =   75
            Width           =   1500
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   0
         TabIndex        =   3
         Top             =   7890
         Width           =   10575
         Begin VB.CommandButton Command1 
            Caption         =   "Employee Menu  >>>"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8295
            MouseIcon       =   "Login_Info.frx":15C4
            MousePointer    =   99  'Custom
            TabIndex        =   5
            Top             =   65
            Width           =   2175
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Exit"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   80
            MouseIcon       =   "Login_Info.frx":1716
            MousePointer    =   99  'Custom
            TabIndex        =   4
            Top             =   65
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Search Name :"
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
            Left            =   2400
            TabIndex        =   10
            Top             =   75
            Width           =   1500
         End
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   495
         Left            =   -69840
         Top             =   4200
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
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
         Connect         =   "Provider=MSDAORA.1;User ID=sts/sts;Persist Security Info=True"
         OLEDBString     =   "Provider=MSDAORA.1;User ID=sts/sts;Persist Security Info=True"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select L.E_id, E.e_nm, L.E_log_ID,L. E_PSWD from Emp E, Emp_login L where L.e_id= E.emp_id"
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
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   495
         Left            =   5760
         Top             =   5040
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
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
         Connect         =   "Provider=MSDAORA.1;User ID=sts/sts;Persist Security Info=True"
         OLEDBString     =   "Provider=MSDAORA.1;User ID=sts/sts;Persist Security Info=True"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "SELECT L.RSTUD_REG_NO, R.RSTUD_NM,L. RSTUD_LOG_ID,  L.RSTUD_PSWD FROM RSTUD R,STUD_LOGIN L WHERE L.RSTUD_REG_NO = R.RSTUD_REG_NO"
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
         Bindings        =   "Login_Info.frx":1868
         Height          =   7435
         Left            =   0
         TabIndex        =   1
         Top             =   455
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   13123
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   25
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "RSTUD_REG_NO"
            Caption         =   " Registration No"
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
            Caption         =   "              Student Name"
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
            DataField       =   "RSTUD_LOG_ID"
            Caption         =   "            Login ID"
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
            DataField       =   "RSTUD_PSWD"
            Caption         =   "              Password"
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
               Alignment       =   2
               ColumnWidth     =   1769.953
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   3254.74
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   2534.74
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   2759.811
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "Login_Info.frx":187D
         Height          =   7435
         Left            =   -75000
         TabIndex        =   2
         Top             =   455
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   13123
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   25
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "E_ID"
            Caption         =   "        Reg. No"
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
            Caption         =   "                  User Name"
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
            DataField       =   "E_LOG_ID"
            Caption         =   "            LogIn ID"
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
            DataField       =   "E_PSWD"
            Caption         =   "           Password"
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
               Alignment       =   2
               ColumnWidth     =   1860.095
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   3270.047
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   2594.835
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   2550.047
            EndProperty
         EndProperty
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   5
      Height          =   430
      Left            =   3840
      Top             =   15
      Width           =   2340
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login Information"
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
      Left            =   3960
      TabIndex        =   14
      Top             =   60
      Width           =   2100
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H80000013&
      FillStyle       =   0  'Solid
      Height          =   465
      Left            =   0
      Top             =   0
      Width           =   10550
   End
End
Attribute VB_Name = "LoginINFO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
SSTab1.Tab = 1
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
SSTab1.Tab = 0
End Sub

Private Sub Command5_Click()
  Text1.Text = ""
  Adodc2.RecordSource = "select L.E_id, E.e_nm, L.E_log_ID,L. E_PSWD from Emp E, Emp_login L where L.e_id= E.emp_id "
  Adodc2.Refresh
  Adodc1.RecordSource = "SELECT L.RSTUD_REG_NO, R.RSTUD_NM,L. RSTUD_LOG_ID,  L.RSTUD_PSWD FROM RSTUD R,STUD_LOGIN L WHERE L.RSTUD_REG_NO = R.RSTUD_REG_No "
  Adodc1.Refresh
End Sub

Private Sub Command6_Click()
If SSTab1.Tab = 0 Then
  RptStudLoginInfo.Show 1, MDI
Else
  RptUserLoginInfo.Show 1, MDI
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Text1.Text = ""
End Sub

Private Sub Form_Unload(cancel As Integer)
admin_dash.Enabled = True
End Sub

Private Sub Text1_Change()
If SSTab1.Tab = 0 Then 'student
  Adodc1.RecordSource = "SELECT L.RSTUD_REG_NO, R.RSTUD_NM,L. RSTUD_LOG_ID,  L.RSTUD_PSWD FROM RSTUD R,STUD_LOGIN L WHERE L.RSTUD_REG_NO = R.RSTUD_REG_NO and upper(R.rstud_nm) like '" & UCase(Trim(Text1.Text)) & "%' "
  Adodc1.Refresh
Else 'Employee
  Adodc2.RecordSource = "select L.E_id, E.e_nm, L.E_log_ID,L. E_PSWD from Emp E, Emp_login L where L.e_id= E.emp_id and upper(E.E_NM)like '" & UCase(Trim(Text1.Text)) & "%' "
  Adodc2.Refresh
End If
End Sub
