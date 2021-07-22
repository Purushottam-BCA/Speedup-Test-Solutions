VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmPackage 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Package"
   ClientHeight    =   9585
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10230
   Icon            =   "Pacakage.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9585
   ScaleWidth      =   10230
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   10560
      Top             =   1560
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1296
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
      RecordSource    =   "select p.pkg_id,initcap(p.Pkg_nm),initcap(C.c_nm), p.pkg_fee,p.pkg_all_tst,p.pkg_dur from pkg p,course C where C.c_id=p.c_id"
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
   Begin VB.ComboBox Combo2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6000
      TabIndex        =   24
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      MouseIcon       =   "Pacakage.frx":09EA
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   360
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Pacakage.frx":0B3C
      Height          =   3950
      Left            =   120
      TabIndex        =   22
      Top             =   5640
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   6959
      _Version        =   393216
      AllowUpdate     =   -1  'True
      ForeColor       =   64
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "PKG_ID"
         Caption         =   "           ID"
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
         DataField       =   "INITCAP(P.PKG_NM)"
         Caption         =   "Package Name"
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
         Caption         =   "Course Name"
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
         DataField       =   "PKG_FEE"
         Caption         =   "        Fee"
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
         DataField       =   "PKG_ALL_TST"
         Caption         =   "   Total Test"
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
         DataField       =   "PKG_DUR"
         Caption         =   "   Duration"
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
            ColumnWidth     =   1335.118
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2324.977
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   1275.024
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Height          =   3805
         Left            =   120
         TabIndex        =   2
         Top             =   885
         Width           =   9855
         Begin VB.TextBox Lbl5 
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
            Height          =   390
            Left            =   7080
            MaxLength       =   3
            TabIndex        =   26
            Top             =   2485
            Width           =   1215
         End
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
            Left            =   3000
            MouseIcon       =   "Pacakage.frx":0B51
            MousePointer    =   99  'Custom
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   320
            Width           =   2535
         End
         Begin VB.TextBox Lbl4 
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
            Height          =   390
            Left            =   3000
            MaxLength       =   3
            TabIndex        =   5
            Text            =   " "
            Top             =   3180
            Width           =   735
         End
         Begin VB.TextBox Lbl2 
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
            Height          =   390
            Left            =   3000
            MaxLength       =   20
            TabIndex        =   4
            Top             =   1740
            Width           =   4095
         End
         Begin VB.TextBox Lbl3 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   365
            Left            =   3365
            MaxLength       =   6
            TabIndex        =   3
            Top             =   2460
            Width           =   1095
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Eg:- Package1, Prayash"
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
            Left            =   7200
            TabIndex        =   28
            Top             =   1800
            Width           =   2070
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Total Tests : "
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
            Left            =   5475
            TabIndex        =   27
            Top             =   2520
            Width           =   1260
         End
         Begin VB.Label lbl 
            BackColor       =   &H8000000D&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   4560
            TabIndex        =   21
            Top             =   -240
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Rs"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3000
            TabIndex        =   19
            Top             =   2460
            Width           =   375
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Days"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   3705
            TabIndex        =   13
            Top             =   3180
            Width           =   735
         End
         Begin VB.Label Lbl1 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   12
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Course               :"
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
            Left            =   945
            TabIndex        =   11
            Top             =   360
            Width           =   1635
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Validity             :"
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
            Left            =   945
            TabIndex        =   10
            Top             =   3165
            Width           =   1575
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Package ID        :"
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
            Left            =   945
            TabIndex        =   9
            Top             =   1080
            Width           =   1620
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Pacakge Name  :"
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
            Left            =   945
            TabIndex        =   8
            Top             =   1725
            Width           =   1620
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Pacakge Fee     :"
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
            Left            =   945
            TabIndex        =   7
            Top             =   2445
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   650
         Left            =   120
         TabIndex        =   1
         Top             =   4750
         Width           =   9855
         Begin VB.CommandButton btnClear 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6720
            MouseIcon       =   "Pacakage.frx":0CA3
            MousePointer    =   99  'Custom
            TabIndex        =   25
            Top             =   150
            Width           =   1695
         End
         Begin VB.CommandButton btnDelete 
            DisabledPicture =   "Pacakage.frx":0DF5
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4920
            MouseIcon       =   "Pacakage.frx":17A3
            MousePointer    =   99  'Custom
            Picture         =   "Pacakage.frx":18F5
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   150
            Width           =   1695
         End
         Begin VB.CommandButton btnBack 
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
            Height          =   375
            Left            =   8520
            MouseIcon       =   "Pacakage.frx":22A3
            MousePointer    =   99  'Custom
            Picture         =   "Pacakage.frx":23F5
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Exit From Here"
            Top             =   150
            Width           =   1215
         End
         Begin VB.CommandButton btnsave 
            DisabledPicture =   "Pacakage.frx":2BF2
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1680
            MouseIcon       =   "Pacakage.frx":364D
            MousePointer    =   99  'Custom
            Picture         =   "Pacakage.frx":379F
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   150
            Width           =   1815
         End
         Begin VB.CommandButton btnUpdate 
            DisabledPicture =   "Pacakage.frx":41FA
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3600
            MouseIcon       =   "Pacakage.frx":488D
            MousePointer    =   99  'Custom
            Picture         =   "Pacakage.frx":49DF
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   150
            Width           =   1215
         End
         Begin VB.CommandButton btnadd 
            DisabledPicture =   "Pacakage.frx":5072
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            MouseIcon       =   "Pacakage.frx":58F5
            MousePointer    =   99  'Custom
            Picture         =   "Pacakage.frx":5A47
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   150
            Width           =   1455
         End
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "...PACKAGE..."
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   20
         Top             =   300
         Width           =   2295
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000013&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   615
         Left            =   120
         Top             =   240
         Width           =   9855
      End
   End
End
Attribute VB_Name = "FrmPackage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBack_Click()
Unload Me
End Sub

Private Sub btnClear_Click()
Fresh
End Sub

Private Sub btnDelete_Click()
If Trim(lbl1.Caption) = "" Then
 MsgBox "Select Corrrect Package", vbCritical + vbOKOnly, "Delete ERROR"
 Exit Sub
Else
opt = MsgBox("Are You Sure to Delete ?", vbQuestion + vbYesNo, "Delete conformation!")
If opt = vbYes Then
Set r1 = New ADODB.Recordset
sql = " delete from pkg where pkg_id='" & lbl1.Caption & "'"
Set r1 = c1.Execute(sql)
MsgBox "1 Package deleted", vbInformation + vbOKOnly, ""
Adodc1.Refresh
Fresh
End If
End If
End Sub

Private Sub btnsave_Click()
If Trim(lbl2.Text) = "" Then
MsgBox "Package Name Blank", vbCritical + vbOKOnly, "warning"
lbl2.SetFocus
ElseIf Trim(lbl3.Text) = "" Or Val(lbl3.Text) <= 0 Then
MsgBox "Invalid Package Fee.....", vbCritical + vbOKOnly, "Warning"
lbl3.SetFocus
ElseIf Trim(lbl4.Text) = "" Or Val(lbl4.Text) <= 0 Then
 MsgBox "Enter Duration Of Package (In Days)..", vbCritical + vbOKOnly, "Warning"
lbl4.SetFocus
ElseIf Trim(lbl5.Text) = "" Or Val(lbl5.Text) <= 0 Then
 MsgBox "Enter Total Number Of Tests Available in Package..", vbCritical + vbOKOnly, "Warning"
lbl5.SetFocus
ElseIf Trim(Combo1.Text) = "" Then
 MsgBox "Select The Course For Package..", vbCritical + vbOKOnly, "Warning"
 Combo1.SetFocus
Else
 Set r2 = c.Execute("select upper(pkg_nm) from pkg")
 While r2.EOF = False
  If UCase(Trim(lbl2.Text)) = r2.Fields(0) Then
   MsgBox "A Package With the Same Name Alreasdy Exist.. Enter Another Name....", vbCritical + vbOKOnly, "Duplicate Name"
   lbl2.SetFocus
   Exit Sub
  End If
 r2.MoveNext
 Wend
 Set r1 = New ADODB.Recordset
 sql = "insert into pkg values ('" & lbl1.Caption & "','" & lbl2.Text & "'," & lbl3.Text & "," & lbl4.Text & "," & lbl5.Text & ",'" & Lbl.Caption & "') "
 Set r1 = c1.Execute(sql)
 MsgBox "New Package Added..", vbInformation + vbOKOnly, ""
 Fresh
 Adodc1.Refresh
End If
End Sub

Private Sub btnUpdate_Click()
If Trim(lbl1.Caption) = "" Then
 MsgBox "Select Corrrect Package..", vbCritical + vbOKOnly, "Update ERROR"
 Exit Sub
 End If
If Trim(lbl2.Text) = "" Or Trim(lbl3.Text) = "" Or Trim(lbl4.Text) = "" Or Trim(lbl5.Text) = "" Or Trim(Lbl.Caption) = "" Then
 MsgBox "Fill all fields. Some Fields May be Empty..", vbInformation + vbOKOnly, ""
 lbl2.SetFocus
Exit Sub
End If
  opt = MsgBox("Are You Sure to Update ?", vbQuestion + vbYesNo, "UPDATE")
  If opt = vbYes Then
  sql = "update pkg set pkg_nm='" & lbl2.Text & "',pkg_fee=" & lbl3.Text & ",pkg_dur=" & Val(lbl4.Text) & ",pkg_all_tst=" & Val(lbl5.Text) & " where pkg_id='" & lbl1.Caption & "'"
   c1.Execute (sql)
   MsgBox "Record SuccessFully updated", vbInformation + vbOKOnly, "Update Success"
   Adodc1.Refresh
   Fresh
  End If
End Sub

Private Sub Combo1_Click()
Set r = c.Execute("select c_id from course where upper(c_nm)='" & UCase(Combo1.Text) & "' ")
If r.EOF = False Then
 Lbl.Caption = r.Fields(0)
End If
lbl2.SetFocus
End Sub

Private Sub btnadd_Click()
Fresh
kauto_id
btnadd.Enabled = False
BtnDelete.Enabled = False
btnUpdate.Enabled = False
btnSave.Enabled = True
Combo1.Enabled = True
Combo1.SetFocus
End Sub
Public Function kauto_id()
Set r1 = New ADODB.Recordset
sql = "select max(to_number(substr(pkg_id,4,length(pkg_id))))from Pkg"
Set r1 = c1.Execute(sql)
If IsNull(r1.Fields(0)) Then
lbl1.Caption = "Pkg001"
Else
t = r1.Fields(0)
If t > 0 And t < 9 Then
lbl1.Caption = "Pkg00" & (t + 1)
ElseIf t < 99 Then
 lbl1.Caption = "Pkg0" & (t + 1)
Else
 lbl1.Caption = "Pkg" & (t + 1)
End If
End If
End Function

Private Sub Command2_Click()
If Trim(Combo2.Text) = "" Then
 Exit Sub
End If
Set r = New ADODB.Recordset
Set r = c.Execute("select * from pkg where upper(pkg_id)='" & UCase(Trim(Combo2.Text)) & "' or upper(pkg_nm)='" & UCase(Trim(Combo2.Text)) & "' ")
If r.EOF = False Then
 lbl1.Caption = r.Fields(0)
 lbl2.Text = r.Fields(1)
 lbl3.Text = r.Fields(2)
 lbl4.Text = r.Fields(3)
 lbl5.Text = r.Fields(4)
 Lbl.Caption = r.Fields(5)
 Set r1 = c.Execute("select initcap(c_nm) from course where c_id='" & Lbl.Caption & "' ")
 Combo1.Text = r1.Fields(0)
 btnUpdate.Enabled = True
 BtnDelete.Enabled = True
 btnSave.Enabled = False
 btnadd.Enabled = False
Else
 MsgBox "Package Not Found ...", vbCritical + vbOKOnly, "Not Found "
 Combo2.SetFocus
End If
End Sub

Private Sub Form_Load()
Me.Top = 500
Me.Left = 5500
conn
Fresh
End Sub

Public Sub Fresh()
btnadd.Enabled = True
btnSave.Enabled = False
btnUpdate.Enabled = False
BtnDelete.Enabled = False
lbl1.Caption = ""
lbl2.Text = ""
lbl3.Text = ""
lbl4.Text = ""
lbl5.Text = ""
Combo1.Clear
Set r = New ADODB.Recordset
Set r = c.Execute("select initcap(c_nm) from course")
While r.EOF = False
 Combo1.AddItem r.Fields(0)
r.MoveNext
Wend
Combo1.Enabled = False
Combo2.Clear
Set r = New ADODB.Recordset
Set r = c.Execute("select pkg_id from pkg")
While r.EOF = False
 Combo2.AddItem r.Fields(0)
 r.MoveNext
Wend
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If ((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8) Then
        Text1.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   Text2.SetFocus
  Else
   KeyAscii = 0
  End If
End Sub

Private Sub lbl2_KeyPress(KeyAscii As Integer)
If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 32 Then
        lbl2.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   lbl3.SetFocus
  Else
   KeyAscii = 0
  End If
End Sub

Private Sub lbl3_KeyPress(KeyAscii As Integer)
If InStr(lbl3.Text, ".") = False Then
  If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Or KeyAscii = 46 Then
   lbl3.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   lbl4.SetFocus
  Else
   KeyAscii = 0
  End If
Else
If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Then
   lbl3.SetFocus
ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   lbl4.SetFocus
Else
  KeyAscii = 0
  End If
End If
End Sub

Private Sub lbl4_KeyPress(KeyAscii As Integer)
 If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Then
     lbl4.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   lbl5.SetFocus
  Else
   KeyAscii = 0
      MsgBox "Date Can only be Number..", vbInformation + vbOKOnly, ""
  End If
End Sub

Private Sub lbl5_KeyPress(KeyAscii As Integer)
  If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Then
   lbl5.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
  Else
   KeyAscii = 0
   MsgBox "Only Number Allow", vbInformation + vbOKOnly, ""
  End If
End Sub
