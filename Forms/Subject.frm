VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCourseMaster 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Course"
   ClientHeight    =   9390
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   7605
   Icon            =   "Subject.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   7605
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Subject.frx":6062
      Height          =   4815
      Left            =   0
      TabIndex        =   9
      Top             =   4560
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   8493
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   12632256
      ForeColor       =   64
      HeadLines       =   1
      RowHeight       =   22
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "C_ID"
         Caption         =   "   Course ID"
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
         DataField       =   "C_NM"
         Caption         =   "Course Code"
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
         DataField       =   "C_FULL_NM"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         ScrollBars      =   2
         Locked          =   -1  'True
         RecordSelectors =   0   'False
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   4050.142
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Course "
      Height          =   2775
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   7335
      Begin VB.ComboBox combo1 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2640
         MouseIcon       =   "Subject.frx":6077
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Text            =   "combo1"
         Top             =   300
         Width           =   2355
      End
      Begin VB.TextBox ctype 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         Left            =   2670
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1815
         Width           =   4270
      End
      Begin VB.TextBox cname 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         Left            =   2715
         MaxLength       =   15
         TabIndex        =   3
         Top             =   1035
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Course Name"
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
         Top             =   1845
         Width           =   1395
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000006&
         Height          =   405
         Left            =   2640
         Shape           =   4  'Rounded Rectangle
         Top             =   1800
         Width           =   4335
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000006&
         Height          =   405
         Left            =   2640
         Shape           =   4  'Rounded Rectangle
         Top             =   1020
         Width           =   2415
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Course Code"
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
         Top             =   1080
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Serial ID"
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
         Top             =   360
         Width           =   840
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   645
      Left            =   15
      TabIndex        =   0
      Top             =   3795
      Width           =   7525
      Begin VB.CommandButton delete 
         DisabledPicture =   "Subject.frx":61C9
         Height          =   390
         Left            =   3135
         MouseIcon       =   "Subject.frx":68F8
         MousePointer    =   99  'Custom
         Picture         =   "Subject.frx":6A4A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   1400
      End
      Begin VB.CommandButton save 
         BackColor       =   &H8000000E&
         DisabledPicture =   "Subject.frx":7179
         Height          =   390
         Left            =   1680
         MouseIcon       =   "Subject.frx":782C
         MousePointer    =   99  'Custom
         Picture         =   "Subject.frx":797E
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   120
         Width           =   1245
      End
      Begin VB.CommandButton addbtn 
         BackColor       =   &H8000000E&
         DisabledPicture =   "Subject.frx":8031
         Height          =   390
         Left            =   240
         MouseIcon       =   "Subject.frx":86CA
         MousePointer    =   99  'Custom
         Picture         =   "Subject.frx":881C
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   1230
      End
      Begin VB.CommandButton backbtn 
         BackColor       =   &H00C0C0FF&
         DisabledPicture =   "Subject.frx":8EB5
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
         Left            =   6150
         MouseIcon       =   "Subject.frx":950F
         MousePointer    =   99  'Custom
         Picture         =   "Subject.frx":9661
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Exit From Here"
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton update 
         DisabledPicture =   "Subject.frx":9CBB
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
         Left            =   4740
         MouseIcon       =   "Subject.frx":A34E
         MousePointer    =   99  'Custom
         Picture         =   "Subject.frx":A4A0
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   1215
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
         TabIndex        =   1
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton cancelbtn 
         DisabledPicture =   "Subject.frx":AB33
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
         Left            =   240
         MouseIcon       =   "Subject.frx":B289
         MousePointer    =   99  'Custom
         Picture         =   "Subject.frx":B3DB
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   1230
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   6480
      Top             =   3120
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Connect         =   "Provider=MSDAORA.1;Password=Sts;User ID=Sts;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;Password=Sts;User ID=Sts;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from course"
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
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "[ Course ]"
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
      Height          =   465
      Left            =   2760
      TabIndex        =   13
      Top             =   120
      Width           =   1590
   End
   Begin VB.Shape Shape9 
      Height          =   825
      Left            =   0
      Top             =   3720
      Width           =   7575
   End
   Begin VB.Shape Shape2 
      Height          =   2985
      Left            =   0
      Top             =   720
      Width           =   7575
   End
   Begin VB.Shape Shape10 
      FillColor       =   &H80000013&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "frmCourseMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim t As Integer
Dim opt As String
Private Sub addbtn_Click()
cauto_id  'Auto ID generate
If combo1.Locked = False Then
 combo1.Locked = True
Else
 combo1.Locked = False
End If
cname.Text = ""
ctype.Text = ""
save.Enabled = True
cancelbtn.Enabled = True
cancelbtn.Visible = True
addbtn.Visible = False
End Sub

Private Sub cancelbtn_Click()
Form_Load
combo1.Text = ""
cancelbtn.Visible = False
addbtn.Visible = True
combo1.Locked = False
cname.Text = ""
ctype.Text = ""
End Sub

Private Sub backbtn_Click()
Unload Me
End Sub

Private Sub cname_KeyPress(KeyAscii As Integer)
If (((KeyAscii >= 65) And (KeyAscii <= 90)) Or KeyAscii = 46 Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or (KeyAscii = 8) Or (KeyAscii = 32)) Or (KeyAscii >= 48 And KeyAscii <= 57) Then
   cname.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   ctype.SetFocus
  Else
   KeyAscii = 0
   MsgBox "Course Code Cannot Contain Special Character", vbInformation + vbOKOnly, ""
  End If
End Sub

Private Sub Combo1_Click() 'Searching Option
Set r1 = New ADODB.Recordset
sql = "select * from course where trim(c_id)=trim('" & combo1.Text & "')"
Set r1 = c1.Execute(sql)
If IsNull(r1.Fields(0)) = False Then
 cname.Enabled = True
 ctype.Enabled = True
 cname.Text = r1.Fields(1)
 ctype.Text = r1.Fields(2)
 update.Enabled = True
 delete.Enabled = True
 save.Enabled = False
Else
MsgBox "Invalid Course ID", vbOKOnly, " "
combo1.SetFocus
End If
End Sub
Private Sub ctype_KeyPress(KeyAscii As Integer)
If (((KeyAscii >= 65) And (KeyAscii <= 90)) Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or (KeyAscii = 8) Or (KeyAscii = 32)) Then
   ctype.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
  Else
   KeyAscii = 0
   MsgBox "Course Name Can Contain Only Characters not The Numbers & Special character..", vbInformation + vbOKOnly, ""
  End If
End Sub

Private Sub Form_Load()
Me.Top = 500
Me.Left = 5500
conn
cancelbtn.Enabled = False
save.Enabled = False
delete.Enabled = False
update.Enabled = False
data_in_combo
End Sub
Public Function data_in_combo() ' For Course ID
Set r1 = New ADODB.Recordset
combo1.Clear
sql = "select c_id from course"
Set r1 = c1.Execute(sql)
While r1.EOF = False
 combo1.AddItem r1.Fields(0)
 r1.MoveNext
Wend
End Function


Public Function cauto_id()
Set r1 = New ADODB.Recordset 'Table c001
sql = "select MAX(to_number(substr(c_id,2,length(c_id))))from course"
Set r1 = c1.Execute(sql)
If IsNull(r1.Fields(0)) Then
 combo1.Text = "C00" & 1
Else
 t = r1.Fields(0)
 If t > 0 And t < 9 Then
  combo1.Text = "C00" & (t + 1)
 ElseIf t < 99 Then
  combo1.Text = "C0" & (t + 1)
 End If
End If
End Function

Private Sub save_Click() 'Save
Set r1 = New ADODB.Recordset
If cname.Text = "" Then
 MsgBox " Course Code Blank", vbCritical + vbOKOnly, "Warning"
 cname.SetFocus
ElseIf ctype.Text = "" Then
 MsgBox " Course Name Cannot be Blank", vbCritical + vbOKOnly, "Warning"
 ctype.SetFocus
ElseIf Trim(cname.Text) <> "" Then
 Set r = New ADODB.Recordset
 Set r = c1.Execute("select * from course")
 While r.EOF = False
  If UCase(Trim(cname.Text)) = UCase(r.Fields(1)) Or UCase(Trim(ctype.Text)) = UCase(r.Fields(2)) Then
   MsgBox "Course Already Exists ", vbCritical + vbOKOnly, "Duplicate Course"
   Exit Sub
   End If
  r.MoveNext
 Wend
 sql = "insert into course values(upper('" & combo1.Text & "'),upper('" & cname.Text & "'),upper('" & ctype.Text & "'))"
 Set r1 = c1.Execute(sql)
 MsgBox "Course Successfully added", vbApplicationModal + vbInformation + vbOKOnly, ""
 Adodc1.Refresh
 data_in_combo
 addbtn_Click
 End If
End Sub

Private Sub delete_Click() 'Delete
If Trim(combo1.Text) = "" Or Trim(cname.Text) = "" Or Trim(ctype.Text) = "" Then
 MsgBox "Select Corrrect Course", vbCritical + vbOKOnly, "Delete ERROR"
Else
Set r1 = New ADODB.Recordset
opt = MsgBox("Are You Sure to Delete ?", vbQuestion + vbYesNo, "Delete conformation!")
If opt = vbYes Then
 Set r1 = New ADODB.Recordset
 sql = "delete from course where trim(c_id) = trim('" & combo1.Text & "')"
 Set r1 = c1.Execute(sql)
 MsgBox "Course Successfully Deleted!!", vbInformation + vbOKOnly, "Delete Course !"
 combo1.Text = ""
 cname.Text = ""
 ctype.Text = ""
 Adodc1.Refresh  'DataGrid Updated
 data_in_combo   'ComboBox Updated
 Form_Load
Else
End If
End If
End Sub

Private Sub update_Click()
 If Trim(combo1.Text) = "" Or Trim(cname.Text) = "" Or Trim(ctype.Text) = "" Then
  MsgBox "Select Corrrect Course", vbCritical + vbOKOnly, "Update ERROR"
 Else
  conn
  opt = MsgBox("Are You Sure to Update ?", vbQuestion + vbYesNo, "UPDATE")
  If opt = vbYes Then
   Set r1 = New ADODB.Recordset
   sql = "update course set c_nm=trim('" + cname.Text + "'),c_type=trim('" + ctype.Text + "') where trim(c_id)=trim('" + combo1.Text + "')"
   Set r1 = c1.Execute(sql)
   MsgBox "Course Successfully Updated!!", vbInformation + vbOKOnly, "Update Course !"
   Adodc1.Refresh
  End If
 End If
End Sub
