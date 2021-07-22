VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmQuesType 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Question Type"
   ClientHeight    =   9465
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8310
   FillStyle       =   0  'Solid
   Icon            =   "Ques_type_now.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9465
   ScaleWidth      =   8310
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Ques_type_now.frx":0ECA
      Height          =   4935
      Left            =   0
      TabIndex        =   10
      Top             =   4560
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   8705
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "Q_TYP_ID"
         Caption         =   "     Q Type ID"
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
         DataField       =   "Q_TYP_NM"
         Caption         =   "Question Type Name"
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
         DataField       =   "Q_TYP_MRK"
         Caption         =   "Marks Per Questions"
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
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         RecordSelectors =   0   'False
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   1769.953
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3825.071
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   2385.071
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
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
      Left            =   6720
      MouseIcon       =   "Ques_type_now.frx":0EDF
      MousePointer    =   99  'Custom
      TabIndex        =   20
      ToolTipText     =   "Exit From Here"
      Top             =   120
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Question Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   8295
      Begin VB.TextBox Text2 
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
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   13
         Top             =   2350
         Width           =   1335
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
         Height          =   375
         Left            =   3000
         TabIndex        =   12
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Height          =   1335
         Left            =   5520
         TabIndex        =   7
         Top             =   1680
         Visible         =   0   'False
         Width           =   2655
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   240
            TabIndex        =   8
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type ID or Name :"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   255
            TabIndex        =   9
            Top             =   315
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Caption         =   "Question Type"
         Height          =   650
         Left            =   0
         TabIndex        =   5
         Top             =   3100
         Width           =   8295
         Begin VB.CommandButton backbtn 
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
            Height          =   375
            Left            =   6840
            MouseIcon       =   "Ques_type_now.frx":1031
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton btnSearch 
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
            Height          =   375
            Left            =   5520
            MouseIcon       =   "Ques_type_now.frx":1183
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton btnUpdate 
            Caption         =   "Update"
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
            MouseIcon       =   "Ques_type_now.frx":12D5
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton BtnDelete 
            Caption         =   "Delete"
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
            Left            =   2880
            MouseIcon       =   "Ques_type_now.frx":1427
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton btnSave 
            Caption         =   "Save"
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
            Left            =   1560
            MouseIcon       =   "Ques_type_now.frx":1579
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton addbtn 
            Caption         =   "Add New"
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
            Left            =   120
            MouseIcon       =   "Ques_type_now.frx":16CB
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   14
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
            TabIndex        =   6
            Top             =   720
            Width           =   1455
         End
         Begin VB.Shape Shape9 
            BackColor       =   &H00000040&
            BackStyle       =   1  'Opaque
            Height          =   360
            Left            =   1680
            Shape           =   4  'Rounded Rectangle
            Top             =   220
            Width           =   1065
         End
         Begin VB.Shape Shape8 
            BackColor       =   &H00000040&
            BackStyle       =   1  'Opaque
            Height          =   360
            Left            =   4320
            Shape           =   4  'Rounded Rectangle
            Top             =   220
            Width           =   1065
         End
         Begin VB.Shape Shape7 
            BackColor       =   &H00000040&
            BackStyle       =   1  'Opaque
            Height          =   360
            Left            =   3000
            Shape           =   4  'Rounded Rectangle
            Top             =   220
            Width           =   1065
         End
         Begin VB.Shape Shape6 
            BackColor       =   &H00000040&
            BackStyle       =   1  'Opaque
            Height          =   360
            Left            =   5640
            Shape           =   4  'Rounded Rectangle
            Top             =   220
            Width           =   1065
         End
         Begin VB.Shape Shape4 
            BackColor       =   &H00000040&
            BackStyle       =   1  'Opaque
            Height          =   360
            Left            =   240
            Shape           =   4  'Rounded Rectangle
            Top             =   225
            Width           =   1185
         End
         Begin VB.Shape Shape3 
            BackColor       =   &H00000040&
            BackStyle       =   1  'Opaque
            Height          =   360
            Left            =   6960
            Shape           =   4  'Rounded Rectangle
            Top             =   220
            Width           =   1065
         End
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
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
         RecordSource    =   "select * from q_typ"
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
      Begin VB.Label id 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Q_TYP_ID"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3000
         TabIndex        =   0
         Top             =   480
         Width           =   1650
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Question Type ID"
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
         Left            =   465
         TabIndex        =   4
         Top             =   480
         Width           =   1920
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Question Type Name"
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
         Left            =   435
         TabIndex        =   3
         Top             =   1355
         Width           =   2340
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Question Type Marks"
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
         Left            =   420
         TabIndex        =   2
         Top             =   2355
         Width           =   2355
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QUESTION  TYPE"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   2700
      TabIndex        =   11
      Top             =   195
      Width           =   2730
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   120
      Picture         =   "Ques_type_now.frx":181D
      Stretch         =   -1  'True
      Top             =   120
      Width           =   525
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000013&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   -2880
      Top             =   0
      Width           =   11185
   End
End
Attribute VB_Name = "FrmQuesType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim opt As Byte
Private Sub addbtn_Click()
Qtpauto_id  'Auto ID generate
btnsave.Enabled = True
addbtn.Enabled = False
btnDelete.Enabled = False
btnUpdate.Enabled = False
btnSearch.Enabled = True
backbtn.Enabled = True
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
End Sub

Private Sub backbtn_Click()
backbtn.Enabled = True
btnsave.Enabled = False
btnDelete.Enabled = False
btnUpdate.Enabled = False
btnSearch.Enabled = True
addbtn.Enabled = True
id.Caption = ""
Text1.Text = ""
Text2.Text = ""
comboLoad
opt = 1

End Sub

Private Sub btnDelete_Click()
If id.Caption = "" Then
 MsgBox "Select Appropriate Id", vbQuestion, ""
 Exit Sub
Else
c.Execute ("delete from q_typ where upper(q_typ_id)='" & UCase(id.Caption) & "' ")
MsgBox "Successfully deleted", vbInformation + vbOKOnly, "Delete"
Adodc1.Refresh
Form_Load
End If
End Sub

Private Sub btnsave_Click()
If Trim(Text1.Text) = "" Then
 MsgBox "Enter Type name ", vbInformation + vbOKOnly, ""
 Text1.SetFocus
 Exit Sub
ElseIf Trim(Text2.Text) = "" Then
 MsgBox "Enter Type Marks ", vbInformation + vbOKOnly, ""
 Text2.SetFocus
 Exit Sub
Else
Set r = c1.Execute("insert into q_typ values('" & id.Caption & "','" & Trim(Text1.Text) & "'," & Trim(Val(Text2.Text)) & ") ")
 MsgBox "Successfully saved ", vbInformation + vbOKOnly, "Saved Data"
 Adodc1.Refresh
 Form_Load
End If
End Sub

Private Sub btnSearch_Click()
If opt = 1 Then
 opt = 0
 Frame3.Visible = True
Else
 opt = 1
 Frame3.Visible = False
 Combo1.Text = ""
 id.Caption = ""
 Text1.Text = ""
 Text2.Text = ""
End If
End Sub

Private Sub Combo1_Change()
On Error Resume Next
Adodc1.RecordSource = "select * from q_typ where upper(q_typ_id) like '" & UCase(Trim(Combo1.Text)) & "%' or upper(q_typ_nm) like '" & UCase(Trim(Combo1.Text)) & "%' "
Adodc1.Refresh
Set r = c.Execute("select * from q_typ where upper(q_typ_id) like '" & UCase(Trim(Combo1.Text)) & "%' or upper(q_typ_nm) like '" & UCase(Trim(Combo1.Text)) & "%' ")
If IsNull(r.Fields(0)) = False Then
id.Caption = r.Fields(0)
Text1.Text = r.Fields(1)
Text2.Text = r.Fields(2)
 btnUpdate.Enabled = True
  btnDelete.Enabled = True
  addbtn.Enabled = False
Else
id.Caption = ""
Text1.Text = ""
Text2.Text = ""
End If
End Sub

Private Sub Combo1_Click()
Set r = c.Execute("select * from q_typ where upper(q_typ_id) ='" & UCase(Trim(Combo1.Text)) & "' ")
If IsNull(r.Fields(0)) = False Then
id.Caption = r.Fields(0)
Text1.Text = r.Fields(1)
Text2.Text = r.Fields(2)
 btnUpdate.Enabled = True
  btnDelete.Enabled = True
  addbtn.Enabled = False
Else
id.Caption = ""
Text1.Text = ""
Text2.Text = ""
  btnUpdate.Enabled = False
  btnDelete.Enabled = False
  addbtn.Enabled = True
End If
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub DataGrid1_Click()
Set id.DataSource = DataGrid1.DataSource
Set Text1.DataSource = DataGrid1.DataSource
Set Text2.DataSource = DataGrid1.DataSource
Text1.Locked = True
Text2.Locked = True

'Set id.DataSource = Nothing
'Set Text1.DataSource = Nothing
'Set Text2.DataSource = Nothing
'id.Caption = ""
'Text1.Text = ""
'Text2.Text = ""
End Sub

Private Sub Form_Load()
conn
Me.Top = 500
Me.Left = 5500
backbtn.Enabled = True
btnsave.Enabled = False
btnDelete.Enabled = False
btnUpdate.Enabled = False
btnSearch.Enabled = True
addbtn.Enabled = True
id.Caption = ""
Text1.Text = ""
Text2.Text = ""
comboLoad
opt = 1
End Sub

Public Function comboLoad()
Set r1 = New ADODB.Recordset
Combo1.Clear
sql = "select q_typ_id from q_typ"
Set r1 = c1.Execute(sql)
While r1.EOF = False
 Combo1.AddItem r1.Fields(0)
 r1.MoveNext
Wend
End Function

Public Function Qtpauto_id()
Set r1 = New ADODB.Recordset
sql = "select MAX(to_number(substr(q_typ_id,3,length(q_typ_id))))from q_typ"
Set r1 = c1.Execute(sql)
If IsNull(r1.Fields(0)) Then
 id.Caption = "QT00" & 1
Else
 t = r1.Fields(0)
 If t > 0 And t < 9 Then
  id.Caption = "QT00" & (t + 1)
 ElseIf t < 99 Then
  id.Caption = "QT0" & (t + 1)
 End If
End If
End Function

Private Sub Text1_KeyPress(KeyAscii As Integer)
 If (((KeyAscii >= 65) And (KeyAscii <= 90)) Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or (KeyAscii = 8) Or (KeyAscii = 32)) Then
        Text1.SetFocus
  ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        Text2.SetFocus
  Else
   KeyAscii = 0
  End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Then
        Text2.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
  Else
   KeyAscii = 0
  End If
End Sub
