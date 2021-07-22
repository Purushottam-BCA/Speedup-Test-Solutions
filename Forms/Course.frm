VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSubMaster 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Subject"
   ClientHeight    =   11520
   ClientLeft      =   120
   ClientTop       =   495
   ClientWidth     =   20490
   Icon            =   "Course.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11520
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Course "
      Height          =   3495
      Left            =   360
      TabIndex        =   6
      Top             =   720
      Width           =   7455
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
         Left            =   2520
         MouseIcon       =   "Course.frx":1E26
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   580
         Width           =   2115
      End
      Begin VB.TextBox sname 
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
         Height          =   350
         Left            =   2535
         MaxLength       =   25
         TabIndex        =   8
         Top             =   2445
         Width           =   4000
      End
      Begin VB.ComboBox Combo2 
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
         Left            =   2520
         MouseIcon       =   "Course.frx":1F78
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   1440
         Width           =   1635
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Subject Name"
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
         Left            =   850
         TabIndex        =   13
         Top             =   2445
         Width           =   1440
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000006&
         Height          =   420
         Left            =   2520
         Shape           =   4  'Rounded Rectangle
         Top             =   2400
         Width           =   4095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Subject ID"
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
         Left            =   850
         TabIndex        =   12
         Top             =   1485
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Course     "
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
         Left            =   850
         TabIndex        =   11
         Top             =   600
         Width           =   1020
      End
      Begin VB.Label Label5 
         Height          =   255
         Left            =   5040
         TabIndex        =   10
         Top             =   1320
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   5910
      Left            =   15
      TabIndex        =   1
      Top             =   4440
      Width           =   7935
      Begin VB.CommandButton srchbtn 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   2640
         MouseIcon       =   "Course.frx":20CA
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton update 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   5640
         MouseIcon       =   "Course.frx":221C
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton clrbtn 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   240
         MouseIcon       =   "Course.frx":236E
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CommandButton save 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   2040
         MouseIcon       =   "Course.frx":24C0
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton backbtn 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Exit"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   5280
         MouseIcon       =   "Course.frx":2612
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Exit From Here"
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton delete 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   3840
         MouseIcon       =   "Course.frx":2764
         MousePointer    =   99  'Custom
         TabIndex        =   22
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton addbtn 
         Caption         =   "Add New"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   240
         MouseIcon       =   "Course.frx":28B6
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   360
         Width           =   1575
      End
      Begin VB.Frame srchframe 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Search"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   3
         Top             =   2040
         Width           =   5415
         Begin VB.CommandButton Command1 
            Caption         =   "Go"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3840
            MouseIcon       =   "Course.frx":2A08
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   425
            Width           =   1335
         End
         Begin VB.ComboBox Combo3 
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
            ForeColor       =   &H00000040&
            Height          =   405
            Left            =   240
            TabIndex        =   4
            Top             =   425
            Width           =   3375
         End
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
         Index           =   0
         Left            =   11475
         ScaleHeight     =   435
         ScaleWidth      =   1395
         TabIndex        =   2
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton cancelbtn 
         Caption         =   "Cancel"
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
         Left            =   240
         MouseIcon       =   "Course.frx":2B5A
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Note :- For Searching , Click On Search Button Then Enter or select Record from list."
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
         Height          =   210
         Left            =   240
         TabIndex        =   23
         Top             =   3240
         Width           =   6675
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   405
         Index           =   0
         Left            =   2160
         Shape           =   4  'Rounded Rectangle
         Top             =   435
         Width           =   1545
      End
      Begin VB.Shape Shape5 
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   405
         Index           =   0
         Left            =   360
         Shape           =   4  'Rounded Rectangle
         Top             =   435
         Width           =   1545
      End
      Begin VB.Shape Shape6 
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   405
         Left            =   5400
         Shape           =   4  'Rounded Rectangle
         Top             =   1155
         Width           =   2265
      End
      Begin VB.Shape Shape7 
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   405
         Left            =   360
         Shape           =   4  'Rounded Rectangle
         Top             =   1155
         Width           =   2025
      End
      Begin VB.Shape Shape8 
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   405
         Left            =   2760
         Shape           =   4  'Rounded Rectangle
         Top             =   1155
         Width           =   2265
      End
      Begin VB.Shape Shape9 
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   400
         Left            =   5760
         Shape           =   4  'Rounded Rectangle
         Top             =   435
         Width           =   1785
      End
      Begin VB.Shape Shape10 
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   405
         Left            =   3960
         Shape           =   4  'Rounded Rectangle
         Top             =   435
         Width           =   1545
      End
   End
   Begin MSComctlLib.ListView lvl1 
      Height          =   10485
      Left            =   8160
      TabIndex        =   0
      Top             =   0
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   18494
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
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
      MouseIcon       =   "Course.frx":2CAC
      NumItems        =   0
   End
   Begin VB.Shape Shape11 
      Height          =   6135
      Left            =   0
      Top             =   4320
      Width           =   8055
   End
   Begin VB.Shape Shape3 
      Height          =   3735
      Left            =   0
      Top             =   600
      Width           =   8055
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-: SUBJECT :-"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2760
      TabIndex        =   14
      Top             =   60
      Width           =   2055
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "frmSubMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim t As Integer
Dim opt As String

 Sub loaddata()
 Dim r As New ADODB.Recordset
 Dim list As ListItem
 lvl1.ListItems.Clear
 Set r = c1.Execute("select initcap(s.sub_id),initcap(s.sub_nm),initcap(c.c_nm)  from sub s,course c where  c.c_id=s.c_id")
 While r.EOF = False
  Set list = lvl1.ListItems.add(, , r.Fields(0))
  list.SubItems(1) = r.Fields(1)
  list.SubItems(2) = r.Fields(2)
  r.MoveNext
  Wend
 End Sub
Private Sub addbtn_Click()
Sauto_id
combo1.Enabled = True
Combo2.Enabled = False
 Label5.Caption = ""
sname.Text = ""
save.Enabled = True
cancelbtn.Enabled = True
cancelbtn.Visible = True
addbtn.Visible = False
delete.Enabled = False
update.Enabled = False
End Sub

Private Sub cancelbtn_Click()
ClrAll
Combo2.Text = ""
sname.Text = ""
cancelbtn.Visible = False
addbtn.Visible = True
Combo2.Locked = False
End Sub

Private Sub backbtn_Click()
Unload Me
End Sub

Private Sub clrbtn_Click()
ClrAll
End Sub

Private Sub Combo1_Click() 'Searching Option by Course
 Label5.Caption = ""
 conn
 Set r1 = New ADODB.Recordset
 Set r1 = c.Execute("select C_id from course where upper(c_nm)='" & UCase(combo1.Text) & "' ")
 If r1.EOF = False Then
 Label5.Caption = r1.Fields(0)
 End If
End Sub
Private Sub combo2_Click() 'Searching Option by Subject
Set r1 = New ADODB.Recordset
sql = "select * from sub where trim(sub_id)=trim('" & Combo2.Text & "')"
Set r1 = c1.Execute(sql)
If IsNull(r1.Fields(0)) = False Then
 combo1.Text = r1.Fields(1)
 sname.Text = r1.Fields(2)
 update.Enabled = True
 delete.Enabled = True
 save.Enabled = False
 sname.SetFocus
Else
MsgBox "Invalid Subject ID", vbOKOnly, " "
Combo2.SetFocus
End If
End Sub

Private Sub Combo3_Change()
Label5.Caption = ""
Set r = c.Execute("select s.sub_id,c.c_nm,s.sub_nm,c.c_id  from sub s,course c where upper(s.sub_id) like '" & UCase(Trim(Combo3.Text)) & "%' ")
If IsNull(r.Fields(0)) = False Then
  combo1.Text = r.Fields(1)
  sname.Text = r.Fields(2)
  Combo2.Text = r.Fields(0)
   Label5.Caption = r.Fields(3)
   combo1.Enabled = True
Else
End If
End Sub

Private Sub combo3_Click() 'Search Click
Dim r1 As New ADODB.Recordset
Set r = c.Execute("select s.sub_id,c.c_nm,s.sub_nm,c.c_id from sub s,course c where s.sub_id='" & Combo3.Text & "' and c.c_id=s.c_id ")
If IsNull(r.Fields(0)) = False Then
 combo1.Text = r.Fields(1)
 sname.Text = r.Fields(2)
 Combo2.Text = r.Fields(0)
 Label5.Caption = r.Fields(3)
   combo1.Enabled = True
Else
 MsgBox "Invalid Id, Select the Appropriate Value ", vbInformation + vbOKOnly, ""
End If
End Sub

Private Sub Command1_Click()
If Trim(Combo3.Text) = "" Then
 MsgBox "Cannot be searched a blank Value ", vbInformation + vbOKOnly, ""
 Combo3.SetFocus
 Exit Sub
Else
Set r = c.Execute("select s.sub_id,c.c_nm,s.sub_nm,c.c_id from sub s,course c where s.sub_id='" & Combo3.Text & "' and c.c_id=s.c_id ")
If IsNull(r.Fields(0)) = False Then
 combo1.Text = r.Fields(1)
 sname.Text = r.Fields(2)
 Combo2.Text = r.Fields(0)
 Label5.Caption = r.Fields(3)
 Combo2.Enabled = False
 Set r1 = c1.Execute("select c_id from course where c_nm='" & combo1.Text & "' ")
 Label5.Caption = r1.Fields(0)
   combo1.Enabled = True
Else
 MsgBox "Invalid Id, Enter the Correct Value ", vbInformation + vbOKOnly, ""
End If
End If
End Sub

Public Sub ClrAll()
Combo3.Clear
Set r = c.Execute("select sub_id from sub")
While r.EOF = False
 Combo3.AddItem r.Fields(0)
 r.MoveNext
Wend
 combo1.Enabled = False
 cancelbtn.Enabled = False
 save.Enabled = False
 delete.Enabled = False
 update.Enabled = False
loaddata
data_in_combo1
data_in_combo2
sname.Text = ""
End Sub
Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
conn
With lvl1.ColumnHeaders
.Clear
.add , "", "Subject ID", Width / 7, lvwColumnLeft
.add , "", "Subject Name", Width / 4, lvwColumnCenter
.add , "", "Course ", Width / 6.1, lvwColumnCenter
End With
ClrAll
End Sub
Public Function data_in_combo1() 'For Course
combo1.Clear
Set r1 = New ADODB.Recordset
sql = "select initcap(c_nm) from course"
Set r1 = c1.Execute(sql)
While r1.EOF = False
 combo1.AddItem r1.Fields(0)
 r1.MoveNext
Wend
End Function

Public Function data_in_combo2() 'For Subject Id
Set r1 = New ADODB.Recordset
Combo2.Clear
sql = "select sub_id from sub"
Set r1 = c1.Execute(sql)
While r1.EOF = False
 Combo2.AddItem r1.Fields(0)
 r1.MoveNext
Wend
End Function

Public Function Sauto_id()
Set r1 = New ADODB.Recordset
sql = "select MAX(to_number(substr(sub_id,2,length(sub_id))))from sub"
Set r1 = c1.Execute(sql)
If IsNull(r1.Fields(0)) Then
 Combo2.Text = "S00" & 1
Else
 t = r1.Fields(0)
 If t > 0 And t < 9 Then
  Combo2.Text = "S00" & (t + 1)
 ElseIf t < 99 Then
  Combo2.Text = "S0" & (t + 1)
 Else
   Combo2.Text = "S" & (t + 1)
 End If
End If
Combo2.Locked = True
End Function

Private Sub Lvl1_Click()
Set r = c.Execute("select s.sub_id,s.sub_nm,c.c_nm,c.c_id from sub s,course c where c.c_id=s.c_id and s.sub_id='" & lvl1.SelectedItem & "' ")
If r.EOF = False Then
 combo1.Text = r.Fields(2)
 Combo2.Text = r.Fields(0)
 sname.Text = r.Fields(1)
 Label5.Caption = r.Fields(3)
 save.Enabled = False
 update.Enabled = True
 delete.Enabled = True
End If
End Sub

Private Sub save_Click() 'Save
Set r1 = New ADODB.Recordset
 If combo1.Text = "" Then
 MsgBox " Select Course ", vbCritical + vbOKOnly, "Warning"
 combo1.SetFocus
ElseIf sname.Text = "" Then
 MsgBox " Subject Name Blank", vbCritical + vbOKOnly, "Warning"
sname.SetFocus
ElseIf Trim(sname.Text) <> "" Then
  Set r = c1.Execute("select * from sub")
  While r.EOF = False
  If UCase(Trim(sname.Text)) = UCase(r.Fields(2)) And UCase(Trim(Label5.Caption)) = UCase(r.Fields(1)) Then
   MsgBox "Subject Already Exists ", vbCritical + vbOKOnly, "Duplicate Subject"
   Exit Sub
   End If
   r.MoveNext
  Wend
  sql = "insert into sub values (upper(trim('" & Combo2.Text & "')),'" & Label5.Caption & "',(upper('" & sname.Text & "')))"
 Set r1 = c1.Execute(sql)
 MsgBox "Subject Successfully added", vbApplicationModal + vbInformation + vbOKOnly, ""
 loaddata
 data_in_combo1
 data_in_combo2
 addbtn_Click
End If
End Sub

Private Sub delete_Click() 'Delete
If Trim(combo1.Text) = "" Or Trim(Combo2.Text) = "" Or Trim(sname.Text) = "" Then
 MsgBox "Select Corrrect Subject", vbCritical + vbOKOnly, "Delete ERROR"
Else
Set r1 = New ADODB.Recordset
opt = MsgBox("Are You Sure to Delete ?", vbQuestion + vbYesNo, "Delete conformation!")
If opt = vbYes Then
 Set r1 = New ADODB.Recordset
  sql = "delete from sub where sub_id='" & Combo2.Text & "' "
 Set r1 = c1.Execute(sql)
 MsgBox "Course Successfully Deleted!!", vbInformation + vbOKOnly, "Delete Course !"
  loaddata
  combo1.Text = ""
  sname.Text = ""
  Combo2.Text = ""
ClrAll
Else
End If
End If
End Sub

Private Sub sname_KeyPress(KeyAscii As Integer)
 If (((KeyAscii >= 65) And (KeyAscii <= 90)) Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or (KeyAscii = 8) Or (KeyAscii = 32)) Then
        sname.SetFocus
  ElseIf KeyAscii = 13 Then
        KeyAscii = 0
  Else
        KeyAscii = 0
  End If
End Sub

Private Sub srchbtn_Click()
srchframe.Enabled = True
save.Enabled = False
delete.Enabled = True
update.Enabled = True
Combo2.Enabled = False
Combo3.SetFocus
End Sub

Private Sub update_Click()
 If Trim(combo1.Text) = "" Or Trim(Combo2.Text) = "" Or Trim(sname.Text) = "" Then
  MsgBox "Select Corrrect Subject or Fill All The Required Fields..", vbCritical + vbOKOnly, "Update ERROR"
  
 Else
  conn
  opt = MsgBox("Are You Sure to Update ?", vbQuestion + vbYesNo, "UPDATE")
   If opt = vbYes Then
    Set r1 = New ADODB.Recordset
   sql = "update sub set c_id='" & Label5.Caption & "',sub_nm=trim('" & sname.Text & "') where sub_id='" & Combo2.Text & "'"
    Set r1 = c1.Execute(sql)
   MsgBox "Subject Successfully Updated!!", vbInformation + vbOKOnly, "Update Course !"
    combo1.Text = ""
    Combo2.Text = ""
    sname.Text = ""
ClrAll
End If
 End If
End Sub
