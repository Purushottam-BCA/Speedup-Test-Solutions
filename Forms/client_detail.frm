VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmClient1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9570
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8670
   Icon            =   "client_detail.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   8670
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1080
      Top             =   6480
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
      Connect         =   "Provider=MSDAORA.1;User ID=sts/sts;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=sts/sts;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from client"
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
      Bindings        =   "client_detail.frx":0EE2
      Height          =   4935
      Left            =   75
      TabIndex        =   0
      Top             =   4635
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   8705
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "CLNT_ID"
         Caption         =   " Client ID"
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
         DataField       =   "CLNT_NM"
         Caption         =   "                             Client Name"
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
         DataField       =   "CLNT_MOB"
         Caption         =   "              Mobile"
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
         DataField       =   "CLNT_GNDR"
         Caption         =   "     Gender"
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
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   3990.047
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   2055.118
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1319.811
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   3510
      Left            =   95
      TabIndex        =   1
      Top             =   1020
      Width           =   8550
      Begin VB.CommandButton cmd5 
         BackColor       =   &H00C0C0FF&
         DisabledPicture =   "client_detail.frx":0EF7
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
         Left            =   150
         MouseIcon       =   "client_detail.frx":1551
         MousePointer    =   99  'Custom
         Picture         =   "client_detail.frx":16A3
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Exit From Here"
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton cmd1 
         BackColor       =   &H8000000E&
         DisabledPicture =   "client_detail.frx":1CFD
         Height          =   390
         Left            =   4080
         MouseIcon       =   "client_detail.frx":2396
         MousePointer    =   99  'Custom
         Picture         =   "client_detail.frx":24E8
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3000
         Width           =   1230
      End
      Begin VB.CommandButton cmd2 
         BackColor       =   &H8000000E&
         DisabledPicture =   "client_detail.frx":2B81
         Height          =   390
         Left            =   5400
         MouseIcon       =   "client_detail.frx":3234
         MousePointer    =   99  'Custom
         Picture         =   "client_detail.frx":3386
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3000
         Width           =   1245
      End
      Begin VB.TextBox t2 
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
         Height          =   330
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   13
         Text            =   "9931293724"
         Top             =   1620
         Width           =   1800
      End
      Begin VB.TextBox t3 
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
         Height          =   350
         Left            =   6240
         TabIndex        =   11
         Text            =   "Mohan Lal Bhagwat"
         Top             =   465
         Width           =   2160
      End
      Begin VB.CommandButton cmd4 
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
         Height          =   330
         Left            =   6840
         MouseIcon       =   "client_detail.frx":3A39
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Return to Main Menu"
         Top             =   935
         Width           =   1000
      End
      Begin VB.CommandButton cmd3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Go For Order >>"
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
         Left            =   6720
         MouseIcon       =   "client_detail.frx":3B8B
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Go for Order"
         Top             =   3000
         Width           =   1695
      End
      Begin VB.TextBox t1 
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
         Left            =   1920
         TabIndex        =   3
         Text            =   "Mohan Lal Bhagwat"
         Top             =   1000
         Width           =   3360
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H80000013&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   390
         Left            =   1920
         MouseIcon       =   "client_detail.frx":3CDD
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Id :-"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6240
         TabIndex        =   12
         Top             =   180
         Width           =   2055
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         Height          =   1300
         Left            =   6120
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label l1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CL001"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1920
         TabIndex        =   8
         Top             =   435
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client ID :"
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
         Left            =   750
         TabIndex        =   7
         Top             =   435
         Width           =   960
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client Name :"
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
         Left            =   390
         TabIndex        =   6
         Top             =   1005
         Width           =   1320
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile :"
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
         Left            =   915
         TabIndex        =   5
         Top             =   1620
         Width           =   795
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gender :"
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
         Left            =   885
         TabIndex        =   4
         Top             =   2160
         Width           =   825
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "-: Client Entry Details :-"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   2190
      TabIndex        =   17
      Top             =   240
      Width           =   3840
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   3450
      Left            =   120
      Top             =   1125
      Width           =   8565
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H80000013&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   8655
   End
End
Attribute VB_Name = "FrmClient1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd1_Click() 'Add Button
idgen
t1.Text = ""
t2.Text = ""
t3.Text = ""
cmd2.Enabled = True
cmd1.Enabled = False
End Sub

Private Sub cmd2_Click()
On Error Resume Next
If Trim(l1.Caption) = "" Or Trim(t1.Text) = "" Or Trim(t2.Text) = "" Then
MsgBox "Please Insert All Field", vbInformation + vbOKOnly, "Blank Field"
t1.SetFocus
Exit Sub
End If
Set r = c.Execute("select upper(clnt_nm) from client ")
While r.EOF = False
 If UCase(Trim(t1.Text)) = r.Fields(0) Then
   MsgBox "Oops.Same Name Already Exist in Database. Please Use Different name or Use Number Eg: Mukesh Kumar 1, Mukesh Kumar 2 etc.", vbInformation + vbOKOnly, "Duplicate Name"
   Exit Sub
 End If
r.MoveNext
Wend
c.Execute ("delete from client where upper(clnt_id) ='" & Trim(UCase(l1.Caption)) & "' ")
c.Execute ("insert into client values('" & l1.Caption & "','" & UCase(t1.Text) & "', " & UCase(t2.Text) & ",'" & Combo1.Text & "' )")
MsgBox "SuccessFullly added", vbInformation + vbOKOnly, "Success"
If MsgBox("Go for Order Now (Yes/No) ??", vbYesNo + vbInformation, "Order Now") = vbYes Then
 cmd3.Enabled = True
 cmd2.Enabled = False
 cmd1.Enabled = True
 MsgBox "Click On GO FOR ORDER button --->", vbInformation + vbOKOnly, "Success"
 cmd3.SetFocus
Else
 Form_Load
End If
End Sub

Private Sub cmd3_Click()
CurrentClient = l1.Caption
Me.Hide
FrmClient2.Show
End Sub

Private Sub cmd4_Click()
If Trim(t3.Text) = "" Then
 t3.SetFocus
Else
 Set r = c.Execute("select * from client where upper(clnt_id) ='" & Trim(UCase(t3.Text)) & "' ")
 If r.EOF = False Then
  l1.Caption = r.Fields(0)
  t1.Text = r.Fields(1)
  t2.Text = r.Fields(2)
  Combo1.Text = r.Fields(3)
  cmd2.Enabled = True
 End If
End If
End Sub

Private Sub cmd5_Click() 'Back Button
Unload Me
End Sub

Private Sub Form_Load()
conn
CenterForm Me
Combo1.Clear
Combo1.AddItem "Male"
Combo1.AddItem "Female"
cmd2.Enabled = False
cmd3.Enabled = False
l1.Caption = ""
t1.Text = ""
t2.Text = ""
t3.Text = ""
cmd1.Enabled = True
cmd2.Enabled = False
cmd3.Enabled = False
End Sub

Public Function idgen()
Dim t As Integer
Set r1 = New ADODB.Recordset
sql = "select MAX(to_number(substr(clnt_id,3,length(clnt_id))))from client"
Set r1 = c1.Execute(sql)
If IsNull(r1.Fields(0)) Then
 l1.Caption = "CL001"
Else
 t = r1.Fields(0)
 If t > 0 And t < 9 Then
  l1.Caption = "CL00" & (t + 1)
 ElseIf t < 99 Then
  l1.Caption = "CL0" & (t + 1)
 Else
  l1.Caption = "CL" & (t + 1)
End If
End If
End Function

Private Sub t1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or (((KeyAscii >= 65) And (KeyAscii <= 90)) Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or (KeyAscii = 8) Or (KeyAscii = 32)) Then
        t1.SetFocus
  ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        t2.SetFocus
  Else
   KeyAscii = 0
   MsgBox "Name Cannot Contains Special Character", vbInformation + vbOKOnly, ""
  End If
End Sub

Private Sub t2_KeyPress(KeyAscii As Integer)
If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Then
       t2.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   Else
   KeyAscii = 0
  End If

End Sub

Private Sub t2_LostFocus()
If (t2.Text <> "") Then
        If (Len(t2.Text) < 10) Then
            MsgBox "Invalid MOBILE NUMBER", vbExclamation + vbOKOnly, "Invalid  Mobile No"
            t2.SetFocus
        End If
End If
End Sub

