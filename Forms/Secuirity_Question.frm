VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Security_Question 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Security Questions"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8010
   Icon            =   "Secuirity_Question.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   8010
   Begin VB.PictureBox Picture1 
      Height          =   6615
      Left            =   0
      ScaleHeight     =   6555
      ScaleWidth      =   7995
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      Begin VB.CommandButton btnadd 
         BackColor       =   &H8000000E&
         DisabledPicture =   "Secuirity_Question.frx":0ECA
         Height          =   390
         Left            =   3810
         MouseIcon       =   "Secuirity_Question.frx":1563
         MousePointer    =   99  'Custom
         Picture         =   "Secuirity_Question.frx":16B5
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2265
         Width           =   1230
      End
      Begin VB.CommandButton btnsave 
         BackColor       =   &H8000000E&
         DisabledPicture =   "Secuirity_Question.frx":1D4E
         Height          =   390
         Left            =   5130
         MouseIcon       =   "Secuirity_Question.frx":2401
         MousePointer    =   99  'Custom
         Picture         =   "Secuirity_Question.frx":2553
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2265
         Width           =   1245
      End
      Begin VB.CommandButton btndel 
         DisabledPicture =   "Secuirity_Question.frx":2C06
         Height          =   390
         Left            =   6450
         MouseIcon       =   "Secuirity_Question.frx":3335
         MousePointer    =   99  'Custom
         Picture         =   "Secuirity_Question.frx":3487
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2265
         Width           =   1400
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1890
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   480
         Width           =   1335
      End
      Begin RichTextLib.RichTextBox rtf1 
         Height          =   855
         Left            =   1890
         TabIndex        =   5
         Top             =   1200
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   1508
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"Secuirity_Question.frx":3BB6
      End
      Begin MSAdodcLib.Adodc seqq 
         Height          =   375
         Left            =   4050
         Top             =   3360
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
         Connect         =   "Provider=MSDAORA.1;User ID=sts/sts;Persist Security Info=True"
         OLEDBString     =   "Provider=MSDAORA.1;User ID=sts/sts;Persist Security Info=True"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from secQues"
         Caption         =   "seqq"
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
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "Secuirity_Question.frx":3C38
         Height          =   3735
         Left            =   0
         TabIndex        =   6
         Top             =   2850
         Width           =   7950
         _ExtentX        =   14023
         _ExtentY        =   6588
         _Version        =   393216
         AllowUpdate     =   0   'False
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   19
         FormatLocked    =   -1  'True
         AllowAddNew     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "SNO"
            Caption         =   "    S.no"
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
            DataField       =   "QUES"
            Caption         =   "Question"
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
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   6765.166
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Serial No:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   330
         TabIndex        =   8
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Question:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   330
         TabIndex        =   7
         Top             =   1200
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Security_Question"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function auto_gen_id()
Set r1 = New ADODB.Recordset
sql = "select count(*) from secques "
Set r1 = c1.Execute(sql)
If IsNull(r1.Fields(0)) Or r.Fields(0) = 0 Then
 Combo1.Text = 1
Else
 Combo1.Text = r.Fields(0) + 1
End If
End Function

Private Sub btnadd_Click()
auto_gen_id
seqq.RecordSource = "select * from secques"
seqq.Refresh
rtf1.Text = ""
btnsave.Enabled = True
End Sub

Private Sub btndel_Click()
Set r = New ADODB.Recordset
If Trim(Combo1.Text) = "" Then
MsgBox "Select Correct number", vbOKOnly + vbInformation, "Invalid Serial No"
Else
sql = "delete from secques where sno=" & Val(Combo1.Text) & ""
Set r = c.Execute(sql)
MsgBox "SuccessFully deleted", vbInformation, ""
End If
seqq.Refresh
Form_Load
End Sub

Private Sub btnsave_Click()
Set r = New ADODB.Recordset
If Trim(rtf1.Text) = "" Then
MsgBox "Insert Question first ", vbOKOnly + vbInformation, "Empty Question"
Else
sql = "insert into secques values(" & Combo1.Text & ",'" & rtf1.Text & "')"
Set r = c.Execute(sql)
MsgBox "SuccessFully saved", vbInformation, ""
r.Close
End If
seqq.Refresh
Form_Load
End Sub

Private Sub Combo1_Click()
Set r = New ADODB.Recordset
Set r = c.Execute("select * from secques where sno=" & Val(Combo1.Text) & " ")
If IsNull(r.Fields(0)) = False Then
 rtf1.Text = r.Fields(1)
 seqq.RecordSource = "select * from secques where sno=" & Val(Combo1.Text) & ""
 seqq.Refresh
End If
End Sub

Private Sub DataGrid2_Click()
On Error Resume Next
Set r = New ADODB.Recordset
Set r = c.Execute("select * from secques where sno=" & Val(DataGrid2.Text) & " or upper(ques)='" & UCase(DataGrid2.Text) & "' ")
If r.EOF = False Then
 rtf1.Text = r.Fields(1)
 Combo1.Text = r.Fields(0)
End If
End Sub

Private Sub Form_Load()
conn
CenterForm Me
btnsave.Enabled = False
Combo1.Clear
Set r = New ADODB.Recordset
Set r = c.Execute("select *from secQues")

seqq.RecordSource = "select * from secques"
seqq.Refresh

If IsNull(r.Fields(0)) = False Then
 While r.EOF = False
  Combo1.AddItem r.Fields(0)
  r.MoveNext
 Wend
End If
rtf1.Text = ""

End Sub

Private Sub picX_Click()
Unload Me
End Sub
