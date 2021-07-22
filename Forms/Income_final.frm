VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmExpense 
   BackColor       =   &H00808080&
   Caption         =   "Expense"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8655
   Icon            =   "Income_final.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9495
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   5400
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
      Connect         =   "Provider=MSDAORA.1;Password=Sts;User ID=Sts;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;Password=Sts;User ID=Sts;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from exp order by EX_NO"
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
   Begin MSDataGridLib.DataGrid dg1 
      Bindings        =   "Income_final.frx":6032
      Height          =   4455
      Left            =   0
      TabIndex        =   1
      Top             =   5040
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   7858
      _Version        =   393216
      AllowUpdate     =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   22
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
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "EX_NO"
         Caption         =   "No"
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
         DataField       =   "EX_WHERE"
         Caption         =   "From Where"
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
         DataField       =   "EX_REASON"
         Caption         =   "Reason"
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
         DataField       =   "EX_AMT"
         Caption         =   "Amount"
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
         DataField       =   "EX_DATE"
         Caption         =   "Date"
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
            ColumnWidth     =   524.976
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1830.047
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3734.929
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1170.142
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Height          =   3555
      Left            =   0
      ScaleHeight     =   3495
      ScaleWidth      =   8595
      TabIndex        =   2
      Top             =   1440
      Width           =   8655
      Begin VB.TextBox n4 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1280
         Left            =   5900
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Text            =   "Income_final.frx":6047
         Top             =   420
         Width           =   2535
      End
      Begin VB.TextBox n5 
         BackColor       =   &H00C0C0C0&
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
         Left            =   5900
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1980
         Width           =   1320
      End
      Begin VB.TextBox n3 
         BackColor       =   &H00C0C0C0&
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
         Left            =   1725
         TabIndex        =   11
         Text            =   "150.60"
         Top             =   1980
         Width           =   1800
      End
      Begin VB.TextBox n2 
         BackColor       =   &H00C0C0C0&
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
         Left            =   1725
         MaxLength       =   30
         TabIndex        =   10
         Text            =   "Mukesh Kumar Sharma"
         Top             =   1110
         Width           =   2880
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   960
         Left            =   0
         TabIndex        =   3
         Top             =   2640
         Width           =   8655
         Begin VB.CommandButton Command5 
            Caption         =   "Report"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   6600
            MouseIcon       =   "Income_final.frx":6059
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Back"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   5040
            MouseIcon       =   "Income_final.frx":61AB
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   3480
            MouseIcon       =   "Income_final.frx":62FD
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Save Entry"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   1920
            MouseIcon       =   "Income_final.frx":644F
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton Command1 
            Caption         =   "New Entry"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   360
            MouseIcon       =   "Income_final.frx":65A1
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   240
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
            TabIndex        =   4
            Top             =   720
            Width           =   1455
         End
         Begin VB.Shape Shape1 
            FillColor       =   &H00404040&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   480
            Shape           =   4  'Rounded Rectangle
            Top             =   330
            Width           =   1215
         End
         Begin VB.Shape Shape3 
            FillColor       =   &H00404040&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   2055
            Shape           =   4  'Rounded Rectangle
            Top             =   330
            Width           =   1215
         End
         Begin VB.Shape Shape6 
            FillColor       =   &H00404040&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   3615
            Shape           =   4  'Rounded Rectangle
            Top             =   345
            Width           =   1215
         End
         Begin VB.Shape Shape7 
            FillColor       =   &H00404040&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   6735
            Shape           =   4  'Rounded Rectangle
            Top             =   345
            Width           =   1455
         End
         Begin VB.Shape Shape9 
            FillColor       =   &H00404040&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   5175
            Shape           =   4  'Rounded Rectangle
            Top             =   345
            Width           =   1215
         End
      End
      Begin VB.Label n1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
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
         Left            =   1725
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Serial No :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   390
         TabIndex        =   18
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date     :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4860
         TabIndex        =   17
         Top             =   1920
         Width           =   840
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount   :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   390
         TabIndex        =   16
         Top             =   1920
         Width           =   1065
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reason :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4785
         TabIndex        =   15
         Top             =   375
         Width           =   885
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Paid To    :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   390
         TabIndex        =   14
         Top             =   1080
         Width           =   1080
      End
   End
   Begin VB.Image Image2 
      Height          =   735
      Left            =   0
      Picture         =   "Income_final.frx":66F3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1155
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "-:  EXPENSE - DETAILS  :-"
      BeginProperty Font 
         Name            =   "FujiyamaExtraBold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   1080
      Width           =   5415
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   0
      Picture         =   "Income_final.frx":B1B3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8655
   End
End
Attribute VB_Name = "FrmExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
ClearA
IdGenerate
n2.SetFocus
Command2.Enabled = True
End Sub
Public Sub IdGenerate()
Set r = New ADODB.Recordset
Set r = c.Execute("select count(*) from exp")
If r.EOF = True Then
n1.Caption = 1
Else
n1.Caption = r.Fields(0) + 1
End If
End Sub

Private Sub Command2_Click()
If n1.Caption = "" Then
 MsgBox "Click on NEW ENTRY For Adding new Record", vbInformation + vbOKOnly, ""
Exit Sub
End If
If Trim(n2.Text) = "" Then
 MsgBox "Enter The Name To Whom Amount is To be Paid..", vbCritical + vbOKOnly, ""
 n2.SetFocus
Exit Sub
ElseIf Trim(n3.Text) = "" Or Val(n3.Text) = 0 Then
 MsgBox "Enter The Amount is To be Paid..", vbCritical + vbOKOnly, ""
 n3.SetFocus
Exit Sub
ElseIf Trim(n4.Text) = "" Then
MsgBox "Enter The Reason for Paying Amount..", vbCritical + vbOKOnly, ""
 n4.SetFocus
Exit Sub
ElseIf Trim(n5.Text) = "" Then
MsgBox "Enter The date on Which Amount is Paid.", vbCritical + vbOKOnly, ""
 n5.SetFocus
Exit Sub
End If
c.Execute ("insert into exp values(" & Val(n1.Caption) & ",'" & Trim(n2.Text) & "','" & Trim(n4.Text) & "'," & Val(n3.Text) & ",'" & Format(n5.Text, "dd-mmm-yyyy") & "')")
MsgBox "SuccessFully Added ..", vbInformation + vbOKOnly, ""
Adodc1.Refresh
ClearA
End Sub

Private Sub Command3_Click()
ClearA
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
FrmReportMain.Show
End Sub

Private Sub Form_Load()
Me.Top = 50
conn
ClearA
End Sub
Public Sub ClearA()
n1.Caption = ""
n2.Text = ""
n3.Text = ""
n4.Text = ""
n5.Text = ""
cld1.Visible = False
Command2.Enabled = False
n5.Text = Format(Date, "DD-MMM-YYYY")
End Sub

Private Sub n2_KeyPress(KeyAscii As Integer)
 If (((KeyAscii >= 65) And (KeyAscii <= 90)) Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or (KeyAscii = 8) Or (KeyAscii = 32)) Then
        n2.SetFocus
  ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        n3.SetFocus
  Else
       KeyAscii = 0
        MsgBox "Name Can Contain Only Characters", vbInformation + vbOKOnly, ""
  End If
End Sub

Private Sub n3_KeyPress(KeyAscii As Integer)
If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Then
        n3.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   n4.SetFocus
  Else
   KeyAscii = 0
  End If
End Sub

Private Sub n4_Change()
If (((KeyAscii >= 65) And (KeyAscii <= 90)) Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or (KeyAscii = 8) Or (KeyAscii = 32)) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 44 Then
   n4.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   n5.SetFocus
  Else
   KeyAscii = 0
  End If
End Sub

Private Sub n5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 KeyAscii = 0
 Command2_Click
End If
End Sub
