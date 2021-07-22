VERSION 5.00
Object = "{BF38D12B-22A9-4B10-B26E-019F2B5F9C22}#1.0#0"; "AniGif.ocx"
Begin VB.Form emp_id_pass 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7425
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   11190
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0FF&
   FillStyle       =   0  'Solid
   Icon            =   "ID_Pass_Generator.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   11190
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   1680
      ScaleHeight     =   4215
      ScaleWidth      =   7695
      TabIndex        =   0
      Top             =   1680
      Width           =   7695
      Begin VB.CommandButton save 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         MouseIcon       =   "ID_Pass_Generator.frx":09EA
         MousePointer    =   99  'Custom
         Picture         =   "ID_Pass_Generator.frx":0B3C
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Click Ok To Save."
         Top             =   3420
         Width           =   1425
      End
      Begin VB.TextBox answer 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2400
         TabIndex        =   2
         Top             =   2640
         Width           =   4935
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
         ItemData        =   "ID_Pass_Generator.frx":1506
         Left            =   2400
         List            =   "ID_Pass_Generator.frx":1508
         MouseIcon       =   "ID_Pass_Generator.frx":150A
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1920
         Width           =   4935
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Height          =   3975
         Left            =   120
         Top             =   120
         Width           =   7455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Security Question  :"
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
         Left            =   255
         TabIndex        =   10
         Top             =   1920
         Width           =   1890
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Secuirity Answer    : "
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
         Left            =   255
         TabIndex        =   9
         Top             =   2640
         Width           =   1965
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LogIn ID                 : "
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
         Left            =   255
         TabIndex        =   8
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password               :"
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
         Left            =   255
         TabIndex        =   7
         Top             =   1200
         Width           =   1890
      End
      Begin VB.Label log_id 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   6
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label Password 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label id 
         BackColor       =   &H80000010&
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Role 
         BackColor       =   &H8000000D&
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin Project1.PictureG PictureG2 
         Height          =   480
         Left            =   5280
         Top             =   1200
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   847
         GIF             =   "ID_Pass_Generator.frx":165C
         Stretch         =   2
         Mirror          =   1
      End
      Begin Project1.PictureG PictureG1 
         Height          =   480
         Left            =   5280
         Top             =   600
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   847
         GIF             =   "ID_Pass_Generator.frx":710E
         Stretch         =   2
         Mirror          =   1
      End
   End
End
Attribute VB_Name = "emp_id_pass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
conn
CenterForm Me
Combo1.Clear
Set r = New ADODB.Recordset
Set r = c.Execute("select initcap(ques) from SECQUES ")
While r.EOF = False
Combo1.AddItem r.Fields(0)
r.MoveNext
Wend
End Sub

Private Sub save_Click()
If id.Caption = "" Or log_id.Caption = "" Or Password.Caption = "" Then
 MsgBox "Cannot Get Data From registration Info", vbInformation, ""
Exit Sub
ElseIf Trim(Combo1.Text) = "" Then
MsgBox "Select security Question !!" & vbCrLf & "It Will help in Future while recovering password", vbInformation + vbOKOnly, "No Question selected"
Combo1.SetFocus
Exit Sub
ElseIf Trim(answer.Text) = "" Then
 MsgBox "Cann't Leave it Blank", vbInformation + vbOKOnly, "Fill Answer"
 answer.SetFocus
 Exit Sub
Else
If UCase(Role.Caption) = "ADMIN" Then
c.Execute ("insert into admin_login values('" & log_id.Caption & "','" & id.Caption & "','" & Password.Caption & "','" & Combo1.Text & "','" & answer.Text & "')")
Else
c.Execute ("insert into emp_login values('" & id.Caption & "','" & log_id.Caption & "','" & Password.Caption & "','" & Combo1.Text & "','" & answer.Text & "')")
End If
MsgBox vbCrLf & "SuccessFully Registered.." & vbCrLf, vbInformation + vbOKOnly, "New Registration"
Role.Caption = ""
id.Caption = ""
Unload Me
End If
End Sub

