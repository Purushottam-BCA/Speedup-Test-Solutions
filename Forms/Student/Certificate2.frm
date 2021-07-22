VERSION 5.00
Begin VB.Form FrmCerti2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Certificate [UnRegistered]"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9720
   Icon            =   "Certificate2.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Certificate2.frx":0FA2
   ScaleHeight     =   7680
   ScaleWidth      =   9720
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
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
      Left            =   7920
      MouseIcon       =   "Certificate2.frx":11494
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   7200
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   7080
      Width           =   9735
      Begin VB.CommandButton Command1 
         Caption         =   "Print"
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
         MouseIcon       =   "Certificate2.frx":115E6
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.Label Rdate 
      BackStyle       =   0  'Transparent
      Caption         =   "24-Dec-2018"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   2335
      TabIndex        =   6
      Top             =   5165
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   4200
      Picture         =   "Certificate2.frx":11738
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1035
   End
   Begin VB.Label SCore 
      BackStyle       =   0  'Transparent
      Caption         =   "73/100"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Top             =   4605
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2019 MCQ Test in Maths"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   4080
      Width           =   6015
   End
   Begin VB.Label UName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mukesh Kumar Sharma"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   3050
      Width           =   6015
   End
End
Attribute VB_Name = "FrmCerti2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Me.Height = 7650
Me.PrintForm
Me.Height = 8250
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
conn
Dim fl As Integer
CenterForm Me
fl = Summary_Test.l6.Caption
SCore.Caption = Summary_Test.l7.Caption & "/" & fl
Set r = New ADODB.Recordset
Set r = c.Execute("select initcap(rstud_nm) from rstud where rstud_reg_no='" & Stu_login_reg_no & "' ")
UName.Caption = r.Fields(0)
Rdate.Caption = Format(Date, "dd-mmm-yyyy")
End Sub

Private Sub Form_Unload(cancel As Integer)
Summary_Test.Enabled = True
End Sub
