VERSION 5.00
Begin VB.Form ResumeTEST 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MCQ Test Resume"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9720
   Icon            =   "ResumeTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   9720
   Begin VB.CommandButton btnadd 
      BackColor       =   &H8000000E&
      DisabledPicture =   "ResumeTest.frx":0FA2
      Height          =   390
      Left            =   3930
      MouseIcon       =   "ResumeTest.frx":1655
      MousePointer    =   99  'Custom
      Picture         =   "ResumeTest.frx":17A7
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Return Back To Test"
      Top             =   2880
      Width           =   2065
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00808080&
      X1              =   80
      X2              =   9640
      Y1              =   1750
      Y2              =   1750
   End
   Begin VB.Label l4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "391"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7455
      TabIndex        =   12
      Top             =   2145
      Width           =   375
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   255
      Left            =   8840
      Shape           =   1  'Square
      Top             =   960
      Width           =   255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   8400
      X2              =   8400
      Y1              =   720
      Y2              =   2740
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Marked and Answered"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   600
      Left            =   7005
      TabIndex        =   11
      Top             =   1220
      Width           =   1380
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   2640
      X2              =   2640
      Y1              =   720
      Y2              =   2740
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   3960
      X2              =   3960
      Y1              =   720
      Y2              =   2740
   End
   Begin VB.Label l5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "391"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8770
      TabIndex        =   10
      Top             =   2145
      Width           =   375
   End
   Begin VB.Label l3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6240
      TabIndex        =   9
      Top             =   2145
      Width           =   135
   End
   Begin VB.Label l2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4800
      TabIndex        =   8
      Top             =   2145
      Width           =   135
   End
   Begin VB.Label l1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3135
      TabIndex        =   7
      Top             =   2145
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800080&
      Height          =   255
      Left            =   7490
      Shape           =   1  'Square
      Top             =   960
      Width           =   255
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00FF80FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF80FF&
      Height          =   255
      Left            =   6120
      Top             =   960
      Width           =   255
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      Height          =   255
      Left            =   4680
      Shape           =   1  'Square
      Top             =   960
      Width           =   255
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080FF80&
      Height          =   255
      Left            =   3120
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Not Seen"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   240
      Left            =   8595
      TabIndex        =   6
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Marked"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   240
      Left            =   5920
      TabIndex        =   5
      Top             =   1320
      Width           =   705
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Not Answered"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   240
      Left            =   4185
      TabIndex        =   4
      Top             =   1320
      Width           =   1275
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Answered"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   240
      Left            =   2835
      TabIndex        =   3
      Top             =   1320
      Width           =   915
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "MCQ Quiz 2019"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      TabIndex        =   2
      Top             =   2110
      Width           =   2460
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Sections"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   300
      Left            =   825
      TabIndex        =   1
      Top             =   1200
      Width           =   945
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808080&
      X1              =   6840
      X2              =   6840
      Y1              =   720
      Y2              =   2740
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   5640
      X2              =   5640
      Y1              =   720
      Y2              =   2740
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00505240&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      FillColor       =   &H00B0C0C0&
      FillStyle       =   0  'Solid
      Height          =   2055
      Left            =   60
      Top             =   720
      Width           =   9585
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Current Status"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   345
      Left            =   3510
      TabIndex        =   0
      Top             =   120
      Width           =   2445
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00BAC0C0&
      Height          =   615
      Left            =   45
      Top             =   0
      Width           =   9615
   End
End
Attribute VB_Name = "ResumeTEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nanswered2 As Integer
Dim answered2 As Integer

Private Sub Form_Load()
conn
CenterForm Me
Label3.Caption = "MCQ Quiz " & Year(Date)
Set r = New ADODB.Recordset
Set r = c.Execute("select count(*) from answerhold where USER_ANS<> 0  and BOOKMRK <>1 ")
l1.Caption = r.Fields(0)
Set r = New ADODB.Recordset
Set r = c.Execute("select count(*) from answerhold where USER_ANS= 0  and BOOKMRK=2 ")
l2.Caption = r.Fields(0)
Set r = New ADODB.Recordset
Set r = c.Execute("select count(*) from answerhold where USER_ANS= 0  and BOOKMRK=1 ")
l3.Caption = r.Fields(0)
Set r = New ADODB.Recordset
Set r = c.Execute("select count(*) from answerhold where USER_ANS<> 0  and BOOKMRK=1 ")
l4.Caption = r.Fields(0)
Set r = New ADODB.Recordset
Set r = c.Execute("select count(*) from answerhold where BOOKMRK=0 ")
l5.Caption = r.Fields(0)
End Sub

Private Sub btnadd_Click()
Unload Me
MCQ.Timer1.Enabled = True
MCQ.Timer2.Enabled = True
End Sub

Private Sub Form_Unload(cancel As Integer)
MCQ.Enabled = True
End Sub
