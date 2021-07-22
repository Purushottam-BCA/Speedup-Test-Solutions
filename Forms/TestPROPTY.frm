VERSION 5.00
Object = "{08654D78-6636-11D3-87BF-B4980CC10374}#2.0#0"; "MyEllipticButton.ocx"
Begin VB.Form FrmTestPrpt1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MCQ Test Properties"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11190
   Icon            =   "TestPROPTY.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   11190
   Begin MyEllipticButton.EllipticButton btn3 
      Height          =   2415
      Left            =   7440
      TabIndex        =   0
      Top             =   2040
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   4260
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "TestPROPTY.frx":0ECA
      DisabledPicture =   "TestPROPTY.frx":2325
      DownPicture     =   "TestPROPTY.frx":2341
      MousePointer    =   99
      MouseIcon       =   "TestPROPTY.frx":235D
      Caption         =   ""
   End
   Begin MyEllipticButton.EllipticButton btn1 
      Height          =   2415
      Left            =   1200
      TabIndex        =   2
      Top             =   2040
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   4260
      BackColor       =   -2147483637
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "TestPROPTY.frx":24BF
      DisabledPicture =   "TestPROPTY.frx":3A94
      DownPicture     =   "TestPROPTY.frx":3AB0
      MousePointer    =   99
      MouseIcon       =   "TestPROPTY.frx":3ACC
      Caption         =   ""
   End
   Begin MyEllipticButton.EllipticButton btn2 
      Height          =   2415
      Left            =   4320
      TabIndex        =   1
      Top             =   2040
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   4260
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "TestPROPTY.frx":3C2E
      DisabledPicture =   "TestPROPTY.frx":4FAF
      DownPicture     =   "TestPROPTY.frx":4FCB
      MousePointer    =   99
      MouseIcon       =   "TestPROPTY.frx":4FE7
      Caption         =   ""
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   5  'Dash-Dot-Dot
      BorderWidth     =   2
      Height          =   2295
      Left            =   7490
      Shape           =   2  'Oval
      Top             =   2091
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   5  'Dash-Dot-Dot
      BorderWidth     =   2
      Height          =   2295
      Left            =   1250
      Shape           =   2  'Oval
      Top             =   2085
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   5  'Dash-Dot-Dot
      BorderWidth     =   2
      Height          =   2300
      Left            =   4365
      Shape           =   2  'Oval
      Top             =   2090
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Full Length Test"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   285
      Left            =   7920
      MouseIcon       =   "TestPROPTY.frx":5149
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   4680
      Width           =   1545
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject Wise Test"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   285
      Left            =   4680
      MouseIcon       =   "TestPROPTY.frx":529B
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   4680
      Width           =   1740
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Topic Wise Test"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   285
      Left            =   1560
      MouseIcon       =   "TestPROPTY.frx":53ED
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   4680
      Width           =   1560
   End
End
Attribute VB_Name = "FrmTestPrpt1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn1_Click()
choiceTST = 1
FrmTestPrpt2.Show 1, MDI
End Sub

Private Sub btn1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.FontBold = True
btn1.BackColor = &H80000005
Shape2.Visible = True
End Sub

Private Sub btn2_Click() 'Subject Wise
choiceTST = 2
FrmTestPrpt2.Show 1, MDI
End Sub

Private Sub btn2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.FontBold = True
btn2.BackColor = &H80000005
Shape1.Visible = True
End Sub

Private Sub btn3_Click() 'Full length
choiceTST = 3
FrmTestPrpt2.Show 1, MDI
End Sub

Private Sub btn3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.FontBold = True
btn3.BackColor = &H80000005
Shape3.Visible = True
End Sub

Private Sub Form_Load()
CenterForm Me
conn
btn1.BackColor = &H8000000B
btn2.BackColor = &H8000000B
btn3.BackColor = &H8000000B
choiceTST = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.FontBold = False
Label3.FontBold = False
Label2.FontBold = False
btn1.BackColor = &H80000005
btn2.BackColor = &H80000005
btn3.BackColor = &H80000005
Shape1.Visible = False
Shape2.Visible = False
Shape3.Visible = False
End Sub

Private Sub Label2_Click()
btn1_Click
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.FontBold = True
btn1.BackColor = &H80000005
Shape2.Visible = True
End Sub

Private Sub Label3_Click()
btn2_Click
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.FontBold = True
btn2.BackColor = &H80000005
Shape1.Visible = True
End Sub

Private Sub Label4_Click()
btn3_Click
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.FontBold = True
btn3.BackColor = &H80000005
Shape3.Visible = True
End Sub

