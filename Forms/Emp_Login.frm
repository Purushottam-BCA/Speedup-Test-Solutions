VERSION 5.00
Begin VB.Form login_EMP 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Log-In"
   ClientHeight    =   5835
   ClientLeft      =   -60
   ClientTop       =   -105
   ClientWidth     =   7980
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Cambria"
      Size            =   12.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Emp_Login.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   Picture         =   "Emp_Login.frx":6062
   ScaleHeight     =   5835
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox vkCheck1 
      BackColor       =   &H80000006&
      Caption         =   "Show Password"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   4320
      MouseIcon       =   "Emp_Login.frx":C9E9
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   3960
      Width           =   2175
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4335
      Picture         =   "Emp_Login.frx":CB3B
      ScaleHeight     =   465
      ScaleWidth      =   495
      TabIndex        =   6
      Top             =   2050
      Width           =   495
   End
   Begin VB.CommandButton LogIn_btn 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4340
      MouseIcon       =   "Emp_Login.frx":D0EC
      MousePointer    =   99  'Custom
      Picture         =   "Emp_Login.frx":D23E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4560
      Width           =   3255
   End
   Begin VB.TextBox pswd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   380
      Left            =   4850
      MaxLength       =   30
      TabIndex        =   1
      Text            =   "hjhjhjhjhj"
      Top             =   3200
      Width           =   2700
   End
   Begin VB.TextBox userID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   4850
      MaxLength       =   30
      TabIndex        =   0
      Text            =   "hvjhbhjhj"
      Top             =   2135
      Width           =   2700
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   8
      Top             =   5400
      Width           =   2055
      Begin VB.Label frgt_pass 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forgot Password"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   120
         MouseIcon       =   "Emp_Login.frx":DD3C
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   0
         Width           =   1620
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   470
      Left            =   4335
      Picture         =   "Emp_Login.frx":DE8E
      ScaleHeight     =   465
      ScaleWidth      =   495
      TabIndex        =   5
      Top             =   3135
      Width           =   495
   End
   Begin VB.Shape pswd_shape 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   4320
      Top             =   3120
      Width           =   3255
   End
   Begin VB.Image Image2 
      Height          =   1300
      Left            =   4800
      Picture         =   "Emp_Login.frx":E998
      Stretch         =   -1  'True
      Top             =   180
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   7500
      MouseIcon       =   "Emp_Login.frx":13458
      MousePointer    =   99  'Custom
      Picture         =   "Emp_Login.frx":135AA
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   4320
      TabIndex        =   4
      Top             =   1680
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   4350
      TabIndex        =   3
      Top             =   2745
      Width           =   930
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   4320
      Top             =   2040
      Width           =   3255
   End
End
Attribute VB_Name = "login_EMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim opt As Integer
Private Sub Form_Load() ' working
conn
CreateRoundRectFromWindow Me
pswd.Text = ""
userID.Text = ""
EMP_login_reg_no = ""
admin_login_reg_no = ""
Stu_login_reg_no = ""
CenterForm Me
pswd.FontName = "Wingdings"
 pswd.FontBold = False
 pswd.FontSize = 11
 pswd.PasswordChar = "l"
opt = 1
userID.MaxLength = 12
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frgt_pass.FontUnderline = True
frgt_pass.ForeColor = &HE0E0E0
End Sub

Private Sub frgt_pass_Click()
userID.Text = ""
pswd.Text = ""
vkCheck1.Value = 0
Owner1 = "Emp"
frmFrgtPswd.Show
End Sub

Private Sub frgt_pass_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frgt_pass.FontUnderline = False
frgt_pass.ForeColor = &H80000014
End Sub

Private Sub Image1_Click()
Unload Me
FrmSelectUser.Show
End Sub

Private Sub LogIn_btn_Click()
On Error Resume Next
If Trim(userID.Text) = "" Then
  MsgBox " Enter User id", vbInformation + vbOKOnly, "Empty User ID"
  userID.SetFocus
ElseIf Trim(pswd.Text) = "" Then
   MsgBox "Password Missing..", vbExclamation + vbOKOnly, "Empty User ID"
   pswd.SetFocus
Else
   Set r1 = New ADODB.Recordset
   Set r1 = c1.Execute("select * from emp_login")
   While r1.EOF = False
   If (UCase(Trim(userID.Text)) = UCase(r1.Fields(1)) And UCase(Trim(pswd.Text)) = UCase(r1.Fields(2))) Then
    EMP_login_reg_no = r1.Fields(0)
    vkCheck1.Value = 0
    Stu_login_reg_no = ""
    admin_login_reg_no = ""
    Me.Hide
   emp_dash.Show
   Exit Sub
   End If
  r1.MoveNext
  Wend
 If r1.EOF = True Then
   MsgBox "Invalid Id/Password", vbInformation + vbOKOnly, "Wrong Id/Password"
   Exit Sub
 End If
End If
End Sub

Private Sub Text1_GotFocus()
If Trim(UCase$(Text1.Text)) <> Trim(UCase$("Password")) Then
If vkCheck1.Value = vbChecked Then
  Text1.FontName = "Cambria"
  Text1.FontSize = 13
  Text1.FontBold = True
  Text1.PasswordChar = ""
Else
Text1.FontName = "Wingdings"
Text1.FontBold = False
Text1.FontSize = 11
Text1.PasswordChar = "l"
End If
Else
End If
End Sub

Private Sub pswd_KeyPress(KeyAscii As Integer)
If KeyAscii >= 65 And KeyAscii <= 90 Then
ElseIf KeyAscii >= 97 And KeyAscii <= 122 Then
 KeyAscii = KeyAscii - 32
ElseIf KeyAscii = 13 Then
KeyAscii = 0
 LogIn_btn_Click
 Exit Sub
End If
End Sub

Private Sub Role_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
userID.Text = ""
pswd.Text = ""
userID.SetFocus
End If
End Sub

Private Sub userID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 KeyAscii = 0
 pswd.SetFocus
End If
End Sub

Private Sub vkCheck1_Click() 'Show Password
If vkCheck1.Value = vbChecked Then
  pswd.FontName = "Cambria"
  pswd.FontSize = 13
  pswd.FontBold = True
  pswd.PasswordChar = ""
Else
 pswd.FontName = "Wingdings"
 pswd.FontBold = False
 pswd.FontSize = 11
 pswd.PasswordChar = "l"
End If
End Sub
