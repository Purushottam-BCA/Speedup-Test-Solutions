VERSION 5.00
Begin VB.Form login_new 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Log-In"
   ClientHeight    =   5790
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
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MouseIcon       =   "login_new.frx":0000
   Moveable        =   0   'False
   Picture         =   "login_new.frx":0152
   ScaleHeight     =   5790
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox vkCheck1 
      BackColor       =   &H00404040&
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
      Left            =   4440
      MouseIcon       =   "login_new.frx":7923
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Frame vkFrame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   6000
      TabIndex        =   7
      Top             =   5235
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
         Left            =   240
         MouseIcon       =   "login_new.frx":7A75
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   120
         Width           =   1620
      End
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
      Left            =   4440
      MouseIcon       =   "login_new.frx":7BC7
      MousePointer    =   99  'Custom
      Picture         =   "login_new.frx":7D19
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4440
      Width           =   3255
   End
   Begin VB.TextBox pswd 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   380
      Left            =   4965
      MaxLength       =   10
      TabIndex        =   1
      Text            =   "hbhbjhbhjb"
      Top             =   2730
      Width           =   2655
   End
   Begin VB.TextBox userID 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   4965
      TabIndex        =   0
      Text            =   "hjhsbjbjjjdbjbj"
      Top             =   1530
      Width           =   2655
   End
   Begin VB.PictureBox Picture2 
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
      Height          =   440
      Left            =   4455
      Picture         =   "login_new.frx":8817
      ScaleHeight     =   435
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   1455
      Width           =   495
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
      Height          =   460
      Left            =   4455
      Picture         =   "login_new.frx":8DC8
      ScaleHeight     =   465
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   2655
      Width           =   495
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
      Left            =   4440
      TabIndex        =   5
      Top             =   1080
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
      Left            =   4470
      TabIndex        =   4
      Top             =   2265
      Width           =   930
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BorderColor     =   &H00400040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   4440
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   7530
      MouseIcon       =   "login_new.frx":98D2
      MousePointer    =   99  'Custom
      Picture         =   "login_new.frx":9A24
      Top             =   15
      Width           =   480
   End
   Begin VB.Shape pswd_shape 
      BackColor       =   &H00E0E0E0&
      BorderColor     =   &H00400040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   4440
      Top             =   2640
      Width           =   3255
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   5895
      Left            =   4200
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "login_new"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim opt As Integer
Private Sub Form_Load() ' working
conn
CreateRoundRectFromWindow Me
Stu_login_reg_no = ""
CenterForm Me
pswd.FontName = "Wingdings"
 pswd.FontBold = False
 pswd.FontSize = 11
 pswd.PasswordChar = "l"
opt = 1
End Sub

Private Sub frgt_pass_Click()
userID.Text = ""
pswd.Text = ""
Owner1 = "Student"
frmFrgtPswd.Show
End Sub

Private Sub frgt_pass_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frgt_pass.FontUnderline = False
frgt_pass.ForeColor = &H80FFFF
frgt_pass.FontSize = 12
End Sub

Private Sub Image1_Click()
Unload Me
FrmSelectUser.Show
End Sub

Private Sub LogIn_btn_Click()
If Trim(userID.Text) = "" Then
  MsgBox " Enter User Login ID..", vbCritical + vbOKOnly, "Empty User ID"
  userID.SetFocus
  Exit Sub
ElseIf Trim(pswd.Text) = "" Then
  MsgBox "Password Missing..", vbExclamation + vbOKOnly, "Empty User ID"
   pswd.SetFocus
   Exit Sub
Else
   Set r1 = New ADODB.Recordset
   Set r1 = c1.Execute("select * from stud_login")
  While r1.EOF = False
   If (UCase(Trim(userID.Text)) = UCase(r1.Fields(1)) And UCase(Trim(pswd.Text)) = UCase(r1.Fields(2))) Then
    Stu_login_reg_no = r1.Fields(0)
    Current_Logged_ID = r1.Fields(0)
    Me.Hide
    stu_dash.Show
    Exit Sub
   End If
  r1.MoveNext
  Wend
 If r1.EOF = True Then
   MsgBox "Invalid Id Or Password..", vbCritical + vbOKOnly, "Wrong Id/Password"
   pswd.SetFocus
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
If KeyAscii >= 65 And KeyAscii <= 90 Then
 KeyAscii = KeyAscii + 32
ElseIf KeyAscii >= 97 And KeyAscii <= 122 Then
 KeyAscii = KeyAscii - 32
ElseIf KeyAscii = 13 Then
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

Private Sub vkCheck1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 LogIn_btn_Click
End If
End Sub

Private Sub vkFrame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frgt_pass.FontUnderline = True
frgt_pass.ForeColor = vbWhite
frgt_pass.FontSize = 11
End Sub
