VERSION 5.00
Begin VB.Form login_Admin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Log-In"
   ClientHeight    =   6015
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
   Moveable        =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox vkCheck1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4320
      MouseIcon       =   "Admin_login.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   4150
      Width           =   2175
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
      Left            =   4320
      MouseIcon       =   "Admin_login.frx":0152
      MousePointer    =   99  'Custom
      Picture         =   "Admin_login.frx":02A4
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4680
      Width           =   3255
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
      Height          =   465
      Left            =   4335
      Picture         =   "Admin_login.frx":0DA2
      ScaleHeight     =   465
      ScaleWidth      =   495
      TabIndex        =   7
      Top             =   3380
      Width           =   495
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
      Height          =   470
      Left            =   4335
      Picture         =   "Admin_login.frx":18AC
      ScaleHeight     =   465
      ScaleWidth      =   495
      TabIndex        =   6
      Top             =   2295
      Width           =   495
   End
   Begin VB.TextBox pswd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   380
      Left            =   4855
      MaxLength       =   20
      TabIndex        =   1
      Text            =   "jhhbhbhbhb"
      Top             =   3450
      Width           =   2700
   End
   Begin VB.TextBox userID 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
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
      Left            =   4900
      MaxLength       =   14
      TabIndex        =   0
      Text            =   "gghghhjvjhvhvhv"
      Top             =   2355
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6075
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4095
      Begin VB.Image Image3 
         Height          =   1935
         Left            =   360
         Picture         =   "Admin_login.frx":1E5D
         Stretch         =   -1  'True
         Top             =   480
         Width           =   3315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "  SPEEDUP   TEST SOLUTIONS"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3255
         Left            =   240
         TabIndex        =   5
         Top             =   2850
         Width           =   3495
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   615
      Left            =   4200
      TabIndex        =   8
      Top             =   5400
      Width           =   3855
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Create User"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   105
         MouseIcon       =   "Admin_login.frx":691D
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   120
         Width           =   1170
      End
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2040
         MouseIcon       =   "Admin_login.frx":6A6F
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   120
         Width           =   1620
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   7560
      MouseIcon       =   "Admin_login.frx":6BC1
      MousePointer    =   99  'Custom
      Picture         =   "Admin_login.frx":6D13
      Top             =   -30
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login ID"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   4320
      TabIndex        =   3
      Top             =   1920
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   4350
      TabIndex        =   2
      Top             =   2985
      Width           =   1080
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0E0FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   4320
      Top             =   2280
      Width           =   3255
   End
   Begin VB.Shape pswd_shape 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   4320
      Top             =   3360
      Width           =   3255
   End
   Begin VB.Image Image2 
      Height          =   1575
      Left            =   3960
      Picture         =   "Admin_login.frx":7334
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "login_Admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim opt As Integer

Private Sub Form_Load() ' working
conn
CreateRoundRectFromWindow Me
Frm_Starting = 2
userID.Text = ""
pswd.Text = ""
EMP_login_reg_no = ""
Stu_login_reg_no = ""
admin_login_reg_no = ""
Current_Logged_ID = ""
CenterForm Me
 pswd.FontName = "Wingdings"
 pswd.FontBold = False
 pswd.FontSize = 11
 pswd.PasswordChar = "l"
 opt = 1
 vkCheck1.Value = vbUnchecked
 Set r = c.Execute("select count(*) from admin_login")
 If r.Fields(0) > 0 Then
  Label4.Visible = False
 Else
  Label4.Visible = True
 End If
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frgt_pass.FontUnderline = True
Label4.FontUnderline = True
End Sub

Private Sub frgt_pass_Click()
userID.Text = ""
pswd.Text = ""
Owner1 = "Admin"
frgt_pass.FontUnderline = True
frmFrgtPswd.Show
MDI.Enabled = True
End Sub

Private Sub frgt_pass_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frgt_pass.FontUnderline = False
End Sub

Private Sub Image1_Click()
Unload Me
FrmSelectUser.Show
End Sub

Private Sub Label4_Click()
userID.Text = ""
pswd.Text = ""
Frm_Starting = 1
FrmEmpMaster.Show
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.FontUnderline = False
End Sub

Private Sub LogIn_btn_Click()
On Error Resume Next
If Trim(userID.Text) = "" Then
  MsgBox " Enter Admin LogIn id", vbInformation + vbOKOnly, "Empty User ID"
  userID.SetFocus
  Exit Sub
ElseIf Trim(pswd.Text) = "" Then
  MsgBox "Password Missing..", vbExclamation + vbOKOnly, "Empty User ID"
  pswd.SetFocus
  Exit Sub
End If
   Set r1 = New ADODB.Recordset
   Set r = New ADODB.Recordset
   Set r = c.Execute("select count(*) from admin_login")
  If r.Fields(0) = 1 Then
   Set r1 = c1.Execute("select * from admin_login")
   If (UCase(Trim(userID.Text)) = UCase(r1.Fields(0)) And UCase(Trim(pswd.Text)) = UCase(r1.Fields(2))) Then
     admin_login_reg_no = r1.Fields(1)
     Frm_Starting = 2
     Me.Hide
     admin_dash.Show
    Exit Sub
   Else
    MsgBox "Invalid Id/Password...", vbCritical + vbOKOnly, "Wrong Id/Password"
    userID.Text = ""
    pswd.Text = ""
    userID.SetFocus
   Exit Sub
   End If
  ElseIf r.Fields(0) > 1 Then
   Set r1 = c1.Execute("select * from admin_login")
   While r1.EOF = False
   If (UCase(Trim(userID.Text)) = UCase(r1.Fields(0)) And UCase(Trim(pswd.Text)) = UCase(r1.Fields(2))) Then
    admin_login_reg_no = r1.Fields(1)
    Frm_Starting = 2
    Me.Hide
    admin_dash.Show
   Exit Sub
   End If
  r1.MoveNext
  Wend
 If r1.EOF = True Then
   MsgBox "Invalid Id/Password...", vbCritical + vbOKOnly, "Wrong Id/Password"
   userID.Text = ""
   pswd.Text = ""
   userID.SetFocus
   Exit Sub
 End If
 Else
   MsgBox "No Admin Available For Application. Click On Create New To Create Admin Account..", vbCritical + vbOKOnly, "Wrong Id/Password"
   userID.Text = ""
   pswd.Text = ""
   userID.SetFocus
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
'If KeyAscii >= 65 And KeyAscii <= 90 Then
'
'ElseIf KeyAscii >= 97 And KeyAscii <= 122 Then
' KeyAscii = KeyAscii - 32
'Endif
If KeyAscii = 13 Then
 KeyAscii = 0
 LogIn_btn_Click
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

Private Sub vkLabel1_MouseDblClick(Button As MouseButtonConstants, Shift As Integer, Control As Integer, X As Long, Y As Long)
MsgBox "Gello"
End Sub
