VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F5E116E1-0563-11D8-AA80-000B6A0D10CB}#1.0#0"; "HookMenu.ocx"
Begin VB.Form emp_dash 
   BackColor       =   &H80000013&
   Caption         =   "Employee"
   ClientHeight    =   10230
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   20250
   ControlBox      =   0   'False
   Icon            =   "emp_dash.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   Picture         =   "emp_dash.frx":0EE2
   ScaleHeight     =   10921.6
   ScaleMode       =   0  'User
   ScaleWidth      =   20340
   WindowState     =   2  'Maximized
   Begin HookMenu.XpMenu XpMenu1 
      Left            =   4680
      Top             =   600
      _ExtentX        =   900
      _ExtentY        =   900
      BitmapSize      =   18
      BmpCount        =   15
      CheckBorderColor=   7021576
      SelMenuBorder   =   7021576
      SelMenuBackColor=   14073525
      SelMenuForeColor=   16646297
      SelCheckBackColor=   14791828
      MenuBorderColor =   6956042
      SeparatorColor  =   -2147483632
      MenuBackColor   =   14609903
      MenuForeColor   =   0
      CheckBackColor  =   15326939
      CheckForeColor  =   10027263
      DisabledMenuBorderColor=   -2147483632
      DisabledMenuBackColor=   15660791
      DisabledMenuForeColor=   -2147483631
      MenuBarBackColor=   15790320
      MenuPopupBackColor=   16777215
      ShortCutNormalColor=   0
      ShortCutSelectColor=   16646297
      ArrowNormalColor=   10027263
      ArrowSelectColor=   12484864
      ShadowColor     =   0
      Bmp:1           =   "emp_dash.frx":28C8B
      Mask:1          =   1909157
      Key:1           =   "#add_ques"
      Bmp:2           =   "emp_dash.frx":290CD
      Mask:2          =   15461863
      Key:2           =   "#mcq_s"
      Bmp:3           =   "emp_dash.frx":29467
      Mask:3          =   16777215
      Key:3           =   "#ques_bnk"
      Bmp:4           =   "emp_dash.frx":29CB9
      Mask:4          =   16449532
      Key:4           =   "#reg_stu"
      Bmp:5           =   "emp_dash.frx":2A07F
      Mask:5          =   10070681
      Key:5           =   "#srch_stu"
      Bmp:6           =   "emp_dash.frx":2A625
      Mask:6          =   1909157
      Key:6           =   "#create_qus_ppr"
      Bmp:7           =   "emp_dash.frx":2AA67
      Mask:7          =   14211288
      Key:7           =   "#view_stu"
      Bmp:8           =   "emp_dash.frx":2ADF9
      Mask:8          =   16776957
      Key:8           =   "#new_ord"
      Bmp:9           =   "emp_dash.frx":2B11B
      Mask:9          =   16777215
      Key:9           =   "#view_all_client"
      Bmp:10          =   "emp_dash.frx":2B6C1
      Mask:10         =   16777215
      Key:10          =   "#calc"
      Bmp:11          =   "emp_dash.frx":2BC67
      Mask:11         =   16777215
      Key:11          =   "#notepad"
      Bmp:12          =   "emp_dash.frx":2C20D
      Mask:12         =   16515071
      Key:12          =   "#emp_prof"
      Bmp:13          =   "emp_dash.frx":2C5DF
      Key:13          =   "#chng_pass"
      Bmp:14          =   "emp_dash.frx":2D067
      Key:14          =   "#lg_out"
      Bmp:15          =   "emp_dash.frx":2DAEF
      Mask:15         =   5405713
      Key:15          =   "#bnjkmkjnhb"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      ClipControls    =   0   'False
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1395
      ScaleWidth      =   3240
      TabIndex        =   12
      Top             =   120
      Width           =   3300
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "    Today"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1335
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   3185
         Begin VB.TextBox Text1 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   840
            TabIndex        =   17
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Left            =   1200
            TabIndex        =   18
            Top             =   -240
            Width           =   1215
         End
         Begin VB.Label sdate 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   840
            TabIndex        =   16
            Top             =   360
            Width           =   1275
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   840
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   540
         End
         Begin VB.Image Image1 
            Height          =   240
            Left            =   120
            Picture         =   "emp_dash.frx":2DDA9
            Top             =   0
            Width           =   240
         End
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4680
      Top             =   600
   End
   Begin MSComctlLib.StatusBar sb1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   9855
      Width           =   20250
      _ExtentX        =   35719
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5576
            MinWidth        =   5576
            Text            =   "         SpeedUp Test Solutions"
            TextSave        =   "         SpeedUp Test Solutions"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Text            =   "Login : ( User )"
            TextSave        =   "Login : ( User )"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Enabled         =   0   'False
            TextSave        =   "SCRL"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   8715
      Left            =   0
      ScaleHeight     =   8655
      ScaleWidth      =   3240
      TabIndex        =   0
      Top             =   1560
      Width           =   3300
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Log Out"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1380
         Left            =   1625
         MouseIcon       =   "emp_dash.frx":2E133
         MousePointer    =   99  'Custom
         Picture         =   "emp_dash.frx":2E43D
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "LogOut"
         Top             =   6880
         Width           =   1575
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "See Order Status"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1380
         Left            =   80
         MouseIcon       =   "emp_dash.frx":2F116
         MousePointer    =   99  'Custom
         Picture         =   "emp_dash.frx":2F268
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   6880
         Width           =   1575
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "User Profile"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1380
         Left            =   1620
         MouseIcon       =   "emp_dash.frx":2F96C
         MousePointer    =   99  'Custom
         Picture         =   "emp_dash.frx":2FABE
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Update Profile"
         Top             =   5535
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Add New Client"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1380
         Left            =   80
         MouseIcon       =   "emp_dash.frx":304B9
         MousePointer    =   99  'Custom
         Picture         =   "emp_dash.frx":3060B
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Register New Client"
         Top             =   5535
         Width           =   1575
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Change Password"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1420
         Left            =   1625
         MouseIcon       =   "emp_dash.frx":314ED
         MousePointer    =   99  'Custom
         Picture         =   "emp_dash.frx":3163F
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Change Password"
         Top             =   4150
         Width           =   1575
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "About Us"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1420
         Left            =   80
         MouseIcon       =   "emp_dash.frx":32509
         MousePointer    =   99  'Custom
         Picture         =   "emp_dash.frx":3265B
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "About Organisation"
         Top             =   4150
         Width           =   1575
      End
      Begin VB.CommandButton Command8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Student Ranking"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1420
         Left            =   1625
         MouseIcon       =   "emp_dash.frx":33525
         MousePointer    =   99  'Custom
         Picture         =   "emp_dash.frx":33677
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Show Student Ranking"
         Top             =   2760
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Question Paper"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1420
         Left            =   80
         MouseIcon       =   "emp_dash.frx":380F9
         MousePointer    =   99  'Custom
         Picture         =   "emp_dash.frx":3824B
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Create Question Paper"
         Top             =   2760
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Question Bank"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1420
         Left            =   1625
         MouseIcon       =   "emp_dash.frx":38AE7
         MousePointer    =   99  'Custom
         Picture         =   "emp_dash.frx":38C39
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "GoTo Question Bank"
         Top             =   1400
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Add Question"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1420
         Left            =   1625
         MouseIcon       =   "emp_dash.frx":39686
         MousePointer    =   99  'Custom
         Picture         =   "emp_dash.frx":397D8
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Add New Questions"
         Top             =   25
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Search Student"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1420
         Left            =   80
         MouseIcon       =   "emp_dash.frx":3A02B
         MousePointer    =   99  'Custom
         Picture         =   "emp_dash.frx":3A17D
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Search and Print Student Record"
         Top             =   1400
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Register Student"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1420
         Left            =   80
         MouseIcon       =   "emp_dash.frx":3AD15
         MousePointer    =   99  'Custom
         Picture         =   "emp_dash.frx":3AE67
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Add a Student"
         Top             =   25
         Width           =   1575
      End
   End
   Begin VB.Label Label2 
      Height          =   135
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   3300
   End
   Begin VB.Menu emp_qus 
      Caption         =   "&Questions"
      Begin VB.Menu hwdg 
         Caption         =   "-"
      End
      Begin VB.Menu add_ques 
         Caption         =   "Add Questions"
         Begin VB.Menu pl 
            Caption         =   "-"
         End
         Begin VB.Menu mcq_s 
            Caption         =   "MCQ Single Choice"
         End
         Begin VB.Menu jjkkj 
            Caption         =   "-"
         End
         Begin VB.Menu mcq_m 
            Caption         =   "MCQ Multiple Choice"
            Enabled         =   0   'False
         End
         Begin VB.Menu shjshj 
            Caption         =   "-"
         End
         Begin VB.Menu fill_in_blank 
            Caption         =   "Fill In The Blanks"
            Enabled         =   0   'False
         End
         Begin VB.Menu opo 
            Caption         =   "-"
         End
         Begin VB.Menu true_false 
            Caption         =   "True / False"
            Enabled         =   0   'False
         End
         Begin VB.Menu qsa 
            Caption         =   "-"
         End
         Begin VB.Menu match_f 
            Caption         =   "Match The Following"
            Enabled         =   0   'False
         End
         Begin VB.Menu hsbsb 
            Caption         =   "-"
         End
      End
      Begin VB.Menu hjhi 
         Caption         =   "-"
      End
      Begin VB.Menu ques_bnk 
         Caption         =   "Question Bank"
      End
      Begin VB.Menu bbbq 
         Caption         =   "-"
      End
   End
   Begin VB.Menu qp_paper 
      Caption         =   "Q&uestion Paper"
      Begin VB.Menu a 
         Caption         =   "-"
      End
      Begin VB.Menu create_qus_ppr 
         Caption         =   "Create Paper"
      End
      Begin VB.Menu as 
         Caption         =   "-"
      End
   End
   Begin VB.Menu reg_by_emp 
      Caption         =   "&Registration"
      Begin VB.Menu ii 
         Caption         =   "-"
      End
      Begin VB.Menu reg_stu 
         Caption         =   "Register New Student"
      End
      Begin VB.Menu jhjh 
         Caption         =   "-"
      End
   End
   Begin VB.Menu emp_stu 
      Caption         =   "&Students"
      Begin VB.Menu wbb 
         Caption         =   "-"
      End
      Begin VB.Menu srch_stu 
         Caption         =   "Search Student"
      End
      Begin VB.Menu bnm 
         Caption         =   "-"
      End
      Begin VB.Menu view_stu 
         Caption         =   "View Students"
      End
      Begin VB.Menu jsj 
         Caption         =   "-"
      End
   End
   Begin VB.Menu client_ord 
      Caption         =   "&Client Order"
      Begin VB.Menu jhj 
         Caption         =   "-"
      End
      Begin VB.Menu new_ord 
         Caption         =   "New Order"
      End
      Begin VB.Menu bwb 
         Caption         =   "-"
      End
      Begin VB.Menu view_all_client 
         Caption         =   "View All Cients"
      End
      Begin VB.Menu hjjbdbd 
         Caption         =   "-"
      End
      Begin VB.Menu bnjkmkjnhb 
         Caption         =   "Client Payment Info"
      End
      Begin VB.Menu jhpudivya 
         Caption         =   "-"
      End
   End
   Begin VB.Menu emp_utility 
      Caption         =   "&Utilities"
      Begin VB.Menu snc 
         Caption         =   "-"
      End
      Begin VB.Menu calc 
         Caption         =   "Calculator"
      End
      Begin VB.Menu cv 
         Caption         =   "-"
      End
      Begin VB.Menu notepad 
         Caption         =   "Notepad"
      End
      Begin VB.Menu c 
         Caption         =   "-"
      End
   End
   Begin VB.Menu emp_profile 
      Caption         =   "User &Profile"
      Begin VB.Menu xc 
         Caption         =   "-"
      End
      Begin VB.Menu emp_prof 
         Caption         =   "See Profile"
      End
      Begin VB.Menu cb 
         Caption         =   "-"
      End
      Begin VB.Menu chng_pass 
         Caption         =   "Change Password"
      End
      Begin VB.Menu ksk 
         Caption         =   "-"
      End
      Begin VB.Menu lg_out 
         Caption         =   "Log Out"
      End
      Begin VB.Menu df 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "emp_dash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bnjkmkjnhb_Click()
FrmClient3.Show 1, MDI
End Sub

Private Sub calc_Click()
On Error GoTo Err
    Shell "calc.exe", vbNormalFocus
    Exit Sub
Err:
    MsgBox "You don't have a Calculator installed in your computer.", vbExclamation, "Calculator Missing"
End Sub

Private Sub chng_pass_Click()
EmpPSWRD.Show
End Sub

Private Sub Command1_Click()
ques_entry_dash.Show
End Sub

Private Sub Command10_Click()
Emp_Profil.Show
End Sub

Private Sub Command11_Click()
FrmClient2.Show
End Sub

Private Sub Command13_Click()
about_org.Show vbModal, MDI
End Sub

Private Sub Command2_Click()
Search_registered.Show
End Sub

Private Sub Command3_Click()
QuesBank.Show
End Sub

Private Sub Command4_Click()
FrmClient1.Show
End Sub

Private Sub Command5_Click()
QuestionPPRdashboard.Show
End Sub

Private Sub Command6_Click()
regstudnt.Show
End Sub

Private Sub Command7_Click() 'Log Out
If MsgBox("Are You Sure to LogOut ?", vbYesNo + vbCritical, "LOGOUT") = vbYes Then
  EMP_login_reg_no = ""
  log_out_Emp
 Else
 End If
End Sub

Private Sub Command8_Click()
Stud_Ranking.Show
End Sub

Private Sub Command9_Click()
EmpPSWRD.Show
End Sub

Private Sub create_qus_ppr_Click()
QuestionPPRdashboard.Show
End Sub

Private Sub emp_prof_Click()
Emp_Profil.Show
End Sub

Private Sub Form_Activate()
Me.WindowState = vbMaximized
End Sub

Private Sub Form_Load()
conn
admin_login_reg_no = ""
Stu_login_reg_no = ""
sdate.Caption = Format(Date, "DD-MMM-YYYY")
Set r = New ADODB.Recordset
Set r = c1.Execute("select initcap(e_nm) from emp where emp_id='" & EMP_login_reg_no & "'")
If r.EOF = False Then
 sb1.Panels(2).Text = " Login :  ( " & r.Fields(0) & " )"
End If
 sb1.Panels(3).Text = " Date : " & Format(Date, "DD-MMM-YYYY")
End Sub

Private Sub lg_out_Click()
Command7_Click
End Sub

Private Sub mcq_s_Click()
ques_entry_dash.Show
End Sub

Private Sub new_ord_Click()
FrmClient1.Show 1, MDI
End Sub

Private Sub notepad_Click()
On Error GoTo Err
    Shell "notepad.exe", vbNormalFocus
    Exit Sub
Err:
    MsgBox "You don't have a Notepad installed in your computer.", vbExclamation, "Calculator Missing"
End Sub

Private Sub ques_bnk_Click()
QuesBank.Show
End Sub

Private Sub reg_stu_Click()
regstudnt.Show
End Sub

Private Sub srch_stu_Click()
Search_registered.Show
End Sub

Private Sub Timer1_Timer()
Label1.Caption = Format(Time$, "hh:mm:ss AM/PM")
Text1.Text = Format(Time$, "hh:mm:ss AM/PM")
sb1.Panels(4).Text = " Time  :  " & Format(Time$, "hh:mm:ss AM/PM")
End Sub

Private Sub view_all_client_Click()
FrmClient1.Show 1, MDI
End Sub

Private Sub view_stu_Click()
Search_registered.Show
End Sub
