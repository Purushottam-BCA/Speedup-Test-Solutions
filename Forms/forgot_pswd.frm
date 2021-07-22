VERSION 5.00
Begin VB.Form frmFrgtPswd 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Forgot Password"
   ClientHeight    =   10950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20250
   Icon            =   "forgot_pswd.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   Picture         =   "forgot_pswd.frx":09EA
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
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
      Height          =   615
      Left            =   240
      MouseIcon       =   "forgot_pswd.frx":190EE
      MousePointer    =   99  'Custom
      Picture         =   "forgot_pswd.frx":19240
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Exit"
      Top             =   240
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3975
      Left            =   13245
      TabIndex        =   16
      Top             =   4905
      Width           =   6420
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   1
         X1              =   215
         X2              =   6220
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   1
         X1              =   55
         X2              =   6315
         Y1              =   150
         Y2              =   150
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   1
         X1              =   75
         X2              =   75
         Y1              =   150
         Y2              =   0
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   1
         X1              =   6315
         X2              =   6315
         Y1              =   150
         Y2              =   0
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   5715
      Left            =   13465
      ScaleHeight     =   5655
      ScaleWidth      =   5925
      TabIndex        =   1
      Top             =   2880
      Width           =   5985
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   495
         Left            =   2400
         TabIndex        =   20
         Top             =   5040
         Width           =   1215
      End
      Begin VB.CommandButton cancel 
         Height          =   550
         Left            =   2885
         MouseIcon       =   "forgot_pswd.frx":19E4E
         MousePointer    =   99  'Custom
         Picture         =   "forgot_pswd.frx":19FA0
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Close and return To Main Menu "
         Top             =   5095
         Width           =   3025
      End
      Begin VB.CommandButton submit 
         Height          =   550
         Left            =   -30
         MouseIcon       =   "forgot_pswd.frx":1A974
         MousePointer    =   99  'Custom
         Picture         =   "forgot_pswd.frx":1AAC6
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   5095
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   14
         Text            =   "forgot_pswd.frx":1B58E
         Top             =   2640
         Width           =   4935
      End
      Begin VB.CommandButton Search_btn 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3960
         MouseIcon       =   "forgot_pswd.frx":1B594
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1350
         Width           =   1335
      End
      Begin VB.TextBox txt3 
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
         Height          =   420
         Left            =   480
         TabIndex        =   7
         Text            =   "19405"
         Top             =   4125
         Width           =   4935
      End
      Begin VB.TextBox Txt1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3015
         MaxLength       =   15
         TabIndex        =   0
         Top             =   720
         Width           =   2280
      End
      Begin VB.Label Label11 
         Height          =   255
         Left            =   960
         TabIndex        =   12
         Top             =   1560
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label10 
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   1560
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label7 
         Height          =   255
         Left            =   1560
         TabIndex        =   10
         Top             =   1560
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   2400
         TabIndex        =   9
         Top             =   3600
         Width           =   135
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   2400
         TabIndex        =   8
         Top             =   2280
         Width           =   135
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Your answer"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   450
         TabIndex        =   6
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Label q_big 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Security Question"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   450
         TabIndex        =   5
         Top             =   2280
         Width           =   1875
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search ID / Reg.No  "
         BeginProperty Font 
            Name            =   "Constantia"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   360
         TabIndex        =   4
         Top             =   720
         Width           =   2100
      End
      Begin VB.Shape pswd_shape 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   420
         Left            =   2975
         Shape           =   4  'Rounded Rectangle
         Top             =   650
         Width           =   2355
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   2805
         TabIndex        =   2
         Top             =   645
         Width           =   135
      End
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   19480
      X2              =   19460
      Y1              =   2280
      Y2              =   8550
   End
   Begin VB.Image Image1 
      Height          =   465
      Left            =   14520
      Picture         =   "forgot_pswd.frx":1B6E6
      Stretch         =   -1  'True
      Top             =   2355
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Forgot Password ??"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   14880
      MouseIcon       =   "forgot_pswd.frx":1BFB0
      TabIndex        =   3
      Top             =   2400
      Width           =   3615
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   13440
      Top             =   2280
      Width           =   6015
   End
   Begin VB.Label Owner 
      Height          =   495
      Left            =   13560
      TabIndex        =   15
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   6555
      Left            =   13320
      Top             =   2160
      Width           =   6255
   End
End
Attribute VB_Name = "frmFrgtPswd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim user As String
Dim psd As String
Dim sk As String
Private Sub cancel_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
conn
Me.Width = MDI.Width
Me.Height = MDI.Height
Me.Top = 0
Me.Left = 0
Owner.Caption = ""
Owner.Caption = Owner1
If Owner.Caption = "Student" Then
 Label2.Caption = "Student ID / Login ID"
ElseIf Owner.Caption = "Emp" Then
 Label2.Caption = "User ID / Login ID :"
ElseIf Owner.Caption = "Admin" Then
 Label2.Caption = "Admin ID / Login ID :"
End If
user = ""
psd = ""
Frame1.Visible = True
Label7.Caption = ""
txt1.Locked = False
End Sub

Private Sub Search_btn_Click()
If Trim(txt1.Text) = "" Then
  MsgBox " Enter " & Label2.Caption & " to Search Your Record..", vbCritical + vbOKOnly, "Empty User ID"
  txt1.SetFocus
  Exit Sub
End If
Set r1 = New ADODB.Recordset
If Owner.Caption = "Student" Then
 Set r1 = c1.Execute("select * from stud_login ")
 While r1.EOF = False
   If UCase(txt1.Text) = UCase(r1.Fields(1)) Or UCase$(txt1.Text) = UCase$(r1.Fields(0)) Then
    MsgBox "Record Exist, Now Answer Security Question To Recover Your Password..", vbInformation + vbOKOnly, "Search Record"
    Frame1.Visible = False
    txt1.Locked = True
    Label10.Caption = r1.Fields(0) 'For Fetching Name
    Set r = c.Execute("select rstud_nm from rstud where rstud_reg_no='" & Label10.Caption & "' ")
    Label10.Caption = r.Fields(0)
    sk = r1.Fields(1)
    Label11.Caption = r1.Fields(2)  'For Fetching Password
    Text1.Text = r1.Fields(3)    'Sequrity Question
    Label7.Caption = r1.Fields(4)  'Password Hint Entered by user
    txt3.Text = ""
    txt3.SetFocus
    Exit Sub
   End If
 r1.MoveNext
 Wend
  MsgBox "Invalid LogIn ID or Student ID .." & vbCrLf & "Record Not Exist In database..", vbCritical + vbOKOnly, "Wrong Id"
 txt1.Text = ""
 txt1.SetFocus
 Exit Sub
ElseIf Owner.Caption = "Emp" Then
 Set r1 = c1.Execute("select * from emp_login ")
 While r1.EOF = False
   If UCase(txt1.Text) = UCase(r1.Fields(1)) Or UCase$(txt1.Text) = UCase$(r1.Fields(0)) Then
    MsgBox "Record Exist, Now Answer Security Question To Recover Your Password..", vbInformation + vbOKOnly, "Search Record"
    Frame1.Visible = False
    txt1.Locked = True
    Label10.Caption = r1.Fields(0) 'For Fetching Name
    Set r = c.Execute("select E_NM from emp where emp_id='" & Label10.Caption & "' ")
    Label10.Caption = r.Fields(0)
    Label11.Caption = r1.Fields(2)  'For Fetching Password
   sk = r1.Fields(1)
    Text1.Text = r1.Fields(3)    'Sequrity Question
    Label7.Caption = r1.Fields(4)  'Password Hint Entered by user
    txt3.Text = ""
    txt3.SetFocus
    Exit Sub
   End If
 r1.MoveNext
 Wend
  MsgBox "Invalid Login ID or User ID .." & vbCrLf & "Record Not Exist In database..", vbCritical + vbOKOnly, "Wrong Id"
 txt1.Text = ""
 txt1.SetFocus
 Exit Sub
ElseIf Owner.Caption = "Admin" Then
 Set r1 = c1.Execute("select * from admin_login ")
 While r1.EOF = False
   If UCase(txt1.Text) = UCase(r1.Fields(0)) Or UCase$(txt1.Text) = UCase$(r1.Fields(1)) Then
    MsgBox "Record Exist, Now Answer Security Question To Recover Your Password..", vbInformation + vbOKOnly, "Search Record"
    Frame1.Visible = False
    txt1.Locked = True
    Label10.Caption = r1.Fields(1) 'For Fetching Name
    Set r = c.Execute("select A_NM from adminTBL where a_id='" & Label10.Caption & "' ")
    Label10.Caption = r.Fields(0)
    Label11.Caption = r1.Fields(2)  'For Fetching Password
    sk = r1.Fields(0)
    Text1.Text = r1.Fields(3)    'Sequrity Question
    Label7.Caption = r1.Fields(4)  'Password Hint Entered by user
    txt3.Text = ""
    txt3.SetFocus
    Exit Sub
   End If
 r1.MoveNext
 Wend
  MsgBox "Invalid Login ID or Admin ID .." & vbCrLf & "Record Not Exist In database..", vbCritical + vbOKOnly, "Wrong Id"
 txt1.Text = ""
 txt1.SetFocus
 Exit Sub
End If
End Sub

Private Sub submit_Click()
If UCase$(Trim(txt3.Text)) = UCase(Label7.Caption) Then
If MsgBox("        Hii, " + Label10.Caption + vbCrLf + vbCrLf & " Login ID    =  " & sk & vbCrLf & " Password  =  " + Label11.Caption, vbInformation + vbOKOnly + vbApplicationModal, "Forgot Password") = vbOK Then
Unload Me
End If
Else
 MsgBox "           Wrong Answer" & vbCrLf & "Please Enter Correct Answer", vbCritical + vbOKOnly, "Wrong Answer"
 txt3.Text = ""
 txt3.SetFocus
End If
End Sub

Private Sub txt1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 KeyAscii = 0
 Search_btn_Click
End If
End Sub


Private Sub txt3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 KeyAscii = 0
 submit_Click
End If
End Sub
