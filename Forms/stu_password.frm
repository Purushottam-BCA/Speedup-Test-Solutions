VERSION 5.00
Begin VB.Form stu_pwsd 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "Change Password"
   ClientHeight    =   10860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20400
   Icon            =   "stu_password.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10860
   ScaleWidth      =   20400
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton ChameleonBtn2 
      Height          =   400
      Left            =   240
      MouseIcon       =   "stu_password.frx":0ECA
      MousePointer    =   99  'Custom
      Picture         =   "stu_password.frx":101C
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   240
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   7260
      Left            =   8160
      ScaleHeight     =   7200
      ScaleWidth      =   5115
      TabIndex        =   0
      Top             =   1800
      Width           =   5175
      Begin VB.Frame frame1 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   735
         Left            =   0
         TabIndex        =   8
         Top             =   6540
         Width           =   5175
         Begin VB.CommandButton btnOk 
            BackColor       =   &H00E0E0E0&
            DisabledPicture =   "stu_password.frx":174F
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   385
            Left            =   4080
            MouseIcon       =   "stu_password.frx":1F75
            MousePointer    =   99  'Custom
            Picture         =   "stu_password.frx":20C7
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Mve Next For Test."
            Top             =   185
            Width           =   855
         End
         Begin VB.CommandButton btnChangePas 
            BackColor       =   &H00E0E0E0&
            DisabledPicture =   "stu_password.frx":2683
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   195
            MouseIcon       =   "stu_password.frx":2EA9
            MousePointer    =   99  'Custom
            Picture         =   "stu_password.frx":2FFB
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Mve Next For Test."
            Top             =   180
            Width           =   2175
         End
      End
      Begin VB.Frame fram2 
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   380
         TabIndex        =   1
         Top             =   1920
         Width           =   3975
         Begin VB.TextBox Text1 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   240
            PasswordChar    =   "*"
            TabIndex        =   4
            Top             =   405
            Width           =   3615
         End
         Begin VB.TextBox lbl5 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            MaxLength       =   15
            TabIndex        =   3
            Top             =   2100
            Width           =   3615
         End
         Begin VB.TextBox lbl4 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   240
            MaxLength       =   15
            PasswordChar    =   "*"
            TabIndex        =   2
            Top             =   1255
            Width           =   3615
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current Password :"
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
            Index           =   2
            Left            =   240
            TabIndex        =   23
            Top             =   0
            Width           =   1830
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "New Password :"
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
            Index           =   3
            Left            =   240
            TabIndex        =   22
            Top             =   810
            Width           =   1545
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Confirm Password :"
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
            Index           =   4
            Left            =   240
            TabIndex        =   21
            Top             =   1695
            Width           =   1875
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   270
            Left            =   2160
            TabIndex        =   7
            Top             =   1725
            Width           =   105
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   270
            Left            =   1815
            TabIndex        =   6
            Top             =   885
            Width           =   105
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   270
            Left            =   2085
            TabIndex        =   5
            Top             =   45
            Width           =   105
         End
      End
      Begin VB.Frame fram1 
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   495
         TabIndex        =   13
         Top             =   1920
         Width           =   4335
         Begin VB.TextBox lbl6 
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   1260
            Width           =   4095
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   14
            Text            =   "Combo1"
            Top             =   405
            Width           =   4095
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Security Question :"
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
            Index           =   5
            Left            =   120
            TabIndex        =   25
            Top             =   0
            Width           =   1830
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Security Answer :"
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
            Index           =   6
            Left            =   120
            TabIndex        =   24
            Top             =   870
            Width           =   1665
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   270
            Left            =   1920
            TabIndex        =   17
            Top             =   870
            Width           =   105
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   270
            Left            =   2040
            TabIndex        =   16
            Top             =   30
            Width           =   105
         End
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Reg. No   :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   630
         TabIndex        =   20
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "LogIn ID   :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   630
         TabIndex        =   19
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H000080FF&
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lbl1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2025
         TabIndex        =   12
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label lbl2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2025
         TabIndex        =   11
         Top             =   1200
         Width           =   2175
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   480
      Top             =   4560
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   480
      Top             =   3720
   End
End
Attribute VB_Name = "stu_pwsd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim btnpass As String
Private Sub btnChangePas_Click()
If btnpass = "Change Password" Then
Timer2.Enabled = False
Timer1.Enabled = True
Fram2.Visible = True
Text1.Enabled = True
lbl4.Enabled = True
lbl5.Enabled = True
Text1.SetFocus
lbl6.Text = ""
 Set r = c.Execute("select QUES from SECQUES ")
 Combo1.Clear
 While r.EOF = False
  Combo1.AddItem r.Fields(0)
  r.MoveNext
 Wend
 btnpass = "Apply Password"
ElseIf btnpass = "Apply Password" Then
If Trim(lbl1.Caption) = "" Or Trim(lbl2.Caption) = "" Or Trim(Text1.Text) = "" Or Trim(lbl4.Text) = "" Or Trim(Combo1.Text) = "" Or Trim(lbl5.Text) = "" Or Trim(lbl6.Text) = "" Then
MsgBox " All Fields are necessary", vbInformation + vbOKOnly, ""
'btnChangePas.Caption = "Change Password"
Exit Sub
ElseIf UCase(Text1.Text) <> UCase(Label4.Caption) Then
MsgBox "Enter Current Password", vbInformation + vbOKOnly, ""
Text1.Text = ""
Text1.SetFocus
Exit Sub

End If
Text1.Enabled = False
lbl4.Enabled = False
lbl5.Enabled = False
c.Execute ("update stud_login set RSTUD_PSWD='" & lbl4.Text & "',RSTUD_HNT='" & Combo1.Text & "', RSTUD_HNT_ANS='" & lbl6.Text & "'  where RSTUD_REG_NO='" & Stu_login_reg_no & "' ")
MsgBox "Password Changed", vbInformation + vbOKOnly, "Changed"
btnpass = "Change Password"
Timer1.Enabled = False
Timer2.Enabled = True
Fram2.Visible = False
End If
End Sub

Private Sub btnOk_Click()
Unload Me
'stu_dash.Show
End Sub
Private Sub ChameleonBtn2_Click()
Unload Me
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 lbl6.SetFocus
End If
End Sub

Private Sub Form_Load()
conn
btnpass = "Change Password"
Fram2.Visible = False
Me.Top = 0 '1080
Me.Left = 0 '5600
lbl1.Caption = Stu_login_reg_no
Set r = c.Execute(" select * from stud_login where RSTUD_REG_NO='" & Stu_login_reg_no & "' ")
If r.EOF = False Then
lbl2.Caption = r.Fields(1)
Label4.Caption = r.Fields(2)
Text1.Text = ""
lbl4.Text = ""
lbl5.Text = ""
Combo1.Text = r.Fields(3)
lbl6.Text = r.Fields(4)
End If
End Sub

Private Sub lbl4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 lbl5.SetFocus
End If
End Sub

Private Sub lbl4_LostFocus()
If Len(Trim(lbl4.Text)) > 0 And Len(Trim(lbl4.Text)) < 8 Then
MsgBox "Password too short (min- 8 char)", vbInformation + vbOKOnly, ""
Exit Sub
lbl4.SetFocus
Else
lbl5.Enabled = True
End If
End Sub

Private Sub lbl5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Combo1.SetFocus
End If
End Sub

Private Sub lbl5_LostFocus()
If Len(Trim(lbl4.Text)) > 8 And Trim(lbl4.Text) <> "" And lbl4.Text <> lbl5.Text Then
MsgBox "ReEnter Correct Password", vbInformation + vbOKOnly, ""
lbl5.Text = ""
lbl5.SetFocus
Exit Sub
Else

End If
End Sub

Private Sub Timer1_Timer()
Fram1.Top = Fram1.Top + 150
If Fram1.Top >= 4560 Then
Fram1.Top = 4560
Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
Fram1.Top = Fram1.Top - 150
If Fram1.Top <= 1920 Then
Fram1.Top = 1920
Timer2.Enabled = False
End If
End Sub
