VERSION 5.00
Begin VB.Form EmpPSWRD 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "Password"
   ClientHeight    =   10725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20415
   ControlBox      =   0   'False
   ForeColor       =   &H00808080&
   Icon            =   "Emp_Password.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10725
   ScaleWidth      =   20415
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton ChameleonBtn2 
      Height          =   400
      Left            =   240
      MouseIcon       =   "Emp_Password.frx":0ECA
      MousePointer    =   99  'Custom
      Picture         =   "Emp_Password.frx":101C
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   255
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   19680
      MouseIcon       =   "Emp_Password.frx":174F
      MousePointer    =   99  'Custom
      Picture         =   "Emp_Password.frx":18A1
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   7
      ToolTipText     =   "Click For Help"
      Top             =   0
      Width           =   735
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   13800
      TabIndex        =   6
      Top             =   1320
      Width           =   6375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   240
      Top             =   3480
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   240
      Top             =   4320
   End
   Begin VB.PictureBox Picture2 
      Height          =   7315
      Left            =   7200
      ScaleHeight     =   7260
      ScaleWidth      =   5235
      TabIndex        =   8
      Top             =   1800
      Width           =   5295
      Begin VB.Frame fram2 
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   585
         TabIndex        =   10
         Top             =   1920
         Width           =   3975
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
            TabIndex        =   13
            Top             =   1255
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
            TabIndex        =   12
            Top             =   2100
            Width           =   3615
         End
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
            TabIndex        =   11
            Top             =   405
            Width           =   3615
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
            TabIndex        =   31
            Top             =   1725
            Width           =   1875
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
            TabIndex        =   30
            Top             =   840
            Width           =   1545
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
            TabIndex        =   29
            Top             =   30
            Width           =   1830
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
            Left            =   2115
            TabIndex        =   16
            Top             =   45
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
            TabIndex        =   15
            Top             =   885
            Width           =   105
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
            TabIndex        =   14
            Top             =   1725
            Width           =   105
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   735
         Left            =   0
         TabIndex        =   9
         Top             =   6540
         Width           =   5325
         Begin VB.CommandButton btnOk 
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
            Height          =   375
            Left            =   4200
            MouseIcon       =   "Emp_Password.frx":23E6
            MousePointer    =   99  'Custom
            Picture         =   "Emp_Password.frx":2538
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Click Ok To Save."
            Top             =   190
            Width           =   825
         End
         Begin VB.CommandButton btnChangePas 
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
            Height          =   375
            Left            =   240
            MouseIcon       =   "Emp_Password.frx":2AF4
            MousePointer    =   99  'Custom
            Picture         =   "Emp_Password.frx":2C46
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Click Ok To Save."
            Top             =   190
            Width           =   2145
         End
         Begin VB.Label Label4 
            BackColor       =   &H000080FF&
            Height          =   255
            Left            =   360
            TabIndex        =   19
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin VB.Frame fram1 
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   720
         TabIndex        =   22
         Top             =   1920
         Width           =   4335
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
            TabIndex        =   24
            Text            =   "Combo1"
            Top             =   405
            Width           =   4095
         End
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
            TabIndex        =   23
            Top             =   1260
            Width           =   4095
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
            TabIndex        =   33
            Top             =   890
            Width           =   1665
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
            TabIndex        =   32
            Top             =   15
            Width           =   1830
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
            TabIndex        =   26
            Top             =   30
            Width           =   105
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
            TabIndex        =   25
            Top             =   870
            Width           =   105
         End
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "LogIn ID  :"
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
         Left            =   840
         TabIndex        =   28
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Emp. No  :"
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
         Left            =   840
         TabIndex        =   27
         Top             =   600
         Width           =   1215
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
         Left            =   2265
         TabIndex        =   18
         Top             =   1200
         Width           =   2175
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
         Left            =   2265
         TabIndex        =   17
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.Image Image2 
      Height          =   405
      Left            =   17160
      Picture         =   "Emp_Password.frx":38FA
      Stretch         =   -1  'True
      Top             =   5700
      Width           =   1680
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "5. Click on Apply Password                                        Button to                   Change Password."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   630
      Left            =   14640
      TabIndex        =   5
      Top             =   5760
      Width           =   5715
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   $"Emp_Password.frx":4483
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   870
      Left            =   14640
      TabIndex        =   4
      Top             =   4800
      Width           =   5535
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   $"Emp_Password.frx":4523
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   870
      Left            =   14640
      TabIndex        =   3
      Top             =   3840
      Width           =   5535
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   $"Emp_Password.frx":45D0
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   750
      Left            =   14640
      TabIndex        =   2
      Top             =   2880
      Width           =   5535
   End
   Begin VB.Image Image1 
      Height          =   405
      Left            =   17370
      Picture         =   "Emp_Password.frx":4658
      Stretch         =   -1  'True
      Top             =   2340
      Width           =   1560
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1. Click on Change Password                                       Button."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   14640
      TabIndex        =   1
      Top             =   2400
      Width           =   5040
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "-:  Steps To Change Password :-"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Left            =   14685
      TabIndex        =   0
      Top             =   1732
      Width           =   5295
   End
End
Attribute VB_Name = "EmpPSWRD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnChangePas_Click()
If btnChangePas.Caption = "Change Password" Then
Timer2.Enabled = False
Timer1.Enabled = True
fram2.Visible = True
Text1.Enabled = True
lbl4.Enabled = True
lbl5.Enabled = True
Text1.SetFocus
lbl6.Text = ""
Combo1.Clear

Set r = c.Execute("select QUES from SECQUES ")
Combo1.Clear
While r.EOF = False
Combo1.AddItem r.Fields(0)
r.MoveNext
Wend

btnChangePas.Caption = "Apply Password"
ElseIf btnChangePas.Caption = "Apply Password" Then
If Trim(lbl1.Caption) = "" Or Trim(lbl2.Caption) = "" Or Trim(Text1.Text) = "" Or Trim(lbl4.Text) = "" Or Trim(Combo1.Text) = "" Or Trim(lbl5.Text) = "" Or Trim(lbl6.Text) = "" Then
MsgBox " All Fields are necessary", vbInformation + vbOKOnly, ""
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
c.Execute ("update emp_login set E_PSWD='" & lbl4.Text & "',E_HNT='" & Combo1.Text & "', E_HNT_ANS='" & lbl6.Text & "'  where E_ID='" & EMP_login_reg_no & "' ")
MsgBox "Password Changed", vbInformation + vbOKOnly, "Changed"
btnChangePas.Caption = "Change Password"
Timer1.Enabled = False
Timer2.Enabled = True
fram2.Visible = False
End If
End Sub

Private Sub btnOk_Click()
Unload Me
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
Frame2.Visible = True
conn
Me.Top = 0
Me.Left = 0
fram2.Visible = False
lbl1.Caption = EMP_login_reg_no
Set r = c.Execute(" select * from emp_login where e_id='" & EMP_login_reg_no & "' ")
If r.EOF = False Then
lbl2.Caption = r.Fields(1)
Label4.Caption = r.Fields(2)
Text1.Text = ""
lbl4.Text = ""
lbl5.Text = ""
lbl6.Text = r.Fields(4)
Combo1.Text = r.Fields(3)
End If
  lbl4.Locked = True
  lbl5.Locked = True
End Sub

Private Sub Form_Unload(cancel As Integer)
 Timer1.Enabled = False
 Timer2.Enabled = False
End Sub

Private Sub lbl4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 lbl5.SetFocus
End If
End Sub

Private Sub lbl4_LostFocus()
If Len(Trim(lbl4.Text)) > 0 And Len(Trim(lbl4.Text)) < 8 Then
MsgBox "Password too short (min- 8 char) Required..", vbInformation + vbOKOnly, ""
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

Private Sub Picture1_Click()
If Frame2.Visible = False Then
Frame2.Visible = True
Else
Frame2.Visible = False
End If
End Sub

Private Sub Text1_LostFocus()
If Trim(Text1.Text) <> "" Then
 If UCase(Text1.Text) <> UCase(Label4.Caption) Then
  lbl4.Locked = True
  lbl5.Locked = True
  MsgBox "Wrong Password. Enter Current Password", vbCritical + vbOKOnly, ""
  Text1.Text = ""
  Exit Sub
  Else
   lbl4.Locked = False
   lbl5.Locked = False
End If
End If
End Sub

Private Sub Timer1_Timer()
fram1.Top = fram1.Top + 150
If fram1.Top >= 4560 Then
fram1.Top = 4560
Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
fram1.Top = fram1.Top - 150
If fram1.Top <= 1920 Then
fram1.Top = 1920
Timer2.Enabled = False
End If
End Sub

