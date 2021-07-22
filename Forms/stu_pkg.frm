VERSION 5.00
Begin VB.Form rstud_pkg 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Package Information"
   ClientHeight    =   10875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "stu_pkg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10875
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   13800
      Picture         =   "stu_pkg.frx":0EE2
      ScaleHeight     =   5895
      ScaleWidth      =   4245
      TabIndex        =   6
      Top             =   2280
      Width           =   4245
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Railway Group D"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         TabIndex        =   7
         Top             =   360
         Width           =   4335
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   6800
      Left            =   7560
      Picture         =   "stu_pkg.frx":8B41
      ScaleHeight     =   6795
      ScaleWidth      =   4755
      TabIndex        =   5
      Top             =   1760
      Width           =   4750
      Begin VB.CommandButton ApplyUpdate 
         Height          =   400
         Left            =   900
         MouseIcon       =   "stu_pkg.frx":CA82
         MousePointer    =   99  'Custom
         Picture         =   "stu_pkg.frx":CBD4
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Click Here To Apply For Packagee Renewel."
         Top             =   6200
         Width           =   2895
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Rs."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   -1  'True
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1335
         TabIndex        =   19
         Top             =   1350
         Width           =   735
      End
      Begin VB.Label lbl3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   2040
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lbl6 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1665
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lbl4 
         BackColor       =   &H0080C0FF&
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label7 
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1695
         Left            =   600
         TabIndex        =   16
         Top             =   2730
         Width           =   4020
      End
      Begin VB.Label label2 
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1335
         Left            =   600
         TabIndex        =   15
         Top             =   4380
         Width           =   4020
      End
      Begin VB.Label lbl2 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "[ Safalta ]"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   17.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1005
         TabIndex        =   14
         Top             =   730
         Width           =   2775
      End
      Begin VB.Label lbl1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Railway Group D"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   300
         Width           =   4575
      End
      Begin VB.Label lbl5 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "21 May 2019"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   2205
         TabIndex        =   12
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label lbl8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   " 300 "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   -1  'True
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1920
         TabIndex        =   11
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lbl7 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "200"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   2655
         TabIndex        =   10
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Valid Till"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Left            =   1200
         TabIndex        =   9
         Top             =   1920
         Width           =   945
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   5775
      Left            =   1800
      Picture         =   "stu_pkg.frx":DBFC
      ScaleHeight     =   5775
      ScaleWidth      =   4230
      TabIndex        =   3
      Top             =   2280
      Width           =   4235
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Railway Group D"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.CommandButton btnOk 
      Height          =   400
      Left            =   240
      MouseIcon       =   "stu_pkg.frx":1597C
      MousePointer    =   99  'Custom
      Picture         =   "stu_pkg.frx":15ACE
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Note :- * In Case of request for updating package , Click on ""Apply for renew Package""  button. this will submit your request ."
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   10125
      Width           =   12615
   End
End
Attribute VB_Name = "rstud_pkg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ApplyUpdate_Click()
On Error Resume Next
Set r = New ADODB.Recordset
Set r = c.Execute("select count(*) from PKG_RENEW where RSTUD_REG_NO='" & Stu_login_reg_no & "' ")
If r.Fields(0) > 0 Then
 MsgBox " Your Previous Request is still in Progress ! " & vbCrLf & " you cannot apply for More than one request at a time" & vbCrLf & " Wait until Your Previous request is acepted" & vbCrLf & " Contact Admin For More Info. ", vbQuestion + vbOKOnly, "Request Page"
Exit Sub
End If
If Val(lbl6.Caption) > 0 And Format(Val(lbl5.Caption), "dd-mm-yyyy") >= Format(Date, "dd-mm-yyyy") Then
  If MsgBox("    Wait ! You Still has " & lbl6.Caption & " Test left in current Package" & vbCrLf & "   Your Current Package will expire on " & lbl5.Caption & vbCrLf & " Do You still want to apply for Renew Your Package ?", vbQuestion + vbYesNo, "Confirm Here") = vbYes Then
    rstud_Pkg_renew.Show
    Else
    Exit Sub
  End If
Else
 Me.Enabled = False
 rstud_Pkg_renew.Show
 End If
 End Sub

Private Sub btnOk_Click()
 Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
conn
Me.Left = 0
Me.Top = 0 '1080
Set r = New ADODB.Recordset
Set r = c.Execute("select c.c_nm,p.pkg_nm,s.strt_time,s.end_time,r.RSTUD_DOJ,r.RSTUD_DOE,r.RSTUD_ALL_TEST,p.PKG_FEE from rstud r,schdl s,course c,pkg p where p.pkg_id=r.pkg_id and p.c_id=r.c_id and s.sch_id=r.sch_id and s.c_id=r.c_id and c.c_id=r.c_id and r.RSTUD_REG_NO='" & Stu_login_reg_no & "' ")
If r.EOF = False Then
lbl1.Caption = r.Fields(0)
lbl2.Caption = "[ " & r.Fields(1) & " ]"
lbl3.Caption = r.Fields(2) & " - " & r.Fields(3)
lbl4.Caption = r.Fields(4)
lbl5.Caption = Format(r.Fields(5), "dd-MMM-YYYY")
lbl6.Caption = r.Fields(6)
lbl7.Caption = r.Fields(7)
lbl8.Caption = " " & Val(lbl7.Caption) + (Val(lbl7.Caption) * 0.25)
If Val(lbl7.Caption) >= 100 And Val(lbl7.Caption) <= 250 Then
 lbl8.Caption = " " & Val(lbl7.Caption) + 60 & " "
ElseIf Val(lbl7.Caption) > 250 And Val(lbl7.Caption) <= 300 Then
 lbl8.Caption = " " & Val(lbl7.Caption) + 100 & " "
ElseIf Val(lbl7.Caption) > 300 And Val(lbl7.Caption) <= 450 Then
 lbl8.Caption = " " & Val(lbl7.Caption) + 150 & " "
ElseIf Val(lbl7.Caption) > 450 Then
 lbl8.Caption = " " & Val(lbl7.Caption) + 200 & " "
Else
End If
Label12.Caption = lbl1.Caption
Label13.Caption = lbl1.Caption
Label7.Caption = "Package Started from " & lbl4.Caption & vbCrLf & "" _
               & "Batch Time is " & lbl3.Caption & vbCrLf & "" _
               & lbl6.Caption & " MCQ's Test for " & lbl1.Caption & vbCrLf & "" _
               & "Topic Wise / Subject Wise Test Facility available. " & vbCrLf & "" _
               & "Full Length Test Facility also available. "
 label2.Caption = "Difficulty Wise Test Option Available." & vbCrLf & "" _
               & "Answer Key and Detailed Analysis of MCQ Test Facility Available. " & vbCrLf & "" _
               & "Certificate Provided by Speedup Test Solutions. "
End If
End Sub
