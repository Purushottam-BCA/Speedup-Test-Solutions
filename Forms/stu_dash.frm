VERSION 5.00
Begin VB.Form stu_dash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Dashboard"
   ClientHeight    =   10560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
   ControlBox      =   0   'False
   Icon            =   "stu_dash.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10560
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1300
      Left            =   30
      TabIndex        =   20
      Top             =   0
      Width           =   10695
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Log Out"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1280
         Index           =   6
         Left            =   9180
         MouseIcon       =   "stu_dash.frx":0EE2
         MousePointer    =   99  'Custom
         Picture         =   "stu_dash.frx":1034
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Log Out"
         Top             =   0
         Width           =   1470
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Change Dashboard"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1280
         Index           =   5
         Left            =   7365
         MouseIcon       =   "stu_dash.frx":1D0D
         MousePointer    =   99  'Custom
         Picture         =   "stu_dash.frx":1E5F
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Change DashBoard Background Image."
         Top             =   0
         Width           =   1830
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Package"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1280
         Index           =   4
         Left            =   5910
         MouseIcon       =   "stu_dash.frx":30FA
         MousePointer    =   99  'Custom
         Picture         =   "stu_dash.frx":324C
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Current Package Details"
         Top             =   0
         Width           =   1470
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Previous Records"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1280
         Index           =   3
         Left            =   4350
         MouseIcon       =   "stu_dash.frx":3AF7
         MousePointer    =   99  'Custom
         Picture         =   "stu_dash.frx":3C49
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Previous Tests Records"
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1280
         Index           =   2
         Left            =   2895
         MouseIcon       =   "stu_dash.frx":4B13
         MousePointer    =   99  'Custom
         Picture         =   "stu_dash.frx":4C65
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Change Passwords"
         Top             =   0
         Width           =   1470
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Profile"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1280
         Index           =   1
         Left            =   1440
         MouseIcon       =   "stu_dash.frx":5B2F
         MousePointer    =   99  'Custom
         Picture         =   "stu_dash.frx":5C81
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "View and Modify Profile"
         Top             =   0
         Width           =   1470
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Start Test"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1280
         Index           =   0
         Left            =   0
         MouseIcon       =   "stu_dash.frx":65A5
         MousePointer    =   99  'Custom
         Picture         =   "stu_dash.frx":66F7
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Click To Begin Test."
         Top             =   0
         Width           =   1455
      End
   End
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
      TabIndex        =   8
      Top             =   1440
      Width           =   2655
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   840
         ScaleHeight     =   495
         ScaleWidth      =   1695
         TabIndex        =   17
         Top             =   720
         Width           =   1695
         Begin VB.Label stime 
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
            Height          =   225
            Left            =   0
            TabIndex        =   18
            Top             =   120
            Width           =   1305
         End
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         Height          =   735
         Left            =   720
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Label11"
         Height          =   15
         Left            =   360
         TabIndex        =   19
         Top             =   600
         Width           =   855
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   120
         Picture         =   "stu_dash.frx":72ED
         Top             =   0
         Width           =   240
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
         TabIndex        =   11
         Top             =   360
         Width           =   540
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
         TabIndex        =   10
         Top             =   840
         Width           =   540
      End
      Begin VB.Label sdate 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         TabIndex        =   9
         Top             =   360
         Width           =   1155
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "    Package"
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
      TabIndex        =   3
      Top             =   3000
      Width           =   2655
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         Left            =   1320
         TabIndex        =   7
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         Left            =   1320
         TabIndex        =   6
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Date :"
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
         Left            =   145
         TabIndex        =   5
         Top             =   840
         Width           =   960
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date :"
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
         TabIndex        =   4
         Top             =   360
         Width           =   1005
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   120
         Picture         =   "stu_dash.frx":7677
         Stretch         =   -1  'True
         Top             =   -20
         Width           =   240
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "    Tests"
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
      TabIndex        =   0
      Top             =   4560
      Width           =   2655
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   225
         Left            =   1680
         TabIndex        =   16
         Top             =   415
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remaining Test :"
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
         Width           =   1455
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   120
         Picture         =   "stu_dash.frx":8061
         Top             =   0
         Width           =   240
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Tests :"
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
         TabIndex        =   2
         Top             =   415
         Width           =   1095
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   225
         Left            =   1680
         TabIndex        =   1
         Top             =   840
         Width           =   495
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   1200
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Education is the Only Weapon To Change The World."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12960
      TabIndex        =   28
      Top             =   360
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Image Image5 
      Height          =   1335
      Left            =   12360
      Picture         =   "stu_dash.frx":83EB
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   5145
   End
   Begin VB.Label Check2 
      Height          =   615
      Left            =   0
      TabIndex        =   14
      Top             =   6600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Checking 
      Height          =   615
      Left            =   0
      TabIndex        =   13
      Top             =   5880
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Image stu_picx 
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   19020
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Reg_no1 
      BackColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   840
      TabIndex        =   12
      Top             =   2280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Image Image4 
      Height          =   12225
      Left            =   -120
      Picture         =   "stu_dash.frx":A075
      Stretch         =   -1  'True
      Top             =   -1680
      Width           =   20640
   End
End
Attribute VB_Name = "stu_dash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pic_name As String

Private Sub Command1_Click(Index As Integer)
On Error Resume Next
 Dim dd As Date
 If Index = 0 Then 'Start Test
  dd = Label5.Caption
  If dd > Date Then
   MsgBox "Your Package has not Been Started Yet." & vbCrLf & "Wait Until your Package Started", vbInformation + vbOKOnly, "Expired package"
   Exit Sub
  End If
  Dim d2d As Date
  If Check2.Caption = 1 Then 'Package Student
   d2d = Label1.Caption
   If d2d < Date Then
    MsgBox "Your Package has Been Expired." & vbCrLf & "Renew your Package", vbInformation + vbOKOnly, "Expired package"
    Exit Sub
   ElseIf Val(Label10.Caption) <= 0 Then
    MsgBox "You didn't have any test left." & vbCrLf & "Renew your Package to give more Test", vbInformation + vbOKOnly, "Expired package"
    Exit Sub
   ElseIf d2d = Date Then
    MsgBox "Your Package is expiring Today." & vbCrLf & "Renew your Package to give more Test", vbInformation + vbOKOnly, "Expired package"
   End If
  Else 'For Non Package
   If Val(Label10.Caption) <= 0 Then
    MsgBox "You didn't have any test left." & vbCrLf & "Renew your Package to give more Test", vbInformation + vbOKOnly, "Expired package"
    Exit Sub
   End If
  End If
 Me.Enabled = False
 Stu_Test_selection.Show
 ElseIf Index = 1 Then 'Profile
  Me.Enabled = False
  stu_profile.Show
 ElseIf Index = 2 Then 'Password
  stu_pwsd.Show
 ElseIf Index = 3 Then
  If Check2.Caption = 2 Then
   MsgBox "You are not authorized to use these Features !!" & vbCrLf & "Only Package Student can Access It ", vbInformation + vbOKOnly, "Restricted Section"
  Else
   stu_prev_record.Show
  End If
 ElseIf Index = 4 Then 'Package Information
  If Check2.Caption = 2 Then
   MsgBox "You are not authorized to use these Features !!" & vbCrLf & "Only Package Student can Access It ", vbInformation + vbOKOnly, "Restricted Section"
  Else
   rstud_pkg.Show
  End If
 ElseIf Index = 5 Then 'Change Dashboard Pic
  If Check2.Caption = 2 Then
   MsgBox "You are not authorized to use these Features !!" & vbCrLf & "Only Package Student can Access It ", vbInformation + vbOKOnly, "Restricted Section"
  Else
   FrmChangeWallppr.Show 'New Form Required
  End If
 ElseIf Index = 6 Then
  If MsgBox("Are You Sure to LogOut ?", vbYesNo + vbCritical, "LOGOUT") = vbYes Then
   Stu_login_reg_no = ""
   Timer1.Enabled = False
   log_out_rstud
  End If
 End If
End Sub

Private Sub Form_Load()
On Error GoTo k:
Me.Top = 0
Me.Left = 0
conn
admin_login_reg_no = ""
EMP_login_reg_no = ""
Checking.Caption = ""
Check2.Caption = ""
sdate.Caption = Date
Reg_no1.Caption = Stu_login_reg_no
Set r = New ADODB.Recordset
Set r = c.Execute("select * from rstud where RSTUD_REG_NO='" & Reg_no1.Caption & "' ")
If r.EOF = False Then
Checking.Caption = r.Fields(9)
Label5.Caption = Format(r.Fields(13), "dd-mm-yyyy")
Label1.Caption = Format(r.Fields(14), "dd-mm-yyyy")
Label10.Caption = r.Fields(15)
End If
If UCase(Checking.Caption) = UCase("Registered") Then
 Check2.Caption = 1
 NonPackage = 0
Else
 Check2.Caption = 2
 NonPackage = 1
End If
If IsNull(r.Fields(17)) = False Then
pic_name = r.Fields(17)
GlobalPic = pic_name
stu_picx.Picture = LoadPicture(pic_name)
Else
pic_name = App.Path & "\Graphics\#\PicNotAvail.jpg"
GlobalPic = pic_name
stu_picx.Picture = LoadPicture(pic_name)
End If
Label9.Caption = r.Fields(18)
Exit Sub
k:
pic_name = App.Path & "\Graphics\#\PicNotAvail.jpg"
GlobalPic = pic_name
stu_picx.Picture = LoadPicture(pic_name)
End Sub

Private Sub Timer1_Timer()
stime.Caption = Format(Time$, "hh:mm:ss AM/PM")
Label11.Caption = Format(Time$, "hh:mm:ss AM/PM")
End Sub
