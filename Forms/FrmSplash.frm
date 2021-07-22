VERSION 5.00
Object = "{BF38D12B-22A9-4B10-B26E-019F2B5F9C22}#1.0#0"; "AniGif.ocx"
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   5145
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8175
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FrmSplash.frx":0000
   ScaleHeight     =   5145
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   240
      Top             =   4080
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      Visible         =   0   'False
      X1              =   1560
      X2              =   6480
      Y1              =   4700
      Y2              =   4700
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      Visible         =   0   'False
      X1              =   1560
      X2              =   6495
      Y1              =   4095
      Y2              =   4095
   End
   Begin Project1.PictureG p 
      Height          =   375
      Left            =   1680
      Top             =   4200
      Visible         =   0   'False
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   661
      GIF             =   "FrmSplash.frx":1BCB
      DelayLoad       =   0
   End
   Begin Project1.PictureG PictureG2 
      Height          =   3570
      Left            =   -120
      Top             =   120
      Width           =   8670
      _ExtentX        =   15293
      _ExtentY        =   6297
      GIF             =   "FrmSplash.frx":3429
      DelayLoad       =   0
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b, ab As Integer
Private Sub Form_Load()
On Error GoTo Errorhandle
Dim TopCorner As Integer
Dim LeftCorner As Integer
If Me.WindowState <> 0 Then Exit Sub
TopCorner = ((Screen.Height - Me.Height) \ 2) - 250
LeftCorner = (Screen.Width - Me.Width) \ 2
Me.Move LeftCorner, TopCorner
CreateRoundRectFromWindow Me
conn
Set r = c.Execute("select count(*) from secques ")
If r.Fields(0) = 0 Then
 c.Execute ("insert into secques values(1,'What Is Your Nick Name ?')")
 c.Execute ("insert into secques values(2,'In Which Year Are You Born ?')")
 c.Execute ("insert into secques values(3,'Which Sport Do You Like Most ?')")
 c.Execute ("insert into secques values(4,'WhIch IS Your favourite Movie ?')")
End If
Set r1 = c.Execute("select count(*) from q_typ ")
If r1.Fields(0) = 0 Then
 c.Execute ("insert into q_typ values('QT001','MCQs',1)")
End If
Timer1.Enabled = True
a = 0
b = 0
ab = 0
EMP_login_reg_no = ""
admin_login_reg_no = ""
Stu_login_reg_no = ""
Current_Logged_ID = ""
Exit Sub
Errorhandle:
MsgBox "Some Error Occured..Either You Are Running a 64-bit Operating System or Oracle 10G is Not Installed Properly.Contact Administrator For Help !!", vbCritical + vbOKOnly, "Connection Error"
End Sub

Private Sub Timer1_Timer()
a = a + 1
If a = 35 Then
 p.Visible = True
 Line1.Visible = True
 Line2.Visible = True
End If
If a = 138 Then
 Timer1.Enabled = False
Unload Me
FrmSelectUser.Show
End If
End Sub

