VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmbackup 
   BorderStyle     =   0  'None
   ClientHeight    =   6405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8790
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Fram1 
      Caption         =   "Backup Remainder"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   5055
      Begin VB.Frame Fram3 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   550
         Left            =   1560
         MouseIcon       =   "Backup1.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   2520
         Width           =   1815
         Begin VB.Label Labe5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Skip"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   120
            MouseIcon       =   "Backup1.frx":0152
            MousePointer    =   99  'Custom
            TabIndex        =   6
            Top             =   120
            Width           =   1590
         End
         Begin VB.Shape Shap3 
            BackColor       =   &H8000000F&
            BackStyle       =   1  'Opaque
            Height          =   550
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.Frame Fram2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   550
         Left            =   1560
         MouseIcon       =   "Backup1.frx":02A4
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   1440
         Width           =   1815
         Begin VB.Label Labe3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Take backup"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Left            =   120
            MouseIcon       =   "Backup1.frx":03F6
            MousePointer    =   99  'Custom
            TabIndex        =   4
            ToolTipText     =   "Click To Take Backup"
            Top             =   120
            Width           =   1620
         End
         Begin VB.Shape Shap1 
            BackColor       =   &H8000000F&
            BackStyle       =   1  'Opaque
            Height          =   550
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.Label Labe2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "23/12/2019"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2640
         TabIndex        =   2
         Top             =   600
         Width           =   4620
      End
      Begin VB.Label Labe1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last backup Taken On : "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   480
         TabIndex        =   1
         Top             =   600
         Width           =   2130
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   4575
      Left            =   480
      TabIndex        =   7
      Top             =   960
      Width           =   7815
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Take Backup"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   415
         Left            =   2040
         MouseIcon       =   "Backup1.frx":0548
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3960
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   415
         Left            =   4200
         MouseIcon       =   "Backup1.frx":069A
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         Caption         =   "Backup"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   7575
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Browse"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   415
            Left            =   6000
            MouseIcon       =   "Backup1.frx":07EC
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   925
            Width           =   1335
         End
         Begin VB.Label Label4 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   14
            Top             =   960
            Width           =   5295
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select the Folder where you  want to take the backup  :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   480
            TabIndex        =   13
            Top             =   480
            Width           =   5070
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Last Backup"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   8
         Top             =   0
         Width           =   7575
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Last backup Taken On : "
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   480
            TabIndex        =   10
            Top             =   480
            Width           =   2130
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "23/12/2019"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2640
            TabIndex        =   9
            Top             =   480
            Width           =   4620
         End
      End
      Begin MSComDlg.CommonDialog ccc 
         Left            =   7200
         Top             =   720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Visible         =   0   'False
         Width           =   7335
      End
      Begin VB.TextBox Text2 
         Height          =   975
         Left            =   840
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   2640
         Visible         =   0   'False
         Width           =   6735
      End
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Select Catagory"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6315
      Left            =   60
      TabIndex        =   20
      Top             =   0
      Width           =   8645
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "<< Back   "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   415
         Left            =   480
         MouseIcon       =   "Backup1.frx":093E
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Go back"
         Top             =   5520
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Next >>"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   415
         Left            =   6840
         MouseIcon       =   "Backup1.frx":0A90
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Take BackUp"
         Top             =   5520
         Width           =   1455
      End
      Begin VB.CommandButton btn7 
         BackColor       =   &H8000000E&
         Caption         =   "Backup ALL"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   6600
         MouseIcon       =   "Backup1.frx":0BE2
         MousePointer    =   99  'Custom
         Picture         =   "Backup1.frx":0D34
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Take Complete Backup."
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton Btn1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Master Entry"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   960
         MouseIcon       =   "Backup1.frx":18AD
         MousePointer    =   99  'Custom
         Picture         =   "Backup1.frx":19FF
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Take Backup of All Master Entries."
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton btn5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Questions"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   2880
         MouseIcon       =   "Backup1.frx":2323
         MousePointer    =   99  'Custom
         Picture         =   "Backup1.frx":2475
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Take Backup of All Questions."
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton btn2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Students"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   2880
         MouseIcon       =   "Backup1.frx":2F35
         MousePointer    =   99  'Custom
         Picture         =   "Backup1.frx":3087
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Take Backup of All Students Records."
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton btn3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Users"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   4800
         MouseIcon       =   "Backup1.frx":3A6F
         MousePointer    =   99  'Custom
         Picture         =   "Backup1.frx":3BC1
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Take Backup of All User's Record"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton btn4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Client Orders"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   960
         MouseIcon       =   "Backup1.frx":5256
         MousePointer    =   99  'Custom
         Picture         =   "Backup1.frx":53A8
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Take Backup of Clients and Thier Recent Orders."
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton btn6 
         BackColor       =   &H8000000E&
         Caption         =   "Account"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   4800
         MouseIcon       =   "Backup1.frx":5C15
         MousePointer    =   99  'Custom
         Picture         =   "Backup1.frx":5D67
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Take Backup of Accounts Dept."
         Top             =   3120
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmbackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sd As Date
Dim choiceBtn As Integer

Private Sub btn1_Click()
choiceBtn = 1
End Sub

Private Sub btn2_Click()
choiceBtn = 2
End Sub

Private Sub btn3_Click()
choiceBtn = 3
End Sub

Private Sub btn4_Click()
choiceBtn = 4
End Sub

Private Sub btn5_Click()
choiceBtn = 5
End Sub

Private Sub btn6_Click()
choiceBtn = 6
End Sub

Private Sub btn7_Click()
choiceBtn = 7
End Sub

Private Sub Command4_Click()
If choiceBtn > 0 And choiceBtn < 8 Then
 Frame4.Visible = False
 Frame3.Visible = True
 Fram1.Visible = False
Else
 MsgBox "Select Any Catagory For Backup..", vbInformation + vbOKOnly, "Select Type"
End If
End Sub

Private Sub Command5_Click()
choiceBtn = 0
Fram1.Visible = True
Frame3.Visible = False
Frame4.Visible = False
End Sub

Private Sub Form_Load()
conn
CenterForm Me
CreateRoundRectFromWindow Me
Fram1.Visible = True
Frame3.Visible = False
Frame4.Visible = False
choiceBtn = 0
Set r = New ADODB.Recordset
Set r = c.Execute("select bdate from Backup1")
If r.EOF = False Then
 sd = r.Fields(0)
 Labe2.Caption = Format(sd, "dd/mm/yyyy")
 Label2.Caption = Labe2.Caption
Else
 Labe2.Caption = "No Backup Taken Yet"
 Label2.Caption = Labe2.Caption
End If
Shap1.BackColor = &H80000005
Shap3.BackColor = &H80000005
Labe3.FontBold = False
Labe5.FontBold = False
Label4.Caption = ""
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub Form_Unload(cancel As Integer)
If adminOpen = 2 Then
adminOpen = 1
admin_dash.Enabled = True
Else
End If
End Sub

Private Sub Fram1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shap1.BackColor = &H80000005
Shap3.BackColor = &H80000005
Labe3.FontBold = False
Labe5.FontBold = False
End Sub

Private Sub Fram3_Click()
Unload Me
End Sub

Private Sub Fram3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shap3.BackColor = &HD0C0C0
Labe5.FontBold = True
End Sub

Private Sub Frame4_Click()
choiceBtn = 0
End Sub

Private Sub Labe5_Click()
Unload Me
End Sub
Private Sub Labe5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shap3.BackColor = &HD0C0C0
Labe5.FontBold = True
End Sub
Private Sub Fram2_Click()
Fram1.Visible = False
Frame3.Visible = True
End Sub
Private Sub Fram2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shap1.BackColor = &HD0C0C0
Labe3.FontBold = True
End Sub

Private Sub Labe3_Click()
choiceBtn = 0
Fram1.Visible = False
Frame4.Visible = True
Frame3.Visible = False
End Sub

Private Sub Labe3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shap1.BackColor = &HD0C0C0
Labe3.FontBold = True
End Sub

'+++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub Command1_Click()
Dim stempDir As String
On Error Resume Next
stempDir = CurDir 'Current Directory
ccc.DialogTitle = "Select A Folder "
ccc.InitDir = App.Path & "\database"
ccc.FileName = "Select Folder"
ccc.Flags = cdlOFNNoValidate + cdlOFNHideReadOnly
ccc.Filter = "Directories|*.~#~"
ccc.CancelError = True
ccc.ShowOpen
If Err <> 32755 Then
 Label4.Caption = CurDir
End If
ChDir stempDir
End Sub

Private Sub Command2_Click()
Fram1.Visible = True
Frame3.Visible = False
Frame4.Visible = False
choiceBtn = 0
End Sub

Private Sub Command3_Click() 'Export Backup File
If Label4.Caption = "" Then
 MsgBox "Select the folder Where You Want to Store the backup File", vbInformation + vbOKOnly, "Empty Location"
 Exit Sub
End If
Text1.Text = ""
If choiceBtn = 1 Then 'Master Entry
 Text1.Text = "exp sts/sts grants=y file=" & Label4.Caption & "\M_Entry_Backup.DMP tables=(q_typ,pkg,schdl,topic,sub,course)"
ElseIf choiceBtn = 2 Then 'Students
 Text1.Text = "exp sts/sts grants=y file=" & Label4.Caption & "\Student_Backup.DMP tables=(STURANK,RSTUD,STUD_LOGIN,STUD_PREV_REC,PKG_RENEW)"
ElseIf choiceBtn = 3 Then 'Users
 Text1.Text = "exp sts/sts grants=y file=" & Label4.Caption & "\User_Backup.DMP tables=(Emp,Emp_login)"
ElseIf choiceBtn = 4 Then 'Clients
 Text1.Text = "exp sts/sts grants=y file=" & Label4.Caption & "\Client_Order_Backup.DMP tables=(Client,CLNT_ORDR_CHLN,CLIENT_PMT,QPAPRDASH)"
ElseIf choiceBtn = 5 Then 'Questions
 Text1.Text = "exp sts/sts grants=y file=" & Label4.Caption & "\Questions_Backup.DMP tables=(quesms)"
ElseIf choiceBtn = 6 Then 'Accounts
 Text1.Text = "exp sts/sts grants=y file=" & Label4.Caption & "\Account_Backup.DMP tables=(incm,exp)"
ElseIf choiceBtn = 7 Then 'Complete
 Text1.Text = "exp sts/sts grants=y file=" & Label4.Caption & "\BackupFile.DMP"
End If
Shell "cmd.exe /c " & Text1.Text
MsgBox "Backup File SuccessFully Created", vbInformation + vbOKOnly, "Backup Success"
c.Execute ("delete from Backup1")
c.Execute ("insert into Backup1 values('" & Format(Date, "dd-mmm-yyyy") & "') ")
Shell "cmd.exe /c explorer.exe " & Label4.Caption
Fram1.Visible = True
Frame3.Visible = False
Frame4.Visible = False
choiceBtn = 0
Set r = New ADODB.Recordset
Set r = c.Execute("select bdate from Backup1")
If r.EOF = False Then
 sd = r.Fields(0)
 Labe2.Caption = Format(sd, "dd/mm/yyyy")
 Label2.Caption = Labe2.Caption
Else
 Labe2.Caption = "No Backup Taken Yet"
 Label2.Caption = Labe2.Caption
End If
Shap1.BackColor = &H80000005
Shap3.BackColor = &H80000005
Labe3.FontBold = False
Labe5.FontBold = False
Label4.Caption = ""
Text1.Text = ""
Text2.Text = ""
End Sub
