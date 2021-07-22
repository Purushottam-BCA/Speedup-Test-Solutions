VERSION 5.00
Begin VB.Form FrmSelectUser 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   ClientHeight    =   10950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10950
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command 
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
      Left            =   19000
      MouseIcon       =   "Select_User.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "Select_User.frx":0152
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Close Application"
      Top             =   120
      Width           =   1300
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   2880
      TabIndex        =   1
      Top             =   2280
      Width           =   14655
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "STUDENT"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Index           =   0
         Left            =   1350
         MouseIcon       =   "Select_User.frx":0D60
         MousePointer    =   99  'Custom
         Picture         =   "Select_User.frx":0EB2
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Login As Student TO Give Test and Evaluating Yourself By Seeing Previous Records"
         Top             =   1350
         Width           =   3045
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "USER"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Index           =   1
         Left            =   5690
         MouseIcon       =   "Select_User.frx":2C00
         MousePointer    =   99  'Custom
         Picture         =   "Select_User.frx":2D52
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Login As An User"
         Top             =   1350
         Width           =   3015
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ADMIN"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Index           =   2
         Left            =   10125
         MouseIcon       =   "Select_User.frx":5B36
         MousePointer    =   99  'Custom
         Picture         =   "Select_User.frx":5C88
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Login As An Admin"
         Top             =   1350
         Width           =   2895
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00808080&
         Height          =   5055
         Left            =   240
         Top             =   360
         Width           =   14175
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Height          =   3200
         Index           =   0
         Left            =   1320
         Top             =   1320
         Width           =   3105
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Height          =   3200
         Index           =   1
         Left            =   5625
         Top             =   1320
         Width           =   3105
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Height          =   3200
         Index           =   2
         Left            =   10065
         Top             =   1320
         Width           =   2985
      End
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   6015
      Left            =   2640
      Top             =   2160
      Width           =   15135
   End
End
Attribute VB_Name = "FrmSelectUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command_Click()
If MsgBox("Are You Sure To Exit Application ?", vbCritical + vbYesNo, "Exit Application") = vbYes Then
 End
End If
End Sub

Private Sub Command1_Click(Index As Integer)
Command.SetFocus
If Index = 0 Then
 Me.Hide
  login_new.userID.Text = ""
  login_new.pswd.Text = ""
  login_new.vkCheck1.Value = vbUnchecked
  login_new.Show
 ElseIf Index = 1 Then
 Me.Hide
  login_EMP.userID.Text = ""
  login_EMP.pswd.Text = ""
  login_EMP.vkCheck1.Value = vbUnchecked
  login_EMP.Show
Else
  Me.Hide
  login_Admin.userID.Text = ""
  login_Admin.pswd.Text = ""
  login_Admin.vkCheck1.Value = vbUnchecked
  login_Admin.Show
End If
End Sub

Private Sub Command1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then
 Shape1(0).Visible = True
 Shape1(1).Visible = False
 Shape1(2).Visible = False
ElseIf Index = 1 Then
 Shape1(0).Visible = False
 Shape1(1).Visible = True
 Shape1(2).Visible = False
ElseIf Index = 2 Then
 Shape1(0).Visible = False
 Shape1(1).Visible = False
 Shape1(2).Visible = True
End If
End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
Me.Height = MDI.Height
Me.Width = MDI.Width
Shape1(0).Visible = False
Shape1(1).Visible = False
Shape1(2).Visible = False
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape1(0).Visible = False
Shape1(1).Visible = False
Shape1(2).Visible = False
End Sub
