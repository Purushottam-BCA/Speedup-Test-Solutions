VERSION 5.00
Begin VB.Form Rstud_Rough 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rough Area "
   ClientHeight    =   8190
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   11550
   Icon            =   "Rough_Stud.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   11550
   Begin VB.PictureBox Picture1 
      Height          =   8175
      Left            =   0
      MouseIcon       =   "Rough_Stud.frx":0FA2
      MousePointer    =   99  'Custom
      ScaleHeight     =   8115
      ScaleWidth      =   11475
      TabIndex        =   0
      Top             =   0
      Width           =   11535
   End
   Begin VB.Menu nh 
      Caption         =   "||"
      Enabled         =   0   'False
   End
   Begin VB.Menu newpg 
      Caption         =   "Blank Page"
   End
   Begin VB.Menu dds 
      Caption         =   "||"
      Enabled         =   0   'False
   End
   Begin VB.Menu clrpg 
      Caption         =   "Clear Page"
   End
   Begin VB.Menu fd 
      Caption         =   "||"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnExit 
      Caption         =   "Exit"
   End
   Begin VB.Menu mk 
      Caption         =   "||"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "Rstud_Rough"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub clrpg_Click()
Picture1.Cls
End Sub

Private Sub Form_Load()
Picture1.AutoRedraw = True
Picture1.DrawWidth = 2
Picture1.ForeColor = vbBlue
Picture1.BackColor = vbWhite
Me.Top = 200
Me.Left = 2000
End Sub

Private Sub Form_Unload(cancel As Integer)
MCQ.Enabled = True
End Sub

Private Sub mnExit_Click()
Unload Me
End Sub

Private Sub newpg_Click()
Picture1.Cls
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Picture1.Line (X, Y)-(X, Y)
End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Picture1.Line -(X, Y)
End If
End Sub
