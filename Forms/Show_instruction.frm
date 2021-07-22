VERSION 5.00
Begin VB.Form Instruction_test 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Instructions"
   ClientHeight    =   9930
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10500
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Show_instruction.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "Show_instruction.frx":0152
   ScaleHeight     =   9930
   ScaleWidth      =   10500
   Begin VB.Timer Timer2 
      Left            =   9960
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Left            =   9960
      Top             =   240
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   7870
      TabIndex        =   4
      Top             =   3925
      Width           =   150
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   2600
      TabIndex        =   3
      Top             =   3900
      Width           =   255
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "   Note  :-  Click anywhere On Page to Close This Page and Return To MCQ Test."
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Left            =   -120
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   9600
      Width           =   10380
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "09"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1380
      TabIndex        =   1
      Top             =   2520
      Width           =   270
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "09"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   3375
      TabIndex        =   0
      Top             =   2850
      Width           =   270
   End
End
Attribute VB_Name = "Instruction_test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bhide As Boolean

Private Sub Form_Click()
Timer2.Interval = 20
bhide = True
End Sub

Private Sub Form_Activate()
Me.Height = 0
bhide = False
Timer2.Interval = 30
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Timer2.Interval = 20
bhide = True
End If
End Sub

Private Sub Form_Load()
Label1.Caption = Total4InstructionPage
Label2.Caption = min4InstructionPage
Label3.Caption = FMRKPERCOR
Label4.Caption = FMRKPERWRONG
Me.Top = -50
Me.Left = 4250
End Sub

Private Sub Label5_Click()
Form_Click
End Sub

Private Sub Timer2_Timer()
If bhide = False Then
    If Me.Height >= 9975 Then
        Me.Width = 10590
        Me.Height = 10350
        Timer1.Interval = 0
        Else
        Me.Height = Me.Height + 300
    End If
Else
 If Me.Height <= 600 Then
        Me.Width = 0
        Me.Height = 0
        Timer1.Interval = 0
        Unload Me
    Else
        Me.Height = Me.Height - 300
    End If
End If
 Me.Top = 300 '((Screen.Height / 2) - (Me.Height / 2)) - 50
End Sub

