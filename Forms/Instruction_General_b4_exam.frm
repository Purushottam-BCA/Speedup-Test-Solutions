VERSION 5.00
Begin VB.Form Instruction_General 
   BackColor       =   &H80000016&
   BorderStyle     =   0  'None
   Caption         =   "Instruction Page"
   ClientHeight    =   10485
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   20490
   Icon            =   "Instruction_General_b4_exam.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   Picture         =   "Instruction_General_b4_exam.frx":0ECA
   ScaleHeight     =   10485
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   11675
      Left            =   0
      Picture         =   "Instruction_General_b4_exam.frx":5831
      ScaleHeight     =   11610
      ScaleWidth      =   16620
      TabIndex        =   0
      Top             =   -100
      Width           =   16680
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0E0E0&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   0
         TabIndex        =   14
         Top             =   9960
         Width           =   16655
         Begin VB.CommandButton Command1 
            Caption         =   "I am ready to begin"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   340
            Left            =   11640
            MouseIcon       =   "Instruction_General_b4_exam.frx":4100E
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Click Here To Begin Test."
            Top             =   120
            Width           =   4815
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0E0E0&
            Caption         =   "I read and agreed with above written statements and hereby i confirm to move next."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   315
            Left            =   240
            MouseIcon       =   "Instruction_General_b4_exam.frx":41160
            MousePointer    =   99  'Custom
            TabIndex        =   15
            Top             =   105
            Width           =   8895
         End
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00800080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800080&
         FillColor       =   &H00004000&
         Height          =   320
         Left            =   840
         Shape           =   3  'Circle
         Top             =   3600
         Width           =   320
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000C000&
         FillColor       =   &H00004000&
         Height          =   320
         Left            =   825
         Shape           =   3  'Circle
         Top             =   2665
         Width           =   320
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00CA30DF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00CA30DF&
         FillColor       =   &H00004000&
         Height          =   320
         Left            =   840
         Shape           =   3  'Circle
         Top             =   3135
         Width           =   320
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         FillColor       =   &H00004000&
         Height          =   320
         Left            =   825
         Shape           =   3  'Circle
         Top             =   1765
         Width           =   320
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H000100FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000100FF&
         FillColor       =   &H00004000&
         Height          =   320
         Left            =   835
         Shape           =   3  'Circle
         Top             =   2200
         Width           =   320
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Minute"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   19540
      TabIndex        =   12
      Top             =   7600
      Width           =   585
   End
   Begin VB.Image img1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3720
      Left            =   17040
      Picture         =   "Instruction_General_b4_exam.frx":412B2
      Stretch         =   -1  'True
      Top             =   360
      Width           =   3060
   End
   Begin VB.Shape Shape11 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   6220
      Left            =   16695
      Top             =   25
      Width           =   3735
   End
   Begin VB.Label lb5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Yes (3/Each)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   19395
      TabIndex        =   11
      Top             =   9720
      Width           =   840
   End
   Begin VB.Label lb4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "45 %"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   19520
      TabIndex        =   10
      Top             =   9000
      Width           =   600
   End
   Begin VB.Label lb3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   19520
      TabIndex        =   9
      Top             =   8160
      Width           =   600
   End
   Begin VB.Label lb2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "60:00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   19480
      TabIndex        =   8
      Top             =   7260
      Width           =   720
   End
   Begin VB.Label lb1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "60"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   19520
      TabIndex        =   7
      Top             =   6480
      Width           =   600
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Negative Mark"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   17040
      TabIndex        =   6
      Top             =   9840
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Passing Marks"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   17040
      TabIndex        =   5
      Top             =   9000
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Full Marks"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   17040
      TabIndex        =   4
      Top             =   8160
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Time "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   17040
      TabIndex        =   3
      Top             =   7320
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Question"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   17040
      TabIndex        =   2
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   19200
      X2              =   19200
      Y1              =   6240
      Y2              =   10490
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H00C0C0C0&
      Height          =   855
      Left            =   16680
      Top             =   8760
      Width           =   3735
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H00C0C0C0&
      Height          =   855
      Left            =   16680
      Top             =   7080
      Width           =   3735
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00808080&
      Height          =   15
      Left            =   16800
      Top             =   5820
      Width           =   2175
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00808080&
      Height          =   15
      Left            =   16800
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Label TstType 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Full Length Test   "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   885
      Left            =   16850
      TabIndex        =   1
      Top             =   4800
      Width           =   2040
   End
   Begin VB.Image Image1 
      Height          =   1980
      Left            =   16855
      Picture         =   "Instruction_General_b4_exam.frx":44BBF
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   3540
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   4265
      Left            =   16695
      Top             =   6240
      Width           =   3735
   End
End
Attribute VB_Name = "Instruction_General"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
If Check1.Value = vbChecked Then
Command1.Enabled = True
 Else
Command1.Enabled = False
End If
End Sub

Private Sub Command1_Click()
Unload Me
MCQ.Show
End Sub

Private Sub Form_Load()
conn
Command1.Enabled = False
c.Execute ("update rstud set rstud_tot_test = rstud_tot_test - 1 where RSTUD_REG_NO='" & Stu_login_reg_no & "' ")
Me.Top = 50
Me.Left = 0 '2200
img1.Picture = LoadPicture(GlobalPic)
TstType.Caption = selectedType 'Test Type
lb1.Caption = FTOTQUESTION 'Total Question
lb2.Caption = FTOTTIMEMINUTE & ":" & FTOTTIMESECOND 'Total Time
lb3.Caption = FTOTMARKS 'Total Marks
lb4.Caption = FPASSPERCENTG & " %" 'Pass Marks
If FMRKPERWRONG = 0 Then
 lb5.Caption = "NO"  'Is Negative Marking
Else
 lb5.Caption = "Yes" & " (" & FMRKPERWRONG & "/Each)" 'Is Negative Marking
End If
End Sub
