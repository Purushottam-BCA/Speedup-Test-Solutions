VERSION 5.00
Begin VB.Form Certificate 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Certificate"
   ClientHeight    =   9735
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   15300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9735
   ScaleWidth      =   15300
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   9735
      Left            =   13320
      TabIndex        =   0
      Top             =   0
      Width           =   2000
      Begin VB.CommandButton cmd1 
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
         Left            =   360
         MouseIcon       =   "Certificate.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "Certificate.frx":0152
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Print Certificate"
         Top             =   1320
         Width           =   1305
      End
      Begin VB.CommandButton Command1 
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
         Left            =   360
         MouseIcon       =   "Certificate.frx":0B07
         MousePointer    =   99  'Custom
         Picture         =   "Certificate.frx":0C59
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Exit"
         Top             =   360
         Width           =   1305
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   2
         Index           =   3
         X1              =   30
         X2              =   2100
         Y1              =   9720
         Y2              =   9720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   2
         Index           =   2
         X1              =   1960
         X2              =   1960
         Y1              =   0
         Y2              =   9750
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   9750
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   2100
         Y1              =   15
         Y2              =   15
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   3
      Height          =   8595
      Left            =   555
      Top             =   525
      Width           =   12150
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "speedup test solution congratulate and best of luck for the future."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1065
      Left            =   1800
      TabIndex        =   6
      Top             =   5850
      Width           =   9705
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "SpeedUp Team"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   5
      Top             =   7320
      Width           =   2175
   End
   Begin VB.Image Image2 
      Height          =   900
      Left            =   5520
      Picture         =   "Certificate.frx":1867
      Stretch         =   -1  'True
      Top             =   1365
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "............................ "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   9720
      TabIndex        =   4
      Top             =   7440
      Width           =   1935
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "04-05-2019"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Left            =   2355
      TabIndex        =   3
      Top             =   7425
      Width           =   1380
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Date :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   1560
      TabIndex        =   2
      Top             =   7410
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "This is Certified That Divya kumari has participated and completed speedup test solutio test and scored"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   2745
      Left            =   1920
      TabIndex        =   1
      Top             =   3240
      Width           =   9465
   End
   Begin VB.Image Image1 
      Height          =   8415
      Left            =   600
      Picture         =   "Certificate.frx":6327
      Stretch         =   -1  'True
      Top             =   600
      Width           =   12060
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   10
      Height          =   9120
      Left            =   285
      Top             =   285
      Width           =   12720
   End
End
Attribute VB_Name = "Certificate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fullmrk, psmrk, mdate As String
Private Sub Form_Load()
conn
 Me.Top = 45
 Me.Left = 2800
 Me.Height = 10155
 Me.Width = 15405
 fullmrk = Summary_Test.l6.Caption
 psmrk = Summary_Test.l7.Caption
 Set r = c.Execute("select rn.rstud_nm,c.c_nm from rstud rn,course c where rn.rstud_reg_no='" & Stu_login_reg_no & "' and c.c_id=(select c_id from rstud where rstud_reg_no='" & Stu_login_reg_no & "')  ")
 mdate = Format(Date, "MMM dd yyyy")
 Label1.Caption = "This is to certify that " & r.Fields(0) & " has participeted and successfully completed  Speedup  Test Solutions's " & r.Fields(1) & " MCQ Test with Score of " & psmrk & " marks out of " & fullmrk & " on " & mdate & "."
 Label6.Caption = "Speedup Test Solutions congratulates " & r.Fields(0) & " and wishes best of luck for future..."
 Label3.Caption = Format(mdate, "DD-mm-yyyy")
End Sub
Private Sub cmd1_Click()
On Error GoTo lm
With Certificate
 .Width = 13425
 .PrintForm
GoTo k
End With
k:
 Certificate.Width = 15405
 MsgBox "Certificate Printed SuccessFully.", vbInformation + vbOKOnly, "Print"
Exit Sub
lm:
MsgBox "No Printer Installed in System..", vbCritical + vbOKOnly, "No printer"
End Sub

Private Sub Command1_Click()
 Unload Me
End Sub

Private Sub Form_Unload(cancel As Integer)
Summary_Test.Enabled = True
End Sub
