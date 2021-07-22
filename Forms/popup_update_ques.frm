VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{EAFDAFBF-1D88-41DD-B117-60ECBC4B8441}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form FrmQuesUpdate 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Question"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11355
   Icon            =   "popup_update_ques.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   11355
   Begin VB.CommandButton Command1 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   430
      Left            =   7920
      MouseIcon       =   "popup_update_ques.frx":09EA
      MousePointer    =   99  'Custom
      TabIndex        =   25
      ToolTipText     =   "Update"
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   430
      Left            =   9600
      MouseIcon       =   "popup_update_ques.frx":0B3C
      MousePointer    =   99  'Custom
      TabIndex        =   0
      ToolTipText     =   "Exit Window"
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FC9090&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   360
      TabIndex        =   1
      Top             =   2340
      Width           =   10710
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   2415
         Left            =   720
         TabIndex        =   2
         Top             =   480
         Width           =   14970
         Begin vkUserContolsXP.vkOptionButton btnOpt 
            Height          =   345
            Index           =   0
            Left            =   360
            TabIndex        =   3
            Top             =   120
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   609
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   1
            Alignment       =   2
         End
         Begin RichTextLib.RichTextBox opt1 
            Height          =   585
            Index           =   0
            Left            =   1080
            TabIndex        =   4
            Top             =   15
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   1032
            _Version        =   393217
            BackColor       =   14737632
            BorderStyle     =   0
            MaxLength       =   200
            Appearance      =   0
            TextRTF         =   $"popup_update_ques.frx":0C8E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin RichTextLib.RichTextBox opt1 
            Height          =   585
            Index           =   1
            Left            =   1080
            TabIndex        =   5
            Top             =   615
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   1032
            _Version        =   393217
            BackColor       =   14737632
            BorderStyle     =   0
            MaxLength       =   200
            Appearance      =   0
            TextRTF         =   $"popup_update_ques.frx":0DBA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin RichTextLib.RichTextBox opt1 
            Height          =   585
            Index           =   2
            Left            =   1080
            TabIndex        =   6
            Top             =   1215
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   1032
            _Version        =   393217
            BackColor       =   14737632
            BorderStyle     =   0
            MaxLength       =   200
            Appearance      =   0
            TextRTF         =   $"popup_update_ques.frx":0EE6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin RichTextLib.RichTextBox opt1 
            Height          =   585
            Index           =   3
            Left            =   1080
            TabIndex        =   7
            Top             =   1815
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   1032
            _Version        =   393217
            BackColor       =   14737632
            BorderStyle     =   0
            MaxLength       =   200
            Appearance      =   0
            TextRTF         =   $"popup_update_ques.frx":1012
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin vkUserContolsXP.vkOptionButton btnOpt 
            Height          =   345
            Index           =   1
            Left            =   360
            TabIndex        =   8
            Top             =   720
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   609
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   1
            Alignment       =   2
         End
         Begin vkUserContolsXP.vkOptionButton btnOpt 
            Height          =   345
            Index           =   3
            Left            =   360
            TabIndex        =   9
            Top             =   1920
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   609
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   1
            Alignment       =   2
         End
         Begin vkUserContolsXP.vkOptionButton btnOpt 
            Height          =   345
            Index           =   2
            Left            =   360
            TabIndex        =   10
            Top             =   1335
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   609
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   1
            Alignment       =   2
         End
         Begin VB.Line Line23 
            X1              =   960
            X2              =   960
            Y1              =   0
            Y2              =   2400
         End
         Begin VB.Line Line11 
            X1              =   0
            X2              =   15000
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line Line9 
            X1              =   0
            X2              =   15000
            Y1              =   2400
            Y2              =   2400
         End
         Begin VB.Line Line12 
            X1              =   0
            X2              =   15000
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Line Line13 
            X1              =   0
            X2              =   15000
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Line Line14 
            X1              =   0
            X2              =   15000
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Line Line8 
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   2400
         End
         Begin VB.Line Line16 
            X1              =   10815
            X2              =   10815
            Y1              =   0
            Y2              =   2400
         End
      End
      Begin VB.Line Line24 
         X1              =   1680
         X2              =   1680
         Y1              =   0
         Y2              =   480
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Correct"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   825
         TabIndex        =   17
         Top             =   120
         Width           =   810
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   240
         TabIndex        =   16
         Top             =   1200
         Width           =   165
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   240
         TabIndex        =   15
         Top             =   600
         Width           =   180
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No."
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   210
         TabIndex        =   14
         Top             =   120
         Width           =   360
      End
      Begin VB.Line Line5 
         X1              =   12550
         X2              =   15690
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   14160
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   15750
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   2880
      End
      Begin VB.Line Line6 
         X1              =   0
         X2              =   14160
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   240
         TabIndex        =   13
         Top             =   1800
         Width           =   165
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   240
         TabIndex        =   12
         Top             =   2385
         Width           =   180
      End
      Begin VB.Line Line7 
         X1              =   0
         X2              =   720
         Y1              =   1085
         Y2              =   1085
      End
      Begin VB.Line Line10 
         X1              =   0
         X2              =   720
         Y1              =   1685
         Y2              =   1685
      End
      Begin VB.Line Line15 
         X1              =   0
         X2              =   720
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line17 
         X1              =   720
         X2              =   720
         Y1              =   0
         Y2              =   480
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Choices"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1920
         TabIndex        =   11
         Top             =   120
         Width           =   840
      End
   End
   Begin RichTextLib.RichTextBox qtext_mcqs 
      Height          =   960
      Left            =   480
      TabIndex        =   18
      Top             =   645
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   1693
      _Version        =   393217
      BorderStyle     =   0
      Appearance      =   0
      TextRTF         =   $"popup_update_ques.frx":113E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox expn_mcq 
      Height          =   720
      Left            =   450
      TabIndex        =   19
      Top             =   5745
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   1270
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      Appearance      =   0
      TextRTF         =   $"popup_update_ques.frx":11BA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   525
      Width           =   10695
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Question :"
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
      Index           =   0
      Left            =   405
      TabIndex        =   24
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Choices : "
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
      Height          =   285
      Left            =   465
      TabIndex        =   23
      Top             =   1965
      Width           =   975
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Explanation : "
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
      Height          =   285
      Left            =   360
      TabIndex        =   22
      Top             =   5325
      Width           =   1410
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   5700
      Width           =   10695
   End
   Begin VB.Label ans_num 
      Caption         =   "ans_Num"
      Height          =   255
      Left            =   3765
      TabIndex        =   21
      Top             =   1725
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label aNswer_txt 
      Caption         =   "ans_txt"
      Height          =   255
      Left            =   4965
      TabIndex        =   20
      Top             =   1725
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "FrmQuesUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnOpt_MouseDown(Index As Integer, Button As MouseButtonConstants, Shift As Integer, Control As Integer, X As Long, Y As Long)
If Trim(opt1(Index).Text) = "" Then
 MsgBox "Enter Option value First in Option box " & vbCrLf & "Then select the correct answer No.", vbInformation + vbOKOnly, "Invalid Answer"
 opt1(Index).SetFocus
 btnOpt(Index).Value = vbUnchecked
 Exit Sub
End If
If Button = vbLeftButton Then
ans_num.Caption = Index + 1
aNswer_txt.Caption = opt1(Index).Text
End If
End Sub

Private Sub Command1_Click()
If Trim(qtext_mcqs.Text) = "" Then
MsgBox "Enter The Question ??? ", vbInformation + vbOKOnly, "Question Empty"
qtext_mcqs.SetFocus
Exit Sub
ElseIf Trim(opt1(0).Text) = "" Or Trim(opt1(1).Text) = "" Or Trim(opt1(2).Text) = "" Or Trim(opt1(3).Text) = "" Then
MsgBox "Enter All The 4 Options ??? ", vbInformation + vbOKOnly, "Fill All Options"
opt1(0).SetFocus
Exit Sub
ElseIf btnOpt(0).Value = vbUnchecked And btnOpt(1).Value = vbUnchecked And btnOpt(2).Value = vbUnchecked And btnOpt(3).Value = vbUnchecked Then
MsgBox "Select the Correct Answer !!! ", vbInformation + vbOKOnly, "Select Correct Answer"
opt1(0).SetFocus
Exit Sub
End If
If Trim(expn_mcq.Text) = "" Then
MsgBox "Please Enter Some Explanation at least Write the exact Answer.. ??? ", vbInformation + vbOKOnly, "Question Explanation"
expn_mcq.SetFocus
Exit Sub
End If
If MsgBox("Are You Sure to Update the question .. ?", vbQuestion + vbYesNo, "Are You Sure") = vbYes Then
c1.Execute ("update quesMS set q_txt='" & qtext_mcqs.Text & "',opt1='" & opt1(0).Text & "',opt2='" & opt1(1).Text & "',opt3='" & opt1(2).Text & "',opt4='" & opt1(3).Text & "',ANS_TXT='" & aNswer_txt.Caption & "',ANS_NO=" & Val(ans_num.Caption) & ",Q_EXPLN='" & expn_mcq.Text & "' where q_id='" & QUESTIONRight & "' ")
MsgBox "Question SuccessFully Updated...", vbInformation + vbOKOnly, "Update Success"
Unload Me
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_DblClick()
Unload Me
End Sub

Private Sub Form_Load()
 conn
 qtext_mcqs.Text = ""
For i = 0 To 3
 btnOpt(i).Value = vbUnchecked
 opt1(i).Text = ""
Next i
expn_mcq.Text = ""
aNswer_txt.Caption = ""
ans_num.Caption = ""

 Set r = New ADODB.Recordset
 Set r = c1.Execute("select * from quesms where q_id='" & QUESTIONRight & "' ")
If r.EOF = False Then
 qtext_mcqs.Text = r.Fields(6)
 opt1(0).Text = r.Fields(7)
 opt1(1).Text = r.Fields(8)
 opt1(2).Text = r.Fields(9)
 opt1(3).Text = r.Fields(10)
 aNswer_txt.Caption = r.Fields(11)
 ans_num.Caption = r.Fields(12)
 btnOpt(r.Fields(12) - 1).Value = vbChecked
 If IsNull(r.Fields(14)) = False Then
  expn_mcq.Text = r.Fields(14)
 Else
  expn_mcq.Text = ""
 End If
End If

End Sub

Private Sub Form_Unload(cancel As Integer)
 QUESTIONRight = ""
 ques_entry_dash.Enabled = True
End Sub

