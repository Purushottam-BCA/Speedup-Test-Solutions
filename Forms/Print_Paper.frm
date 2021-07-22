VERSION 5.00
Begin VB.Form Paper_Preview 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paper Print Preview"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6090
   ControlBox      =   0   'False
   Icon            =   "Print_Paper.frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   6090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton HTML_Print 
      BackColor       =   &H00FFFFFF&
      DisabledPicture =   "Print_Paper.frx":1E26
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   4200
      MouseIcon       =   "Print_Paper.frx":2893
      MousePointer    =   99  'Custom
      Picture         =   "Print_Paper.frx":29E5
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Currently Not Available"
      Top             =   1125
      Width           =   895
   End
   Begin VB.CommandButton PDF_print 
      BackColor       =   &H00FFFFFF&
      DisabledPicture =   "Print_Paper.frx":3452
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   2520
      MouseIcon       =   "Print_Paper.frx":3AAC
      MousePointer    =   99  'Custom
      Picture         =   "Print_Paper.frx":3BFE
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Click To Print Question Paper."
      Top             =   1125
      Width           =   980
   End
   Begin VB.CommandButton TXT_Print 
      BackColor       =   &H00FFFFFF&
      DisabledPicture =   "Print_Paper.frx":4656
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   840
      MouseIcon       =   "Print_Paper.frx":4CB0
      MousePointer    =   99  'Custom
      Picture         =   "Print_Paper.frx":4E02
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Currently Not Available"
      Top             =   1125
      Width           =   880
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      DisabledPicture =   "Print_Paper.frx":586F
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4800
      MouseIcon       =   "Print_Paper.frx":5EC9
      MousePointer    =   99  'Custom
      Picture         =   "Print_Paper.frx":601B
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Exit From Here"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Print Question Paper"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   1920
      TabIndex        =   10
      Top             =   120
      Width           =   2190
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HTML "
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4245
      TabIndex        =   5
      Top             =   2175
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PDF"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2580
      MouseIcon       =   "Print_Paper.frx":6675
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2175
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TEXT"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   2175
      Width           =   855
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Answer Key"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   225
      Left            =   4200
      MouseIcon       =   "Print_Paper.frx":67C7
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2385
      Width           =   990
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Answer Key"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   2535
      MouseIcon       =   "Print_Paper.frx":6919
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2400
      Width           =   990
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Answer Key"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   225
      Left            =   840
      MouseIcon       =   "Print_Paper.frx":6A6B
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   2385
      Width           =   990
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000013&
      FillColor       =   &H80000013&
      FillStyle       =   0  'Solid
      Height          =   530
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   6090
   End
End
Attribute VB_Name = "Paper_Preview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
Unload Question_PPR
Unload QuestionPPRdashboard
emp_dash.Hide
QuestionPPRdashboard.Show
End Sub

Private Sub Form_Load()
conn
CenterForm Me
 Label6.Enabled = False
 Label7.Enabled = False
 Label8.Enabled = False
 
If Ques_Include_Ans = 1 Then 'Check from Form 5
 Label6.Enabled = True
 Label7.Enabled = True
 Label8.Enabled = True
 End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.FontBold = False
Label6.FontUnderline = False
Label6.ForeColor = &H8080FF

Label7.FontBold = False
Label7.FontUnderline = False
Label7.ForeColor = &H8080FF

Label8.FontBold = False
Label8.FontUnderline = False
Label8.ForeColor = &H8080FF

End Sub

Private Sub HTML_Print_Click()
Dim s As String 'Using Date and Time as report name
s = "\Question_Paper\Paper_" & Format(Date, "dd-mm-yy") & "_" & Format(Now, "hh-mm") & ".html"
Ques_PPR.ExportReport rptKeyHTML, App.Path & s, , True, rptRangeAllPages
MsgBox "Question Printed"
End Sub

Private Sub Label2_Click()
PDF_print_Click
End Sub

Private Sub Label6_Click()
Dim s As String 'Using Date and Time as report name
Dim obj As Object
s = "\Question_Paper\Ans_" & Format(Date, "dd-mm-yy") & "_" & Format(Now, "hh-mm") & ".txt"
Ans_Key.ExportReport rptKeyText, App.Path & s, , True, rptRangeAllPages
MsgBox "Answer Key Printed"
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.FontBold = True
Label6.FontUnderline = True
Label6.ForeColor = vbRed
End Sub

Private Sub Label7_Click()
Ans_Key.Show vbModal, MDI
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.FontBold = True
Label7.FontUnderline = True
Label7.ForeColor = vbRed
End Sub

Private Sub Label8_Click()
Dim s As String 'Using Date and Time as report name
s = "\Question_Paper\ANS_" & Format(Date, "dd-mm-yy") & "_" & Format(Now, "hh-mm") & ".html"
Ans_Key.ExportReport rptKeyHTML, App.Path & s, , True, rptRangeAllPages
MsgBox "Answer Key Printed"
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.FontBold = True
Label8.FontUnderline = True
Label8.ForeColor = vbRed
End Sub

Private Sub PDF_print_Click()
On Error Resume Next
Command1.Enabled = True
If WantInstruction = 1 Then
 DV.rsQpaper.Open
 Set Ques_PPR.Sections("section4").Controls.Item("image1").Picture = LoadPicture(school_pic)
 Ques_PPR.Sections("section4").Controls("org_name").Caption = school_nm
 Ques_PPR.Sections("section4").Controls("org_address").Caption = school_add
 Ques_PPR.Sections("section4").Controls("test_name").Caption = Test_nm
 Ques_PPR.Sections("section4").Controls("org_sub").Caption = testSUB_nm
 Ques_PPR.Sections("section4").Controls("org_class").Caption = testclass_nm
 Ques_PPR.Sections("section4").Controls("org_full_mrk").Caption = testFULLmrk
 Ques_PPR.Sections("section4").Controls("org_total_time").Caption = testTOTALtime
 Ques_PPR.Sections("section4").Controls("org_totques").Caption = testTotQues
 Ques_PPR.Sections("section4").Controls("org_mark_p_ques").Caption = testCorrectMRK
 Ques_PPR.Sections("section4").Controls("org_total_mrk").Caption = testFULLmrk
 Ques_PPR.Sections("section4").Controls("label12").Caption = instructionSET
 Ques_PPR.Refresh
 Ques_PPR.Show vbModal, MDI
 Ques_PPR.Refresh
 DV.rsQpaper.Close
Else
 DV.rsQpaper.Open
 Set Ques_PPR.Sections("section4").Controls.Item("image1").Picture = LoadPicture(school_pic)
 Ques_ppr2.Sections("section4").Controls("org_name").Caption = school_nm
 Ques_ppr2.Sections("section4").Controls("org_address").Caption = school_add
 Ques_ppr2.Sections("section4").Controls("test_name").Caption = Test_nm
 Ques_ppr2.Sections("section4").Controls("org_sub").Caption = testSUB_nm
 Ques_ppr2.Sections("section4").Controls("org_class").Caption = testclass_nm
 Ques_ppr2.Sections("section4").Controls("org_full_mrk").Caption = testFULLmrk
 Ques_ppr2.Sections("section4").Controls("org_total_time").Caption = testTOTALtime
 Ques_ppr2.Sections("section4").Controls("org_totques").Caption = testTotQues
 Ques_ppr2.Sections("section4").Controls("org_mark_p_ques").Caption = testCorrectMRK
 Ques_ppr2.Sections("section4").Controls("org_total_mrk").Caption = testFULLmrk
 Ques_ppr2.Refresh
 Ques_ppr2.Show vbModal, MDI
 Ques_ppr2.Refresh
 DV.rsQpaper.Close
End If
End Sub

Private Sub TXT_Print_Click()
Dim s As String 'Using Date and Time as report name
Dim obj As Object
s = "\Question_Paper\Paper_" & Format(Date, "dd-mm-yy") & "_" & Format(Now, "hh-mm") & ".txt"
Ques_PPR.ExportReport rptKeyText, App.Path & s, , True, rptRangeAllPages
MsgBox "Question Printed"
'Set obj = CreateObject("Word.basic")
'obj.fileopen App.Path & "\Question_Paper\.txt"
'obj.filesaveas App.Path & "\Question_Paper\Paper_19-04-19_03-33.Doc", 6
'obj.filedelete
'obj.quit
'Set obj = Nothing
End Sub
