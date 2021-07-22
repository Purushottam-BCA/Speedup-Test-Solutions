VERSION 5.00
Begin VB.Form Question_PPR 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Question Paper Generator"
   ClientHeight    =   10695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20400
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10695
   ScaleWidth      =   20400
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton xpButton2 
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
      Left            =   18840
      MouseIcon       =   "Question_PPR.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "Question_PPR.frx":0152
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Print Certificate"
      Top             =   9930
      Width           =   1305
   End
   Begin VB.CommandButton btnbck 
      Height          =   400
      Left            =   120
      MouseIcon       =   "Question_PPR.frx":0B07
      MousePointer    =   99  'Custom
      Picture         =   "Question_PPR.frx":0C59
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   10000
      Width           =   1215
   End
   Begin VB.CommandButton btnRandom 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Randomize again"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13800
      MouseIcon       =   "Question_PPR.frx":138C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   9980
      Width           =   2535
   End
   Begin VB.ListBox FrontList1 
      BackColor       =   &H00DCDCDC&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8460
      ItemData        =   "Question_PPR.frx":14DE
      Left            =   120
      List            =   "Question_PPR.frx":14E0
      MouseIcon       =   "Question_PPR.frx":14E2
      MousePointer    =   99  'Custom
      Style           =   1  'Checkbox
      TabIndex        =   37
      Top             =   1440
      Width           =   9855
   End
   Begin VB.TextBox txt6 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   18435
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   150
      Width           =   480
   End
   Begin VB.TextBox txt7 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   19320
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   150
      Width           =   480
   End
   Begin VB.TextBox txt1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1035
      TabIndex        =   14
      Top             =   150
      Width           =   2160
   End
   Begin VB.TextBox txt2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4995
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   150
      Width           =   2520
   End
   Begin VB.TextBox txt4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11835
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   150
      Width           =   2040
   End
   Begin VB.TextBox txt5 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15315
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   150
      Width           =   2040
   End
   Begin VB.TextBox txt3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9315
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   150
      Width           =   1320
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404080&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   10440
      TabIndex        =   3
      Top             =   840
      Width           =   9615
      Begin VB.Label Top2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   4815
         TabIndex        =   4
         Top             =   120
         Width           =   105
      End
   End
   Begin VB.ListBox FrontList2 
      BackColor       =   &H00DCDCDC&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8460
      Left            =   10320
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   1440
      Width           =   9855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404080&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   9615
      Begin VB.Label Top1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL  QUESTIONS"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   3360
         TabIndex        =   2
         Top             =   120
         Width           =   2535
      End
   End
   Begin VB.ListBox List1 
      Height          =   1410
      Left            =   10560
      Style           =   1  'Checkbox
      TabIndex        =   29
      Top             =   5160
      Width           =   975
   End
   Begin VB.ListBox List2 
      Height          =   1410
      Left            =   11520
      Style           =   1  'Checkbox
      TabIndex        =   30
      Top             =   5160
      Width           =   2535
   End
   Begin VB.ListBox List3 
      Height          =   1410
      Left            =   14040
      Style           =   1  'Checkbox
      TabIndex        =   31
      Top             =   5280
      Width           =   1575
   End
   Begin VB.ListBox List4 
      Height          =   1410
      Left            =   15720
      Style           =   1  'Checkbox
      TabIndex        =   32
      Top             =   5280
      Width           =   1455
   End
   Begin VB.ListBox List5 
      Height          =   1410
      Left            =   17160
      Style           =   1  'Checkbox
      TabIndex        =   33
      Top             =   5280
      Width           =   1335
   End
   Begin VB.ListBox List6 
      Height          =   1410
      Left            =   18480
      Style           =   1  'Checkbox
      TabIndex        =   34
      Top             =   5280
      Width           =   1455
   End
   Begin VB.ListBox List7 
      Height          =   1410
      Left            =   17520
      Style           =   1  'Checkbox
      TabIndex        =   35
      Top             =   6840
      Width           =   1215
   End
   Begin VB.ListBox List8 
      Height          =   1410
      Left            =   18720
      Style           =   1  'Checkbox
      TabIndex        =   36
      Top             =   6840
      Width           =   1215
   End
   Begin VB.ListBox ListAns_No 
      Height          =   1185
      Left            =   8760
      Style           =   1  'Checkbox
      TabIndex        =   28
      Top             =   6720
      Width           =   1095
   End
   Begin VB.ListBox Listopt4 
      Height          =   1410
      Left            =   8400
      Style           =   1  'Checkbox
      TabIndex        =   26
      Top             =   5160
      Width           =   1455
   End
   Begin VB.ListBox Listopt3 
      Height          =   1410
      Left            =   7200
      Style           =   1  'Checkbox
      TabIndex        =   25
      Top             =   5160
      Width           =   1215
   End
   Begin VB.ListBox Listopt2 
      Height          =   1410
      Left            =   5760
      Style           =   1  'Checkbox
      TabIndex        =   24
      Top             =   5160
      Width           =   1455
   End
   Begin VB.ListBox Listopt1 
      Height          =   1410
      Left            =   4440
      Style           =   1  'Checkbox
      TabIndex        =   23
      Top             =   5160
      Width           =   1335
   End
   Begin VB.ListBox ListQ_text 
      Height          =   1410
      Left            =   1440
      Style           =   1  'Checkbox
      TabIndex        =   22
      Top             =   5160
      Width           =   3015
   End
   Begin VB.ListBox ListQ_No 
      Height          =   1410
      Left            =   240
      Style           =   1  'Checkbox
      TabIndex        =   21
      Top             =   5280
      Width           =   1215
   End
   Begin VB.ListBox ListAns_text 
      Height          =   1185
      Left            =   7680
      Style           =   1  'Checkbox
      TabIndex        =   27
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00733C00&
      Height          =   270
      Left            =   17640
      TabIndex        =   20
      Top             =   195
      Width           =   525
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "hr"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00733C00&
      Height          =   270
      Left            =   19005
      TabIndex        =   19
      Top             =   195
      Width           =   225
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "min"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00733C00&
      Height          =   270
      Left            =   19845
      TabIndex        =   18
      Top             =   195
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date "
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00733C00&
      Height          =   270
      Left            =   315
      TabIndex        =   15
      Top             =   210
      Width           =   555
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Test Name "
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00733C00&
      Height          =   270
      Left            =   3600
      TabIndex        =   13
      Top             =   210
      Width           =   1185
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Class"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00733C00&
      Height          =   270
      Left            =   11040
      TabIndex        =   11
      Top             =   210
      Width           =   585
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00733C00&
      Height          =   270
      Left            =   14280
      TabIndex        =   9
      Top             =   210
      Width           =   765
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Max. Marks"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00733C00&
      Height          =   270
      Left            =   7920
      TabIndex        =   7
      Top             =   210
      Width           =   1185
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   11655
      TabIndex        =   6
      Top             =   915
      Width           =   105
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   735
      Left            =   10320
      Top             =   720
      Width           =   9855
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   735
      Left            =   120
      Top             =   720
      Width           =   9855
   End
End
Attribute VB_Name = "Question_PPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim totQues As Integer

Private Sub btnRandom_Click()
 FrontList2.Clear
 List1.Clear
 List2.Clear
 List3.Clear
 List4.Clear
 List5.Clear
 List6.Clear
 List7.Clear
 List8.Clear
 c.Execute ("drop sequence QNOGENERATOR2") 'Deleting old Sequence if exist
 c.Execute ("create sequence QNOGENERATOR2 increment by 1 start with 1 cache 100") 'Creating New Sequence start with 1

 c.Execute ("delete from QPAPERFINAL")
 c.Execute ("INSERT INTO QPAPERFINAL select * from ( select * from QTESTBANK order by dbms_random.value) where rownum<=" & testTotQues & " ")
 c.Execute ("update QPAPERFINAL set Q_NO = QNOGENERATOR2.nextval")
 Set r = New ADODB.Recordset
 Set r = c.Execute("select * from QPAPERFINAL order by Q_NO")
 While r.EOF = False
  FrontList2.AddItem r.Fields(0) & Space(7) & r.Fields(1)
 List1.AddItem r.Fields(0)
 List2.AddItem r.Fields(1)
 List3.AddItem r.Fields(2)
 List4.AddItem r.Fields(3)
 List5.AddItem r.Fields(4)
 List6.AddItem r.Fields(5)
 List7.AddItem r.Fields(6)
 List8.AddItem r.Fields(7)
  r.MoveNext
 Wend
End Sub



Private Sub cmd1_Click()

End Sub

Private Sub FrontList1_Click()

If FrontList1.SelCount > testTotQues Then
 MsgBox "Cannot Insert More Questions " & vbCrLf & "Remove Some Quesions then add new ", vbInformation + vbOKOnly, "Question Overflow"
Exit Sub
End If
 List1.Clear 'Q_No
 List2.Clear 'Q_text
 List3.Clear 'Opt1
 List4.Clear 'Opt2
 List5.Clear 'Opt3
 List6.Clear 'Opt4
 List7.Clear 'Ans_text
 List8.Clear 'Ans_No
 FrontList2.Clear

 Dim n As Integer
 For n = 0 To FrontList1.ListCount - 1
 If FrontList1.Selected(n) Then
  List1.AddItem ListQ_No.list(n)
  List2.AddItem ListQ_text.list(n)
  List3.AddItem Listopt1.list(n)
  List4.AddItem Listopt2.list(n)
  List5.AddItem Listopt3.list(n)
  List6.AddItem Listopt4.list(n)
  List7.AddItem ListAns_text.list(n)
  List8.AddItem ListAns_No.list(n)
  FrontList2.AddItem FrontList1.list(n)
 End If
Next n

  c.Execute ("delete from QPAPERFINAL") 'Now This Is Blank
  List1.Clear
  For n = 0 To FrontList2.ListCount - 1
  List1.AddItem n + 1  'Updating With Sequence Number Questions
  sql = "insert into QPAPERFINAL values(" & List1.list(n) & ",'" & List2.list(n) & "','" & List3.list(n) & "','" & List4.list(n) & "','" & List5.list(n) & "','" & List6.list(n) & "','" & List7.list(n) & "','" & List8.list(n) & "')"
  Set r = c.Execute(sql)
 Next n
End Sub

Private Sub btnbck_Click()
Unload Me
QpaprSetup.Show
End Sub

Private Sub Form_Load()
Dim i As Integer
On Error Resume Next
conn
Me.Top = 0
Me.Left = 0
FrontList1.Enabled = True
btnRandom.Visible = False

txt1.Text = Format(Date, "dd-mmm-yyyy")
txt2.Text = Test_nm
txt3.Text = testFULLmrk
txt4.Text = testclass_nm
txt5.Text = testSUB_nm
txt6.Text = ppr_Time_hr
txt7.Text = ppr_time_min

Top2.Caption = Ques_Cat 'Type of question Selected (random or selected)
Set r1 = New ADODB.Recordset
c.Execute ("delete from QPAPERFINAL")
c.Execute ("delete from QTESTBANK")   'Erase Previous data
Set r = New ADODB.Recordset
Set r = c1.Execute("select count(*)from user_sequences where sequence_name='QNOGENERATOR1' ")
If r.Fields(0) > 0 Then
 c.Execute ("drop sequence QNOGENERATOR1") 'Deleting old Sequence if exist
 c.Execute (" create sequence QNOGENERATOR1 increment by 1 start with 1 cache 100") 'Creating New Sequence start with 1
Else
 c.Execute (" create sequence QNOGENERATOR1 increment by 1 start with 1 cache 100") 'Creating New Sequence start with 1
End If

Set r = New ADODB.Recordset
Set r = c1.Execute("select count(*)from user_sequences where sequence_name='QNOGENERATOR2' ")
If r.Fields(0) > 0 Then
 c.Execute ("drop sequence QNOGENERATOR2") 'Deleting old Sequence if exist
 c.Execute ("create sequence QNOGENERATOR2 increment by 1 start with 1 cache 100") 'Creating New Sequence start with 1
Else
 c.Execute ("create sequence QNOGENERATOR2 increment by 1 start with 1 cache 100") 'Creating New Sequence start with 1
End If
r.Close

If Ques_Cat = "Selected Questions" Then 'Select Catogary
 If Ques_selection = "Subject Wise Question" Then  'select Mode is subject wise
   If Ques_diff_leval = "" Then 'All Type of questions
            c.Execute ("insert into QTESTBANK select q_no,q_txt,opt1,opt2,opt3,opt4,ans_txt,ans_no from quesms where c_id=(select c_id from course where c_nm='" & Ques_Course & "') and sub_id=(select sub_id from sub where sub_nm='" & Ques_Subject & "' and c_id =(select c_id from course where c_nm='" & Ques_Course & "'))")
            c.Execute ("update QTESTBANK set Q_NO = QNOGENERATOR1.nextval")
   Else 'Easy/Medium/Hard Leval is selected
            c.Execute ("insert into QTESTBANK select q_no,q_txt,opt1,opt2,opt3,opt4,ans_txt,ans_no from quesms where c_id=(select c_id from course where c_nm='" & Ques_Course & "') and sub_id=(select sub_id from sub where sub_nm='" & Ques_Subject & "' and c_id =(select c_id from course where c_nm='" & Ques_Course & "'))and q_dif_lvl='" & Ques_diff_leval & "' ")
            c.Execute ("update QTESTBANK set Q_NO = QNOGENERATOR1.nextval")
   End If
 Else 'Topic Wise Question paper Generation
  If Ques_diff_leval = "" Then 'All Type of questions
   For i = 0 To QpaprSetup.List1.ListCount - 1
   If QpaprSetup.List1.Selected(i) Then
    c.Execute ("insert into QTESTBANK select q_no,q_txt,opt1,opt2,opt3,opt4,ans_txt,ans_no from quesms where c_id=(select c_id from course where c_nm='" & Ques_Course & "') and sub_id=(select sub_id from sub where sub_nm='" & Ques_Subject & "' and c_id =(select c_id from course where c_nm='" & Ques_Course & "')) and tp_id=(select tp_id from topic where tp_nm='" & QpaprSetup.List1.list(i) & "' and sub_id=(select sub_id from sub where sub_nm='" & Ques_Subject & "' and c_id =(select c_id from course where c_nm='" & Ques_Course & "'))) ")
   End If
   Next i
    c.Execute ("update QTESTBANK set Q_NO = QNOGENERATOR1.nextval")
  Else 'Easy/Medium/Hard Leval is selected
   For i = 0 To QpaprSetup.List1.ListCount - 1
   If QpaprSetup.List1.Selected(i) Then
    c.Execute ("insert into QTESTBANK select q_no,q_txt,opt1,opt2,opt3,opt4,ans_txt,ans_no from quesms where c_id=(select c_id from course where c_nm='" & Ques_Course & "') and sub_id=(select sub_id from sub where sub_nm='" & Ques_Subject & "' and c_id =(select c_id from course where c_nm='" & Ques_Course & "')) and tp_id=(select tp_id from topic where tp_nm='" & QpaprSetup.List1.list(i) & "' and sub_id=(select sub_id from sub where sub_nm='" & Ques_Subject & "' and c_id =(select c_id from course where c_nm='" & Ques_Course & "')))and q_dif_lvl='" & Ques_diff_leval & "'")
   End If
   Next i
    c.Execute ("update QTESTBANK set Q_NO = QNOGENERATOR1.nextval")
   End If
 End If
Else 'Random Questions
'Disable the first list so that user cannot add more
FrontList1.Enabled = False
btnRandom.Visible = True

   If Ques_selection = "Subject Wise Question" Then  'select Mode is subject wise
   If Ques_diff_leval = "" Then 'All Type of questions
    c.Execute ("INSERT INTO  QTESTBANK Select q_no,q_txt,opt1,opt2,opt3,opt4,ans_txt,ans_no from quesms where c_id=(select c_id from course where c_nm='" & Ques_Course & "') and sub_id=(select sub_id from sub where sub_nm='" & Ques_Subject & "' and c_id =(select c_id from course where c_nm='" & Ques_Course & "'))")
    c.Execute ("update QTESTBANK set Q_NO = QNOGENERATOR1.nextval")
    c.Execute ("INSERT INTO QPAPERFINAL select * from ( select * from QTESTBANK order by dbms_random.value) where rownum<=" & testTotQues & " ")
    c.Execute ("update QPAPERFINAL set Q_NO = QNOGENERATOR2.nextval")
   Else 'Easy/Medium/Hard Leval is selected
     c.Execute ("INSERT INTO  QTESTBANK select q_no,q_txt,opt1,opt2,opt3,opt4,ans_txt,ans_no from quesms where c_id=(select c_id from course where c_nm='" & Ques_Course & "') and sub_id=(select sub_id from sub where sub_nm='" & Ques_Subject & "' and c_id =(select c_id from course where c_nm='" & Ques_Course & "'))and q_dif_lvl='" & Ques_diff_leval & "' ")
     c.Execute ("update QTESTBANK set Q_NO = QNOGENERATOR1.nextval")
     c.Execute ("INSERT INTO QPAPERFINAL select * from ( select * from QTESTBANK order by dbms_random.value) where rownum<=" & testTotQues & " ")
     c.Execute ("update QPAPERFINAL set Q_NO = QNOGENERATOR2.nextval")
   End If
 Else 'Topic Wise Question paper Generation
   If Ques_diff_leval = "" Then 'All Type of questions
    For i = 0 To QpaprSetup.List1.ListCount - 1
   If QpaprSetup.List1.Selected(i) Then
    c.Execute ("insert into QTESTBANK select q_no,q_txt,opt1,opt2,opt3,opt4,ans_txt,ans_no from quesms where c_id=(select c_id from course where c_nm='" & Ques_Course & "') and sub_id=(select sub_id from sub where sub_nm='" & Ques_Subject & "' and c_id =(select c_id from course where c_nm='" & Ques_Course & "')) and tp_id=(select tp_id from topic where tp_nm='" & QpaprSetup.List1.list(i) & "' and sub_id=(select sub_id from sub where sub_nm='" & Ques_Subject & "' and c_id =(select c_id from course where c_nm='" & Ques_Course & "'))) ")
   End If
   Next i
    c.Execute ("update QTESTBANK set Q_NO = QNOGENERATOR1.nextval")
    c.Execute ("INSERT INTO QPAPERFINAL select * from ( select * from QTESTBANK order by dbms_random.value) where rownum<=" & testTotQues & " ")
    c.Execute ("update QPAPERFINAL set Q_NO = QNOGENERATOR2.nextval")
   Else 'Easy/Medium/Hard Leval is selected
     For i = 0 To QpaprSetup.List1.ListCount - 1
   If QpaprSetup.List1.Selected(i) Then
    c.Execute ("insert into QTESTBANK select q_no,q_txt,opt1,opt2,opt3,opt4,ans_txt,ans_no from quesms where c_id=(select c_id from course where c_nm='" & Ques_Course & "') and sub_id=(select sub_id from sub where sub_nm='" & Ques_Subject & "' and c_id =(select c_id from course where c_nm='" & Ques_Course & "')) and tp_id=(select tp_id from topic where tp_nm='" & QpaprSetup.List1.list(i) & "' and sub_id=(select sub_id from sub where sub_nm='" & Ques_Subject & "' and c_id =(select c_id from course where c_nm='" & Ques_Course & "')))and q_dif_lvl='" & Ques_diff_leval & "'")
   End If
   Next i
    c.Execute ("update QTESTBANK set Q_NO = QNOGENERATOR1.nextval")
    c.Execute ("INSERT INTO QPAPERFINAL select * from ( select * from QTESTBANK order by dbms_random.value) where rownum<=" & testTotQues & " ")
    c.Execute ("update QPAPERFINAL set Q_NO = QNOGENERATOR2.nextval")
   End If
 End If
End If

'++++++ Now Showing Questions in ListBox ++++++'
'--------- Now adding for selected choices---------'

ListQ_No.Clear  'Clearing Existing record
ListQ_text.Clear
Listopt1.Clear
Listopt2.Clear
Listopt3.Clear
Listopt4.Clear
 ListAns_text.Clear
 ListAns_No.Clear
FrontList1.Clear

Set r = New ADODB.Recordset
Set r = c.Execute("select * from QTESTBANK")
While r.EOF = False
 'Actual ListBox Containg Questions and answers But Back Side
  ListQ_No.AddItem r.Fields(0)
  ListQ_text.AddItem r.Fields(1)
  Listopt1.AddItem r.Fields(2)
  Listopt2.AddItem r.Fields(3)
  Listopt3.AddItem r.Fields(4)
  Listopt4.AddItem r.Fields(5)
  ListAns_text.AddItem r.Fields(6)
  ListAns_No.AddItem r.Fields(7)
 r.MoveNext
Wend
'r.Close

Set r = New ADODB.Recordset
Set r = c.Execute("select * from QTESTBANK")
While r.EOF = False
 FrontList1.AddItem "#" & Space(2) & r.Fields(1)
r.MoveNext
Wend


'--------- Now adding for Random choices---------'
If Ques_Cat = "Random Questions" Then
 Set r = New ADODB.Recordset
 Set r = c.Execute("select * from QPAPERFINAL order by Q_NO")

 List1.Clear
 List2.Clear
 List3.Clear
 List4.Clear
 List5.Clear
 List6.Clear
 List7.Clear
 List8.Clear

 While r.EOF = False
  FrontList2.AddItem r.Fields(0) & Space(7) & r.Fields(1)

 List1.AddItem r.Fields(0)
 List2.AddItem r.Fields(1)
 List3.AddItem r.Fields(2)
 List4.AddItem r.Fields(3)
 List5.AddItem r.Fields(4)
 List6.AddItem r.Fields(5)
 List7.AddItem r.Fields(6)
 List8.AddItem r.Fields(7)
  r.MoveNext
 Wend
End If

End Sub

'autoId For DashBoard
Public Function autoGenNum() As Integer
Set r1 = New ADODB.Recordset
Set r1 = c1.Execute("select count(*) from qpaprdash")
 t = r1.Fields(0)
 autoGenNum = t + 1
End Function

Private Sub xpButton2_Click()
Dim serial As Integer
If FrontList2.ListCount < testTotQues Then
 MsgBox "Some Questions Left !!Please Insert more questions to continue", vbExclamation + vbOKOnly, " "
Exit Sub
Else
'++++++++++++++++++++++++++Inserting For Dashboard++++++++++++++++++++++++++++++++++
serial = autoGenNum()
Set r = c.Execute("insert into qpaprdash values(" & serial & ",'" & Format(ordrdt, "dd-mmm-yyyy") & "','" & Format(delivrdt, "dd-mmm-yyyy") & "','" & purposedash & "', '" & tstTypDash & "','" & classdash & "', '" & subdash & "'," & totqsdash & "," & totmrkdash & ")")
Paper_Preview.Show vbModal, MDI
End If
End Sub
