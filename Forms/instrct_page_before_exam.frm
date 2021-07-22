VERSION 5.00
Begin VB.Form Stu_Test_selection 
   BackColor       =   &H0080FF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Instructions"
   ClientHeight    =   9000
   ClientLeft      =   -15
   ClientTop       =   330
   ClientWidth     =   10815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   10815
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   9090
      Left            =   0
      Picture         =   "instrct_page_before_exam.frx":0000
      ScaleHeight     =   9030
      ScaleWidth      =   10770
      TabIndex        =   0
      Top             =   -120
      Width           =   10830
      Begin VB.CommandButton btnnext 
         Height          =   400
         Left            =   9480
         MouseIcon       =   "instrct_page_before_exam.frx":248C
         MousePointer    =   99  'Custom
         Picture         =   "instrct_page_before_exam.frx":25DE
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   8400
         Width           =   1215
      End
      Begin VB.CommandButton ChameleonBtn1 
         Height          =   400
         Left            =   215
         MouseIcon       =   "instrct_page_before_exam.frx":2D1A
         MousePointer    =   99  'Custom
         Picture         =   "instrct_page_before_exam.frx":2E6C
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   8400
         Width           =   1215
      End
      Begin VB.Timer Timer3 
         Left            =   4920
         Top             =   360
      End
      Begin VB.Timer Timer2 
         Left            =   4440
         Top             =   360
      End
      Begin VB.Timer Timer1 
         Left            =   3960
         Top             =   360
      End
      Begin VB.Frame FulFrame 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   975
         Left            =   6000
         TabIndex        =   16
         Top             =   4920
         Width           =   3375
         Begin VB.ComboBox lcombo3 
            BackColor       =   &H80000006&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   405
            Left            =   1320
            MouseIcon       =   "instrct_page_before_exam.frx":359F
            MousePointer    =   99  'Custom
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   240
            Width           =   1935
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
            Left            =   1150
            TabIndex        =   25
            Top             =   240
            Width           =   105
         End
         Begin VB.Label Label6 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Level"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   18
            Top             =   280
            Width           =   975
         End
      End
      Begin VB.Frame TPFrame 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "1"
         Height          =   3855
         Left            =   600
         TabIndex        =   10
         Top             =   4800
         Width           =   4455
         Begin VB.ListBox List1 
            BackColor       =   &H80000006&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1860
            Left            =   1320
            MouseIcon       =   "instrct_page_before_exam.frx":36F1
            MousePointer    =   99  'Custom
            Style           =   1  'Checkbox
            TabIndex        =   29
            Top             =   1680
            Width           =   3015
         End
         Begin VB.ComboBox scombo1 
            BackColor       =   &H80000006&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   405
            Left            =   1320
            MouseIcon       =   "instrct_page_before_exam.frx":3843
            MousePointer    =   99  'Custom
            TabIndex        =   12
            Top             =   1030
            Width           =   2655
         End
         Begin VB.ComboBox lcombo1 
            BackColor       =   &H80000006&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   405
            Left            =   1320
            MouseIcon       =   "instrct_page_before_exam.frx":3995
            MousePointer    =   99  'Custom
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label11 
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
            Left            =   1150
            TabIndex        =   23
            Top             =   1680
            Width           =   105
         End
         Begin VB.Label Label7 
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
            Left            =   1150
            TabIndex        =   20
            Top             =   1080
            Width           =   105
         End
         Begin VB.Label Label17 
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
            Left            =   1150
            TabIndex        =   19
            Top             =   360
            Width           =   105
         End
         Begin VB.Label Label3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Topics"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   15
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Subject"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   14
            Top             =   1035
            Width           =   975
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Level"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   13
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame SubFrame 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1815
         Left            =   3480
         TabIndex        =   5
         Top             =   4800
         Width           =   4215
         Begin VB.ComboBox lcombo2 
            BackColor       =   &H80000006&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   405
            Left            =   1320
            MouseIcon       =   "instrct_page_before_exam.frx":3AE7
            MousePointer    =   99  'Custom
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   360
            Width           =   1935
         End
         Begin VB.ComboBox Scombo2 
            BackColor       =   &H80000006&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   405
            Left            =   1320
            MouseIcon       =   "instrct_page_before_exam.frx":3C39
            MousePointer    =   99  'Custom
            TabIndex        =   6
            Top             =   1080
            Width           =   2655
         End
         Begin VB.Label Label12 
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
            Left            =   1150
            TabIndex        =   24
            Top             =   1080
            Width           =   105
         End
         Begin VB.Label Label8 
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
            Left            =   1150
            TabIndex        =   21
            Top             =   360
            Width           =   105
         End
         Begin VB.Label Label5 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Level"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   400
            TabIndex        =   9
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Subject"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   300
            TabIndex        =   7
            Top             =   1080
            Width           =   975
         End
      End
      Begin VB.CommandButton SubjectFrame 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Subject Wise Test"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   4440
         MouseIcon       =   "instrct_page_before_exam.frx":3D8B
         MousePointer    =   99  'Custom
         Picture         =   "instrct_page_before_exam.frx":3EDD
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Subject Wise Test"
         Top             =   2640
         Width           =   2055
      End
      Begin VB.CommandButton TopicFrame 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Topic Wise Test"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   1560
         MouseIcon       =   "instrct_page_before_exam.frx":524E
         MousePointer    =   99  'Custom
         Picture         =   "instrct_page_before_exam.frx":53A0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Topic Wise Test"
         Top             =   2640
         Width           =   2055
      End
      Begin VB.CommandButton FullFrame 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Full Length test"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   7200
         MouseIcon       =   "instrct_page_before_exam.frx":6965
         MousePointer    =   99  'Custom
         Picture         =   "instrct_page_before_exam.frx":6AB7
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Full Mock Test"
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Image Image1 
         Height          =   1050
         Left            =   0
         Picture         =   "instrct_page_before_exam.frx":7F02
         Stretch         =   -1  'True
         Top             =   0
         Width           =   10815
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Height          =   480
         Left            =   20
         Top             =   1070
         Width           =   10755
      End
      Begin VB.Label label10 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5160
         TabIndex        =   28
         Top             =   2040
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lbl3 
         BackColor       =   &H0080FF80&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5160
         TabIndex        =   27
         Top             =   1680
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lbl2 
         BackColor       =   &H00FFC0FF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5160
         TabIndex        =   26
         Top             =   1920
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lbl1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5160
         TabIndex        =   22
         Top             =   1800
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " -: Choose a Test Type to begin the test :-"
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
         Left            =   0
         TabIndex        =   1
         Top             =   1155
         Width           =   4365
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000080FF&
         Height          =   520
         Left            =   0
         Top             =   1030
         Width           =   10935
      End
   End
End
Attribute VB_Name = "Stu_Test_selection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnnext_Click()
 Set r = New ADODB.Recordset
If Val(Label10.Caption) = 1 Then 'Topic Wise Test
 selectedType = "Topic Wise Test"
 Set r = c.Execute("select * from TOPICWISETEST where c_id='" & lbl1.Caption & "' ")
 If r.EOF = False Then
  FTOTQUESTION = r.Fields(1)
  FTOTTIMEMINUTE = r.Fields(2)
  FTOTTIMESECOND = r.Fields(3)
  FTOTMARKS = r.Fields(4)
  FPASSPERCENTG = r.Fields(5)
  FMRKPERCOR = r.Fields(6)
  FMRKPERWRONG = r.Fields(7)
 Else
  MsgBox "Oops, No SetUp For Topic Wise Test.. Contact Admin to Set Test Properties Of Topic Wise Test", vbCritical + vbOKOnly, "Topic Wise Test"
  Refr 'Function With Refresh
  Exit Sub
 End If
 selectedlvl = lcombo1.Text
 IsFullLengthSelected = 0
 If lcombo1.ListIndex <> 3 Then
   For i = 0 To List1.ListCount - 1
   If List1.Selected(i) Then
     CurrentTopic = CurrentTopic & List1.list(i)
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id=(select sub_id from sub where sub_nm='" & scombo1.Text & "' and c_id='" & lbl1.Caption & "')and tp_id=(select tp_id from topic where tp_nm='" & List1.list(i) & "' and sub_id=(select sub_id from sub where sub_nm='" & scombo1.Text & "' and c_id='" & lbl1.Caption & "')and c_id='" & lbl1.Caption & "') and upper(q_dif_lvl)='" & UCase(lcombo1.Text) & "'order by dbms_random.value)where rownum <" & (FTOTQUESTION / List1.SelCount) + 1 & " ")
   End If
   Next i
  chkupdate
  c.Execute ("delete from mcqtest where q_no >" & FTOTQUESTION & " ")
  Set r1 = New ADODB.Recordset
  Set r1 = c1.Execute("select count(*) from mcqtest")
  If FTOTQUESTION > r1.Fields(0) Then
   MsgBox "Not Enough Questions Available in Question Bank for this Property,Try Other", vbCritical + vbOKOnly, ""
   Refr 'Function With Refresh
   Exit Sub
  End If
  Timer1.Enabled = False
  Timer2.Enabled = False
  Timer3.Enabled = False
  Unload stu_dash
  Unload Me
  Instruction_General.Show
 ElseIf lcombo1.ListIndex = 3 Then 'All Level
   For i = 0 To List1.ListCount - 1
   If List1.Selected(i) Then
     CurrentTopic = CurrentTopic & List1.list(i)
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id=(select sub_id from sub where sub_nm='" & scombo1.Text & "' and c_id='" & lbl1.Caption & "')and tp_id=(select tp_id from topic where tp_nm='" & List1.list(i) & "' and sub_id=(select sub_id from sub where sub_nm='" & scombo1.Text & "' and c_id='" & lbl1.Caption & "')and c_id='" & lbl1.Caption & "') order by dbms_random.value)where rownum <=" & (FTOTQUESTION / List1.SelCount) + 1 & " ")
   End If
   Next i
   chkupdate
   c.Execute ("delete from mcqtest where q_no >" & FTOTQUESTION & " ")
   Set r1 = New ADODB.Recordset
   Set r1 = c1.Execute("select count(*) from mcqtest")
   If FTOTQUESTION > r1.Fields(0) Then
    MsgBox "Not Enough Questions Available in Question Bank for this Property,Try Other", vbCritical + vbOKOnly, ""
    Refr 'Function With Refresh
    Exit Sub
   End If
  Timer1.Enabled = False
  Timer2.Enabled = False
  Timer3.Enabled = False
  Unload stu_dash
  Unload Me
  Instruction_General.Show
 End If
ElseIf Val(Label10.Caption) = 2 Then 'Subject Wise Test
 selectedType = "Subject Wise Test"
 selectedlvl = lcombo2.Text
 IsFullLengthSelected = 0
 Set r = c.Execute("select * from SUBWISETEST where c_id='" & lbl1.Caption & "' ")
 If r.EOF = False Then
  FTOTQUESTION = r.Fields(1)
  FTOTTIMEMINUTE = r.Fields(2)
  FTOTTIMESECOND = r.Fields(3)
  FTOTMARKS = r.Fields(4)
  FPASSPERCENTG = r.Fields(5)
  FMRKPERCOR = r.Fields(6)
  FMRKPERWRONG = r.Fields(7)
 Else
  MsgBox "Oops, No SetUp For Subject Wise Test.. Contact Admin to Set Test Properties Of Subject Wise Test", vbCritical + vbOKOnly, "Topic Wise Test"
 Refr 'Function With Refresh
 Exit Sub
 End If
 If lcombo2.ListIndex <> 3 Then
  c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id=(select sub_id from sub where sub_nm='" & Scombo2.Text & "' and c_id='" & lbl1.Caption & "') and upper(q_dif_lvl)='" & UCase(lcombo2.Text) & "'order by dbms_random.value)where rownum <" & FTOTQUESTION + 1 & " ")
  chkupdate
  c.Execute ("delete from mcqtest where q_no >" & FTOTQUESTION & " ")
  Set r1 = New ADODB.Recordset
  Set r1 = c1.Execute("select count(*) from mcqtest")
  If FTOTQUESTION > r1.Fields(0) Then
   MsgBox "Not Enough Questions Available in Question Bank for this Property,Try Other", vbCritical + vbOKOnly, ""
   Refr 'Function With Refresh
   Exit Sub
  End If
  Timer1.Enabled = False
  Timer2.Enabled = False
  Timer3.Enabled = False
  Unload stu_dash
  Unload Me
  Instruction_General.Show
 ElseIf lcombo2.ListIndex = 3 Then 'All Level
  c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id=(select sub_id from sub where sub_nm='" & Scombo2.Text & "' and c_id='" & lbl1.Caption & "')order by dbms_random.value)where rownum <" & FTOTQUESTION + 1 & "  ")
  chkupdate
  c.Execute ("delete from mcqtest where q_no >" & FTOTQUESTION & " ")
   Set r1 = New ADODB.Recordset
   Set r1 = c.Execute("select count(*) from mcqtest")
   If FTOTQUESTION > r1.Fields(0) Then
    MsgBox "Not Enough Questions Available in Question Bank for this Property,Try Other", vbCritical + vbOKOnly, ""
    Refr 'Function With Refresh
    Exit Sub
   End If
  Timer1.Enabled = False
  Timer2.Enabled = False
  Timer3.Enabled = False
  Unload stu_dash
  Unload Me
  Instruction_General.Show
 End If
ElseIf Val(Label10.Caption) = 3 Then 'full length test
selectedType = "Full Length Test"
selectedlvl = lcombo3.Text
IsFullLengthSelected = 1
'calling Test Properties of Full Mock test
Set r = c.Execute("select * from FULLMOCKTEST where c_id='" & lbl1.Caption & "' ")
If r.EOF = False Then
 FTOTSUB = r.Fields(1)
 FTOTQUESTION = r.Fields(2)
 FTOTTIMEMINUTE = r.Fields(3)
 FTOTTIMESECOND = r.Fields(4)
 FTOTMARKS = r.Fields(5)
 FPASSPERCENTG = r.Fields(6)
 FMRKPERCOR = r.Fields(7)
 FMRKPERWRONG = r.Fields(8)
 If IsNull(r.Fields(9)) = False Then
  Fsub1 = r.Fields(9)
 End If
 If IsNull(r.Fields(10)) = False Then
  Fsub2 = r.Fields(10)
 End If
 If IsNull(r.Fields(11)) = False Then
  Fsub3 = r.Fields(11)
 End If
 If IsNull(r.Fields(12)) = False Then
  Fsub4 = r.Fields(12)
 End If
 If IsNull(r.Fields(13)) = False Then
  Fsub5 = r.Fields(13)
 End If
 FNOQ1 = r.Fields(14)
 FNOQ2 = r.Fields(15)
 FNOQ3 = r.Fields(16)
 FNOQ4 = r.Fields(17)
 FNOQ5 = r.Fields(18)
Else
 MsgBox "Oops, No SetUp For Full Mock Test.. Contact Admin to Set Test Properties Of Full length Test", vbCritical + vbOKOnly, "Topic Wise Test"
Refr 'Function With Refresh
Exit Sub
End If
'+++++++++++++++++++++++++++++++++++++++++ Main task
If lcombo3.ListIndex <> 3 Then 'Difficulty Wise
 If FTOTSUB = 1 Then 'If Only 1 Subject is Selected as Full Mock test
  If FNOQ1 <> 0 Then
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub1 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ1 + 1 & "  ")
     subName11 = Fsub1
     Question11 = 1
  ElseIf FNOQ2 <> 0 Then
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub2 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ2 + 1 & "  ")
     subName11 = Fsub2
     Question11 = 1
  ElseIf FNOQ3 <> 0 Then
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub3 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ3 + 1 & "  ")
     subName11 = Fsub3
     Question11 = 1
  ElseIf FNOQ4 <> 0 Then
      subName11 = Fsub4
      Question11 = 1
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub4 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ4 + 1 & "  ")
  ElseIf FNOQ5 <> 0 Then
      subName11 = Fsub5
      Question11 = 1
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub5 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ5 + 1 & "  ")
  End If
 ElseIf FTOTSUB = 2 Then 'If Only 2 Subject is Selected as Full Mock test
     If FNOQ1 <> 0 And FNOQ2 <> 0 Then
       subName21 = Fsub1
       Question21 = 1
       subName22 = Fsub2
       Question22 = FNOQ1 + 1
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub1 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ1 + 1 & "  ")
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub2 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ2 + 1 & "  ")
     ElseIf FNOQ1 <> 0 And FNOQ3 <> 0 Then
       subName21 = Fsub1
       Question21 = 1
       subName22 = Fsub3
       Question22 = FNOQ1 + 1
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub1 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ1 + 1 & "  ")
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub3 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ3 + 1 & "  ")
     ElseIf FNOQ1 <> 0 And FNOQ4 <> 0 Then
       subName21 = Fsub1
       Question21 = 1
       subName22 = Fsub4
       Question22 = FNOQ1 + 1
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub1 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ1 + 1 & "  ")
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub4 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ4 + 1 & "  ")
     ElseIf FNOQ1 <> 0 And FNOQ5 <> 0 Then
       subName21 = Fsub1
       Question21 = 1
       subName22 = Fsub5
       Question22 = FNOQ1 + 1
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub1 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ1 + 1 & "  ")
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub5 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ5 + 1 & "  ")
     ElseIf FNOQ2 <> 0 And FNOQ3 <> 0 Then
       subName21 = Fsub2
       Question21 = 1
       subName22 = Fsub3
       Question22 = FNOQ2 + 1
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub2 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ2 + 1 & "  ")
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub3 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ3 + 1 & "  ")
     ElseIf FNOQ2 <> 0 And FNOQ4 <> 0 Then
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub2 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ2 + 1 & "  ")
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub4 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ4 + 1 & "  ")
       subName21 = Fsub2
       Question21 = 1
       subName22 = Fsub4
       Question22 = FNOQ2 + 1
     ElseIf FNOQ2 <> 0 And FNOQ5 <> 0 Then
       subName21 = Fsub2
       Question21 = 1
       subName22 = Fsub5
       Question22 = FNOQ2 + 1
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub2 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ2 + 1 & "  ")
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub5 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ5 + 1 & "  ")
     ElseIf FNOQ3 <> 0 And FNOQ4 <> 0 Then
       subName21 = Fsub3
       Question21 = 1
       subName22 = Fsub4
       Question22 = FNOQ3 + 1
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub3 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ3 + 1 & "  ")
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub4 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ4 + 1 & "  ")
     ElseIf FNOQ3 <> 0 And FNOQ5 <> 0 Then
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub3 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ3 + 1 & "  ")
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub5 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ5 + 1 & "  ")
       subName21 = Fsub3
       Question21 = 1
       subName22 = Fsub5
       Question22 = FNOQ3 + 1
     ElseIf FNOQ4 <> 0 And FNOQ5 <> 0 Then
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub4 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ4 + 1 & "  ")
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub5 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ5 + 1 & "  ")
       subName21 = Fsub4
       Question21 = 1
       subName22 = Fsub5
       Question22 = FNOQ4 + 1
     End If
 ElseIf FTOTSUB = 3 Then 'If Only 3 Subject is Selected as Full Mock test
    If FNOQ1 <> 0 And FNOQ2 <> 0 And FNOQ3 <> 0 Then
      subName31 = Fsub1
      Question31 = 1
      subName32 = Fsub2
      Question32 = FNOQ1 + 1
      subName33 = Fsub3
      Question33 = FNOQ1 + FNOQ2 + 1
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub1 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ1 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub2 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ2 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub3 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ3 + 1 & "  ")
    ElseIf FNOQ1 <> 0 And FNOQ2 <> 0 And FNOQ4 <> 0 Then
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub1 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ1 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub2 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ2 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub4 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ4 + 1 & "  ")
      subName31 = Fsub1
      Question31 = 1
      subName32 = Fsub2
      Question32 = FNOQ1 + 1
      subName33 = Fsub4
      Question33 = FNOQ1 + FNOQ2 + 1
    ElseIf FNOQ1 <> 0 And FNOQ2 <> 0 And FNOQ5 <> 0 Then
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub1 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ1 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub2 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ2 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub5 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ5 + 1 & "  ")
      subName31 = Fsub1
      Question31 = 1
      subName32 = Fsub2
      Question32 = FNOQ1 + 1
      subName33 = Fsub5
      Question33 = FNOQ1 + FNOQ2 + 1
    ElseIf FNOQ1 <> 0 And FNOQ3 <> 0 And FNOQ4 <> 0 Then
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub1 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ1 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub3 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ3 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub4 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ4 + 1 & "  ")
      subName31 = Fsub1
      Question31 = 1
      subName32 = Fsub3
      Question32 = FNOQ1 + 1
      subName33 = Fsub4
      Question33 = FNOQ1 + FNOQ3 + 1
    ElseIf FNOQ1 <> 0 And FNOQ3 <> 0 And FNOQ5 <> 0 Then
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub1 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ1 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub3 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ3 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub5 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ5 + 1 & "  ")
      subName31 = Fsub1
      Question31 = 1
      subName32 = Fsub3
      Question32 = FNOQ1 + 1
      subName33 = Fsub5
      Question33 = FNOQ1 + FNOQ3 + 1
    ElseIf FNOQ1 <> 0 And FNOQ4 <> 0 And FNOQ5 <> 0 Then
      subName31 = Fsub1
      Question31 = 1
      subName32 = Fsub4
      Question32 = FNOQ1 + 1
      subName33 = Fsub5
      Question33 = FNOQ1 + FNOQ4 + 1
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub1 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ1 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub4 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ4 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub5 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ5 + 1 & "  ")
    ElseIf FNOQ2 <> 0 And FNOQ3 <> 0 And FNOQ4 <> 0 Then
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub2 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ2 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub3 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ3 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub4 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ4 + 1 & "  ")
      subName31 = Fsub2
      Question31 = 1
      subName32 = Fsub3
      Question32 = FNOQ2 + 1
      subName33 = Fsub4
      Question33 = FNOQ2 + FNOQ3 + 1
    ElseIf FNOQ2 <> 0 And FNOQ3 <> 0 And FNOQ5 <> 0 Then
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub2 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ2 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub3 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ3 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub5 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ5 + 1 & "  ")
      subName31 = Fsub2
      Question31 = 1
      subName32 = Fsub3
      Question32 = FNOQ2 + 1
      subName33 = Fsub5
      Question33 = FNOQ2 + FNOQ3 + 1
    ElseIf FNOQ2 <> 0 And FNOQ4 <> 0 And FNOQ5 <> 0 Then
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub2 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ2 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub4 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ4 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub5 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ5 + 1 & "  ")
      subName31 = Fsub2
      Question31 = 1
      subName32 = Fsub4
      Question32 = FNOQ2 + 1
      subName33 = Fsub5
      Question33 = FNOQ2 + FNOQ4 + 1
    ElseIf FNOQ3 <> 0 And FNOQ4 <> 0 And FNOQ5 <> 0 Then
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub3 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ3 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub4 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ4 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub5 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ5 + 1 & "  ")
      subName31 = Fsub3
      Question31 = 1
      subName32 = Fsub4
      Question32 = FNOQ3 + 1
      subName33 = Fsub5
      Question33 = FNOQ3 + FNOQ4 + 1
    End If
 ElseIf FTOTSUB = 4 Then 'If Only 4 Subject is Selected as Full Mock test
    If FNOQ1 = 0 Then
      subName41 = Fsub2
      Question41 = 1
      subName42 = Fsub3
      Question42 = FNOQ2 + 1
      subName43 = Fsub4
      Question43 = FNOQ2 + FNOQ3 + 1
      subName44 = Fsub5
      Question44 = FNOQ2 + FNOQ3 + FNOQ4 + 1
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub2 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ2 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub3 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ3 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub4 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ4 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub5 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ5 + 1 & "  ")
    ElseIf FNOQ2 = 0 Then
      subName41 = Fsub1
      Question41 = 1
      subName42 = Fsub3
      Question42 = FNOQ1 + 1
      subName43 = Fsub4
      Question43 = FNOQ1 + FNOQ3 + 1
      subName44 = Fsub5
      Question44 = FNOQ1 + FNOQ3 + FNOQ4 + 1
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub1 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ1 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub3 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ3 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub4 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ4 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub5 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ5 + 1 & "  ")
    ElseIf FNOQ3 = 0 Then
      subName41 = Fsub1
      Question41 = 1
      subName42 = Fsub2
      Question42 = FNOQ1 + 1
      subName43 = Fsub4
      Question43 = FNOQ2 + FNOQ1 + 1
      subName44 = Fsub5
      Question44 = FNOQ2 + FNOQ1 + FNOQ4 + 1
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub1 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ1 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub2 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ2 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub4 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ4 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub5 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ5 + 1 & "  ")
    ElseIf FNOQ4 = 0 Then
      subName41 = Fsub1
      Question41 = 1
      subName42 = Fsub2
      Question42 = FNOQ1 + 1
      subName43 = Fsub3
      Question43 = FNOQ2 + FNOQ1 + 1
      subName44 = Fsub5
      Question44 = FNOQ2 + FNOQ1 + FNOQ3 + 1
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub1 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ1 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub2 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ2 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub3 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ3 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub5 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ5 + 1 & "  ")
    ElseIf FNOQ5 = 0 Then
      subName41 = Fsub1
      Question41 = 1
      subName42 = Fsub2
      Question42 = FNOQ1 + 1
      subName43 = Fsub3
      Question43 = FNOQ1 + FNOQ2 + 1
      subName44 = Fsub4
      Question44 = FNOQ2 + FNOQ3 + FNOQ1 + 1
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub1 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ1 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub2 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ2 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub3 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ3 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub4 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ4 + 1 & "  ")
    End If
 ElseIf FTOTSUB = 5 Then 'If All 5 Subject is Selected as Full Mock test
      subName51 = Fsub1
      Question51 = 1
      subName52 = Fsub2
      Question52 = FNOQ1 + 1
      subName53 = Fsub3
      Question53 = FNOQ1 + FNOQ2 + 1
      subName54 = Fsub4
      Question54 = FNOQ2 + FNOQ3 + FNOQ1 + 1
      subName55 = Fsub5
      Question55 = FNOQ1 + FNOQ2 + FNOQ3 + FNOQ4 + 1
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub1 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ1 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub2 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ2 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub3 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ3 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub4 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ4 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub5 & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' order by dbms_random.value)where rownum <" & FNOQ5 + 1 & "  ")
 End If
 c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and upper(q_dif_lvl)='" & UCase(lcombo3.Text) & "' ")
  Set r1 = New ADODB.Recordset
  Set r1 = c.Execute("select count(*) from mcqtest")
  If FTOTQUESTION > r1.Fields(0) Then
   MsgBox "Not Enough Questions Available in Question Bank for this Property,Try Other", vbCritical + vbOKOnly, ""
  Refr 'Function With Refresh
  Exit Sub
  End If
  chkupdate
  Timer1.Enabled = False
  Timer2.Enabled = False
  Timer3.Enabled = False
   Unload stu_dash
   Unload Me
   Instruction_General.Show
 ElseIf lcombo3.ListIndex = 3 Then 'All Level
 If FTOTSUB = 1 Then 'If Only 1 Subject is Selected as Full Mock test
  If FNOQ1 <> 0 Then
     subName11 = Fsub1
     Question11 = 1
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub1 & "' order by dbms_random.value)where rownum <" & FNOQ1 + 1 & "  ")
  ElseIf FNOQ2 <> 0 Then
      subName11 = Fsub2
     Question11 = 1
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub2 & "' order by dbms_random.value)where rownum <" & FNOQ2 + 1 & "  ")
  ElseIf FNOQ3 <> 0 Then
    subName11 = Fsub3
     Question11 = 1
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub3 & "' order by dbms_random.value)where rownum <" & FNOQ3 + 1 & "  ")
  ElseIf FNOQ4 <> 0 Then
   subName11 = Fsub4
     Question11 = 1
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub4 & "' order by dbms_random.value)where rownum <" & FNOQ4 + 1 & "  ")
  ElseIf FNOQ5 <> 0 Then
   subName11 = Fsub5
     Question11 = 1
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub5 & "' order by dbms_random.value)where rownum <" & FNOQ5 + 1 & "  ")
  End If
 ElseIf FTOTSUB = 2 Then 'If Only 2 Subject is Selected as Full Mock test
 If FNOQ1 <> 0 And FNOQ2 <> 0 Then
       subName21 = Fsub1
       Question21 = 1
       subName22 = Fsub2
       Question22 = FNOQ1 + 1
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub1 & "' order by dbms_random.value)where rownum <" & FNOQ1 + 1 & "  ")
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub2 & "' order by dbms_random.value)where rownum <" & FNOQ2 + 1 & "  ")
     ElseIf FNOQ1 <> 0 And FNOQ3 <> 0 Then
       subName21 = Fsub1
       Question21 = 1
       subName22 = Fsub3
       Question22 = FNOQ1 + 1
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub1 & "' order by dbms_random.value)where rownum <" & FNOQ1 + 1 & "  ")
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub3 & "' order by dbms_random.value)where rownum <" & FNOQ3 + 1 & "  ")
     ElseIf FNOQ1 <> 0 And FNOQ4 <> 0 Then
       subName21 = Fsub1
       Question21 = 1
       subName22 = Fsub4
       Question22 = FNOQ1 + 1
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub1 & "' order by dbms_random.value)where rownum <" & FNOQ1 + 1 & "  ")
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub4 & "' order by dbms_random.value)where rownum <" & FNOQ4 + 1 & "  ")
     ElseIf FNOQ1 <> 0 And FNOQ5 <> 0 Then
       subName21 = Fsub1
       Question21 = 1
       subName22 = Fsub5
       Question22 = FNOQ1 + 1
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub1 & "' order by dbms_random.value)where rownum <" & FNOQ1 + 1 & "  ")
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub5 & "' order by dbms_random.value)where rownum <" & FNOQ5 + 1 & "  ")
     ElseIf FNOQ2 <> 0 And FNOQ3 <> 0 Then
       subName21 = Fsub2
       Question21 = 1
       subName22 = Fsub3
       Question22 = FNOQ2 + 1
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub2 & "' order by dbms_random.value)where rownum <" & FNOQ2 + 1 & "  ")
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub3 & "' order by dbms_random.value)where rownum <" & FNOQ3 + 1 & "  ")
     ElseIf FNOQ2 <> 0 And FNOQ4 <> 0 Then
       subName21 = Fsub2
       Question21 = 1
       subName22 = Fsub4
       Question22 = FNOQ2 + 1
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub2 & "' order by dbms_random.value)where rownum <" & FNOQ2 + 1 & "  ")
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub4 & "' order by dbms_random.value)where rownum <" & FNOQ4 + 1 & "  ")
     ElseIf FNOQ2 <> 0 And FNOQ5 <> 0 Then
       subName21 = Fsub2
       Question21 = 1
       subName22 = Fsub5
       Question22 = FNOQ2 + 1
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub2 & "' order by dbms_random.value)where rownum <" & FNOQ2 + 1 & "  ")
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub5 & "' order by dbms_random.value)where rownum <" & FNOQ5 + 1 & "  ")
     ElseIf FNOQ3 <> 0 And FNOQ4 <> 0 Then
       subName21 = Fsub3
       Question21 = 1
       subName22 = Fsub4
       Question22 = FNOQ3 + 1
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub3 & "' order by dbms_random.value)where rownum <" & FNOQ3 + 1 & "  ")
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub4 & "' order by dbms_random.value)where rownum <" & FNOQ4 + 1 & "  ")
     ElseIf FNOQ3 <> 0 And FNOQ5 <> 0 Then
       subName21 = Fsub3
       Question21 = 1
       subName22 = Fsub5
       Question22 = FNOQ3 + 1
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub3 & "' order by dbms_random.value)where rownum <" & FNOQ3 + 1 & "  ")
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub5 & "' order by dbms_random.value)where rownum <" & FNOQ5 + 1 & "  ")
     ElseIf FNOQ4 <> 0 And FNOQ5 <> 0 Then
       subName21 = Fsub4
       Question21 = 1
       subName22 = Fsub5
       Question22 = FNOQ4 + 1
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub4 & "' order by dbms_random.value)where rownum <" & FNOQ4 + 1 & "  ")
      c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub5 & "' order by dbms_random.value)where rownum <" & FNOQ5 + 1 & "  ")
     End If
 ElseIf FTOTSUB = 3 Then 'If Only 3 Subject is Selected as Full Mock test
    If FNOQ1 <> 0 And FNOQ2 <> 0 And FNOQ3 <> 0 Then
      subName31 = Fsub1
      Question31 = 1
      subName32 = Fsub2
      Question32 = FNOQ1 + 1
      subName33 = Fsub3
      Question33 = FNOQ1 + FNOQ2 + 1
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub1 & "' order by dbms_random.value)where rownum <" & FNOQ1 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub2 & "' order by dbms_random.value)where rownum <" & FNOQ2 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub3 & "' order by dbms_random.value)where rownum <" & FNOQ3 + 1 & "  ")
    ElseIf FNOQ1 <> 0 And FNOQ2 <> 0 And FNOQ4 <> 0 Then
      subName31 = Fsub1
      Question31 = 1
      subName32 = Fsub2
      Question32 = FNOQ1 + 1
      subName33 = Fsub4
      Question33 = FNOQ1 + FNOQ2 + 1
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub1 & "' order by dbms_random.value)where rownum <" & FNOQ1 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub2 & "' order by dbms_random.value)where rownum <" & FNOQ2 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub4 & "' order by dbms_random.value)where rownum <" & FNOQ4 + 1 & "  ")
    ElseIf FNOQ1 <> 0 And FNOQ2 <> 0 And FNOQ5 <> 0 Then
      subName31 = Fsub1
      Question31 = 1
      subName32 = Fsub2
      Question32 = FNOQ1 + 1
      subName33 = Fsub5
      Question33 = FNOQ1 + FNOQ2 + 1
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub1 & "' order by dbms_random.value)where rownum <" & FNOQ1 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub2 & "' order by dbms_random.value)where rownum <" & FNOQ2 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub5 & "' order by dbms_random.value)where rownum <" & FNOQ5 + 1 & "  ")
    ElseIf FNOQ1 <> 0 And FNOQ3 <> 0 And FNOQ4 <> 0 Then
      subName31 = Fsub1
      Question31 = 1
      subName32 = Fsub3
      Question32 = FNOQ1 + 1
      subName33 = Fsub4
      Question33 = FNOQ1 + FNOQ3 + 1
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub1 & "' order by dbms_random.value)where rownum <" & FNOQ1 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub3 & "' order by dbms_random.value)where rownum <" & FNOQ3 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub4 & "' order by dbms_random.value)where rownum <" & FNOQ4 + 1 & "  ")
    ElseIf FNOQ1 <> 0 And FNOQ3 <> 0 And FNOQ5 <> 0 Then
      subName31 = Fsub1
      Question31 = 1
      subName32 = Fsub3
      Question32 = FNOQ1 + 1
      subName33 = Fsub5
      Question33 = FNOQ1 + FNOQ3 + 1
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub1 & "' order by dbms_random.value)where rownum <" & FNOQ1 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub3 & "' order by dbms_random.value)where rownum <" & FNOQ3 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub5 & "' order by dbms_random.value)where rownum <" & FNOQ5 + 1 & "  ")
    ElseIf FNOQ1 <> 0 And FNOQ4 <> 0 And FNOQ5 <> 0 Then
      subName31 = Fsub1
      Question31 = 1
      subName32 = Fsub4
      Question32 = FNOQ1 + 1
      subName33 = Fsub5
      Question33 = FNOQ1 + FNOQ4 + 1
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub1 & "' order by dbms_random.value)where rownum <" & FNOQ1 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub4 & "' order by dbms_random.value)where rownum <" & FNOQ4 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub5 & "' order by dbms_random.value)where rownum <" & FNOQ5 + 1 & "  ")
    ElseIf FNOQ2 <> 0 And FNOQ3 <> 0 And FNOQ4 <> 0 Then
      subName31 = Fsub2
      Question31 = 1
      subName32 = Fsub3
      Question32 = FNOQ2 + 1
      subName33 = Fsub4
      Question33 = FNOQ2 + FNOQ3 + 1
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub2 & "' order by dbms_random.value)where rownum <" & FNOQ2 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub3 & "' order by dbms_random.value)where rownum <" & FNOQ3 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub4 & "' order by dbms_random.value)where rownum <" & FNOQ4 + 1 & "  ")
    ElseIf FNOQ2 <> 0 And FNOQ3 <> 0 And FNOQ5 <> 0 Then
    subName31 = Fsub2
      Question31 = 1
      subName32 = Fsub3
      Question32 = FNOQ2 + 1
      subName33 = Fsub5
      Question33 = FNOQ2 + FNOQ3 + 1
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub2 & "' order by dbms_random.value)where rownum <" & FNOQ2 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub3 & "' order by dbms_random.value)where rownum <" & FNOQ3 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub5 & "' order by dbms_random.value)where rownum <" & FNOQ5 + 1 & "  ")
    ElseIf FNOQ2 <> 0 And FNOQ4 <> 0 And FNOQ5 <> 0 Then
      subName31 = Fsub2
      Question31 = 1
      subName32 = Fsub4
      Question32 = FNOQ2 + 1
      subName33 = Fsub5
      Question33 = FNOQ2 + FNOQ4 + 1
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub2 & "' order by dbms_random.value)where rownum <" & FNOQ2 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub4 & "' order by dbms_random.value)where rownum <" & FNOQ4 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub5 & "' order by dbms_random.value)where rownum <" & FNOQ5 + 1 & "  ")
    ElseIf FNOQ3 <> 0 And FNOQ4 <> 0 And FNOQ5 <> 0 Then
      subName31 = Fsub3
      Question31 = 1
      subName32 = Fsub4
      Question32 = FNOQ3 + 1
      subName33 = Fsub5
      Question33 = FNOQ3 + FNOQ4 + 1
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub3 & "' order by dbms_random.value)where rownum <" & FNOQ3 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub4 & "' order by dbms_random.value)where rownum <" & FNOQ4 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub5 & "' order by dbms_random.value)where rownum <" & FNOQ5 + 1 & "  ")
    End If
 ElseIf FTOTSUB = 4 Then 'If Only 4 Subject is Selected as Full Mock test
    If FNOQ1 = 0 Then
      subName41 = Fsub2
      Question41 = 1
      subName42 = Fsub3
      Question42 = FNOQ2 + 1
      subName43 = Fsub4
      Question43 = FNOQ2 + FNOQ3 + 1
      subName44 = Fsub5
      Question44 = FNOQ2 + FNOQ3 + FNOQ4 + 1
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub2 & "' order by dbms_random.value)where rownum <" & FNOQ2 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub3 & "' order by dbms_random.value)where rownum <" & FNOQ3 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub4 & "' order by dbms_random.value)where rownum <" & FNOQ4 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub5 & "' order by dbms_random.value)where rownum <" & FNOQ5 + 1 & "  ")
    ElseIf FNOQ2 = 0 Then
      subName41 = Fsub1
      Question41 = 1
      subName42 = Fsub3
      Question42 = FNOQ1 + 1
      subName43 = Fsub4
      Question43 = FNOQ1 + FNOQ3 + 1
      subName44 = Fsub5
      Question44 = FNOQ1 + FNOQ3 + FNOQ4 + 1
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub1 & "' order by dbms_random.value)where rownum <" & FNOQ1 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub3 & "' order by dbms_random.value)where rownum <" & FNOQ3 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub4 & "' order by dbms_random.value)where rownum <" & FNOQ4 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub5 & "' order by dbms_random.value)where rownum <" & FNOQ5 + 1 & "  ")
    ElseIf FNOQ3 = 0 Then
      subName41 = Fsub1
      Question41 = 1
      subName42 = Fsub2
      Question42 = FNOQ1 + 1
      subName43 = Fsub4
      Question43 = FNOQ2 + FNOQ1 + 1
      subName44 = Fsub5
      Question44 = FNOQ2 + FNOQ1 + FNOQ4 + 1
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub1 & "' order by dbms_random.value)where rownum <" & FNOQ1 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub2 & "' order by dbms_random.value)where rownum <" & FNOQ2 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub4 & "' order by dbms_random.value)where rownum <" & FNOQ4 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub5 & "' order by dbms_random.value)where rownum <" & FNOQ5 + 1 & "  ")
    ElseIf FNOQ4 = 0 Then
      subName41 = Fsub1
      Question41 = 1
      subName42 = Fsub2
      Question42 = FNOQ1 + 1
      subName43 = Fsub3
      Question43 = FNOQ2 + FNOQ1 + 1
      subName44 = Fsub5
      Question44 = FNOQ2 + FNOQ1 + FNOQ3 + 1
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub1 & "' order by dbms_random.value)where rownum <" & FNOQ1 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub2 & "' order by dbms_random.value)where rownum <" & FNOQ2 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub3 & "' order by dbms_random.value)where rownum <" & FNOQ3 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub5 & "' order by dbms_random.value)where rownum <" & FNOQ5 + 1 & "  ")
    ElseIf FNOQ5 = 0 Then
      subName41 = Fsub1
      Question41 = 1
      subName42 = Fsub2
      Question42 = FNOQ1 + 1
      subName43 = Fsub3
      Question43 = FNOQ1 + FNOQ2 + 1
      subName44 = Fsub4
      Question44 = FNOQ2 + FNOQ3 + FNOQ1 + 1
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub1 & "' order by dbms_random.value)where rownum <" & FNOQ1 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub2 & "' order by dbms_random.value)where rownum <" & FNOQ2 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub3 & "' order by dbms_random.value)where rownum <" & FNOQ3 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub4 & "' order by dbms_random.value)where rownum <" & FNOQ4 + 1 & "  ")
    End If
 ElseIf FTOTSUB = 5 Then 'If All 5 Subject is Selected as Full Mock test
      subName51 = Fsub1
      Question51 = 1
      subName52 = Fsub2
      Question52 = FNOQ1 + 1
      subName53 = Fsub3
      Question53 = FNOQ1 + FNOQ2 + 1
      subName54 = Fsub4
      Question54 = FNOQ2 + FNOQ3 + FNOQ1 + 1
      subName55 = Fsub5
      Question55 = FNOQ1 + FNOQ2 + FNOQ3 + FNOQ4 + 1
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub1 & "' order by dbms_random.value)where rownum <" & FNOQ1 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub2 & "' order by dbms_random.value)where rownum <" & FNOQ2 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub3 & "' order by dbms_random.value)where rownum <" & FNOQ3 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub4 & "' order by dbms_random.value)where rownum <" & FNOQ4 + 1 & "  ")
     c.Execute ("insert into MCQTEST select * from ( select Q_NO,Q_TXT,opt1,opt2,opt3,opt4,ANS_TXT,ANS_NO,Q_PIC from quesms where c_id='" & lbl1.Caption & "' and sub_id='" & Fsub5 & "' order by dbms_random.value)where rownum <" & FNOQ5 + 1 & "  ")
 End If
   chkupdate
   Set r1 = New ADODB.Recordset
   Set r1 = c.Execute("select count(*) from mcqtest")
   If FTOTQUESTION > r1.Fields(0) Then
    MsgBox "Not Enough Questions Available in Question Bank for this Property,Try Other", vbCritical + vbOKOnly, ""
    Refr 'Function With Refresh
    Exit Sub
   End If
   Timer1.Enabled = False
   Timer2.Enabled = False
   Timer3.Enabled = False
   Unload stu_dash
   Unload Me
   Instruction_General.Show
 End If
End If
End Sub

Public Function chkupdate()
Set r = New ADODB.Recordset
Set r = c1.Execute("select count(*)from user_sequences where sequence_name='QNOGENERATOR1' ")
If r.Fields(0) <= 0 Or IsNull(r.Fields(0)) = False Then
 c.Execute ("drop sequence QNOGENERATOR1") 'Deleting old Sequence if exist
 c.Execute ("create sequence QNOGENERATOR1 increment by 1 start with 1 cache 100") 'Creating New Sequence start with 1
 c1.Execute ("update MCQTEST set Q_NO=QNOGENERATOR1.Nextval")
Else
 c.Execute ("create sequence QNOGENERATOR1 increment by 1 start with 1 cache 100") 'Creating New Sequence start with 1
 c1.Execute ("update MCQTEST set Q_NO=QNOGENERATOR1.Nextval")
End If
End Function
Private Sub ChameleonBtn1_Click() 'Back button
Unload Me
stu_dash.Enabled = True
End Sub
Public Function Refr()
c.Execute ("delete from MCQTEST")
c.Execute ("delete from answerhold")
IsFullLengthSelected = 0
btnnext.Enabled = False
Label10.Caption = ""
TPFrame.Visible = False
SubFrame.Visible = False
FulFrame.Visible = False
'Clearing The Combo Record
scombo1.Clear
Scombo2.Clear

lcombo1.Clear
lcombo2.Clear
lcombo3.Clear
'++++++++++++++++++++++++++
lcombo1.AddItem "EASY"
lcombo1.AddItem "MEDIUM"
lcombo1.AddItem "HARD"
lcombo1.AddItem "Mix (All)"

lcombo2.AddItem "EASY"
lcombo2.AddItem "MEDIUM"
lcombo2.AddItem "HARD"
lcombo2.AddItem "Mix (All)"

lcombo3.AddItem "EASY"
lcombo3.AddItem "MEDIUM"
lcombo3.AddItem "HARD"
lcombo3.AddItem "Mix (All)"

Set r = New ADODB.Recordset
Set r = c.Execute("select c_id,pkg_id,sch_id from rstud where rstud_reg_no='" & Stu_login_reg_no & "' ")
If r.EOF = False Then
 lbl1.Caption = r.Fields(0)
 If IsNull(r.Fields(1)) = False Then
  lbl2.Caption = r.Fields(1)
 End If
 lbl3.Caption = r.Fields(2)
 End If
Set r = New ADODB.Recordset
Set r = c.Execute("select sub_nm from sub where c_id='" & lbl1.Caption & "'")
While r.EOF = False
 Scombo2.AddItem r.Fields(0)
 scombo1.AddItem r.Fields(0)
 r.MoveNext
Wend
End Function
Private Sub Form_Load()
On Error Resume Next
conn
Me.Left = 2300
Me.Top = 800
Timer1.Interval = 30
Timer2.Interval = 50
Timer3.Interval = 50
Stu_login_reg_no = Current_Logged_ID
Refr 'Function With Refresh
End Sub

Private Sub FullFrame_Click()
Label10.Caption = 3
TPFrame.Visible = False
SubFrame.Visible = False
FulFrame.Visible = True
End Sub

Private Sub scombo1_Click()
List1.Clear
CurrentSub = scombo1.Text '4444444444444444
Set r1 = New ADODB.Recordset
Set r1 = c1.Execute("select tp_nm from topic where sub_id =(select sub_id from sub where sub_nm='" & scombo1.Text & "' and c_id='" & lbl1.Caption & "')and c_id='" & lbl1.Caption & "' ")
While r1.EOF = False
 List1.AddItem r1.Fields(0)
 r1.MoveNext
Wend
End Sub

Private Sub Scombo2_Click()
CurrentSub = Scombo2.Text
End Sub

Private Sub SubjectFrame_Click()
Label10.Caption = 2
TPFrame.Visible = False
SubFrame.Visible = True
FulFrame.Visible = False
End Sub

Private Sub Timer1_Timer()
 If Label10.Caption = "1" And scombo1.Text <> "" And List1.SelCount <> 0 And lcombo1.Text <> "" Then
  btnnext.Enabled = True
  ElseIf Label10.Caption = "2" And Scombo2.Text <> "" And lcombo2.Text <> "" Then
    btnnext.Enabled = True
  ElseIf Label10.Caption = "3" And lcombo3.Text <> "" Then
    btnnext.Enabled = True
  Else
     btnnext.Enabled = False
  End If
End Sub

Private Sub TopicFrame_Click()
Label10.Caption = 1
TPFrame.Visible = True
SubFrame.Visible = False
FulFrame.Visible = False
End Sub

Private Sub Timer2_Timer()
Label9.Left = Label9.Left - 50 '
If Label9.Left < 0 Then ' If label1's left position is smaller than 0 then
Label9.Left = Label9.Left + 30
Timer3.Enabled = True 'Timer3 = true
Timer2.Enabled = False 'Timer2 = false
End If
End Sub

Private Sub Timer3_Timer()
Label9.Left = Label9.Left + 50
If Label9.Left > Me.Width - Label9.Width Then ' If label1's left position is bigger than label1 width
Label9.Left = Label9.Left - 30
Timer2.Enabled = True 'Timer2 = True
Timer3.Enabled = False 'Timer3 = False
End If
End Sub
