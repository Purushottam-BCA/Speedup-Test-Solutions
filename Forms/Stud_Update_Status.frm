VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form StuPendingReq 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "Student Pending Request"
   ClientHeight    =   10680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20250
   Icon            =   "Stud_Update_Status.frx":0000
   LinkTopic       =   "Form8"
   MDIChild        =   -1  'True
   ScaleHeight     =   10680
   ScaleWidth      =   20250
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00DCDCDC&
      BorderStyle     =   0  'None
      Height          =   9150
      Left            =   16250
      ScaleHeight     =   9150
      ScaleWidth      =   4200
      TabIndex        =   4
      Top             =   650
      Width           =   4200
      Begin VB.Frame Frame1 
         BackColor       =   &H00DCDCDC&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   460
         Left            =   1800
         TabIndex        =   18
         Top             =   8635
         Width           =   450
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000C000&
         DisabledPicture =   "Stud_Update_Status.frx":0ECA
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         MouseIcon       =   "Stud_Update_Status.frx":1791
         MousePointer    =   99  'Custom
         Picture         =   "Stud_Update_Status.frx":18E3
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   8635
         Width           =   2070
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0C0&
         DisabledPicture =   "Stud_Update_Status.frx":21AA
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         MouseIcon       =   "Stud_Update_Status.frx":284F
         MousePointer    =   99  'Custom
         Picture         =   "Stud_Update_Status.frx":29A1
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   8635
         Width           =   2015
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact No :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   405
         TabIndex        =   17
         Top             =   6600
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Course :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   810
         TabIndex        =   16
         Top             =   6000
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Join date :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   615
         TabIndex        =   15
         Top             =   5400
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Father Name :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   14
         Top             =   4800
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Name :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   930
         TabIndex        =   13
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label l5 
         BackStyle       =   0  'Transparent
         Caption         =   "Not Available"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1890
         TabIndex        =   12
         Top             =   6600
         Width           =   2055
      End
      Begin VB.Label l4 
         BackStyle       =   0  'Transparent
         Caption         =   "Not Available"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1890
         TabIndex        =   11
         Top             =   6000
         Width           =   2055
      End
      Begin VB.Label l3 
         BackStyle       =   0  'Transparent
         Caption         =   "Not Available"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1890
         TabIndex        =   10
         Top             =   5400
         Width           =   2055
      End
      Begin VB.Label l2 
         BackStyle       =   0  'Transparent
         Caption         =   "Not Available"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1890
         TabIndex        =   9
         Top             =   4800
         Width           =   2055
      End
      Begin VB.Label l1 
         BackStyle       =   0  'Transparent
         Caption         =   "Not Availble"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1890
         TabIndex        =   8
         Top             =   4200
         Width           =   2055
      End
      Begin VB.Image Image1 
         Height          =   3120
         Left            =   750
         Picture         =   "Stud_Update_Status.frx":3046
         Stretch         =   -1  'True
         Top             =   480
         Width           =   2400
      End
      Begin VB.Label lbl1 
         Height          =   375
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   16320
      MouseIcon       =   "Stud_Update_Status.frx":49B3
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9900
      Width           =   4010
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   9300
      Left            =   -1035
      TabIndex        =   1
      ToolTipText     =   "Click To see details."
      Top             =   600
      Width           =   17250
      _ExtentX        =   30427
      _ExtentY        =   16404
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14737632
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "Stud_Update_Status.frx":4B05
      NumItems        =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   16220
      X2              =   16220
      Y1              =   0
      Y2              =   600
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Student Info"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   390
      Left            =   17310
      TabIndex        =   19
      Top             =   120
      Width           =   1635
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   $"Stud_Update_Status.frx":4C67
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
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   10035
      Width           =   15015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Student  Pending  Request"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   16095
   End
End
Attribute VB_Name = "StuPendingReq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim img1 As String
Dim temfee As Double
Private Sub Command1_Click()
If Trim(lbl1.Caption) <> "" Then
 If MsgBox("Are You Sure To Confirm ???", vbQuestion + vbYesNo, "Confirmation") = vbYes Then
 Set r1 = New ADODB.Recordset
 Set r1 = c1.Execute("select strt_dt, expr_dt,tot_test, tot_fee,pkg_id  from PKG_RENEW where rstud_reg_no='" & lbl1.Caption & "' ")
 temfee = r1.Fields(3)
 c.Execute ("update rstud  set RSTUD_DOJ='" & r1.Fields(0) & "',RSTUD_DOE='" & r1.Fields(1) & "',RSTUD_TOT_TEST=" & r1.Fields(2) & ",RSTUD_ALL_TEST=" & r1.Fields(2) & ", RSTUD_AMNT=" & r1.Fields(3) & ",PKG_ID='" & r1.Fields(4) & "' where rstud_reg_no='" & lbl1.Caption & "'  ")
 c.Execute ("delete from pkg_renew where RSTUD_REG_NO = '" & lbl1.Caption & "' ")
 MsgBox "Package SuccessFully Updated", vbInformation + vbOKOnly, "Update success"
 'Inserting into Account
Dim statement As String
Set r = New ADODB.Recordset
statement = l1.Caption & "'s package has been updated ."
Set r = c.Execute("select count(*) from incm")
c.Execute ("insert into incm values (" & r.Fields(0) + 1 & ",'" & l1.Caption & "','" & statement & "'," & temfee & ",'" & Format(Date, "dd-mmm-yyyy") & "' )")
 admin_dash.PkgReq.Caption = Val(admin_dash.PkgReq.Caption) - 1
 lbl1.Caption = ""
loaddata1
If lv1.ListItems.Count = 0 Then
 Command1.Enabled = False
 Else
 Command1.Enabled = True
 End If
l1.Caption = "Not Available"
l2.Caption = "Not Available"
l3.Caption = "Not Available"
l4.Caption = "Not Available"
l5.Caption = "Not Available"
Image1.Picture = LoadPicture(App.Path & "\Graphics\Main_Screen_Icon\PicNotAvail.jpg")
 Else
 End If
Else
MsgBox "Please Select A Record from List ", vbInformation + vbOKOnly, "Invalid Record"
End If
End Sub

Private Sub Command2_Click() 'Clear Button
'lv1.ListItems.Remove (lv1.SelectedItem.Index) 'For Delete from listview
img1 = App.Path & "\Graphics\Main_Screen_Icon\PicNotAvail.jpg"
lbl1.Caption = ""
loaddata1
If lv1.ListItems.Count = 0 Then
 Command1.Enabled = False
 Else
 Command1.Enabled = True
 End If
l1.Caption = "Not Available"
l2.Caption = "Not Available"
l3.Caption = "Not Available"
l4.Caption = "Not Available"
l5.Caption = "Not Available"
Image1.Picture = LoadPicture(App.Path & "\Graphics\Main_Screen_Icon\PicNotAvail.jpg")
End Sub

Private Sub Command3_Click()
Unload Me
End Sub


Private Sub Form_Load()
conn
Me.Width = MDI.Width
Me.Height = MDI.Height
Me.Top = 0
Me.Left = 0
img1 = App.Path & "\Graphics\Main_Screen_Icon\PicNotAvail.jpg"
lbl1.Caption = ""
With lv1.ColumnHeaders
.Clear
.add , "", "REGNo", Width / 20, lvwColumnLeft
.add , "", "No.", Width / 30, lvwColumnCenter
.add , "", "Registration No", Width / 11, lvwColumnCenter
.add , "", "Student Name ", Width / 5.7, lvwColumnCenter
.add , "", "Request date ", Width / 11, lvwColumnCenter
.add , "", "Package ", Width / 10.6, lvwColumnCenter
.add , "", "Start From ", Width / 11.5, lvwColumnCenter
.add , "", "Expire On ", Width / 11.5, lvwColumnCenter
.add , "", "Total test ", Width / 14, lvwColumnCenter
.add , "", "Fee", Width / 17.3, lvwColumnCenter
End With
loaddata1
If lv1.ListItems.Count = 0 Then
 Command1.Enabled = False
 Else
 Command1.Enabled = True
 End If
l1.Caption = "Not Available"
l2.Caption = "Not Available"
l3.Caption = "Not Available"
l4.Caption = "Not Available"
l5.Caption = "Not Available"
Image1.Picture = LoadPicture(App.Path & "\Graphics\Main_Screen_Icon\PicNotAvail.jpg")
End Sub

 Sub loaddata1()
 Dim r As New ADODB.Recordset
 Dim list As ListItem
 lv1.ListItems.Clear
 Set r = c1.Execute("select P.rstud_reg_no,P.sno,P.rstud_reg_no, r.rstud_nm, P.req_dt, pk.pkg_nm , p.strt_dt,p.expr_dt,p.tot_test, p.tot_fee  from PKG_RENEW P, rstud r, pkg pk where p.rstud_reg_no=r.rstud_reg_no and p.pkg_id=pk.pkg_id order by sno")
 While r.EOF = False
  Set list = lv1.ListItems.add(, , r.Fields(0))
  list.SubItems(1) = r.Fields(1)
  list.SubItems(2) = r.Fields(2)
  list.SubItems(3) = r.Fields(3)
  list.SubItems(4) = r.Fields(4)
  list.SubItems(5) = r.Fields(5)
  list.SubItems(6) = r.Fields(6)
  list.SubItems(7) = r.Fields(7)
  list.SubItems(8) = r.Fields(8)
list.SubItems(9) = r.Fields(9)
  r.MoveNext
  Wend
 End Sub

Private Sub lv1_Click()
On Error Resume Next
lbl1.Caption = lv1.SelectedItem
Set r = New ADODB.Recordset
Set r = c.Execute("select Rn.rstud_pic, Rn.rstud_nm,Rn.RSTUD_FATHER_NM,Rn.rstud_doj,C.c_nm,Rn.RSTUD_MOB from rstud Rn, Course C where Rn.rstud_reg_no='" & lv1.SelectedItem & "' and Rn.c_id=C.c_id")
img1 = r.Fields(0)
Image1.Picture = LoadPicture(img1)
l1.Caption = r.Fields(1)
l2.Caption = r.Fields(2)
l3.Caption = r.Fields(3)
l4.Caption = r.Fields(4)
l5.Caption = r.Fields(5)
End Sub
