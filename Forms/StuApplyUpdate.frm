VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form rstud_Pkg_renew 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Package Renew Application"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6420
   Icon            =   "StuApplyUpdate.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   6420
   Begin MSACAL.Calendar cld1 
      Height          =   2895
      Left            =   2400
      TabIndex        =   18
      Top             =   2880
      Width           =   3135
      _Version        =   524288
      _ExtentX        =   5530
      _ExtentY        =   5106
      _StockProps     =   1
      BackColor       =   14737632
      Year            =   2019
      Month           =   5
      Day             =   9
      DayLength       =   1
      MonthLength     =   0
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   0   'False
      ShowTitle       =   0   'False
      ShowVerticalGrid=   0   'False
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox vkCheck1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "I Confirm to apply for renewing the package."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   240
      MouseIcon       =   "StuApplyUpdate.frx":08CA
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   5640
      Width           =   4455
   End
   Begin VB.TextBox lbl2 
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
      IMEMode         =   3  'DISABLE
      Left            =   240
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   9
      Top             =   2925
      Width           =   2175
   End
   Begin VB.TextBox lbl3 
      Enabled         =   0   'False
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
      Left            =   240
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   8
      Top             =   3765
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      MouseIcon       =   "StuApplyUpdate.frx":0A1C
      MousePointer    =   99  'Custom
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   980
      Width           =   2295
   End
   Begin VB.TextBox lbl1 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1950
      Width           =   2295
   End
   Begin VB.TextBox lbl5 
      Enabled         =   0   'False
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
      Left            =   4200
      MaxLength       =   15
      TabIndex        =   5
      Top             =   990
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6495
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Mukesh Kumar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   285
         TabIndex        =   4
         Top             =   135
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   6165
      Width           =   6495
      Begin VB.CommandButton btnapply 
         BackColor       =   &H00C0C0FF&
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
         Left            =   120
         MouseIcon       =   "StuApplyUpdate.frx":0B6E
         MousePointer    =   99  'Custom
         Picture         =   "StuApplyUpdate.frx":0CC0
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Exit From Here"
         Top             =   150
         Width           =   1575
      End
      Begin VB.CommandButton btnCancel 
         BackColor       =   &H00C0C0FF&
         DisabledPicture =   "StuApplyUpdate.frx":161F
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   5160
         MouseIcon       =   "StuApplyUpdate.frx":1C79
         MousePointer    =   99  'Custom
         Picture         =   "StuApplyUpdate.frx":1DCB
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Exit From Here"
         Top             =   150
         Width           =   1095
      End
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Date :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   16
      Top             =   630
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Price :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Test :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Package :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Expiary Date :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Date :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label lbl4 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "200"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   10
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Image Image2 
      Height          =   195
      Left            =   360
      Picture         =   "StuApplyUpdate.frx":2425
      Stretch         =   -1  'True
      Top             =   4660
      Width           =   150
   End
End
Attribute VB_Name = "rstud_Pkg_renew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tempPKG As String
Dim tempDURATION As Integer
Private Sub btnapply_Click()
If Trim(Combo1.Text) = "" Then
MsgBox "Select Any package ", vbInformation + vbOKOnly, "Select Package"
Exit Sub
ElseIf vkCheck1.Value = vbUnchecked Then
 MsgBox "Check The Confirmation info", vbInformation + vbOKOnly, "Unselected Confirmation"
 Exit Sub
 Else
  Set r1 = c.Execute("select count(*) from PKG_RENEW")
  c.Execute ("insert  into PKG_RENEW values(" & r1.Fields(0) + 1 & ",'" & Format(Date$, "dd-mmm-yyyy") & "','" & tempPKG & "','" & Format(lbl2.Text, "dd-mmm-yyyy") & "','" & Format(lbl3.Text, "dd-mmm-yyyy") & "'," & Val(lbl1.Text) & "," & Val(lbl4.Caption) & ",'" & Stu_login_reg_no & "' )")
  MsgBox "   Request SuccessFully Submmitted" & vbCrLf & " Wait Until Admin Activate Your Package ", vbInformation + vbOKOnly, "Request Submitted"
  Unload Me
End If
End Sub

Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub cld1_Click()
If cld1.Value < Date Then
 MsgBox "You cannot Select the date that " & vbCrLf & "  is already passed away." & vbCrLf & "  Please select valid date", vbExclamation + vbOKOnly, "Invalid date"
 lbl2.Text = ""
 lbl3.Text = ""
 lbl2.SetFocus
Else
 lbl3.Text = ""
 lbl2.Text = cld1.Day & "-" & cld1.Month & "-" & cld1.Year
 lbl3.Text = Format(cld1.Value + tempDURATION, "dd-mm-yyyy")
 cld1.Visible = False
 End If
End Sub

Private Sub Combo1_Click()
Set r = New ADODB.Recordset
Set r = c1.Execute("select * from pkg where upper(pkg_nm) ='" & UCase(Combo1.Text) & "' ")
If r.EOF = False Then
 tempPKG = r.Fields(0)
 lbl1.Text = r.Fields("PKG_ALL_TST")
 tempDURATION = r.Fields(3)
 lbl2.Text = Format(Date$, "dd-mm-yyyy")
 lbl3.Text = Format(Date + tempDURATION, "dd-mm-yyyy")
 lbl4.Caption = r.Fields(2)
End If
End Sub

Private Sub Form_Load()
conn
cld1.Value = Format(Date, "DD-MMM-YY")
CenterForm Me
lbl1.Text = ""
lbl2.Text = ""
lbl3.Text = ""
lbl4.Caption = ""
vkCheck1.Value = vbUnchecked
btnapply.Enabled = False
lbl5.Text = Date$
 Combo1.Clear
 Set r = New ADODB.Recordset
 Set r = c.Execute("select pkg_nm from pkg where c_id=(select c_id from rstud where rstud_reg_no='" & Stu_login_reg_no & "') ")
  While r.EOF = False
   Combo1.AddItem r.Fields(0)
   r.MoveNext
  Wend
 cld1.Visible = False
 tempDURATION = 0
End Sub

Private Sub Form_Unload(cancel As Integer)
rstud_pkg.Enabled = True
End Sub

Private Sub lbl2_GotFocus()
lbl2.Text = ""
cld1.Visible = True
End Sub

Private Sub lbl2_LostFocus()
 cld1.Visible = False
End Sub

Private Sub vkCheck1_Click()
If vkCheck1.Value = vbChecked Then
btnapply.Enabled = True
ElseIf vkCheck1.Value = vbUnchecked Then
btnapply.Enabled = False
End If
End Sub
