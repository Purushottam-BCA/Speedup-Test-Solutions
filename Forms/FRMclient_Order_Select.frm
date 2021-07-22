VERSION 5.00
Begin VB.Form FrmClient4 
   BorderStyle     =   0  'None
   Caption         =   "Form6"
   ClientHeight    =   3585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6705
   LinkTopic       =   "Form6"
   ScaleHeight     =   3585
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmb2 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2760
      MouseIcon       =   "FRMclient_Order_Select.frx":0000
      MousePointer    =   99  'Custom
      Style           =   2  'Dropdown List
      TabIndex        =   4
      ToolTipText     =   "Select Client Name"
      Top             =   840
      Width           =   3015
   End
   Begin VB.ComboBox cmb1 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2760
      MouseIcon       =   "FRMclient_Order_Select.frx":0152
      MousePointer    =   99  'Custom
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "Select Order No"
      Top             =   1560
      Width           =   3015
   End
   Begin VB.CommandButton btn1 
      BackColor       =   &H80000016&
      Caption         =   "Confirm"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      MouseIcon       =   "FRMclient_Order_Select.frx":02A4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton btn2 
      BackColor       =   &H80000016&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      MouseIcon       =   "FRMclient_Order_Select.frx":03F6
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd-MMM-yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2760
      TabIndex        =   13
      Top             =   2230
      Width           =   1455
   End
   Begin VB.Label Label1 
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
      Left            =   2520
      TabIndex        =   12
      Top             =   2160
      Width           =   105
   End
   Begin VB.Label Label25 
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
      Left            =   2520
      TabIndex        =   11
      Top             =   1560
      Width           =   105
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client Name"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      TabIndex        =   10
      Top             =   840
      Width           =   1200
   End
   Begin VB.Label Label23 
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
      Left            =   2520
      TabIndex        =   9
      Top             =   840
      Width           =   105
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "Order No."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      TabIndex        =   8
      Top             =   1560
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order Date "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      TabIndex        =   7
      Top             =   2280
      Width           =   1140
   End
   Begin VB.Label info1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "* Select Order No"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4440
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label info2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "* Select Client Name "
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4320
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Image picx 
      Height          =   480
      Left            =   6240
      MouseIcon       =   "FRMclient_Order_Select.frx":0548
      MousePointer    =   99  'Custom
      Picture         =   "FRMclient_Order_Select.frx":069A
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client Order Information"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   2370
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C85A21&
      FillColor       =   &H00C85A21&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   6700
   End
End
Attribute VB_Name = "FrmClient4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Btn1_Click() 'Confirm Button
If cmb1.Text = "" Then
 info2.Visible = False
 info1.Visible = True
ElseIf cmb2.Text = "" Then
 info2.Visible = True
 info1.Visible = False
Else
ClintOrdrDate = Format(Label3.Caption, "dd-mm-yyyy") 'For Dashboard
ordrdt = ClintOrdrDate 'For Dashboard
Set r = New ADODB.Recordset
Set r = c.Execute("select * from clnt_ordr_chln where ord_no='" & cmb1.Text & "' ")
If r.EOF = False Then
 QpaprSetup.txt1.Text = r.Fields(3)
 QpaprSetup.txt2.Text = r.Fields(4)
 QpaprSetup.txt4.Text = r.Fields(5)
 QpaprSetup.txt5.Text = r.Fields(6)
 QpaprSetup.Text4.Text = r.Fields(7)
 QpaprSetup.txt6.Text = r.Fields(8)
 QpaprSetup.txt7.Text = r.Fields(9)
 QpaprSetup.txt8.Text = r.Fields(10)
 QpaprSetup.txt3.Text = r.Fields(11)
 QpaprSetup.txt9.Text = r.Fields(12)
 QpaprSetup.chk1.Value = vbChecked
 QpaprSetup.Text5.Text = r.Fields(13)
 QpaprSetup.Label27.Caption = r.Fields(15)
 QpaprSetup.Image1.Picture = LoadPicture(r.Fields(15))
 QpaprSetup.Fram2.Enabled = True
 Unload Me
 End If
End If
End Sub

Private Sub btn2_Click()
Unload Me
QpaprSetup.opt1.Value = False
QpaprSetup.opt2.Value = False
 QpaprSetup.Text1.Text = ""
 QpaprSetup.txt1.Text = ""
 QpaprSetup.txt2.Text = ""
 QpaprSetup.txt4.Text = ""
End Sub

Private Sub cmb1_Click()
Set r1 = New ADODB.Recordset
Set r1 = c.Execute("select ORD_DATE from clnt_ordr_chln where ord_no='" & cmb1.Text & "' ")
If r1.EOF = False Then
Label3.Caption = r1.Fields(0)
End If

End Sub

Private Sub cmb2_Click()
cmb1.Clear
 Set r = New ADODB.Recordset
 Set r = c.Execute("select ord_no from clnt_ordr_chln where CLNT_ID=(select CLNT_ID from client where upper(clnt_nm)='" & UCase(cmb2.Text) & "') and upper(CSTATUS)='COMPLETED' ")
 While r.EOF = False
  cmb1.AddItem r.Fields(0)
 r.MoveNext
 Wend
End Sub

Private Sub Form_Load()
CreateRoundRectFromWindow Me
Me.Top = 2500
Me.Left = 4180
info2.Visible = True
info1.Visible = False
conn
cmb1.Clear
cmb2.Clear
Set r = New ADODB.Recordset
Set r = c.Execute("Select distinct(initcap(CLNT_NM)) from client")
While r.EOF = False
 cmb2.AddItem r.Fields(0)
r.MoveNext
Wend
End Sub
Private Sub picX_Click()
Unload Me
End Sub
