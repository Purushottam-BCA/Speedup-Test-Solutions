VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student ID Card"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6000
   Icon            =   "Stu_ID_Card.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6000
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Print"
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
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "# 9931293724, 8002878845"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3840
      TabIndex        =   17
      Top             =   15
      Width           =   1980
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   5880
      Y1              =   1235
      Y2              =   1235
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   0
      X2              =   5880
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   0
      X2              =   4200
      Y1              =   1310
      Y2              =   1310
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   0
      X2              =   4200
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Authority Signature"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4200
      TabIndex        =   15
      Top             =   3095
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   1680
      Left            =   4320
      Picture         =   "Stu_ID_Card.frx":09EA
      Stretch         =   -1  'True
      Top             =   1340
      Width           =   1440
   End
   Begin VB.Label m6 
      BackStyle       =   0  'Transparent
      Caption         =   "05:30 to 06:30"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   2160
      TabIndex        =   13
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label m5 
      BackStyle       =   0  'Transparent
      Caption         =   "24-12-2019"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   2160
      TabIndex        =   12
      Top             =   2715
      Width           =   2055
   End
   Begin VB.Label m4 
      BackStyle       =   0  'Transparent
      Caption         =   "Railway (D)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   2160
      TabIndex        =   11
      Top             =   2415
      Width           =   2055
   End
   Begin VB.Label m3 
      BackStyle       =   0  'Transparent
      Caption         =   "Ram Nivas Sharma"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   2160
      TabIndex        =   10
      Top             =   2115
      Width           =   2055
   End
   Begin VB.Label m1 
      BackStyle       =   0  'Transparent
      Caption         =   "RS0003"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   2160
      TabIndex        =   9
      Top             =   1490
      Width           =   2055
   End
   Begin VB.Label m2 
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
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   2160
      TabIndex        =   8
      Top             =   1815
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Batch Time        :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Adm. Date         :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2715
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Course              :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2415
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Father's Name   :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2115
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Reg No              :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1485
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Student's Name :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1815
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Chini mill buxar, 802103"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   465
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SpeedUp Test Solutions"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   5880
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Image Image2 
      Height          =   735
      Left            =   15
      Picture         =   "Stu_ID_Card.frx":24C4
      Stretch         =   -1  'True
      Top             =   15
      Width           =   1275
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Computerized Test Series For Railway, SSC, 11th and 12th"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   840
      Width           =   5895
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   5895
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   3495
      Left            =   0
      Top             =   0
      Width           =   5905
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() 'Print Button
DV.Stud_Id_Gen regstudnt.reg_no
Set idcard.Sections("section1").Controls.Item("stphoto").Picture = LoadPicture(stuPicPath)
idcard.Sections("section1").Controls("text2").Caption = m1.Caption
idcard.Sections("section1").Controls("text3").Caption = stuname
idcard.Sections("section1").Controls("text4").Caption = stufather
idcard.Sections("section1").Controls("text5").Caption = stuCourse
idcard.Sections("section1").Controls("text6").Caption = stuIddate
idcard.Sections("section1").Controls("text7").Caption = stuBatch
idcard.Show vbModal, Me
idcard.Refresh
DV.rsStud_Id_Gen.Close
Unload Me
End Sub

Private Sub Form_Load()
conn
Image1.Picture = LoadPicture(stuPicPath)
m1.Caption = stud_id_pass.id.Caption
m2.Caption = stuname
m3.Caption = stufather
m4.Caption = stuCourse
m5.Caption = stuIddate
m6.Caption = stuBatch
End Sub
