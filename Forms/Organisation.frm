VERSION 5.00
Begin VB.Form FrmOrganisation 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Organisation Info"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8190
   Icon            =   "Organisation.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      MouseIcon       =   "Organisation.frx":0ECA
      MousePointer    =   99  'Custom
      Picture         =   "Organisation.frx":101C
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7920
      Width           =   1320
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      MouseIcon       =   "Organisation.frx":1819
      MousePointer    =   99  'Custom
      Picture         =   "Organisation.frx":196B
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7920
      Width           =   1335
   End
   Begin VB.TextBox t8 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3360
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   6960
      Width           =   3855
   End
   Begin VB.TextBox t7 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3360
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   6120
      Width           =   3855
   End
   Begin VB.TextBox t6 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3360
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   5280
      Width           =   3855
   End
   Begin VB.TextBox t5 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3360
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   4440
      Width           =   3855
   End
   Begin VB.TextBox t4 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3360
      MaxLength       =   20
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3600
      Width           =   3855
   End
   Begin VB.TextBox t3 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3360
      MaxLength       =   20
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2760
      Width           =   3855
   End
   Begin VB.TextBox t2 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3360
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1920
      Width           =   3855
   End
   Begin VB.TextBox t1 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1080
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Owner Mobile No        :"
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
      Index           =   7
      Left            =   600
      TabIndex        =   18
      Top             =   7005
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Owner Name               :"
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
      Index           =   6
      Left            =   600
      TabIndex        =   17
      Top             =   6165
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Organisation Email ID :"
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
      Index           =   5
      Left            =   600
      TabIndex        =   16
      Top             =   5325
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Organisation Contact  :"
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
      Index           =   4
      Left            =   600
      TabIndex        =   15
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Org. GST IN Number    :"
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
      Index           =   3
      Left            =   600
      TabIndex        =   14
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Org. Registration No    :"
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
      Index           =   2
      Left            =   600
      TabIndex        =   13
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Organisation Address  :"
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
      Index           =   1
      Left            =   600
      TabIndex        =   12
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Organisation Name     :"
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
      Index           =   0
      Left            =   600
      TabIndex        =   11
      Top             =   1150
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Organisation Information"
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
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   120
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000C&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   8175
   End
End
Attribute VB_Name = "FrmOrganisation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Trim(t1.Text) <> "" And Trim(t2.Text) <> "" And Trim(t3.Text) <> "" And Trim(t4.Text) <> "" And Trim(t5.Text) <> "" And Trim(t6.Text) <> "" And Trim(t7.Text) <> "" And Trim(t8.Text) <> "" Then
 c.Execute ("delete from org")
 c.Execute ("insert into org values('" & t3.Text & "','" & t4.Text & "','" & t1.Text & "','" & t2.Text & "','" & t5.Text & "','" & t6.Text & "','" & t7.Text & "','" & t8.Text & "')")
 MsgBox "SuccessFully Saved", vbInformation + vbOKOnly, ""
Else
MsgBox "All Fields are Compulsory.Please Fill All Fields.", vbInformation + vbOKOnly, ""
t1.SetFocus
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
conn
Set r = c.Execute("select * from org")
If r.EOF = False Then
 t1.Text = r.Fields(2)
 t2.Text = r.Fields(3)
 t3.Text = r.Fields(0)
 t4.Text = r.Fields(1)
 t5.Text = r.Fields(4)
 t6.Text = r.Fields(5)
 t7.Text = r.Fields(6)
 t8.Text = r.Fields(7)
Else
 t1.Text = ""
 t2.Text = ""
 t3.Text = ""
 t4.Text = ""
 t5.Text = ""
 t6.Text = ""
 t7.Text = ""
 t8.Text = ""
End If
End Sub

Private Sub t7_KeyPress(KeyAscii As Integer)
  If (((KeyAscii >= 65) And (KeyAscii <= 90)) Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or (KeyAscii = 8) Or (KeyAscii = 32)) Then
        t7.SetFocus
  ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        t8.SetFocus
  Else
   KeyAscii = 0
  End If
End Sub

Private Sub t8_KeyPress(KeyAscii As Integer)
If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Then
        t8.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
  Else
   KeyAscii = 0
End If
End Sub

Private Sub t8_LostFocus()
If (t8.Text <> "") Then
        If (Len(t8.Text) < 10) Then
            MsgBox "Invalid MOBILE NUMBER", vbExclamation + vbOKOnly, "Invalid  Mobile No"
            t8.Text = ""
            t8.SetFocus
        End If
End If
End Sub


Private Sub t5_KeyPress(KeyAscii As Integer)
If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Then
        t5.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   t6.SetFocus
  Else
   KeyAscii = 0
End If
End Sub

Private Sub t5_LostFocus()
If (t5.Text <> "") Then
        If (Len(t5.Text) < 10) Then
            MsgBox "Invalid MOBILE NUMBER", vbExclamation + vbOKOnly, "Invalid  Mobile No"
            t5.Text = ""
            t5.SetFocus
        End If
End If
End Sub



Private Sub t3_KeyPress(KeyAscii As Integer)
If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Then
        t3.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   t4.SetFocus
  Else
   KeyAscii = 0
End If
End Sub
Private Sub t4_KeyPress(KeyAscii As Integer)
If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Then
        t4.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   t5.SetFocus
  Else
   KeyAscii = 0
End If
End Sub
