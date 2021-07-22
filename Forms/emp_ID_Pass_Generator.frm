VERSION 5.00
Object = "{BF38D12B-22A9-4B10-B26E-019F2B5F9C22}#1.0#0"; "AniGif.ocx"
Begin VB.Form stud_id_pass 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9570
   FillColor       =   &H00C0C0FF&
   FillStyle       =   0  'Solid
   Icon            =   "emp_ID_Pass_Generator.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   9570
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   5175
      Left            =   480
      ScaleHeight     =   5115
      ScaleWidth      =   8505
      TabIndex        =   0
      Top             =   480
      Width           =   8565
      Begin VB.TextBox answer 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3240
         TabIndex        =   3
         Top             =   2760
         Width           =   4455
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FFFFFF&
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
         ItemData        =   "emp_ID_Pass_Generator.frx":0ECA
         Left            =   3240
         List            =   "emp_ID_Pass_Generator.frx":0ECC
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2040
         Width           =   4455
      End
      Begin VB.CommandButton vkCommand1 
         Caption         =   "Generate ID card"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5760
         MouseIcon       =   "emp_ID_Pass_Generator.frx":0ECE
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   3900
         Width           =   1935
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H00404040&
         Height          =   450
         Left            =   5740
         Top             =   3890
         Width           =   1960
      End
      Begin VB.Shape Shape1 
         Height          =   4935
         Left            =   120
         Top             =   120
         Width           =   8295
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   $"emp_ID_Pass_Generator.frx":1020
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   555
         Left            =   240
         TabIndex        =   12
         Top             =   4560
         Width           =   8220
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Security Question :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         TabIndex        =   11
         Top             =   2040
         Width           =   1980
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Answer : "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         TabIndex        =   10
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Log_In ID : "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         TabIndex        =   9
         Top             =   720
         Width           =   1140
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         TabIndex        =   8
         Top             =   1320
         Width           =   1065
      End
      Begin VB.Label log_id 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3240
         TabIndex        =   7
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Password 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3240
         TabIndex        =   6
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label id 
         BackColor       =   &H80000010&
         Height          =   255
         Left            =   960
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Query 
         BackColor       =   &H8000000D&
         Height          =   375
         Left            =   2520
         TabIndex        =   4
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin Project1.PictureG PictureG1 
         Height          =   420
         Left            =   6120
         Top             =   675
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   741
         GIF             =   "emp_ID_Pass_Generator.frx":10AF
         Stretch         =   2
         Mirror          =   1
      End
      Begin Project1.PictureG PictureG2 
         Height          =   420
         Left            =   6120
         Top             =   1335
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   741
         GIF             =   "emp_ID_Pass_Generator.frx":6B61
         Stretch         =   2
         Mirror          =   1
      End
   End
End
Attribute VB_Name = "stud_id_pass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
conn
CenterForm Me
Set r = c.Execute("select ques from SECQUES")
Combo1.Clear
While r.EOF = False
  Combo1.AddItem r.Fields(0)
 r.MoveNext
Wend
If regstudnt.Combo6.ListIndex = 1 Then
vkCommand1.Caption = "Ok"
Else
vkCommand1.Caption = "Generate ID Card "
End If
End Sub

Private Sub vkCommand1_Click()
If vkCommand1.Caption = "Ok" Then
GoTo kml:
 Exit Sub
End If
kml:
If id.Caption = "" Or log_id.Caption = "" Or Password.Caption = "" Then
 MsgBox "Cannot Get Data From registration Info", vbInformation, ""
ElseIf answer.Text = "" Or Combo1.Text = "" Then
 MsgBox "Cann't Leave it Blank", vbCritical + vbOKOnly, ""
 answer.SetFocus
Else
Set r = New ADODB.Recordset
If vkCommand1.Caption <> "Ok" Then
Form1.Show 1, Me
End If
Set r = c.Execute("insert into stud_login values('" & id.Caption & "','" & log_id.Caption & "','" & Password.Caption & "','" & Combo1.Text & "','" & answer.Text & "')")
MsgBox "Student SuccessFully Registered. Now He/She can login with his/her id and password to attempt Tests.", vbInformation + vbOKOnly, "New Registration"
Unload Me
End If
End Sub
