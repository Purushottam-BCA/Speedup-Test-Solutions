VERSION 5.00
Object = "{BF38D12B-22A9-4B10-B26E-019F2B5F9C22}#1.0#0"; "AniGif.ocx"
Begin VB.Form formConfirm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   1185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3030
   LinkTopic       =   "Form5"
   ScaleHeight     =   1185
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   600
      Top             =   960
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   1080
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "QUESTION SUCCESSFULLY ADDED"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   1455
         Left            =   -240
         TabIndex        =   1
         Top             =   120
         Width           =   2460
      End
   End
   Begin Project1.PictureG PictureG1 
      Height          =   1755
      Left            =   -600
      Top             =   -360
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   3096
      GIF             =   "formConfirm.frx":0000
      Delay           =   10
      Stretch         =   2
      DelayLoad       =   0
   End
   Begin VB.Image Image2 
      Height          =   1140
      Left            =   0
      Picture         =   "formConfirm.frx":41CA2
      Stretch         =   -1  'True
      Top             =   75
      Width           =   1140
   End
End
Attribute VB_Name = "formConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Private Sub Form_Load()
a = 0
CenterForm Me
CreateRoundRectFromWindow Me
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
a = a + 50
If a > 1900 Then
Timer1.Enabled = False
Unload Me
End If
End Sub
