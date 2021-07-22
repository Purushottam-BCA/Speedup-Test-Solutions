VERSION 5.00
Begin VB.Form FrmChangeWallppr 
   BackColor       =   &H80000010&
   BorderStyle     =   0  'None
   Caption         =   "Change Wallpaper"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   MouseIcon       =   "ChangeWallpaper.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   11520
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
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
      Left            =   18960
      MouseIcon       =   "ChangeWallpaper.frx":0152
      MousePointer    =   99  'Custom
      Picture         =   "ChangeWallpaper.frx":02A4
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Exit"
      Top             =   2280
      Width           =   1305
   End
   Begin VB.CommandButton Command1 
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
      Height          =   495
      Left            =   18960
      MouseIcon       =   "ChangeWallpaper.frx":0EB2
      MousePointer    =   99  'Custom
      Picture         =   "ChangeWallpaper.frx":1004
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Exit"
      Top             =   1515
      Width           =   1305
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000010&
      Caption         =   "Pictures"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   5790
      Width           =   990
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000010&
      Caption         =   "Solid Colors"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   1110
      Width           =   1350
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Height          =   4095
      Index           =   1
      Left            =   360
      Top             =   6000
      Width           =   18375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Height          =   4095
      Index           =   0
      Left            =   360
      Top             =   1320
      Width           =   18375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Background Image :-"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   390
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3555
   End
   Begin VB.Image Image7 
      BorderStyle     =   1  'Fixed Single
      Height          =   3375
      Left            =   14640
      Picture         =   "ChangeWallpaper.frx":19CE
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   3735
   End
   Begin VB.Image Image6 
      BorderStyle     =   1  'Fixed Single
      Height          =   3375
      Left            =   720
      Picture         =   "ChangeWallpaper.frx":2A67
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   3855
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  'Fixed Single
      Height          =   3375
      Left            =   5280
      Picture         =   "ChangeWallpaper.frx":E688
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   4095
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  'Fixed Single
      Height          =   3375
      Left            =   10080
      Picture         =   "ChangeWallpaper.frx":175AA
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   3855
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   3375
      Left            =   12600
      Picture         =   "ChangeWallpaper.frx":1A6FC
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   5775
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   3375
      Left            =   6600
      Picture         =   "ChangeWallpaper.frx":4239C
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   5775
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3375
      Left            =   720
      Picture         =   "ChangeWallpaper.frx":5DA42
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   5655
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000013&
      FillColor       =   &H80000013&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   20415
   End
End
Attribute VB_Name = "FrmChangeWallppr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
stu_dash.Image4.Picture = LoadPicture(CurrentDashPic)
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
Me.Width = MDI.Width
Me.Height = MDI.Height
End Sub

Private Sub Image1_Click()
CurrentDashPic = App.Path & "\Graphics\#new\#4.jpg"
End Sub

Private Sub Image2_Click()
CurrentDashPic = App.Path & "\Graphics\#new\#5.jpg"
End Sub

Private Sub Image3_Click()
CurrentDashPic = App.Path & "\Graphics\#new\#6.jpg"
End Sub

Private Sub Image2_DblClick()
Command1_Click
End Sub

Private Sub Image1_DblClick()
Command1_Click
End Sub

Private Sub Image3_DblClick()
Command1_Click
End Sub

Private Sub Image4_Click()
CurrentDashPic = App.Path & "\Graphics\#new\#2.jpg"
End Sub

Private Sub Image5_Click()
CurrentDashPic = App.Path & "\Graphics\#new\#1.jpg"
End Sub

Private Sub Image6_Click()
CurrentDashPic = App.Path & "\Graphics\#new\#32.jpg"
End Sub

Private Sub Image7_Click()
CurrentDashPic = App.Path & "\Graphics\#new\#3.jpg"
End Sub

Private Sub Image7_DblClick()
Command1_Click
End Sub

Private Sub Image6_DblClick()
Command1_Click
End Sub

Private Sub Image5_DblClick()
Command1_Click
End Sub

Private Sub Image4_DblClick()
Command1_Click
End Sub

