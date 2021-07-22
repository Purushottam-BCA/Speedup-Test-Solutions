VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form about_org 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Us"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8805
   Icon            =   "about_org.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   735
      Left            =   0
      TabIndex        =   13
      Top             =   6810
      Width           =   8895
      Begin VB.CommandButton ChameleonBtn1 
         Height          =   375
         Left            =   120
         MouseIcon       =   "about_org.frx":0ECA
         MousePointer    =   99  'Custom
         Picture         =   "about_org.frx":101C
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   145
         Width           =   1695
      End
      Begin VB.CommandButton btnadd 
         Height          =   375
         Left            =   7200
         MouseIcon       =   "about_org.frx":1A33
         MousePointer    =   99  'Custom
         Picture         =   "about_org.frx":1B85
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   145
         Width           =   1455
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   0
         X2              =   8860
         Y1              =   45
         Y2              =   45
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         BorderWidth     =   3
         X1              =   0
         X2              =   8780
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5655
      Left            =   3950
      TabIndex        =   5
      Top             =   840
      Width           =   4455
      Begin VB.Frame Frame5 
         Height          =   2655
         Left            =   600
         TabIndex        =   14
         Top             =   1440
         Width           =   3375
         Begin VB.Frame Frame2 
            Caption         =   "Owner-Info"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2415
            Left            =   120
            TabIndex        =   15
            Top             =   100
            Width           =   3135
            Begin VB.Label lbl3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mr. Rakesh Sinha"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   225
               Left            =   1320
               TabIndex        =   25
               Top             =   1200
               Width           =   1470
            End
            Begin VB.Label lbl4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "9931293724"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   225
               Left            =   1320
               TabIndex        =   24
               Top             =   1560
               Width           =   1050
            End
            Begin VB.Label lbl5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Rakesh@gmail.com"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   225
               Left            =   1320
               TabIndex        =   23
               Top             =   1920
               Width           =   1680
            End
            Begin VB.Label lbl1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "1J254LNJ906125"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   225
               Left            =   1320
               TabIndex        =   22
               Top             =   480
               Width           =   1470
            End
            Begin VB.Label lbl2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "1J254LNJ905"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   225
               Left            =   1320
               TabIndex        =   21
               Top             =   840
               Width           =   1155
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Reg. No :"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   375
               Left            =   240
               TabIndex        =   20
               Top             =   480
               Width           =   1335
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "GST In No :"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   375
               Left            =   240
               TabIndex        =   19
               Top             =   840
               Width           =   1335
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Owner :"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   375
               Left            =   240
               TabIndex        =   18
               Top             =   1200
               Width           =   1335
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Mobile No :"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   375
               Left            =   240
               TabIndex        =   17
               Top             =   1560
               Width           =   1335
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "Email ID :"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   375
               Left            =   240
               TabIndex        =   16
               Top             =   1920
               Width           =   1335
            End
         End
      End
      Begin VB.Label ORGMob 
         AutoSize        =   -1  'True
         Caption         =   "9931293724, 8002878845"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1800
         TabIndex        =   12
         Top             =   5160
         Width           =   2190
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Conatct  :"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   11
         Top             =   5160
         Width           =   810
      End
      Begin VB.Label OrgMail 
         AutoSize        =   -1  'True
         Caption         =   "admin@STS.com"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1800
         MouseIcon       =   "about_org.frx":22DE
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Tag             =   "admin@emran-hasan.com"
         Top             =   4680
         Width           =   1605
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "E-mail ID :"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   9
         Top             =   4680
         Width           =   900
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version 1.0 "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3345
         TabIndex        =   8
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label lblCopyright 
         Alignment       =   1  'Right Justify
         Caption         =   "Copyright © 2019-20,SpeedUp Test Solutions."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -270
         TabIndex        =   7
         Top             =   600
         Width           =   4650
      End
      Begin VB.Label lblProduct 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Speedup Test Solutions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2130
         TabIndex        =   6
         Top             =   240
         Width           =   2160
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   6865
      Left            =   0
      ScaleHeight     =   6810
      ScaleWidth      =   3435
      TabIndex        =   3
      Top             =   0
      Width           =   3495
      Begin VB.Image Image3 
         Height          =   740
         Left            =   0
         Picture         =   "about_org.frx":2430
         Stretch         =   -1  'True
         Top             =   0
         Width           =   720
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         X1              =   0
         X2              =   3480
         Y1              =   6785
         Y2              =   6775
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         X1              =   3415
         X2              =   3415
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Speedup Test  Solutions"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   2415
         Left            =   240
         TabIndex        =   4
         Top             =   3120
         Width           =   2925
      End
      Begin VB.Image Image1 
         Height          =   1695
         Left            =   360
         Picture         =   "about_org.frx":54D4
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   2775
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   3600
      TabIndex        =   1
      Top             =   840
      Width           =   4815
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5655
         Left            =   360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   2
         Text            =   "about_org.frx":A2B9
         Top             =   0
         Width           =   4395
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6745
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   11906
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "About Organisation"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Readme"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "about_org"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnadd_Click()
Unload Me
End Sub

Private Sub ChameleonBtn1_Click()
FrmOrganisation.Show 1, MDI
End Sub

Private Sub Form_Load()
conn
CenterForm Me
Text1.Text = "Speedup test Solutions is an easy to use MCQ quiz application and MCQ Question Paper Generator that can be implemented in any kind of quiz taking purpose or Question paper Generating Process. Whether it is a quiz contest or a quiz for your institute, This System can help you organize a quiz quickly and easily. This software is divided into two part. The first part is the Administrator program which is used to create, modify or set a quiz. And the second part is the Quiz program which lets the users take the quiz. For security purpose, the users do not have the choice of running the administrator program." & vbCrLf & vbCrLf & "Currently Maximum Of 60 Questions Can Be Conducted Through test but In Future it may be Inhanced as per requirement. This Application Performs several Other Tasks Like Creating Student ID Card, Generating progress Report Etc."
Set r = c.Execute("select * from org")
If r.EOF = False Then
lblProduct.Caption = r.Fields(2)
 Label1.Caption = r.Fields(2)
 lbl1.Caption = r.Fields(0)
 lbl2.Caption = r.Fields(1)
 lbl3.Caption = r.Fields(6)
 lbl4.Caption = r.Fields(7)
 lbl5.Caption = r.Fields(5)
 OrgMail.Caption = r.Fields(5)
 ORGMob.Caption = r.Fields(4) & "," & r.Fields(7)
End If
End Sub

Private Sub TabStrip1_Click()
If TabStrip1.SelectedItem.Caption = "Readme" Then
    Frame1.Visible = False
    Frame3.Visible = True
Else
    Frame3.Visible = False
    Frame1.Visible = True
End If
End Sub
