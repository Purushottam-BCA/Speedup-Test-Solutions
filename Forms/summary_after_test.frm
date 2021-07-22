VERSION 5.00
Begin VB.Form Summary_Test 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Summary"
   ClientHeight    =   9465
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15630
   Icon            =   "summary_after_test.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9465
   ScaleWidth      =   15630
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1095
      Left            =   1180
      TabIndex        =   38
      Top             =   7875
      Width           =   13250
      Begin VB.CommandButton certi 
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
         Height          =   640
         Left            =   150
         MouseIcon       =   "summary_after_test.frx":602E
         MousePointer    =   99  'Custom
         Picture         =   "summary_after_test.frx":6180
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Certificate For Test."
         Top             =   240
         Width           =   2070
      End
      Begin VB.CommandButton anaReport 
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
         Height          =   640
         Left            =   2400
         MouseIcon       =   "summary_after_test.frx":6D85
         MousePointer    =   99  'Custom
         Picture         =   "summary_after_test.frx":6ED7
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Detailed Analysis And Report"
         Top             =   240
         Width           =   2190
      End
      Begin VB.CommandButton ans_ky 
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
         Height          =   640
         Left            =   7080
         MouseIcon       =   "summary_after_test.frx":7D8A
         MousePointer    =   99  'Custom
         Picture         =   "summary_after_test.frx":7EDC
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Only Answer Key."
         Top             =   240
         Width           =   1830
      End
      Begin VB.CommandButton Qppr 
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
         Height          =   640
         Left            =   4800
         MouseIcon       =   "summary_after_test.frx":8ABD
         MousePointer    =   99  'Custom
         Picture         =   "summary_after_test.frx":8C0F
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Only Question Paper"
         Top             =   240
         Width           =   2070
      End
      Begin VB.CommandButton compreAns 
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
         Height          =   640
         Left            =   9105
         MouseIcon       =   "summary_after_test.frx":9A21
         MousePointer    =   99  'Custom
         Picture         =   "summary_after_test.frx":9B73
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Compare Answer With Correct One."
         Top             =   240
         Width           =   2430
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
         Height          =   640
         Left            =   11700
         MouseIcon       =   "summary_after_test.frx":AACD
         MousePointer    =   99  'Custom
         Picture         =   "summary_after_test.frx":AC1F
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   1345
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Height          =   1080
         Left            =   15
         Top             =   15
         Width           =   13215
      End
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "100 %"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12675
      TabIndex        =   42
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Accuracy : "
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   1
      Left            =   11475
      TabIndex        =   41
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Line Line24 
      BorderColor     =   &H00FFFFFF&
      X1              =   1150
      X2              =   3500
      Y1              =   7755
      Y2              =   7755
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Summary Options"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1200
      TabIndex        =   40
      Top             =   7395
      Width           =   2100
   End
   Begin VB.Line Line23 
      BorderColor     =   &H00FFFFFF&
      X1              =   1380
      X2              =   3360
      Y1              =   4755
      Y2              =   4755
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Overall Report"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1440
      TabIndex        =   39
      Top             =   4395
      Width           =   1800
   End
   Begin VB.Image Image5 
      Height          =   420
      Left            =   11400
      Picture         =   "summary_after_test.frx":B82D
      Stretch         =   -1  'True
      Top             =   5115
      Width           =   435
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-: Test Summary :-"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   405
      Left            =   6300
      TabIndex        =   37
      Top             =   1245
      Width           =   2670
   End
   Begin VB.Line Line22 
      X1              =   -600
      X2              =   -600
      Y1              =   -480
      Y2              =   7560
   End
   Begin VB.Line Line21 
      BorderColor     =   &H00808080&
      X1              =   1200
      X2              =   14400
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "SSC CGL "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   375
      Left            =   2640
      TabIndex        =   36
      Top             =   1275
      Width           =   3495
   End
   Begin VB.Line Line20 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   14370
      X2              =   14370
      Y1              =   1080
      Y2              =   4185
   End
   Begin VB.Line Line19 
      BorderColor     =   &H00808080&
      X1              =   1200
      X2              =   14350
      Y1              =   4185
      Y2              =   4185
   End
   Begin VB.Line Line18 
      BorderColor     =   &H00808080&
      X1              =   10200
      X2              =   10200
      Y1              =   1815
      Y2              =   4205
   End
   Begin VB.Line Line17 
      BorderColor     =   &H00808080&
      X1              =   3360
      X2              =   3360
      Y1              =   1815
      Y2              =   4205
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00808080&
      X1              =   1200
      X2              =   14350
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00808080&
      X1              =   1200
      X2              =   14350
      Y1              =   2970
      Y2              =   2970
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00808080&
      X1              =   1200
      X2              =   14350
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Minute"
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
      Left            =   11160
      TabIndex        =   35
      Top             =   1995
      Width           =   1095
   End
   Begin VB.Label l8 
      BackStyle       =   0  'Transparent
      Caption         =   "Pass"
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
      Left            =   10440
      TabIndex        =   34
      Top             =   3720
      Width           =   2775
   End
   Begin VB.Label l7 
      BackStyle       =   0  'Transparent
      Caption         =   "25"
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
      Left            =   10440
      TabIndex        =   33
      Top             =   3120
      Width           =   3495
   End
   Begin VB.Label l6 
      BackStyle       =   0  'Transparent
      Caption         =   "60"
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
      Left            =   10440
      TabIndex        =   32
      Top             =   2565
      Width           =   3495
   End
   Begin VB.Label l5 
      BackStyle       =   0  'Transparent
      Caption         =   "60 : 00"
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
      Left            =   10440
      TabIndex        =   31
      Top             =   1995
      Width           =   735
   End
   Begin VB.Label l4 
      BackStyle       =   0  'Transparent
      Caption         =   "Random"
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
      Left            =   3600
      TabIndex        =   30
      Top             =   3720
      Width           =   2775
   End
   Begin VB.Label l3 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject Wise Test"
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
      Left            =   3600
      TabIndex        =   29
      Top             =   3120
      Width           =   3495
   End
   Begin VB.Label l2 
      BackStyle       =   0  'Transparent
      Caption         =   "Mukesh Kumar"
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
      Left            =   3600
      TabIndex        =   28
      Top             =   2565
      Width           =   3495
   End
   Begin VB.Label l1 
      BackStyle       =   0  'Transparent
      Caption         =   "28-11-2019"
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
      Left            =   3600
      TabIndex        =   27
      Top             =   1995
      Width           =   2175
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00B0B0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Course : "
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   360
      Left            =   1440
      TabIndex        =   26
      Top             =   1260
      Width           =   1140
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      X1              =   -600
      X2              =   -600
      Y1              =   -480
      Y2              =   240
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Final Result : "
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   0
      Left            =   8400
      TabIndex        =   25
      Top             =   3720
      Width           =   2655
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Obtained Marks :"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   8400
      TabIndex        =   24
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Full Marks : "
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   8400
      TabIndex        =   23
      Top             =   2565
      Width           =   2655
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Time :"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   8400
      TabIndex        =   22
      Top             =   1995
      Width           =   2655
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Diffficulty Level : "
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   1680
      TabIndex        =   21
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Test Type :"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   1680
      TabIndex        =   20
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Name : "
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   1680
      TabIndex        =   19
      Top             =   2565
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date : "
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   1680
      TabIndex        =   18
      Top             =   1995
      Width           =   2655
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   7800
      X2              =   7800
      Y1              =   1815
      Y2              =   4205
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00808080&
      X1              =   1200
      X2              =   14400
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      Height          =   3135
      Left            =   1200
      Top             =   1080
      Width           =   13185
   End
   Begin VB.Image Image6 
      Height          =   420
      Left            =   13200
      Picture         =   "summary_after_test.frx":C0F7
      Stretch         =   -1  'True
      Top             =   5115
      Width           =   435
   End
   Begin VB.Image Image4 
      Height          =   315
      Left            =   10200
      Picture         =   "summary_after_test.frx":C9C1
      Stretch         =   -1  'True
      Top             =   5145
      Width           =   315
   End
   Begin VB.Image Image3 
      Height          =   315
      Left            =   9015
      Picture         =   "summary_after_test.frx":CF64
      Stretch         =   -1  'True
      Top             =   5145
      Width           =   315
   End
   Begin VB.Image Image2 
      Height          =   315
      Left            =   7635
      Picture         =   "summary_after_test.frx":D5C6
      Stretch         =   -1  'True
      Top             =   5145
      Width           =   315
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   1635
      Picture         =   "summary_after_test.frx":DADD
      Stretch         =   -1  'True
      Top             =   5145
      Width           =   345
   End
   Begin VB.Label Rem_Tim_Smry 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   360
      Left            =   12480
      TabIndex        =   17
      Top             =   6480
      Width           =   1845
   End
   Begin VB.Label score 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "30"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   360
      Left            =   10200
      TabIndex        =   16
      Top             =   6480
      Width           =   345
   End
   Begin VB.Label total 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   360
      Left            =   1680
      TabIndex        =   15
      Top             =   6480
      Width           =   510
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00E0E0E0&
      X1              =   10920
      X2              =   10920
      Y1              =   4920
      Y2              =   7200
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   1560
      TabIndex        =   14
      Top             =   5550
      Width           =   555
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00E0E0E0&
      X1              =   9840
      X2              =   9840
      Y1              =   4920
      Y2              =   7200
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Correct"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   7440
      TabIndex        =   13
      Top             =   5550
      Width           =   765
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "  Wrong"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   8760
      TabIndex        =   12
      Top             =   5550
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "  Score"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   9960
      TabIndex        =   11
      Top             =   5550
      Width           =   855
   End
   Begin VB.Label correct 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "27"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   360
      Left            =   7680
      TabIndex        =   10
      Top             =   6480
      Width           =   345
   End
   Begin VB.Label wrong 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "23"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   360
      Left            =   9120
      TabIndex        =   9
      Top             =   6480
      Width           =   345
   End
   Begin VB.Label Passd_time 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "05:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   360
      Left            =   11040
      TabIndex        =   8
      Top             =   6480
      Width           =   1245
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H008130FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H008130FF&
      Height          =   255
      Left            =   4800
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label rTime 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Remaining Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   12480
      TabIndex        =   7
      Top             =   5550
      Width           =   2055
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   11070
      TabIndex        =   6
      Top             =   5550
      Width           =   1140
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00E0E0E0&
      X1              =   12360
      X2              =   12360
      Y1              =   4920
      Y2              =   7200
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00E0E0E0&
      X1              =   8520
      X2              =   8520
      Y1              =   4920
      Y2              =   7200
   End
   Begin VB.Label nseen 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   360
      Left            =   6360
      TabIndex        =   5
      Top             =   6480
      Width           =   180
   End
   Begin VB.Label nanswered 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   360
      Left            =   4800
      TabIndex        =   4
      Top             =   6480
      Width           =   180
   End
   Begin VB.Label answered 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   360
      Left            =   3000
      TabIndex        =   3
      Top             =   6480
      Width           =   345
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   255
      Left            =   6360
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080FF80&
      Height          =   255
      Left            =   3120
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Not Seen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   6000
      TabIndex        =   2
      Top             =   5550
      Width           =   975
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Not Answered"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   4200
      TabIndex        =   1
      Top             =   5550
      Width           =   1455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Answered"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   2760
      TabIndex        =   0
      Top             =   5550
      Width           =   1035
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00E0E0E0&
      X1              =   1200
      X2              =   14400
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00E0E0E0&
      X1              =   7200
      X2              =   7200
      Y1              =   4920
      Y2              =   7200
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00E0E0E0&
      X1              =   5760
      X2              =   5760
      Y1              =   4920
      Y2              =   7200
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00E0E0E0&
      X1              =   3960
      X2              =   3960
      Y1              =   4920
      Y2              =   7200
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      X1              =   2400
      X2              =   2400
      Y1              =   4920
      Y2              =   7200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   1200
      X2              =   14400
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      Height          =   2295
      Left            =   1200
      Top             =   4920
      Width           =   13215
   End
End
Attribute VB_Name = "Summary_Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NotAns As Integer, comnd As String
Dim answered1 As Integer, cor_anss As Integer, wrong_anss As Integer
Dim nvisiting As Integer, actualno As Integer
Dim pos_mrk As Integer, neg_mrk As Integer, snumb As Integer

Private Sub anaReport_Click() 'Analysis Report
If NonPackage = 1 Then
 MsgBox "This Facility is Provided only For Package Registered Student..", vbInformation + vbOKOnly, "Non Registered Student"
 Exit Sub
Else
 DV.rsanalysisAfterTest.Open
 TestAnalysis.Sections("section4").Controls("StuNm").Caption = l2.Caption
 TestAnalysis.Show 1, MDI
 TestAnalysis.Refresh
 DV.rsanalysisAfterTest.Close
End If
End Sub

Private Sub ans_ky_Click()
FrmAnsKey.Show 1, MDI
End Sub

Private Sub certi_Click()
Me.Enabled = False
If NonPackage = 1 Then 'Individual
 FrmCerti2.Show
Else 'Package
 Certificate.Show
End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
c.Execute ("delete from answerhold")
c.Execute ("delete from mcqtest")
Unload Me
stu_dash.Show
End Sub

Private Sub compreAns_Click()
If NonPackage = 1 Then
 MsgBox "This Facility is Provided only For Package Registered Student..", vbInformation + vbOKOnly, "Non Registered Student"
 Exit Sub
Else
 CmpreChrt.Show 1, MDI
End If
End Sub

Private Sub Form_Load()
Dim tmrk As Single
Me.Top = 400
Me.Left = 2200
conn
NotAns = 0
answered1 = 0
cor_anss = 0
wrong_anss = 0
nvisiting = 0
pos_mrk = FMRKPERCOR
neg_mrk = FMRKPERWRONG

Label20.Caption = GivenTESTCourse  'Course
l1.Caption = Format(Date, "DD-MM-YYYY")
l2.Caption = StuNam 'Name
l3.Caption = selectedType 'Tst Type
l4.Caption = selectedlvl 'Diff Level
l5.Caption = ToTaTiMe 'Time
l6.Caption = "" 'Full Marks
l7.Caption = "" 'Obtained Marks
l8.Caption = "" 'Pass or Fail
Set r1 = New ADODB.Recordset
Set r1 = c.Execute("select count(*) from answerhold ")
If IsNull(r1.Fields(0)) = False Then
 total.Caption = r1.Fields(0)
End If
Set r = New ADODB.Recordset
Set r = c1.Execute("select * from answerhold ")
If IsNull(r.Fields(0)) = False Then
While r.EOF = False
 If r.Fields(2) <> 0 Then 'answered
  answered1 = answered1 + 1
 ElseIf (r.Fields(2) = 0 And r.Fields(3) = 2) Then 'Not Answered
  NotAns = NotAns + 1
 End If
 If r.Fields(2) <> 0 And r.Fields(1) = r.Fields(2) Then 'Correct Answer
  cor_anss = cor_anss + 1
 ElseIf r.Fields(2) <> 0 And r.Fields(1) <> r.Fields(2) Then 'Wrong Answer
  wrong_anss = wrong_anss + 1
 End If
  If r.Fields(3) = 0 Then 'not visiting question
   nvisiting = nvisiting + 1
  End If
 r.MoveNext
Wend
End If
Rem_Tim_Smry.Caption = remainTIM
Passd_time.Caption = ToTaTiMe

answered.Caption = answered1
nanswered.Caption = NotAns
nseen.Caption = nvisiting
correct.Caption = cor_anss
wrong.Caption = wrong_anss
If Val(correct.Caption) = Val(answered.Caption) Then
 Label23.Caption = "100 %"
Else
 Label23.Caption = Format((Val(correct.Caption) * 100) / Val(answered.Caption), "00.00") & " %"
End If
SCore.Caption = (cor_anss * pos_mrk) - (wrong_anss * neg_mrk)
l6.Caption = FTOTMARKS
l7.Caption = SCore.Caption
If Val(l6.Caption) <> 0 Then
tmrk = (Val(l7.Caption) * 100) / Val(l6.Caption)
End If
If tmrk >= FPASSPERCENTG Then
 l8.Caption = "Pass ( " & Format(tmrk, "00.00") & " %)"
Else
 l8.Caption = "Fail ( " & Format(tmrk, "00.00") & " %)"
End If
Stu_login_reg_no = Current_Logged_ID
anaReport.Enabled = True
compreAns.Enabled = True
If NonPackage <> 1 Then 'Registered
 Set r = New ADODB.Recordset
 Set r = c.Execute("select count(*) from STUD_PREV_REC where rstud_reg_no='" & Stu_login_reg_no & "' ")
 snumb = r.Fields(0)
 snumb = snumb + 1
 actualno = (Val(nanswered.Caption) + Val(nseen.Caption))
 c1.Execute ("insert into STUD_PREV_REC values(" & snumb & ",'" & Format(l1.Caption, "dd-mmm-yyyy") & "','" & l3.Caption & "'," & Val(l6.Caption) & "," & Val(l7.Caption) & ",'" & l4.Caption & "','" & Left$(l8.Caption, 4) & "','" & Stu_login_reg_no & "'," & Val(total.Caption) & "," & Val(correct.Caption) & "," & Val(wrong.Caption) & "," & actualno & ",'" & l5.Caption & "','" & Rem_Tim_Smry.Caption & "' )")
End If
If UCase(Left$(l8.Caption, 4)) = "FAIL" Then
 certi.Enabled = False
Else
 certi.Enabled = True
End If
End Sub

Private Sub Form_Unload(cancel As Integer)
On Error Resume Next
stu_dash.Show
Unload Summary_Test
End Sub

Private Sub Qppr_Click()
mcqTestRunPPr.Orientation = rptOrientPortrait
mcqTestRunPPr.Show 1
mcqTestRunPPr.Refresh
DV.rsTest_Ppr_mcq.Close
End Sub
