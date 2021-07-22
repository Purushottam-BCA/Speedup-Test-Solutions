VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F5E116E1-0563-11D8-AA80-000B6A0D10CB}#1.0#0"; "HookMenu.ocx"
Begin VB.Form admin_dash 
   BackColor       =   &H00F0E0E0&
   Caption         =   "Dashboard"
   ClientHeight    =   10455
   ClientLeft      =   225
   ClientTop       =   3540
   ClientWidth     =   19950
   ControlBox      =   0   'False
   Icon            =   "admin_dash.frx":0000
   LinkTopic       =   "Form10"
   MDIChild        =   -1  'True
   ScaleHeight     =   10455
   ScaleWidth      =   19950
   WindowState     =   2  'Maximized
   Begin VB.Timer DashTiMEr 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   3240
      Top             =   2640
   End
   Begin VB.Timer Timer2 
      Interval        =   120
      Left            =   3240
      Top             =   2280
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3240
      Top             =   3000
   End
   Begin HookMenu.XpMenu XpMenu1 
      Left            =   3240
      Top             =   1800
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   54
      CheckBorderColor=   7021576
      SelMenuBorder   =   7021576
      SelMenuBackColor=   14073525
      SelMenuForeColor=   16646297
      SelCheckBackColor=   14791828
      MenuBorderColor =   6956042
      SeparatorColor  =   -2147483632
      MenuBackColor   =   16777215
      MenuForeColor   =   0
      CheckBackColor  =   15326939
      CheckForeColor  =   10027263
      DisabledMenuBorderColor=   -2147483632
      DisabledMenuBackColor=   16777215
      DisabledMenuForeColor=   -2147483631
      MenuBarBackColor=   15790320
      MenuPopupBackColor=   16777215
      ShortCutNormalColor=   0
      ShortCutSelectColor=   16646297
      ArrowNormalColor=   10027263
      ArrowSelectColor=   12484864
      ShadowColor     =   0
      Bmp:1           =   "admin_dash.frx":6062
      Mask:1          =   16776951
      Key:1           =   "#l_info"
      Bmp:2           =   "admin_dash.frx":63C0
      Mask:2          =   16775423
      Key:2           =   "#abut_org"
      Bmp:3           =   "admin_dash.frx":6712
      Key:3           =   "#crse"
      Bmp:4           =   "admin_dash.frx":6B3A
      Mask:4          =   14258990
      Key:4           =   "#sub"
      Bmp:5           =   "admin_dash.frx":6F4C
      Mask:5          =   13692927
      Key:5           =   "#quesbank"
      Bmp:6           =   "admin_dash.frx":735E
      Mask:6          =   1909157
      Key:6           =   "#cnewset"
      Bmp:7           =   "admin_dash.frx":77A0
      Mask:7          =   15461863
      Key:7           =   "#vallset"
      Bmp:8           =   "admin_dash.frx":7B3A
      Mask:8          =   1909157
      Key:8           =   "#create"
      Bmp:9           =   "admin_dash.frx":7F7C
      Mask:9          =   16776703
      Key:9           =   "#Viw"
      Bmp:10          =   "admin_dash.frx":8342
      Mask:10         =   8288628
      Key:10          =   "#delete"
      Bmp:11          =   "admin_dash.frx":876A
      Mask:11         =   8288628
      Key:11          =   "#delset"
      Bmp:12          =   "admin_dash.frx":8B92
      Mask:12         =   8288628
      Key:12          =   "#dlt"
      Bmp:13          =   "admin_dash.frx":8FBA
      Mask:13         =   16515071
      Key:13          =   "#pkgstd"
      Bmp:14          =   "admin_dash.frx":938C
      Mask:14         =   16514559
      Key:14          =   "#npkgstud"
      Bmp:15          =   "admin_dash.frx":96DE
      Mask:15         =   16449532
      Key:15          =   "#emp"
      Bmp:16          =   "admin_dash.frx":9AA4
      Mask:16         =   16777215
      Key:16          =   "#salary"
      Bmp:17          =   "admin_dash.frx":9E66
      Mask:17         =   14211288
      Key:17          =   "#detl"
      Bmp:18          =   "admin_dash.frx":A1F8
      Mask:18         =   16777215
      Key:18          =   "#payment"
      Bmp:19          =   "admin_dash.frx":A5BA
      Mask:19         =   16777214
      Key:19          =   "#viewed"
      Bmp:20          =   "admin_dash.frx":A918
      Mask:20         =   16777215
      Key:20          =   "#calculator"
      Bmp:21          =   "admin_dash.frx":AEBE
      Mask:21         =   16777215
      Key:21          =   "#notepad"
      Bmp:22          =   "admin_dash.frx":B464
      Mask:22         =   13424076
      Key:22          =   "#exprep"
      Bmp:23          =   "admin_dash.frx":BA0A
      Mask:23         =   13424076
      Key:23          =   "#increp"
      Bmp:24          =   "admin_dash.frx":BFB0
      Mask:24         =   16777215
      Key:24          =   "#income"
      Bmp:25          =   "admin_dash.frx":C556
      Mask:25         =   13412966
      Key:25          =   "#expense"
      Bmp:26          =   "admin_dash.frx":CB10
      Mask:26         =   16777215
      Key:26          =   "#overall"
      Bmp:27          =   "admin_dash.frx":D0B6
      Mask:27         =   14678015
      Key:27          =   "#tpc"
      Bmp:28          =   "admin_dash.frx":D4F8
      Mask:28         =   16777215
      Key:28          =   "#pkg1"
      Bmp:29          =   "admin_dash.frx":DAB2
      Mask:29         =   14456648
      Key:29          =   "#questp"
      Bmp:30          =   "admin_dash.frx":DE4C
      Mask:30         =   16184823
      Key:30          =   "#schdl"
      Bmp:31          =   "admin_dash.frx":E256
      Mask:31         =   1909157
      Key:31          =   "#add"
      Bmp:32          =   "admin_dash.frx":E698
      Mask:32         =   16645629
      Key:32          =   "#view"
      Bmp:33          =   "admin_dash.frx":E9BA
      Mask:33         =   15461863
      Key:33          =   "#see"
      Bmp:34          =   "admin_dash.frx":ED54
      Mask:34         =   16316150
      Key:34          =   "#quespaperbank"
      Bmp:35          =   "admin_dash.frx":F11A
      Mask:35         =   16777215
      Key:35          =   "#attendence"
      Bmp:36          =   "admin_dash.frx":F6C0
      Mask:36         =   14145495
      Key:36          =   "#stud"
      Bmp:37          =   "admin_dash.frx":FB02
      Mask:37         =   10070681
      Key:37          =   "#srstu"
      Bmp:38          =   "admin_dash.frx":100A8
      Mask:38         =   16777215
      Key:38          =   "#viewe"
      Bmp:39          =   "admin_dash.frx":1064E
      Mask:39         =   16316150
      Key:39          =   "#tpc1"
      Bmp:40          =   "admin_dash.frx":10A14
      Mask:40         =   16777215
      Key:40          =   "#schdlee"
      Bmp:41          =   "admin_dash.frx":10FBA
      Mask:41         =   14456648
      Key:41          =   "#quebank"
      Bmp:42          =   "admin_dash.frx":11354
      Mask:42         =   5405713
      Key:42          =   "#tstdft"
      Bmp:43          =   "admin_dash.frx":1160E
      Mask:43         =   16449532
      Key:43          =   "#empd"
      Bmp:44          =   "admin_dash.frx":119D4
      Mask:44         =   3415795
      Key:44          =   "#pntcheck"
      Bmp:45          =   "admin_dash.frx":11C8E
      Mask:45         =   16777215
      Key:45          =   "#secQu"
      Bmp:46          =   "admin_dash.frx":127A8
      Mask:46         =   5405713
      Key:46          =   "#cl"
      Bmp:47          =   "admin_dash.frx":12A62
      Mask:47         =   5405713
      Key:47          =   "#act"
      Bmp:48          =   "admin_dash.frx":12D1C
      Mask:48         =   5405713
      Key:48          =   "#mzone"
      Bmp:49          =   "admin_dash.frx":12FD6
      Mask:49         =   5405713
      Key:49          =   "#sturep"
      Bmp:50          =   "admin_dash.frx":13290
      Mask:50         =   5405713
      Key:50          =   "#empe"
      Bmp:51          =   "admin_dash.frx":1354A
      Mask:51         =   5405713
      Key:51          =   "#mnuLgOt"
      Bmp:52          =   "admin_dash.frx":13804
      Mask:52         =   5405713
      Key:52          =   "#lgout"
      Bmp:53          =   "admin_dash.frx":13ABE
      Mask:53         =   16775423
      Key:53          =   "#edtorgdtls"
      Bmp:54          =   "admin_dash.frx":13E10
      Mask:54         =   16777215
      Key:54          =   "#rstallsystm"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Height          =   9960
      Left            =   0
      ScaleHeight     =   9900
      ScaleWidth      =   3195
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin VB.CommandButton Command10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Log Out"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1380
         Left            =   1620
         MouseIcon       =   "admin_dash.frx":14362
         MousePointer    =   99  'Custom
         Picture         =   "admin_dash.frx":144B4
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Log Out "
         Top             =   8320
         Width           =   1575
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Add New User"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1380
         Left            =   80
         MouseIcon       =   "admin_dash.frx":1518D
         MousePointer    =   99  'Custom
         Picture         =   "admin_dash.frx":152DF
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Change Password"
         Top             =   8320
         Width           =   1575
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Backup"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1420
         Left            =   1620
         MouseIcon       =   "admin_dash.frx":15B21
         MousePointer    =   99  'Custom
         Picture         =   "admin_dash.frx":15C73
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Take Backup Of DataBase"
         Top             =   6965
         Width           =   1575
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Login Info"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1420
         Left            =   80
         MouseIcon       =   "admin_dash.frx":16B55
         MousePointer    =   99  'Custom
         Picture         =   "admin_dash.frx":16CA7
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "See All User LogIn ID and Password"
         Top             =   6965
         Width           =   1575
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "About Organisation"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1420
         Left            =   1620
         MouseIcon       =   "admin_dash.frx":17671
         MousePointer    =   99  'Custom
         Picture         =   "admin_dash.frx":177C3
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "About Organisation"
         Top             =   5585
         Width           =   1575
      End
      Begin VB.CommandButton Command14 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Test Properties"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1420
         Left            =   80
         MouseIcon       =   "admin_dash.frx":1868D
         MousePointer    =   99  'Custom
         Picture         =   "admin_dash.frx":187DF
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Set Test Properties"
         Top             =   5585
         Width           =   1575
      End
      Begin VB.CommandButton Command8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Student Ranking"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1420
         Left            =   1620
         MouseIcon       =   "admin_dash.frx":19372
         MousePointer    =   99  'Custom
         Picture         =   "admin_dash.frx":194C4
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Show  Student ranking"
         Top             =   4190
         Width           =   1575
      End
      Begin VB.CommandButton Command7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Income-Expense"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1420
         Left            =   80
         MouseIcon       =   "admin_dash.frx":1DF46
         MousePointer    =   99  'Custom
         Picture         =   "admin_dash.frx":1E098
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "All Accounts Information"
         Top             =   4190
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Register Student"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1420
         Left            =   1620
         MouseIcon       =   "admin_dash.frx":1EF7A
         MousePointer    =   99  'Custom
         Picture         =   "admin_dash.frx":1F0CC
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Add A Student"
         Top             =   2790
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Question Paper"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1420
         Left            =   80
         MouseIcon       =   "admin_dash.frx":1FD86
         MousePointer    =   99  'Custom
         Picture         =   "admin_dash.frx":1FED8
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Create Question paper"
         Top             =   2790
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Report && Tools"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1420
         Left            =   1620
         MouseIcon       =   "admin_dash.frx":20774
         MousePointer    =   99  'Custom
         Picture         =   "admin_dash.frx":208C6
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1400
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Question Bank"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1420
         Left            =   80
         MouseIcon       =   "admin_dash.frx":23A9E
         MousePointer    =   99  'Custom
         Picture         =   "admin_dash.frx":23BF0
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Question bank"
         Top             =   1400
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Search Student"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1420
         Left            =   1620
         MouseIcon       =   "admin_dash.frx":2463D
         MousePointer    =   99  'Custom
         Picture         =   "admin_dash.frx":2478F
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Search and Print Student Record"
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Add Question"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1420
         Left            =   80
         MouseIcon       =   "admin_dash.frx":25327
         MousePointer    =   99  'Custom
         Picture         =   "admin_dash.frx":25479
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Add new Questions"
         Top             =   0
         Width           =   1575
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   9960
      Width           =   19950
      _ExtentX        =   35190
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5680
            MinWidth        =   5680
            Text            =   "SPEEDUP - TEST - SOLUTIONS"
            TextSave        =   "SPEEDUP - TEST - SOLUTIONS"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8643
            MinWidth        =   8643
            Text            =   "STATUS : "
            TextSave        =   "STATUS : "
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   3881
            MinWidth        =   3881
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   3881
            MinWidth        =   3881
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   3881
            MinWidth        =   3881
            TextSave        =   "SCRL"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4939
            MinWidth        =   4939
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   1455
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   32
      Text            =   "admin_dash.frx":25CCC
      Top             =   600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Shape shp10 
      BorderColor     =   &H80000012&
      BorderWidth     =   3
      Height          =   2280
      Left            =   16200
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Shape shp12 
      BorderColor     =   &H80000013&
      BorderWidth     =   3
      Height          =   2290
      Left            =   16200
      Shape           =   4  'Rounded Rectangle
      Top             =   6960
      Visible         =   0   'False
      Width           =   3515
   End
   Begin VB.Shape shp11 
      BorderColor     =   &H80000013&
      BorderWidth     =   3
      Height          =   2280
      Left            =   16200
      Shape           =   4  'Rounded Rectangle
      Top             =   4025
      Visible         =   0   'False
      Width           =   3525
   End
   Begin VB.Shape shp8 
      BorderColor     =   &H80000013&
      BorderWidth     =   3
      Height          =   2290
      Left            =   12000
      Shape           =   4  'Rounded Rectangle
      Top             =   3985
      Visible         =   0   'False
      Width           =   3540
   End
   Begin VB.Shape shp7 
      BorderColor     =   &H80000013&
      BorderWidth     =   3
      Height          =   2280
      Left            =   12000
      Shape           =   4  'Rounded Rectangle
      Top             =   980
      Visible         =   0   'False
      Width           =   3540
   End
   Begin VB.Shape shp9 
      BorderColor     =   &H80000013&
      BorderWidth     =   3
      Height          =   2265
      Left            =   12000
      Shape           =   4  'Rounded Rectangle
      Top             =   6960
      Visible         =   0   'False
      Width           =   3540
   End
   Begin VB.Shape Shp5 
      BorderColor     =   &H80000013&
      BorderWidth     =   3
      Height          =   2305
      Left            =   7940
      Shape           =   4  'Rounded Rectangle
      Top             =   3960
      Visible         =   0   'False
      Width           =   3555
   End
   Begin VB.Shape Shp6 
      BorderColor     =   &H80000013&
      BorderWidth     =   3
      Height          =   2265
      Left            =   7920
      Shape           =   4  'Rounded Rectangle
      Top             =   6960
      Visible         =   0   'False
      Width           =   3535
   End
   Begin VB.Shape Shp4 
      BorderColor     =   &H80000013&
      BorderWidth     =   3
      Height          =   2280
      Left            =   7935
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Visible         =   0   'False
      Width           =   3555
   End
   Begin VB.Shape shp3 
      BorderColor     =   &H8000000D&
      BorderWidth     =   3
      Height          =   2280
      Left            =   3870
      Shape           =   4  'Rounded Rectangle
      Top             =   6960
      Visible         =   0   'False
      Width           =   3435
   End
   Begin VB.Shape shp2 
      BorderColor     =   &H8000000D&
      BorderWidth     =   3
      FillColor       =   &H8000000D&
      Height          =   2280
      Left            =   3840
      Shape           =   4  'Rounded Rectangle
      Top             =   3960
      Visible         =   0   'False
      Width           =   3450
   End
   Begin VB.Shape shp1 
      BorderColor     =   &H8000000D&
      BorderWidth     =   3
      Height          =   2280
      Left            =   3870
      Shape           =   4  'Rounded Rectangle
      Top             =   970
      Visible         =   0   'False
      Width           =   3455
   End
   Begin VB.Label Sub_Dash 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "42"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5160
      TabIndex        =   31
      Top             =   5480
      Width           =   810
   End
   Begin VB.Label Course_Dash 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5280
      TabIndex        =   30
      Top             =   2480
      Width           =   585
   End
   Begin VB.Label topic_Dash 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5280
      TabIndex        =   29
      Top             =   8450
      Width           =   690
   End
   Begin VB.Image Image2 
      Height          =   2400
      Left            =   3840
      MouseIcon       =   "admin_dash.frx":25CD2
      MousePointer    =   99  'Custom
      Picture         =   "admin_dash.frx":25E24
      Top             =   6900
      Width           =   3540
   End
   Begin VB.Image Image28 
      Height          =   2385
      Left            =   3840
      MouseIcon       =   "admin_dash.frx":27C42
      MousePointer    =   99  'Custom
      Picture         =   "admin_dash.frx":27D94
      Top             =   920
      Width           =   3555
   End
   Begin VB.Image Image26 
      Height          =   2340
      Left            =   3840
      MouseIcon       =   "admin_dash.frx":29DA4
      MousePointer    =   99  'Custom
      Picture         =   "admin_dash.frx":29EF6
      Top             =   3930
      Width           =   3510
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Rs"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   18120
      TabIndex        =   28
      Top             =   2750
      Width           =   375
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Rs"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   16500
      TabIndex        =   27
      Top             =   2750
      Width           =   375
   End
   Begin VB.Label Expenselbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1100"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   18420
      TabIndex        =   26
      Top             =   2745
      Width           =   900
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Admin Dashboard"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   10320
      TabIndex        =   11
      Top             =   80
      Width           =   2895
   End
   Begin VB.Label sch_dash 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   9360
      TabIndex        =   10
      Top             =   2510
      Width           =   660
   End
   Begin VB.Label ques_dash 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "65"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   9360
      TabIndex        =   9
      Top             =   5510
      Width           =   660
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   9360
      TabIndex        =   8
      Top             =   8505
      Width           =   660
   End
   Begin VB.Label incomeLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   16845
      TabIndex        =   7
      Top             =   2745
      Width           =   780
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   17640
      TabIndex        =   6
      Top             =   5510
      Width           =   660
   End
   Begin VB.Label Package_dash 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "26"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   17820
      TabIndex        =   5
      Top             =   8505
      Width           =   300
   End
   Begin VB.Label Student_dash 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   13440
      TabIndex        =   4
      Top             =   2510
      Width           =   660
   End
   Begin VB.Label PkgReq 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8020&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   13440
      TabIndex        =   3
      Top             =   5505
      Width           =   510
   End
   Begin VB.Label Emp_dashes 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   13440
      TabIndex        =   2
      Top             =   8505
      Width           =   510
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   525
      Left            =   3240
      Top             =   0
      Width           =   17235
   End
   Begin VB.Image Image24 
      Height          =   2400
      Left            =   16155
      MouseIcon       =   "admin_dash.frx":2BC5C
      MousePointer    =   99  'Custom
      Picture         =   "admin_dash.frx":2BDAE
      Top             =   900
      Width           =   3585
   End
   Begin VB.Image Image6 
      Height          =   2415
      Left            =   16150
      MouseIcon       =   "admin_dash.frx":2F1F7
      MousePointer    =   99  'Custom
      Picture         =   "admin_dash.frx":2F349
      Top             =   3960
      Width           =   3570
   End
   Begin VB.Image Image8 
      Height          =   2400
      Left            =   16150
      MouseIcon       =   "admin_dash.frx":3141B
      MousePointer    =   99  'Custom
      Picture         =   "admin_dash.frx":3156D
      Top             =   6915
      Width           =   3600
   End
   Begin VB.Image Image22 
      Height          =   2340
      Left            =   11985
      MouseIcon       =   "admin_dash.frx":335BA
      MousePointer    =   99  'Custom
      Picture         =   "admin_dash.frx":3370C
      Top             =   6915
      Width           =   3600
   End
   Begin VB.Image Image25 
      Height          =   2355
      Left            =   11980
      MouseIcon       =   "admin_dash.frx":35F58
      MousePointer    =   99  'Custom
      Picture         =   "admin_dash.frx":360AA
      Top             =   3960
      Width           =   3570
   End
   Begin VB.Image Image29 
      Height          =   2400
      Left            =   11985
      MouseIcon       =   "admin_dash.frx":38F9E
      MousePointer    =   99  'Custom
      Picture         =   "admin_dash.frx":390F0
      Top             =   920
      Width           =   3570
   End
   Begin VB.Image Image12 
      Height          =   2385
      Left            =   7920
      MouseIcon       =   "admin_dash.frx":3B653
      MousePointer    =   99  'Custom
      Picture         =   "admin_dash.frx":3B7A5
      Top             =   6900
      Width           =   3570
   End
   Begin VB.Image Image27 
      Height          =   2415
      Left            =   7905
      MouseIcon       =   "admin_dash.frx":3E11F
      MousePointer    =   99  'Custom
      Picture         =   "admin_dash.frx":3E271
      Top             =   3915
      Width           =   3600
   End
   Begin VB.Image Image23 
      Height          =   2415
      Left            =   7920
      MouseIcon       =   "admin_dash.frx":4083F
      MousePointer    =   99  'Custom
      Picture         =   "admin_dash.frx":40991
      Top             =   920
      Width           =   3615
   End
   Begin VB.Menu Master 
      Caption         =   "&Master Entry"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu vc 
         Caption         =   "-"
      End
      Begin VB.Menu crse 
         Caption         =   "Course"
      End
      Begin VB.Menu gghhh 
         Caption         =   "-"
      End
      Begin VB.Menu sub 
         Caption         =   "Subject"
      End
      Begin VB.Menu dfg 
         Caption         =   "-"
      End
      Begin VB.Menu tpc1 
         Caption         =   "Topic"
      End
      Begin VB.Menu cvb 
         Caption         =   "-"
      End
      Begin VB.Menu pkg1 
         Caption         =   "Package"
      End
      Begin VB.Menu bn 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu questp 
         Caption         =   "Ques_type"
      End
      Begin VB.Menu asd 
         Caption         =   "-"
      End
      Begin VB.Menu schdlee 
         Caption         =   "Schedule"
      End
      Begin VB.Menu nnnnn 
         Caption         =   "-"
      End
   End
   Begin VB.Menu question 
      Caption         =   "&Questions"
      Begin VB.Menu lk5 
         Caption         =   "-"
      End
      Begin VB.Menu add 
         Caption         =   "Add"
      End
      Begin VB.Menu as 
         Caption         =   "-"
      End
      Begin VB.Menu delete 
         Caption         =   "Delete"
      End
      Begin VB.Menu gh 
         Caption         =   "-"
      End
      Begin VB.Menu view 
         Caption         =   "View"
      End
      Begin VB.Menu adfgj 
         Caption         =   "-"
      End
      Begin VB.Menu quebank 
         Caption         =   "Question Bank"
      End
      Begin VB.Menu rtl5 
         Caption         =   "-"
      End
   End
   Begin VB.Menu tstproperties 
      Caption         =   "Test Properties"
      Begin VB.Menu tstdft 
         Caption         =   "Set Test Property"
      End
   End
   Begin VB.Menu questionpaper 
      Caption         =   "Q&uestion Paper"
      Begin VB.Menu llklb 
         Caption         =   "-"
      End
      Begin VB.Menu create 
         Caption         =   "Create Question Paper"
      End
      Begin VB.Menu kv 
         Caption         =   "-"
      End
      Begin VB.Menu Viw 
         Caption         =   "View Previous Details"
      End
      Begin VB.Menu qc 
         Caption         =   "-"
      End
   End
   Begin VB.Menu employee 
      Caption         =   "&New User "
      Begin VB.Menu hjhu56 
         Caption         =   "-"
      End
      Begin VB.Menu empd 
         Caption         =   "Add New User"
      End
      Begin VB.Menu cv 
         Caption         =   "-"
      End
   End
   Begin VB.Menu student 
      Caption         =   "&Students"
      Begin VB.Menu l1 
         Caption         =   "-"
      End
      Begin VB.Menu pkgstd 
         Caption         =   "Register New Student"
      End
      Begin VB.Menu fgg123 
         Caption         =   "-"
      End
      Begin VB.Menu stud 
         Caption         =   "View All Students"
      End
      Begin VB.Menu we 
         Caption         =   "-"
      End
      Begin VB.Menu srstu 
         Caption         =   "Update Student Details"
      End
      Begin VB.Menu zx 
         Caption         =   "-"
      End
   End
   Begin VB.Menu client 
      Caption         =   "&Client Order Info"
      Begin VB.Menu nb128 
         Caption         =   "-"
      End
      Begin VB.Menu detl 
         Caption         =   "Add New Client"
      End
      Begin VB.Menu al 
         Caption         =   "-"
      End
      Begin VB.Menu payment 
         Caption         =   "Check  Order Status"
      End
      Begin VB.Menu fn 
         Caption         =   "-"
      End
      Begin VB.Menu pntcheck 
         Caption         =   "Payment Check"
      End
      Begin VB.Menu munna 
         Caption         =   "-"
      End
   End
   Begin VB.Menu account 
      Caption         =   "&Accounts"
      Begin VB.Menu bjb98 
         Caption         =   "-"
      End
      Begin VB.Menu income 
         Caption         =   "Income"
      End
      Begin VB.Menu yz 
         Caption         =   "-"
      End
      Begin VB.Menu expense 
         Caption         =   "Expense"
      End
      Begin VB.Menu qh 
         Caption         =   "-"
      End
      Begin VB.Menu overall 
         Caption         =   "Overall"
      End
      Begin VB.Menu mkk25 
         Caption         =   "-"
      End
   End
   Begin VB.Menu reports 
      Caption         =   "&Reports"
      Begin VB.Menu mm1 
         Caption         =   "-"
      End
      Begin VB.Menu mzone 
         Caption         =   "Master Zone"
         Begin VB.Menu lk12 
            Caption         =   "-"
         End
         Begin VB.Menu crs 
            Caption         =   "Course"
         End
         Begin VB.Menu a1 
            Caption         =   "-"
         End
         Begin VB.Menu subrep 
            Caption         =   "Subject"
         End
         Begin VB.Menu a2 
            Caption         =   "-"
         End
         Begin VB.Menu tpc 
            Caption         =   "Topic"
         End
         Begin VB.Menu a3 
            Caption         =   "-"
         End
         Begin VB.Menu pkgrep 
            Caption         =   "Package"
         End
         Begin VB.Menu a4 
            Caption         =   "-"
         End
         Begin VB.Menu schdl_rep 
            Caption         =   "Schedule"
         End
         Begin VB.Menu mnjk1 
            Caption         =   "-"
         End
      End
      Begin VB.Menu a7 
         Caption         =   "-"
      End
      Begin VB.Menu sturep 
         Caption         =   "Student Zone"
         Begin VB.Menu l99n 
            Caption         =   "-"
         End
         Begin VB.Menu regsturep 
            Caption         =   "Registered Student"
         End
         Begin VB.Menu l105 
            Caption         =   "-"
         End
         Begin VB.Menu uregsturep 
            Caption         =   "Unregistered Student"
         End
         Begin VB.Menu l112 
            Caption         =   "-"
         End
      End
      Begin VB.Menu kloiu 
         Caption         =   "-"
      End
      Begin VB.Menu empe 
         Caption         =   "Employee"
      End
      Begin VB.Menu aqs 
         Caption         =   "-"
      End
      Begin VB.Menu cl 
         Caption         =   "Client"
      End
      Begin VB.Menu adf 
         Caption         =   "-"
      End
      Begin VB.Menu act 
         Caption         =   "Account Zone"
         Begin VB.Menu l106 
            Caption         =   "-"
         End
         Begin VB.Menu increp 
            Caption         =   "Income"
         End
         Begin VB.Menu bgc 
            Caption         =   "-"
         End
         Begin VB.Menu exprep 
            Caption         =   "Expense"
         End
         Begin VB.Menu bvfdgt 
            Caption         =   "-"
         End
      End
      Begin VB.Menu j56 
         Caption         =   "-"
      End
   End
   Begin VB.Menu utilities 
      Caption         =   "&Utilities"
      Begin VB.Menu kj90 
         Caption         =   "-"
      End
      Begin VB.Menu calculator 
         Caption         =   "Calculator"
      End
      Begin VB.Menu au 
         Caption         =   "-"
      End
      Begin VB.Menu notepad 
         Caption         =   "Notepad"
      End
      Begin VB.Menu jpkpk 
         Caption         =   "-"
      End
      Begin VB.Menu secQu 
         Caption         =   "Security Questions"
      End
      Begin VB.Menu bnoo 
         Caption         =   "-"
      End
      Begin VB.Menu rstallsystm 
         Caption         =   "Reset All System"
      End
   End
   Begin VB.Menu abut 
      Caption         =   "About"
      Begin VB.Menu bbkivines 
         Caption         =   "-"
      End
      Begin VB.Menu abut_org 
         Caption         =   "About Organisation"
      End
      Begin VB.Menu cfcgv 
         Caption         =   "-"
      End
      Begin VB.Menu l_info 
         Caption         =   "Log In Info Table"
      End
      Begin VB.Menu bmnkn 
         Caption         =   "-"
      End
      Begin VB.Menu edtorgdtls 
         Caption         =   "Edit Organisation Details"
      End
      Begin VB.Menu jhorgsdtls 
         Caption         =   "-"
      End
   End
   Begin VB.Menu lgout 
      Caption         =   "Log Out"
      Begin VB.Menu mnuLgOt 
         Caption         =   "Log Out"
      End
   End
   Begin VB.Menu alPkalPk 
      Caption         =   "All"
      Visible         =   0   'False
      Begin VB.Menu l91 
         Caption         =   "-"
      End
      Begin VB.Menu refrshPk 
         Caption         =   "Refresh  Now"
      End
      Begin VB.Menu l92 
         Caption         =   "-"
      End
      Begin VB.Menu qupk 
         Caption         =   "Questions"
         Begin VB.Menu addPk 
            Caption         =   "Add New Question"
         End
         Begin VB.Menu delPk 
            Caption         =   "Delete Questions"
         End
         Begin VB.Menu updtPk 
            Caption         =   "Update Question"
         End
         Begin VB.Menu QBPK 
            Caption         =   "Question Bank"
         End
      End
      Begin VB.Menu mstrPK 
         Caption         =   "Master Entry"
         Begin VB.Menu ncrsPk 
            Caption         =   "New Course"
         End
         Begin VB.Menu nsubPk 
            Caption         =   "New Subject"
         End
         Begin VB.Menu ntpcPk 
            Caption         =   "New Topic"
         End
         Begin VB.Menu npkgPK 
            Caption         =   "New Package"
         End
      End
      Begin VB.Menu nReprtPK 
         Caption         =   "Reports"
      End
      Begin VB.Menu addStdPK 
         Caption         =   "Add New Student"
      End
      Begin VB.Menu adnwusrPK 
         Caption         =   "Add New User"
      End
      Begin VB.Menu ttth 
         Caption         =   "-"
      End
      Begin VB.Menu clntPKd 
         Caption         =   "Client"
         Begin VB.Menu adnwclnPK 
            Caption         =   "Add New Client"
         End
         Begin VB.Menu ordrnewPk 
            Caption         =   "Add New Order"
         End
         Begin VB.Menu pmtPK 
            Caption         =   "Check Payment Status"
         End
      End
      Begin VB.Menu l96 
         Caption         =   "-"
      End
      Begin VB.Menu bkupPk 
         Caption         =   "Backup "
      End
      Begin VB.Menu pndngReqPk 
         Caption         =   "Pending Request"
      End
      Begin VB.Menu l99 
         Caption         =   "-"
      End
      Begin VB.Menu IncInfoPk 
         Caption         =   "Income Info"
      End
      Begin VB.Menu exPkInfo 
         Caption         =   "Expense Info"
      End
      Begin VB.Menu l100 
         Caption         =   "-"
      End
      Begin VB.Menu stRnkPk 
         Caption         =   "Student Ranking"
      End
      Begin VB.Menu lginPkInfo 
         Caption         =   "Login Info"
      End
      Begin VB.Menu orgeditPkInfo 
         Caption         =   "Edit Organisation Info"
      End
   End
End
Attribute VB_Name = "admin_dash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub abut_org_Click()
Command13_Click
End Sub

Private Sub add_Click()
ques_entry_dash.Show
End Sub

Private Sub addPk_Click()
ques_entry_dash.Show
End Sub

Private Sub addStdPK_Click()
regstudnt.Show
End Sub

Private Sub adnwclnPK_Click()
FrmClient1.Show
End Sub

Private Sub adnwusrPK_Click()
FrmEmpMaster.Show
End Sub

Private Sub bkupPk_Click()
frmbackup.Show 1, MDI
End Sub

Private Sub calculator_Click()
On Error GoTo Err
    Shell "calc.exe", vbNormalFocus
    Exit Sub
Err:
    MsgBox "You don't have a Calculator installed in your computer.", vbExclamation, "Calculator Missing"
End Sub

Private Sub cl_Click()
FrmReportMain.Show
End Sub

Private Sub Command1_Click()
ques_entry_dash.Show
End Sub

Private Sub Command10_Click()
If MsgBox("   Are You Sure to LogOut ?   ", vbYesNo + vbQuestion, "LOGOUT") = vbYes Then
 log_out_Admin
End If
End Sub

Private Sub Command11_Click()
LoginINFO.Show 1, MDI
End Sub

Private Sub Command12_Click()
adminOpen = 2
frmbackup.Show 1, MDI
End Sub

Private Sub Command13_Click()
about_org.Show 1, MDI
End Sub

Private Sub Command14_Click()
FrmTestPrpt1.Show 1, MDI
End Sub


Private Sub Command2_Click()
Search_registered.Show
End Sub

Private Sub Command3_Click()
QuesBank.Show
End Sub

Private Sub Command4_Click()
Set r = New ADODB.Recordset
Set r = c.Execute("select * from backup1")
If r.EOF = False Then
If IsNull(r.Fields(0)) = False Then
 If r.Fields(0) < Date - 2 Then
  MsgBox "You Didn't Have taken backup From Last  2 days .So Take Backup Regularly."
 End If
End If
End If
FrmReportMain.Show
End Sub

Private Sub Command5_Click() 'Question Paper
QuestionPPRdashboard.Show
End Sub

Private Sub Command6_Click()
regstudnt.Show
End Sub

Private Sub Command7_Click()
FrmIncmExpense.Show
End Sub

Private Sub Command8_Click()
Stud_Ranking.Show
End Sub

Private Sub Command9_Click() 'add user
FrmEmpMaster.Show
End Sub

Private Sub create_Click() 'Paper from menu
QuestionPPRdashboard.Show
End Sub

Private Sub crs_Click()
FrmReportMain.Show
End Sub

Private Sub crse_Click()
frmCourseMaster.Show 1, MDI
End Sub

Private Sub emp_Click()

End Sub

Private Sub delete_Click()
QuesBank.Show
End Sub

Private Sub delPk_Click()
QuesBank.Show
End Sub

Private Sub detl_Click()
FrmClient1.Show
End Sub

Private Sub edtorgdtls_Click()
FrmOrganisation.Show 1, MDI
End Sub

Private Sub empd_Click()
FrmEmpMaster.Show
End Sub

Private Sub empe_Click()
FrmReportMain.Show
End Sub

Private Sub expense_Click()
FrmExpense.Show 1, MDI
End Sub

Private Sub exPkInfo_Click()
FrmExpense.Show 1, MDI
End Sub

Private Sub exprep_Click()
FrmReportMain.Show
End Sub

Sub rfrshnow()
Set r = c.Execute("select count(*) from client ")
 Label21.Caption = r.Fields(0)
Set r = c.Execute("select count(*) from course")
 Course_Dash.Caption = r.Fields(0)
Set r1 = c.Execute("select count(*) from sub")
 Sub_Dash.Caption = r1.Fields(0)
Set r = c.Execute("select count(*) from topic")
 topic_Dash.Caption = r.Fields(0)
Set r = c.Execute("select count(*) from schdl")
 sch_dash.Caption = r.Fields(0)
Set r = c.Execute("select count(*) from quesms")
 ques_dash.Caption = r.Fields(0)
Set r = c.Execute("select  count(*) from rstud")
 Student_dash.Caption = r.Fields(0)
Set r = c.Execute("select  count(*) from emp")
 Emp_dashes.Caption = r.Fields(0)
Set r = c.Execute("select count(*) from pkg")
 Package_dash.Caption = r.Fields(0)
Set r = c.Execute("select count(*) from qpaprdash")
 Label23.Caption = r.Fields(0)
Set r = c.Execute("  select sum(INC_AMT) from incm where inc_date='" & Date & "' ")
If IsNull(r.Fields(0)) = False Then
 incomeLbl.Caption = r.Fields(0)
Else
 incomeLbl.Caption = 0
End If
Set r = c.Execute(" select sum(EX_AMT) from exp where ex_date='" & Date & "' ")
If IsNull(r.Fields(0)) = False Then
 Expenselbl.Caption = r.Fields(0)
Else
 Expenselbl.Caption = 0
End If
Set r = New ADODB.Recordset
Set r = c.Execute("select count(*) from PKG_RENEW")
If r.Fields(0) = 0 Then
 PkgReq.Font.Size = 16
 PkgReq.ForeColor = vbWhite
 PkgReq.Caption = r.Fields(0)
 Timer2.Enabled = False
ElseIf r.Fields(0) > 0 Then
 PkgReq.ForeColor = vbYellow
 PkgReq.Caption = r.Fields(0)
 Timer2.Enabled = True
End If
End Sub
Private Sub Form_Activate()
DashTiMEr.Enabled = True
Me.WindowState = vbMaximized
End Sub

Private Sub Form_Load()
conn
DashTiMEr.Enabled = True
Stu_login_reg_no = ""
EMP_login_reg_no = ""
adminOpen = 1
CenterForm Me
Set r = New ADODB.Recordset
Set r = c1.Execute("select initcap(a_nm) from admintbl where a_id='" & admin_login_reg_no & "'")
If r.EOF = False Then
 StatusBar1.Panels(2).Text = "( Admin )" & " Logged In : " & r.Fields(0)
End If
StatusBar1.Panels(6).Text = " DATE :   " & Format$(Date, "dd-mm-yyyy")
Set r = c.Execute("select count(*) from course")
 Course_Dash.Caption = r.Fields(0)
Set r1 = c.Execute("select count(*) from sub")
 Sub_Dash.Caption = r1.Fields(0)
Set r = c.Execute("select count(*) from topic")
 topic_Dash.Caption = r.Fields(0)
Set r = c.Execute("select count(*) from schdl")
 sch_dash.Caption = r.Fields(0)
Set r = c.Execute("select count(*) from quesms")
 ques_dash.Caption = r.Fields(0)
Set r = c.Execute("select  count(*) from rstud")
 Student_dash.Caption = r.Fields(0)
Set r = c.Execute("select  count(*) from emp")
 Emp_dashes.Caption = r.Fields(0)
Set r = c.Execute("select count(*) from pkg")
 Package_dash.Caption = r.Fields(0)
Set r = c.Execute("select count(*) from qpaprdash")
 Label23.Caption = r.Fields(0)
Set r = c.Execute("  select sum(INC_AMT) from incm where inc_date='" & Date & "' ")
If r.Fields(0) > 0 Then
 incomeLbl.Caption = r.Fields(0)
Else
 incomeLbl.Caption = 0
End If
Set r = c.Execute("select count(*) from client ")
Label21.Caption = r.Fields(0)
'chkadmn admin_dash, login_Admin
Set r = c.Execute(" select sum(EX_AMT) from exp where ex_date='" & Date & "' ")
If IsNull(r.Fields(0)) = False Then
 Expenselbl.Caption = r.Fields(0)
Else
 Expenselbl.Caption = 0
End If
Set r = New ADODB.Recordset
Set r = c.Execute("select count(*) from PKG_RENEW")
If r.Fields(0) = 0 Then
PkgReq.Font.Size = 16
PkgReq.ForeColor = vbWhite
PkgReq.Caption = r.Fields(0)
Timer2.Enabled = False
ElseIf r.Fields(0) > 0 Then
PkgReq.ForeColor = vbYellow
PkgReq.Caption = r.Fields(0)
Timer2.Enabled = True
End If
End Sub

Private Sub Form_LostFocus()
DashTiMEr.Enabled = False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
 PopupMenu alPkalPk
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
all_hide
End Sub
Public Sub all_hide()
shp1.Visible = False
shp2.Visible = False
shp3.Visible = False
Shp4.Visible = False
Shp5.Visible = False
Shp6.Visible = False
shp7.Visible = False
shp8.Visible = False
shp9.Visible = False
shp10.Visible = False
shp11.Visible = False
shp12.Visible = False
End Sub
Private Sub Image12_Click()
QuestionPPRdashboard.Show
End Sub

Private Sub Image14_Click()
frmSubMaster.Show
End Sub

Private Sub Image12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
shp1.Visible = False
shp2.Visible = False
shp3.Visible = False
Shp4.Visible = False
Shp5.Visible = False
shp7.Visible = False
shp8.Visible = False
shp9.Visible = False
shp10.Visible = False
shp11.Visible = False
shp12.Visible = False
Shp6.Visible = True
End Sub

Private Sub Image2_Click()
FrmTopicMaster.Show
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
shp1.Visible = False
shp2.Visible = False
Shp4.Visible = False
Shp5.Visible = False
Shp6.Visible = False
shp7.Visible = False
shp8.Visible = False
shp9.Visible = False
shp10.Visible = False
shp11.Visible = False
shp12.Visible = False
shp3.Visible = True
End Sub

Private Sub Image22_Click()
FrmEmpMaster.Show
End Sub

Private Sub Image22_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
shp1.Visible = False
shp2.Visible = False
shp3.Visible = False
Shp4.Visible = False
Shp5.Visible = False
Shp6.Visible = False
shp7.Visible = False
shp8.Visible = False
shp10.Visible = False
shp11.Visible = False
shp12.Visible = False

shp9.Visible = True
End Sub

Private Sub Image23_Click()
FrmSchedule.Show 1, MDI
End Sub

Private Sub Image23_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
shp1.Visible = False
shp2.Visible = False
shp3.Visible = False
Shp5.Visible = False
Shp6.Visible = False
shp7.Visible = False
shp8.Visible = False
shp9.Visible = False
shp10.Visible = False
shp11.Visible = False
shp12.Visible = False

Shp4.Visible = True
End Sub

Private Sub Image24_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
shp1.Visible = False
shp2.Visible = False
shp3.Visible = False
Shp4.Visible = False
Shp5.Visible = False
Shp6.Visible = False
shp7.Visible = False
shp8.Visible = False
shp9.Visible = False
shp11.Visible = False
shp12.Visible = False
shp10.Visible = True
End Sub

Private Sub Image25_Click()
StuPendingReq.Show
End Sub

Private Sub Image25_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
shp1.Visible = False
shp2.Visible = False
shp3.Visible = False
Shp4.Visible = False
Shp5.Visible = False
Shp6.Visible = False
shp7.Visible = False
shp9.Visible = False
shp10.Visible = False
shp11.Visible = False
shp12.Visible = False

shp8.Visible = True
End Sub

Private Sub Image26_Click()
frmSubMaster.Show
End Sub

Private Sub Image26_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
shp1.Visible = False
shp3.Visible = False
Shp4.Visible = False
Shp5.Visible = False
Shp6.Visible = False
shp7.Visible = False
shp8.Visible = False
shp9.Visible = False
shp10.Visible = False
shp11.Visible = False
shp12.Visible = False
shp2.Visible = True
End Sub

Private Sub Image27_Click()
ques_entry_dash.Show
End Sub

Private Sub Image27_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
shp1.Visible = False
shp2.Visible = False
shp3.Visible = False
Shp4.Visible = False
Shp6.Visible = False
shp7.Visible = False
shp8.Visible = False
shp9.Visible = False
shp10.Visible = False
shp11.Visible = False
shp12.Visible = False
Shp5.Visible = True
End Sub

Private Sub Image28_Click()
frmCourseMaster.Show 1, MDI
End Sub

Private Sub Image28_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
shp2.Visible = False
shp3.Visible = False
Shp4.Visible = False
Shp5.Visible = False
Shp6.Visible = False
shp7.Visible = False
shp8.Visible = False
shp9.Visible = False
shp10.Visible = False
shp11.Visible = False
shp12.Visible = False
shp1.Visible = True
End Sub

Private Sub Image29_Click()
Search_registered.Show
End Sub

Private Sub Image29_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
shp1.Visible = False
shp2.Visible = False
shp3.Visible = False
Shp4.Visible = False
Shp5.Visible = False
Shp6.Visible = False
shp8.Visible = False
shp9.Visible = False
shp10.Visible = False
shp11.Visible = False
shp12.Visible = False

shp7.Visible = True
End Sub

Private Sub Image6_Click()
FrmClient1.Show
End Sub

Private Sub Image6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
shp1.Visible = False
shp2.Visible = False
shp3.Visible = False
Shp4.Visible = False
Shp5.Visible = False
Shp6.Visible = False
shp7.Visible = False
shp8.Visible = False
shp9.Visible = False
shp10.Visible = False
shp12.Visible = False
shp11.Visible = True
End Sub

Private Sub Image8_Click()
FrmPackage.Show 1, MDI
End Sub

Private Sub Image8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
shp1.Visible = False
shp2.Visible = False
shp3.Visible = False
Shp4.Visible = False
Shp5.Visible = False
Shp6.Visible = False
shp7.Visible = False
shp8.Visible = False
shp9.Visible = False
shp10.Visible = False
shp11.Visible = False
shp12.Visible = True
End Sub

Private Sub IncInfoPk_Click()
FrmIncome.Show 1, MDI
End Sub

Private Sub income_Click()
FrmIncome.Show 1, MDI
End Sub

Private Sub increp_Click()
FrmReportMain.Show
End Sub

Private Sub l_info_Click()
Command11_Click
End Sub

Private Sub mstu_Click()
FrmReportMain.Show
End Sub

Private Sub msturepr_Click()
FrmReportMain.Show
End Sub

Private Sub lginPkInfo_Click()
LoginINFO.Show 1, MDI
End Sub

Private Sub mnuLgOt_Click()
Command10_Click
End Sub

Private Sub ncrsPk_Click()
frmCourseMaster.Show 1, MDI
End Sub

Private Sub notepad_Click()
On Error GoTo Err
    Shell "Notepad.exe", vbNormalFocus
    Exit Sub
Err:
    MsgBox "You don't have Notepad installed in your computer.", vbExclamation, "Notepad Missing"
End Sub

Private Sub npkgPK_Click()
FrmPackage.Show 1, MDI
End Sub

Private Sub nReprtPK_Click()
FrmReportMain.Show
End Sub

Private Sub nsubPk_Click()
frmSubMaster.Show
End Sub

Private Sub ntpcPk_Click()
FrmTopicMaster.Show
End Sub

Private Sub ordrnewPk_Click()
FrmClient2.Show
End Sub

Private Sub orgeditPkInfo_Click()
about_org.Show 1, MDI
End Sub

Private Sub overall_Click()
FrmIncmExpense.Show
End Sub

Private Sub payment_Click()
FrmClient2.Show
End Sub

Private Sub pkg1_Click()
FrmPackage.Show 1, MDI
End Sub

Private Sub pkgrep_Click()
FrmReportMain.Show
End Sub

Private Sub pkgstd_Click()
regstudnt.Show
End Sub

Private Sub quesbank_Click()
ques_entry_dash.Show
End Sub

Private Sub pmtPK_Click()
FrmClient3.Show
End Sub

Private Sub pndngReqPk_Click()
StuPendingReq.Show
End Sub

Private Sub pntcheck_Click()
FrmClient3.Show 1, MDI
End Sub

Private Sub QBPK_Click()
ques_entry_dash.Show
End Sub

Private Sub quebank_Click()
QuesBank.Show
End Sub

Private Sub quespaperbank_Click()
QuestionPPRdashboard.Show
End Sub

Private Sub questp_Click()
FrmQuesType.Show 1, MDI
End Sub

Private Sub refrshPk_Click()
rfrshnow
End Sub

Private Sub rstallsystm_Click()
If MsgBox("Are You Sure To RESET the System. It Can cause you data loss , Make Sure You Have Taken Backup Before Reset System...", vbCritical + vbYesNo, "RESET SYSTEM") = vbYes Then
 Text1.Text = "sqlplus system/manager @ " & App.Path & "\database\DB2.sql"
 Shell "Cmd.exe /c " & Text1.Text
 MsgBox "System Reset Completed..", vbInformation + vbOKOnly, "System Reset"
 rfrshnow
End If
End Sub

Private Sub schdl_rep_Click()
FrmReportMain.Show
End Sub

Private Sub schdlee_Click()
FrmSchedule.Show 1, MDI
End Sub
Private Sub see_Click()
QuestionPPRdashboard.Show
End Sub

Private Sub secQu_Click()
Security_Question.Show 1, MDI
End Sub

Private Sub srstu_Click()
regstudnt.Show
End Sub

Private Sub sstu_Click()
FrmReportMain.Show
End Sub

Private Sub sstur_Click()
FrmReportMain.Show
End Sub

Private Sub stRnkPk_Click()
Stud_Ranking.Show
End Sub

Private Sub stud_Click()
Search_registered.Show
End Sub

Private Sub sub_Click()
frmSubMaster.Show
End Sub

Private Sub subrep_Click()
FrmReportMain.Show
End Sub

Private Sub Timer1_Timer()
StatusBar1.Panels(7).Text = "  Time :" & Space(6) & Format$(Time, "hh:mm:ss  AM/PM")
End Sub

Private Sub Timer2_Timer()
Static k As Integer, Max As Boolean, b As Integer
If Val(PkgReq.Caption) > 0 Then
If Max = False Then
k = k + 1
PkgReq.Font.Size = k
If k > 16 Then Max = True
Else
k = k - 1
PkgReq.Font.Size = k
If k = 13 Then Max = False: b = 5
End If
Else
PkgReq.Font.Size = 16
PkgReq.ForeColor = vbWhite
'PkgReq.Caption = Val(PkgReq.Caption) - 1
End If
End Sub

Private Sub DashTiMEr_Timer()
rfrshnow
End Sub

Private Sub tpc_Click()
FrmReportMain.Show
End Sub

Private Sub tpc1_Click()
FrmTopicMaster.Show
End Sub

Private Sub tstdft_Click()
FrmTestPrpt1.Show 1, MDI
End Sub

Private Sub updtPk_Click()
QuesBank.Show
End Sub

Private Sub view_Click()
ques_entry_dash.Show
End Sub

Private Sub Viw_Click()
QuestionPPRdashboard.Show
End Sub
