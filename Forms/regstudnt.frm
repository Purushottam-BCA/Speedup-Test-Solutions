VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{BF38D12B-22A9-4B10-B26E-019F2B5F9C22}#1.0#0"; "AniGif.ocx"
Begin VB.Form regstudnt 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   10215
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   18960
   Icon            =   "regstudnt.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10215
   ScaleWidth      =   18960
   WindowState     =   2  'Maximized
   Begin MSACAL.Calendar cldr1 
      Height          =   3015
      Left            =   8640
      TabIndex        =   72
      Top             =   6480
      Width           =   3255
      _Version        =   524288
      _ExtentX        =   5741
      _ExtentY        =   5318
      _StockProps     =   1
      BackColor       =   14737632
      Year            =   2019
      Month           =   6
      Day             =   10
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   0   'False
      ShowVerticalGrid=   -1  'True
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
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Student Type"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1455
      Left            =   7320
      TabIndex        =   47
      Top             =   1680
      Width           =   6735
      Begin VB.ComboBox Combo6 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   2520
         MouseIcon       =   "regstudnt.frx":0ECA
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   535
         Width           =   3015
      End
      Begin Project1.PictureG PictureG1 
         Height          =   405
         Left            =   5520
         Top             =   520
         Width           =   640
         _ExtentX        =   1138
         _ExtentY        =   714
         GIF             =   "regstudnt.frx":101C
         Stretch         =   2
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "( Registered = Student With Package &&  Unregistered = Student Without Package )"
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
         Height          =   195
         Left            =   120
         TabIndex        =   50
         Top             =   1180
         Width           =   5895
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student Type  :"
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
         Left            =   360
         TabIndex        =   48
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Search Student"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   4800
      TabIndex        =   38
      Top             =   120
      Width           =   9255
      Begin VB.ComboBox Combo5 
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
         Left            =   3360
         TabIndex        =   41
         Top             =   735
         Width           =   2775
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Search"
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
         Left            =   6600
         MouseIcon       =   "regstudnt.frx":35EF6
         MousePointer    =   99  'Custom
         TabIndex        =   40
         Top             =   720
         Width           =   1815
      End
      Begin VB.ComboBox Combo2 
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
         Left            =   600
         MouseIcon       =   "regstudnt.frx":36048
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search by :-"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   645
         TabIndex        =   43
         Top             =   375
         Width           =   1080
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter or Select Record :-"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3645
         TabIndex        =   42
         Top             =   375
         Width           =   2220
      End
   End
   Begin MSComDlg.CommonDialog Cd1 
      Left            =   6960
      Top             =   8280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Personal Details"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   7455
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   6855
      Begin VB.OptionButton Female 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Female"
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
         Left            =   3530
         TabIndex        =   74
         Top             =   3240
         Width           =   1095
      End
      Begin VB.OptionButton Male 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Male"
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
         Left            =   2400
         TabIndex        =   73
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2460
         MaxLength       =   35
         TabIndex        =   29
         Text            =   "Pkr.bca@gmail.com"
         Top             =   6555
         Width           =   2880
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2460
         MaxLength       =   12
         TabIndex        =   28
         Text            =   "125468954126"
         Top             =   5835
         Width           =   1800
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   2460
         MaxLength       =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Text            =   "regstudnt.frx":3619A
         Top             =   4060
         Width           =   3960
      End
      Begin VB.TextBox mobNo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2460
         MaxLength       =   10
         TabIndex        =   7
         Text            =   "9998454554"
         Top             =   2595
         Width           =   1680
      End
      Begin VB.TextBox fname 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2460
         MaxLength       =   40
         TabIndex        =   6
         Text            =   "Gorakh Nath Lal"
         Top             =   1155
         Width           =   3960
      End
      Begin VB.TextBox reg_name 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2460
         MaxLength       =   35
         TabIndex        =   5
         Text            =   "Purushottam Kumar"
         Top             =   500
         Width           =   3960
      End
      Begin MSComCtl2.DTPicker date1 
         Height          =   375
         Left            =   2400
         TabIndex        =   8
         Top             =   1800
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "regstudnt.frx":361CF
         Format          =   102498307
         CurrentDate     =   43565
      End
      Begin VB.Label Label34 
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
         Left            =   2160
         TabIndex        =   67
         Top             =   5760
         Width           =   105
      End
      Begin VB.Label Label33 
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
         Left            =   2160
         TabIndex        =   66
         Top             =   3360
         Width           =   105
      End
      Begin VB.Label Label32 
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
         Left            =   2160
         TabIndex        =   65
         Top             =   4080
         Width           =   105
      End
      Begin VB.Label Label31 
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
         Left            =   2160
         TabIndex        =   64
         Top             =   6480
         Width           =   105
      End
      Begin VB.Label Label30 
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
         Left            =   2160
         TabIndex        =   63
         Top             =   1800
         Width           =   105
      End
      Begin VB.Label Label29 
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
         Left            =   2160
         TabIndex        =   62
         Top             =   2520
         Width           =   105
      End
      Begin VB.Label Label28 
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
         Left            =   2160
         TabIndex        =   61
         Top             =   1080
         Width           =   105
      End
      Begin VB.Label Label27 
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
         Left            =   2160
         TabIndex        =   60
         Top             =   480
         Width           =   105
      End
      Begin VB.Label genderlbl 
         BackColor       =   &H0000FFFF&
         Height          =   375
         Left            =   5520
         TabIndex        =   32
         Top             =   1800
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Shape Shape5 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   405
         Index           =   13
         Left            =   2400
         Shape           =   4  'Rounded Rectangle
         Top             =   6480
         Width           =   3015
      End
      Begin VB.Shape Shape5 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   405
         Index           =   12
         Left            =   2400
         Shape           =   4  'Rounded Rectangle
         Top             =   5760
         Width           =   1935
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email ID  :"
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
         Left            =   885
         TabIndex        =   21
         Top             =   6500
         Width           =   990
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Adhar No.  :"
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
         Left            =   735
         TabIndex        =   20
         Top             =   5760
         Width           =   1140
      End
      Begin VB.Shape Shape5 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   1365
         Index           =   4
         Left            =   2400
         Shape           =   4  'Rounded Rectangle
         Top             =   4000
         Width           =   4075
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address  :"
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
         Left            =   915
         TabIndex        =   18
         Top             =   4100
         Width           =   960
      End
      Begin VB.Shape Shape5 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   405
         Index           =   1
         Left            =   2400
         Shape           =   4  'Rounded Rectangle
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Shape Shape5 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   405
         Index           =   0
         Left            =   2400
         Shape           =   4  'Rounded Rectangle
         Top             =   1080
         Width           =   4095
      End
      Begin VB.Shape Shape5 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   405
         Index           =   3
         Left            =   2400
         Shape           =   4  'Rounded Rectangle
         Top             =   420
         Width           =   4095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name  :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1125
         TabIndex        =   13
         Top             =   480
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Father Name  :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   455
         TabIndex        =   12
         Top             =   1080
         Width           =   1425
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile No.  :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   625
         TabIndex        =   11
         Top             =   2520
         Width           =   1245
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date of birth  :"
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
         Left            =   465
         TabIndex        =   10
         Top             =   1800
         Width           =   1410
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   " Gender  :"
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
         Left            =   935
         TabIndex        =   9
         Top             =   3300
         Width           =   945
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Test && Package information"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5955
      Left            =   7320
      TabIndex        =   1
      Top             =   3180
      Width           =   6735
      Begin VB.CommandButton Uploadbtn 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Browse..."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         MouseIcon       =   "regstudnt.frx":36331
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   3000
         Width           =   2055
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1980
         TabIndex        =   30
         Top             =   4360
         Width           =   1080
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2340
         TabIndex        =   27
         Text            =   "100"
         Top             =   5200
         Width           =   960
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd.MM.yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   2900
         Width           =   1920
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1980
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   25
         Text            =   "24-12-1998"
         Top             =   3555
         Width           =   1920
      End
      Begin VB.ComboBox Combo4 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   1920
         MouseIcon       =   "regstudnt.frx":36483
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2170
         Width           =   1935
      End
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   1920
         MouseIcon       =   "regstudnt.frx":365D5
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1335
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   1920
         MouseIcon       =   "regstudnt.frx":36727
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   620
         Width           =   1935
      End
      Begin VB.Label Label38 
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
         Left            =   1800
         TabIndex        =   71
         Top             =   1320
         Width           =   105
      End
      Begin VB.Label Label37 
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
         Left            =   1800
         TabIndex        =   70
         Top             =   2160
         Width           =   105
      End
      Begin VB.Label Label36 
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
         Left            =   1800
         TabIndex        =   69
         Top             =   600
         Width           =   105
      End
      Begin VB.Label Label35 
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
         Left            =   1800
         TabIndex        =   68
         Top             =   2880
         Width           =   105
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   6480
         X2              =   4440
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   6480
         X2              =   6480
         Y1              =   360
         Y2              =   2880
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "( Only JPG or GIF Images Allow )  Recommended Size : 185 x 275"
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
         Height          =   435
         Left            =   4320
         TabIndex        =   59
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Rs"
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
         Left            =   2040
         TabIndex        =   51
         Top             =   5190
         Width           =   375
      End
      Begin VB.Label Label7 
         Height          =   375
         Left            =   5040
         TabIndex        =   36
         Top             =   3960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   2535
         Left            =   4440
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label26 
         BackColor       =   &H0000FF00&
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label25 
         BackColor       =   &H0000FF00&
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   1080
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label24 
         BackColor       =   &H0000FF00&
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   1800
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Test  : "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   590
         TabIndex        =   31
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Shape Shape5 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   360
         Index           =   11
         Left            =   1920
         Shape           =   4  'Rounded Rectangle
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Shape Shape5 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   360
         Index           =   9
         Left            =   1920
         Shape           =   4  'Rounded Rectangle
         Top             =   5160
         Width           =   1455
      End
      Begin VB.Shape Shape5 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   380
         Index           =   8
         Left            =   1920
         Shape           =   4  'Rounded Rectangle
         Top             =   2850
         Width           =   2055
      End
      Begin VB.Shape Shape5 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   405
         Index           =   7
         Left            =   1920
         Shape           =   4  'Rounded Rectangle
         Top             =   3480
         Width           =   2055
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount  :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   815
         TabIndex        =   24
         Top             =   5160
         Width           =   930
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "End date   :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   650
         TabIndex        =   23
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start date  :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   615
         TabIndex        =   22
         Top             =   2880
         Width           =   1125
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Schedule  :"
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
         Left            =   690
         TabIndex        =   15
         Top             =   2160
         Width           =   1050
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Package  :"
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
         Left            =   750
         TabIndex        =   14
         Top             =   1320
         Width           =   990
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Course   :"
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
         Left            =   820
         TabIndex        =   3
         Top             =   600
         Width           =   915
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   9240
      Width           =   13935
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   12360
         MouseIcon       =   "regstudnt.frx":36879
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   8400
         MouseIcon       =   "regstudnt.frx":369CB
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Show_btn 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Show All Students"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   10200
         MouseIcon       =   "regstudnt.frx":36B1D
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton dl_btn 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Remove"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   6240
         MouseIcon       =   "regstudnt.frx":36C6F
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton update_btn 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   4080
         MouseIcon       =   "regstudnt.frx":36DC1
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton sv_btn 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   2040
         MouseIcon       =   "regstudnt.frx":36F13
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton add_btn 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Add new"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   0
         MouseIcon       =   "regstudnt.frx":37065
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   240
         Width           =   1695
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H80000007&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000006&
         Height          =   360
         Left            =   12480
         Top             =   360
         Width           =   1410
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000007&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000006&
         Height          =   360
         Left            =   120
         Top             =   360
         Width           =   1645
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H80000007&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000006&
         Height          =   360
         Left            =   8520
         Top             =   360
         Width           =   1410
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H80000007&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000006&
         Height          =   360
         Left            =   4200
         Top             =   360
         Width           =   1770
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H80000007&
         BackStyle       =   1  'Opaque
         Height          =   360
         Left            =   2160
         Top             =   360
         Width           =   1650
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H80000007&
         BackStyle       =   1  'Opaque
         Height          =   360
         Left            =   6360
         Top             =   360
         Width           =   1785
      End
      Begin VB.Shape Shape10 
         BackColor       =   &H80000007&
         BackStyle       =   1  'Opaque
         Height          =   360
         Left            =   10320
         Top             =   360
         Width           =   1875
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Registration No"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   240
      TabIndex        =   44
      Top             =   120
      Width           =   4335
      Begin VB.TextBox reg_no 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   290
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   45
         Text            =   "RS0001"
         Top             =   620
         Width           =   1320
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registration No  :"
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
         Left            =   240
         TabIndex        =   46
         Top             =   620
         Width           =   1695
      End
      Begin VB.Shape Shape5 
         BackStyle       =   1  'Opaque
         Height          =   405
         Index           =   6
         Left            =   2520
         Shape           =   4  'Rounded Rectangle
         Top             =   550
         Width           =   1455
      End
   End
End
Attribute VB_Name = "regstudnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pic_name As String, pic_ext As String, pic_changed As Boolean
Dim course_id As String, package As String, timetbl As String

Private Sub add_btn_Click()
cauto_id
reg_name.Text = blank
fname.Text = blank
mobNo.Text = blank
Text3.Text = blank
Text8.Text = blank
Text9.Text = blank
Text5.Text = blank
Text4.Text = blank
Text7.Text = blank
Text6.Text = blank
Label26.Caption = ""
Label25.Caption = ""
Label24.Caption = ""
genderlbl.Caption = ""
pic_name = ""
Image1.Picture = Nothing
add_btn.Enabled = False
sv_btn.Enabled = True
End Sub
Public Function cauto_id()
Set r1 = New ADODB.Recordset
sql = "select max(to_number(substr(rstud_reg_no,3,length(rstud_reg_no))))from rstud"
Set r1 = c1.Execute(sql)
If IsNull(r1.Fields(0)) Then
reg_no.Text = "RS000" & 1
Else
t = r1.Fields(0)
If t > 0 And t < 9 Then
reg_no.Text = "RS000" & (t + 1)
ElseIf t < 99 Then
reg_no.Text = "RS00" & (t + 1)
ElseIf t < 999 Then
reg_no.Text = "RS0" & (t + 1)
ElseIf t < 9999 Then
reg_no.Text = "RS" & (t + 1)
End If
End If
End Function

Private Sub cldr1_Click()
Dim sdat As Date
If cldr1.Value < Date Then
 MsgBox "You cannot Select the date that " & vbCrLf & "is already passed away." & vbCrLf & "Please select valid date", vbExclamation + vbOKOnly, "Invalid date"
 Text5.SetFocus
Else
 Text5.Text = cldr1.Day & "-" & cldr1.Month & "-" & cldr1.Year
 sdat = Format(Text5.Text, "dd-mm-yyyy")
 If Combo6.ListIndex = 0 Then
   If Trim(Label7.Caption) = "" Then
   MsgBox "Select Package First", vbInformation + vbOKOnly, "Package"
   Combo3.SetFocus
   Exit Sub
   End If
   Text4.Text = Format(sdat + Val(Label7.Caption), "DD-MM-YYYY")
 Else
  Text4.Text = Format(sdat + 7, "DD-MM-YYYY")
 End If
  cldr1.Visible = False
 End If
End Sub

Private Sub Combo6_Click()
Frame2.Enabled = True
If Combo6.ListIndex = 0 Then 'Registered
 Combo3.Enabled = True
 Combo4.Enabled = True
 Text6.Locked = True
 Text7.Locked = True
 Text4.Locked = True
Else
 Combo3.Enabled = False
 Text4.Locked = True
 Text7.Locked = False
 Text6.Locked = False
End If
End Sub

Private Sub Command2_Click() 'Exit Button
Unload Me
End Sub
Private Sub date1_Click()
Male.SetFocus
End Sub

Private Sub Female_Click()
genderlbl.Caption = "FEMALE"
 Combo6.SetFocus
End Sub

Private Sub Form_Activate()
Me.WindowState = vbMaximized
End Sub

Private Sub Male_Click()
 genderlbl.Caption = "MALE"
 Combo6.SetFocus
End Sub

Private Sub Show_btn_Click() 'Student Show All
Search_registered.Show
End Sub

Private Sub Command7_Click()
Form_Load
End Sub

Private Sub Text5_GotFocus()
Text5.Text = ""
Text4.Text = ""
cldr1.Visible = True
End Sub

Private Sub Text5_LostFocus()
cldr1.Visible = False
End Sub

Private Sub Text9_LostFocus()
Dim domain As String
If Len(Trim(Text9.Text)) <> 0 Then
 If Len(Trim(Text9.Text)) <= 12 Then
 MsgBox "Invalid Email, Too Short Email", vbCritical + vbOKOnly, "Email"
Text9.SetFocus
Exit Sub
End If
If InStr(Text9.Text, "@") = False Then
 MsgBox "Invalid Email, It Must contain @..", vbCritical + vbOKOnly, "Email"
 Text9.SetFocus
Exit Sub
End If
domain = Right(Text9.Text, 4)
If UCase(domain) = UCase(".COM") Or UCase(domain) = UCase(".NET") Then
Exit Sub
Else
 MsgBox "Invalid Email", vbCritical + vbOKOnly, "Email"
 Text9.SetFocus
 Exit Sub
 End If
domain = Right(Text9.Text, 3)
If UCase(domain) = UCase(".TK") Or UCase(domain) = UCase(".IN") Then
Exit Sub
Else
 MsgBox "Invalid Email", vbCritical + vbOKOnly, "Email"
 Text9.SetFocus
 Exit Sub
End If
 End If

End Sub

Private Sub uploadbtn_Click()
Dim filefilter As String
filefilter = "JPEG Image (*.jpg)|*.jpg|All Files (*.*)|*.*"
On Error Resume Next
Cd1.Filter = filefilter
Cd1.ShowOpen
If Cd1.FileName <> "" Then
pic_ext = Right(Cd1.FileName, 3)
If UCase(Trim(pic_ext)) = "GIF" Or UCase(Trim(pic_ext)) = "JPG" Then
pic_changed = True
Else
 MsgBox "Invalid Image !! Please Select JPG Image Only", vbCritical + vbOKOnly, "Image"
 pic_name = ""
 Exit Sub
End If
Shell "cmd.exe /c del " & pic_name
Image1.Picture = LoadPicture(Cd1.FileName)
Shell "cmd.exe /c copy " & Cd1.FileName & " C:\STS\Student_Pic\" & Cd1.FileTitle
'pic_name = Cd1.FileName
pic_name = "C:\STS\Student_Pic\" & Cd1.FileTitle
Else
Exit Sub
End If
End Sub


Private Sub Combo1_Click()
On Error Resume Next
Label7.Caption = "" 'For Total days of package
Set r = New ADODB.Recordset
Set r = c.Execute("select c_id from course where c_nm='" & Combo1.Text & "' ")
Label26.Caption = r.Fields(0)

Set r1 = New ADODB.Recordset
sql = "select pkg_id,pkg_nm from pkg where c_id='" & Label26.Caption & "' "
Set r1 = c1.Execute(sql)
Combo3.Clear
While r1.EOF = False
 Combo3.AddItem r1.Fields(1)
 r1.MoveNext
Wend
r1.Close

Combo4.Clear
Set r1 = New ADODB.Recordset
sql = "select SCH_TIMING from schdl where c_id='" & Label26.Caption & "' "
Set r1 = c1.Execute(sql)
While r1.EOF = False
Combo4.AddItem r1.Fields(0)
r1.MoveNext
Wend
End Sub

Private Sub combo2_Click()
Combo5.Clear
If Combo2.ListIndex = 0 Then
Set r1 = New ADODB.Recordset
sql = "select distinct (rstud_reg_no) from rstud order by rstud_reg_no "
Set r1 = c1.Execute(sql)
While r1.EOF = False
Combo5.AddItem r1.Fields(0)
r1.MoveNext
Wend
ElseIf Combo2.ListIndex = 1 Then
Set r1 = New ADODB.Recordset
sql = "select rstud_nm from rstud order by rstud_reg_no "
Set r1 = c1.Execute(sql)
While r1.EOF = False
Combo5.AddItem r1.Fields(0)
r1.MoveNext
Wend
End If
End Sub

Private Sub combo3_Click()
On Error Resume Next
Set r1 = New ADODB.Recordset
Set r1 = c1.Execute("select pkg_id,pkg_fee,pkg_all_tst,PKG_DUR from pkg where pkg_nm='" & Combo3.Text & "' and c_id='" & Label26.Caption & "'")
Label25.Caption = r1.Fields(0)
Text6.Text = r1.Fields(1)
Text7.Text = r1.Fields(2)
Label7.Caption = r1.Fields(3)
End Sub

Private Sub Combo4_Click()
On Error Resume Next
Set r = New ADODB.Recordset
Set r = c.Execute("select sch_id from schdl where SCH_TIMING ='" & Combo4.Text & "' and c_id='" & Label26.Caption & "' ")
If r.EOF = False Then
 Label24.Caption = r.Fields(0)
End If
End Sub
Private Sub Command1_Click()
conn
If Combo2.Text = "" Or Trim(Combo5.Text) = "" Then
 MsgBox "Invalid Searching,Type Some Value to search", vbQuestion + vbOKOnly, "Empty Search"
Combo2.SetFocus
Exit Sub
End If

On Error GoTo kp:
Set r4 = New ADODB.Recordset
sql = " select * from rstud where upper(rstud_reg_no)='" & UCase(Trim(Combo5.Text)) & "' or upper(rstud_nm)='" & UCase(Trim(Combo5.Text)) & "' "
Set r4 = c.Execute(sql)
If IsNull(r4.Fields(0)) = False Then
 reg_no.Text = r4.Fields(0)
 reg_name.Text = r4.Fields(1)
 fname.Text = r4.Fields(2)
 date1.Value = r4.Fields(3)
 mobNo.Text = r4.Fields(4)
 genderlbl.Caption = r4.Fields(5)
 Text3.Text = r4.Fields(6)
 Text8.Text = r4.Fields(7)
 Text9.Text = r4.Fields(8)
Combo6.Text = r4.Fields(9)
Label26.Caption = r4.Fields(10)
If IsNull(r4.Fields(11)) = False Then
 Label25.Caption = r4.Fields(11)
Else
 Label25.Caption = ""
End If
If IsNull(r4.Fields(12)) = False Then
 Label24.Caption = r4.Fields(12)
Else
 Label24.Caption = ""
End If
Text5.Text = r4.Fields(13)
Text4.Text = r4.Fields(14)
Text7.Text = r4.Fields(15)
Text6.Text = r4.Fields(16)
If IsNull(r4.Fields(17)) = False Then
pic_name = r4.Fields(17)
Else
pic_name = App.Path & "\Graphics\Main_Screen_Icon\PicNotAvail.jpg"
End If
If UCase(genderlbl.Caption) = "MALE" Then
Male.Value = vbChecked
Else
Female.Value = vbChecked
End If
Set r = New ADODB.Recordset
Set r = c.Execute("select c_nm from course where c_id='" & Label26.Caption & "' ")
Combo1.Text = r.Fields(0)
If Label25.Caption <> "" Then
Set r = New ADODB.Recordset
Set r = c.Execute("select pkg_nm from pkg where c_id='" & Label26.Caption & "' and pkg_id='" & Label25.Caption & "' ")
Combo3.Text = r.Fields(0)
End If
If Label24.Caption <> "" Then
Set r = New ADODB.Recordset
Set r = c.Execute("select SCH_TIMING from schdl where c_id='" & Label26.Caption & "' and sch_id='" & Label24.Caption & "' ")
Combo4.Text = r.Fields(0)
End If
If pic_name <> "" Then
Image1.Picture = LoadPicture(pic_name)
End If
add_btn.Enabled = False
sv_btn.Enabled = False
dl_btn.Enabled = True
update_btn.Enabled = True
Exit Sub
 Else
 MsgBox "record Not Found", vbCritical + vbOKOnly, "Record Not Found"
 Exit Sub
End If
kp:
 Image1.Picture = LoadPicture(App.Path & "\Graphics\Main_Screen_Icon\PicNotAvail.jpg")
End Sub

Private Sub dl_btn_Click()
If Trim(Combo2.Text) = "" Or Trim(Combo5.Text) = "" Then
 MsgBox "First Search the record", vbExclamation + vbOKOnly, "No Record"
 Combo2.SetFocus
 Exit Sub
End If
If MsgBox("Are you sure To Remove The Record from database ???", vbYesNo + vbQuestion, "Delete Preview") = vbYes Then
sql = " delete from rstud where rstud_reg_no='" & reg_no.Text & "'"
c1.Execute (sql)
MsgBox "Record SuccessFully deleted", vbInformation + vbOKOnly, "Deleted"
Form_Load
End If
End Sub

Private Sub fname_KeyPress(KeyAscii As Integer)
If (((KeyAscii >= 65) And (KeyAscii <= 90)) Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or (KeyAscii = 8) Or (KeyAscii = 32)) Then
        fname.SetFocus
    ElseIf KeyAscii = 13 Then
     KeyAscii = 0
     mobNo.SetFocus
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
conn
Frame2.Enabled = False
Label7.Caption = ""
cldr1.Visible = False
cldr1.Value = Format(Date, "DD-MMM-YY")
Combo6.Clear
Combo6.AddItem "Registered"
Combo6.AddItem "UnRegistered"

date1.MaxDate = Date - (15 * 365)
date1.MinDate = Date - (40 * 365)
date1.Value = date1.MaxDate
Combo2.Clear
Combo2.AddItem "Reg. No "
Combo2.AddItem "Name "
add_btn.Enabled = True
pic_name = ""
stuPicPath = ""
Text4.Locked = True
Text6.Locked = True
Text7.Locked = True
Male.Value = vbUnchecked
Female.Value = vbUnchecked
reg_no.Text = blank
Label24.Caption = ""
Label25.Caption = ""
Label26.Caption = ""
reg_name.Text = blank
fname.Text = blank
mobNo.Text = blank
Text3.Text = blank
Text8.Text = blank
Text9.Text = blank
Text5.Text = blank
Text4.Text = blank
Text7.Text = blank
Text6.Text = blank
Combo1.Clear
Combo5.Clear
Combo3.Clear
Combo4.Clear
Image1.Picture = Nothing
Set r1 = New ADODB.Recordset
sql = "select distinct (c_nm) from course"
Set r1 = c1.Execute(sql)
While r1.EOF = False
 Combo1.AddItem r1.Fields(0)
 r1.MoveNext
Wend
sv_btn.Enabled = False
dl_btn.Enabled = False
update_btn.Enabled = False
date1.MinDate = Date - (60 * 365)
date1.MaxDate = Date - (18 * 365)
End Sub

Private Sub mobNo_KeyPress(KeyAscii As Integer)
If Len(Trim(mobNo.Text)) = 0 Then
If KeyAscii >= 48 And KeyAscii <= 53 Then
 MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
KeyAscii = 0
mobNo.SetFocus
Exit Sub
End If
End If
If Len(Trim(mobNo.Text)) = 1 Then
 If mobNo.Text = 6 Then
  If KeyAscii = 50 Or KeyAscii = 8 Then
  Else
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   mobNo.SetFocus
  Exit Sub
 End If
End If
End If
If Len(Trim(mobNo.Text)) = 6 Then
 If Right(mobNo.Text, 4) = "0000" Then
  If KeyAscii = 48 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   mobNo.SetFocus
  Exit Sub
 End If
End If
End If
If Len(Trim(mobNo.Text)) = 7 Then
 If Right(mobNo.Text, 5) = "00000" Then
  If KeyAscii = 48 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   mobNo.SetFocus
  Exit Sub
 End If
End If
If Right(mobNo.Text, 5) = "11111" Then
  If KeyAscii = 49 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   mobNo.SetFocus
  Exit Sub
 End If
End If
If Right(mobNo.Text, 5) = "22222" Then
  If KeyAscii = 50 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   mobNo.SetFocus
  Exit Sub
 End If
End If
If Right(mobNo.Text, 5) = "55555" Then
  If KeyAscii = 53 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   mobNo.SetFocus
  Exit Sub
 End If
End If
If Right(mobNo.Text, 5) = "66666" Then
  If KeyAscii = 54 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   mobNo.SetFocus
  Exit Sub
 End If
End If
If Right(mobNo.Text, 5) = "77777" Then
  If KeyAscii = 55 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   mobNo.SetFocus
  Exit Sub
 End If
End If
If Right(mobNo.Text, 5) = "88888" Then
  If KeyAscii = 56 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   mobNo.SetFocus
  Exit Sub
 End If
End If
If Right(mobNo.Text, 5) = "99999" Then
  If KeyAscii = 57 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   mobNo.SetFocus
  Exit Sub
 End If
End If

End If
If Len(Trim(mobNo.Text)) = 8 Then
 If Right(mobNo.Text, 6) = "000000" Then
  If KeyAscii = 48 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   mobNo.SetFocus
  Exit Sub
 End If
End If
 If Right(mobNo.Text, 6) = "111111" Then
  If KeyAscii = 48 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   mobNo.SetFocus
  Exit Sub
 End If
End If
End If

If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Then
        mobNo.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   Text3.SetFocus
  Else
   KeyAscii = 0
  End If

End Sub

Private Sub mobNo_LostFocus()
 If (mobNo.Text <> "") Then
        If (Len(mobNo.Text) < 10) Then
            MsgBox "Invalid MOBILE NUMBER", vbExclamation + vbOKOnly, "Wrong Mobile No"
            mobNo.Text = ""
            mobNo.SetFocus
        End If
    End If
End Sub

Private Sub reg_name_KeyPress(KeyAscii As Integer)
    If (((KeyAscii >= 65) And (KeyAscii <= 90)) Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or (KeyAscii = 8) Or (KeyAscii = 32)) Then
        reg_name.SetFocus
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        fname.SetFocus
    Else
 KeyAscii = 0
    End If
End Sub

Private Sub sv_btn_Click()
If Trim(reg_no.Text) = "" Then
MsgBox "Student Reg No Blank", vbCritical + vbOKOnly, "Warning"
Exit Sub
ElseIf Trim(reg_name.Text) = "" Then
MsgBox "Student Name Blank", vbCritical + vbOKOnly, "Warning"
reg_name.SetFocus
Exit Sub
ElseIf Trim(fname.Text) = "" Then
MsgBox "Student father Name Blank", vbCritical + vbOKOnly, "Warning"
fname.SetFocus
Exit Sub
ElseIf Trim(Text3.Text) = "" Then
 MsgBox "Enter Student Address", vbCritical + vbOKOnly, "Warning"
Text3.SetFocus
Exit Sub
ElseIf Trim(mobNo.Text) = "" Then
 MsgBox "Enter Mobile No", vbCritical + vbOKOnly, "Warning"
mobNo.SetFocus
Exit Sub
ElseIf Trim(genderlbl.Caption) = "" Then
 MsgBox "Gender field cann't be blank", vbCritical + vbOKOnly, "Warning"
Male.SetFocus
Exit Sub
ElseIf Trim(Text8.Text) = "" Then
 MsgBox "Adhar No is Mandatory", vbCritical + vbOKOnly, "Warning"
Text8.SetFocus
Exit Sub
ElseIf Trim(Text9.Text) = "" Then
 MsgBox "Enter Email ID, If Not Then Use Demo@gmail.com ", vbCritical + vbOKOnly, "Warning"
Text9.SetFocus
Exit Sub
ElseIf Trim(Combo6.Text) = "" Then
MsgBox "Enter Student Type (Package / Without package)", vbCritical + vbOKOnly, "Warning"
Combo6.SetFocus
Exit Sub
ElseIf Trim(Combo1.Text) = "" Then
MsgBox "Select Correspondent Course for The Student", vbCritical + vbOKOnly, "Warning"
Combo1.SetFocus
Exit Sub
End If
If Combo6.ListIndex = 0 Then
 If Trim(Combo3.Text) = "" Then
  MsgBox "Select package For student", vbCritical + vbOKOnly, "Warning"
  Combo3.SetFocus
 Exit Sub
 End If
 If pic_name = "" Then
  If MsgBox("Student Photo is Required But If Photo Is Not Available Then Click On Yes To Select Default Image, You Can Upload Image Later. Do You want To Select Default Image ? ", vbCritical + vbYesNo, "No Image") = vbYes Then
  pic_name = App.Path & "\Graphics\#\PicNotAvail.jpg"
  Image1.Picture = LoadPicture(pic_name)
 Else
  Uploadbtn.SetFocus
 Exit Sub
 End If
End If
End If
If Trim(Combo4.Text) = "" Then
  MsgBox "Select Batch Time For student..", vbCritical + vbOKOnly, "Warning"
  Combo4.SetFocus
 Exit Sub
 End If
If Trim(Text5.Text) = "" Then
  MsgBox "Enter start date..", vbCritical + vbOKOnly, "Warning"
  Text5.SetFocus
 Exit Sub
ElseIf Trim(Text7.Text) = "" Then
  MsgBox "How many Test ??", vbCritical + vbOKOnly, "Warning"
  Text7.SetFocus
 Exit Sub
ElseIf Trim(Text6.Text) = "" Then
  MsgBox "Enter Amount For Test..", vbCritical + vbOKOnly, "Warning"
  Text6.SetFocus
 Exit Sub
 End If
'Inserting here
stuname = reg_name.Text
stufather = fname.Text
stuCourse = Combo1.Text
stuBatch = Combo4.Text
stuIddate = Text5.Text
stuPicPath = pic_name
If Combo6.ListIndex = 0 Then
  sql = " insert into rstud values ('" & reg_no.Text & "','" & reg_name.Text & "','" & fname.Text & "','" & Format(date1, "dd/mmm/yyyy") & "'," & mobNo.Text & ",'" & genderlbl.Caption & "','" & Text3.Text & "','" & Text8.Text & "','" & Text9.Text & "','" & Combo6.Text & "','" & Label26.Caption & "','" & Label25.Caption & "','" & Label24.Caption & "','" & Format(Text5.Text, "dd/mmm/yyyy") & "','" & Format(Text4.Text, "dd/mmm/yyyy") & "'," & Text7.Text & "," & Text6.Text & ",'" & pic_name & "'," & Text7.Text & ")"
Else
 If pic_name = "" Then
  sql = " insert into rstud values ('" & reg_no.Text & "','" & reg_name.Text & "','" & fname.Text & "','" & Format(date1, "dd/mmm/yyyy") & "'," & mobNo.Text & ",'" & genderlbl.Caption & "','" & Text3.Text & "','" & Text8.Text & "','" & Text9.Text & "','" & Combo6.Text & "','" & Label26.Caption & "',NULL,'" & Label24.Caption & "','" & Format(Text5.Text, "dd/mmm/yyyy") & "','" & Format(Text4.Text, "dd/mmm/yyyy") & "'," & Text7.Text & "," & Text6.Text & ",NULL," & Text7.Text & ")"
 Else
  sql = " insert into rstud values ('" & reg_no.Text & "','" & reg_name.Text & "','" & fname.Text & "','" & Format(date1, "dd/mmm/yyyy") & "'," & mobNo.Text & ",'" & genderlbl.Caption & "','" & Text3.Text & "','" & Text8.Text & "','" & Text9.Text & "','" & Combo6.Text & "','" & Label26.Caption & "',NULL,'" & Label24.Caption & "','" & Format(Text5.Text, "dd/mmm/yyyy") & "','" & Format(Text4.Text, "dd/mmm/yyyy") & "'," & Text7.Text & "," & Text6.Text & ",'" & pic_name & "'," & Text7.Text & ")"
 End If
End If
c1.Execute (sql)
'Inserting into Account
Dim statement As String
Set r = New ADODB.Recordset
statement = reg_name.Text & " Has Enrolled in Course " & Combo1.Text & " As a " & Combo6.Text & " Student "
Set r = c.Execute("select count(*) from incm")
c.Execute ("insert into incm values (" & r.Fields(0) + 1 & ",'" & reg_name.Text & "','" & statement & "'," & Val(Text6.Text) & ",'" & Format(Date, "dd-mmm-yyyy") & "' )")
stud_id_pass.id.Caption = reg_no.Text
stud_id_pass.log_id.Caption = UCase(Trim(Mid$(reg_name.Text, 1, 3))) & Trim(Mid$(reg_no.Text, 4, 4))
stud_id_pass.Password.Caption = date1.Year & UCase(Trim(Mid$(Combo1.Text, 1, 3))) & UCase(Trim(Mid$(reg_name.Text, 1, 3)))
stud_id_pass.Show vbModal, MDI
Label7.Caption = ""
cldr1.Visible = False

Combo6.Clear
Combo6.AddItem "Registered"
Combo6.AddItem "UnRegistered"
Frame2.Enabled = False
date1.MaxDate = Date - (15 * 365)
date1.MinDate = Date - (40 * 365)
date1.Value = date1.MaxDate
Combo2.Clear
Combo2.AddItem "Reg. No "
Combo2.AddItem "Name "
add_btn.Enabled = True
pic_name = ""
stuPicPath = ""
Text4.Locked = True
Text6.Locked = True
Text7.Locked = True
Male.Value = vbUnchecked
Female.Value = vbUnchecked
reg_no.Text = blank
Label24.Caption = ""
Label25.Caption = ""
Label26.Caption = ""
reg_name.Text = blank
fname.Text = blank
mobNo.Text = blank
Text3.Text = blank
Text8.Text = blank
Text9.Text = blank
Text5.Text = blank
Text4.Text = blank
Text7.Text = blank
Text6.Text = blank
Combo1.Clear
Combo5.Clear
Combo3.Clear
Combo4.Clear
Image1.Picture = Nothing
Set r1 = New ADODB.Recordset
sql = "select distinct (c_nm) from course"
Set r1 = c1.Execute(sql)
While r1.EOF = False
 Combo1.AddItem r1.Fields(0)
 r1.MoveNext
Wend
sv_btn.Enabled = False
dl_btn.Enabled = False
update_btn.Enabled = False
date1.MinDate = Date - (60 * 365)
date1.MaxDate = Date - (18 * 365)
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If (((KeyAscii >= 65) And (KeyAscii <= 90)) Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or (KeyAscii = 8) Or (KeyAscii = 32) Or (KeyAscii >= 48 And keyaascii <= 57) Or KeyAscii = 44 Or KeyAscii = 46) Then
       Text3.SetFocus
    ElseIf KeyAscii = 13 Then
       KeyAscii = 0
       Text8.SetFocus
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
 If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Then
        Text8.SetFocus
    ElseIf KeyAscii = 13 Then
       KeyAscii = 0
       Text9.SetFocus
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Text8_LostFocus()
  If (Text8.Text <> "") Then
        If (Len(Text8.Text) < 12) Then
            MsgBox "Adhar Number must be of 12 digits !", vbQuestion + vbOKOnly, "Error Adhar"
            Text8.Text = ""
            Text8.SetFocus
            Exit Sub
        End If
        If Len(Text8.Text) = 12 Then
         If Val(Text8.Text) = 0 Then
             MsgBox "Invalid Adhar card no !!", vbInformation + vbOKOnly, "Invalid Adhar"
             Text8.SetFocus
             Exit Sub
         End If
         If Val(Left(Text8.Text, 4)) = 0 Or Val(Mid(Text8.Text, 4, 4)) = 0 Or Val(Mid(Text8.Text, 8, 4)) = 0 Or Val(Right(Text8.Text, 4)) = 0 Then
         MsgBox "Invalid Adhar card no !!", vbInformation + vbOKOnly, "Invalid Adhar"
             Text8.SetFocus
             Exit Sub
         End If
         Set r = c.Execute("select  RSTUD_ADHR from rstud")
         While r.EOF = False
          If Trim(Text8.Text) = r.Fields(0) Then
           MsgBox "Adhar no Already Exist!!", vbInformation + vbOKOnly, "Duplicate Adhar"
           Text8.SetFocus
           Exit Sub
          End If
         r.MoveNext
         Wend
    End If
End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer) 'Email
If Len(Trim(Text9.Text)) = 0 Then
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 32 Then
Else
 MsgBox "Email Id Must start With Character!!", vbInformation + vbOKOnly, "Email"
 KeyAscii = 0
 Text9.SetFocus
Exit Sub
End If
End If
If InStr(Text9.Text, "@") = False Then
 If KeyAscii = 95 Or KeyAscii = 46 Or KeyAscii = 64 Or (((KeyAscii >= 65) And (KeyAscii <= 90)) Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or (KeyAscii = 8) Or (KeyAscii = 32)) Or (KeyAscii >= 48 And KeyAscii <= 57) Then
   Text9.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   Combo1.SetFocus
  Else
   KeyAscii = 0
  End If
Else
  If KeyAscii = 95 Or KeyAscii = 46 Or (((KeyAscii >= 65) And (KeyAscii <= 90)) Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or (KeyAscii = 8) Or (KeyAscii = 32)) Or (KeyAscii >= 48 And KeyAscii <= 57) Then
   Text9.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   Combo6.SetFocus
  Else
   KeyAscii = 0
  End If
End If
End Sub

Private Sub update_btn_Click()
If Trim(reg_no.Text) = "" Then
MsgBox "Student Reg No Blank", vbCritical + vbOKOnly, "Warning"
Exit Sub
ElseIf Trim(reg_name.Text) = "" Then
MsgBox "Student Name Blank", vbCritical + vbOKOnly, "Warning"
reg_name.SetFocus
Exit Sub
ElseIf Trim(fname.Text) = "" Then
MsgBox "Student father Name Blank", vbCritical + vbOKOnly, "Warning"
fname.SetFocus
Exit Sub
ElseIf Trim(Text3.Text) = "" Then
 MsgBox "Enter Student Address", vbCritical + vbOKOnly, "Warning"
Text3.SetFocus
Exit Sub
ElseIf Trim(mobNo.Text) = "" Then
 MsgBox "Enter Mobile No", vbCritical + vbOKOnly, "Warning"
mobNo.SetFocus
Exit Sub
ElseIf Trim(genderlbl.Caption) = "" Then
 MsgBox "Gender field cann't be blank", vbCritical + vbOKOnly, "Warning"
Male.SetFocus
Exit Sub
ElseIf Trim(Text8.Text) = "" Then
 MsgBox "Adhar No is Mandatory", vbCritical + vbOKOnly, "Warning"
Text8.SetFocus
Exit Sub
ElseIf Trim(Text9.Text) = "" Then
 MsgBox "Enter Email ID, If Not Then Use Demo@gmail.com ", vbCritical + vbOKOnly, "Warning"
Text9.SetFocus
Exit Sub
ElseIf Trim(Combo6.Text) = "" Then
MsgBox "Enter Student Type (Package / Without package)", vbCritical + vbOKOnly, "Warning"
Combo6.SetFocus
Exit Sub
ElseIf Trim(Combo1.Text) = "" Then
MsgBox "Select Correspondent Course for The Student", vbCritical + vbOKOnly, "Warning"
Combo1.SetFocus
Exit Sub
End If
If Combo6.ListIndex = 0 Then
 If Trim(Combo3.Text) = "" Then
  MsgBox "Select package For student", vbCritical + vbOKOnly, "Warning"
  Combo3.SetFocus
 Exit Sub
 ElseIf Trim(Combo4.Text) = "" Then
  MsgBox "Select Schedule For student..", vbCritical + vbOKOnly, "Warning"
  Combo4.SetFocus
 Exit Sub
 ElseIf pic_name = "" Then
  MsgBox "Student Photo is Required", vbCritical + vbOKOnly, "Warning"
  Uploadbtn.SetFocus
 Exit Sub
 End If
End If
If Trim(Text5.Text) = "" Then
  MsgBox "Enter start date..", vbCritical + vbOKOnly, "Warning"
  Text5.SetFocus
 Exit Sub
ElseIf Trim(Text7.Text) = "" Then
  MsgBox "How many Test ??", vbCritical + vbOKOnly, "Warning"
  Text7.SetFocus
 Exit Sub
ElseIf Trim(Text6.Text) = "" Then
  MsgBox "Enter Amount For Test..", vbCritical + vbOKOnly, "Warning"
  Text6.SetFocus
 Exit Sub
 End If
If MsgBox("Are You Sure Update This Record ??", vbQuestion + vbYesNo, "Update ") = vbYes Then
Set r1 = New ADODB.Recordset
sql = " update rstud set rstud_nm='" & reg_name.Text & "',rstud_father_nm='" & fname.Text & "',RSTUD_DOB='" & Format(date1.Value, "dd-mmm-yyyy") & "',rstud_mob=" & mobNo.Text & ",rstud_add='" & Text3.Text & "',rstud_adhr='" & Text8.Text & "',rstud_email='" & Text9.Text & "',c_id='" & Label26.Caption & "',pkg_id='" & Label25.Caption & "',sch_id='" & Label24.Caption & "',rstud_doj='" & Format(Text5.Text, "dd-mmm-yyyy") & "',rstud_doe='" & Format(Text4.Text, "dd-mmm-yyyy") & "',rstud_tot_test=" & Text7.Text & ",rstud_amnt=" & Text6.Text & ",RSTUD_PIC='" & pic_name & "' where rstud_reg_no='" & reg_no.Text & "'"
Set r1 = c1.Execute(sql)
MsgBox "Record SuccessFully Updated", vbInformation + vbOKOnly, "Updated"
Form_Load
Else
End If
End Sub
