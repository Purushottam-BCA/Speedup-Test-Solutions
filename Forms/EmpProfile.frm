VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Emp_Profil 
   BorderStyle     =   0  'None
   Caption         =   "Profile"
   ClientHeight    =   10710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20415
   Icon            =   "EmpProfile.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10710
   ScaleWidth      =   20415
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   10575
      Left            =   0
      ScaleHeight     =   10515
      ScaleWidth      =   11715
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      Begin VB.Frame Frame1 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   735
         Left            =   0
         TabIndex        =   17
         Top             =   9720
         Width           =   11775
         Begin VB.CommandButton modfy 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Modify"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   10200
            MouseIcon       =   "EmpProfile.frx":076A
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Click To Modify"
            Top             =   180
            Width           =   1305
         End
         Begin VB.CommandButton btnadd 
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
            Height          =   375
            Left            =   120
            MouseIcon       =   "EmpProfile.frx":08BC
            MousePointer    =   99  'Custom
            Picture         =   "EmpProfile.frx":0A0E
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Return Back."
            Top             =   180
            Width           =   1065
         End
      End
      Begin VB.TextBox lbl5 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   3795
         Width           =   1320
      End
      Begin VB.TextBox lbl13 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   7755
         Width           =   3360
      End
      Begin VB.TextBox lbl9 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   2580
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   13
         Top             =   5475
         Width           =   2880
      End
      Begin VB.TextBox lbl8 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2580
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   4515
         Width           =   8760
      End
      Begin VB.TextBox lbl15 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   3060
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   11
         Top             =   8595
         Width           =   2400
      End
      Begin VB.TextBox lbl3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2235
         Width           =   4560
      End
      Begin VB.TextBox lbl2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1515
         Width           =   4560
      End
      Begin VB.TextBox lbl6 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   6585
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   3795
         Width           =   1320
      End
      Begin VB.TextBox lbl7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   10230
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   3795
         Width           =   360
      End
      Begin VB.TextBox lbl16 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   8925
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   6
         Top             =   8595
         Width           =   2400
      End
      Begin VB.TextBox lbl4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2955
         Width           =   4560
      End
      Begin VB.TextBox lbl14 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   8460
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   7755
         Width           =   2160
      End
      Begin VB.TextBox lbl11 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   2580
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   3
         Top             =   6915
         Width           =   1800
      End
      Begin VB.TextBox lbl10 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   8445
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   2
         Top             =   6315
         Width           =   2880
      End
      Begin VB.TextBox lbl12 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   8445
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   1
         Top             =   7035
         Width           =   2160
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-MMMM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   16
         Top             =   6120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "EmpProfile.frx":10D7
         Format          =   121438209
         CurrentDate     =   40178
         MaxDate         =   40178
         MinDate         =   32874
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Gender  :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         TabIndex        =   47
         Top             =   3720
         Width           =   945
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "State                 :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6480
         TabIndex        =   46
         Top             =   6240
         Width           =   1590
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Years"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10620
         TabIndex        =   45
         Top             =   3795
         Width           =   615
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H80000004&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   405
         Index           =   10
         Left            =   8385
         Shape           =   4  'Rounded Rectangle
         Top             =   8520
         Width           =   3015
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   2535
         Left            =   9375
         Stretch         =   -1  'True
         Top             =   105
         Width           =   2055
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Birth  :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         TabIndex        =   44
         Top             =   6120
         Width           =   1515
      End
      Begin VB.Label lbl1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2520
         TabIndex        =   43
         Top             =   720
         Width           =   1455
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   405
         Index           =   6
         Left            =   2520
         Shape           =   4  'Rounded Rectangle
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H80000004&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   405
         Index           =   5
         Left            =   2520
         Shape           =   4  'Rounded Rectangle
         Top             =   7680
         Width           =   3495
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H80000004&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   405
         Index           =   4
         Left            =   2520
         Shape           =   4  'Rounded Rectangle
         Top             =   5400
         Width           =   3015
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H80000004&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   645
         Index           =   2
         Left            =   2520
         Shape           =   4  'Rounded Rectangle
         Top             =   4440
         Width           =   8895
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email ID  :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   360
         TabIndex        =   42
         Top             =   7680
         Width           =   1140
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Adhar No.  :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   360
         TabIndex        =   41
         Top             =   5400
         Width           =   1380
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   360
         TabIndex        =   40
         Top             =   4470
         Width           =   975
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registration No :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         TabIndex        =   39
         Top             =   720
         Width           =   1740
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H80000004&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   405
         Index           =   1
         Left            =   2520
         Shape           =   4  'Rounded Rectangle
         Top             =   8520
         Width           =   3015
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   405
         Index           =   0
         Left            =   2520
         Shape           =   4  'Rounded Rectangle
         Top             =   2160
         Width           =   4695
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   405
         Index           =   3
         Left            =   2520
         Shape           =   4  'Rounded Rectangle
         Top             =   1440
         Width           =   4695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name  :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   360
         TabIndex        =   38
         Top             =   1440
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Father's name :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   360
         TabIndex        =   37
         Top             =   2160
         Width           =   1605
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile No (1)  :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   360
         TabIndex        =   36
         Top             =   8520
         Width           =   1620
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   405
         Index           =   7
         Left            =   6525
         Shape           =   4  'Rounded Rectangle
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Salary : "
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5355
         TabIndex        =   35
         Top             =   3750
         Width           =   825
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   405
         Index           =   8
         Left            =   10170
         Shape           =   4  'Rounded Rectangle
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Age : "
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9255
         TabIndex        =   34
         Top             =   3720
         Width           =   570
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile No (2)  :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6480
         TabIndex        =   33
         Top             =   8520
         Width           =   1620
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   405
         Index           =   11
         Left            =   2520
         Shape           =   4  'Rounded Rectangle
         Top             =   2880
         Width           =   4695
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mother's name :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   360
         TabIndex        =   32
         Top             =   2880
         Width           =   1680
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Qualification    : "
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6480
         TabIndex        =   31
         Top             =   7680
         Width           =   1650
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "+91"
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
         Left            =   2640
         TabIndex        =   30
         Top             =   8595
         Width           =   495
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "+91"
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
         Left            =   8520
         TabIndex        =   29
         Top             =   8595
         Width           =   495
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H80000004&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   405
         Index           =   9
         Left            =   8400
         Shape           =   4  'Rounded Rectangle
         Top             =   7680
         Width           =   2295
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Joining Date : "
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         TabIndex        =   28
         Top             =   6840
         Width           =   1440
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H80000004&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   405
         Index           =   12
         Left            =   2520
         Shape           =   4  'Rounded Rectangle
         Top             =   6840
         Width           =   1935
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H80000004&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   405
         Index           =   14
         Left            =   8385
         Shape           =   4  'Rounded Rectangle
         Top             =   6240
         Width           =   3015
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pin Code          :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6480
         TabIndex        =   27
         Top             =   6960
         Width           =   1575
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H80000004&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   405
         Index           =   15
         Left            =   8385
         Shape           =   4  'Rounded Rectangle
         Top             =   6960
         Width           =   2295
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "* ( All Details cannot be modified by user. only few Can Be Updated By User itself. For more      updation Contact To Admin..)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   26
         Top             =   0
         Width           =   8055
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   360
         Left            =   2280
         TabIndex        =   25
         Top             =   1440
         Width           =   105
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   360
         Left            =   2280
         TabIndex        =   24
         Top             =   4440
         Width           =   105
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   255
         Left            =   2280
         TabIndex        =   23
         Top             =   7680
         Width           =   135
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   255
         Left            =   2280
         TabIndex        =   22
         Top             =   2880
         Width           =   135
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   360
         Left            =   2280
         TabIndex        =   21
         Top             =   2160
         Width           =   105
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   255
         Left            =   8200
         TabIndex        =   20
         Top             =   8520
         Width           =   135
      End
   End
End
Attribute VB_Name = "Emp_Profil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pic_name As Variant

Private Sub btnadd_Click()
Unload Me
End Sub

Private Sub DTPicker1_Change()
lbl7.Text = Format((Date - DTPicker1.Value) / 365, "00")
End Sub

Private Sub Form_Load()
conn
On Error Resume Next
Me.Top = 0
Me.Left = 0
Set r = New ADODB.Recordset
Set r = c.Execute("select * from emp where emp_id='" & EMP_login_reg_no & "' ")
If r.EOF = False Then
lbl1.Caption = r.Fields(0)
lbl2.Text = r.Fields(1) 'Name
lbl3.Text = r.Fields(2) 'Father
lbl4.Text = r.Fields(3) 'Mother
lbl8.Text = r.Fields(4) 'Address
lbl10.Text = r.Fields(5) 'State
lbl15.Text = r.Fields(6) 'Mobile 1
If IsNull(r.Fields(7)) = False Then
 lbl16.Text = r.Fields(7) 'Mobile 2
End If
DTPicker1.Value = Format(r.Fields(8), "dd-mmm-yyyy") 'DOB
lbl5.Text = r.Fields(9) 'Gender
 lbl6.Text = r.Fields(10) 'Salary
 lbl9.Text = r.Fields(11) 'Adhar
 lbl13.Text = r.Fields(12) 'Email
 lbl12.Text = r.Fields(13) 'Pin Code
 lbl11.Text = r.Fields(14) 'Join Date
 lbl14.Text = r.Fields(15) 'Qualification
If IsNull(r.Fields(16)) = False Then
 pic_name = r.Fields(16)
 Image1.Picture = LoadPicture(pic_name)
Else
 pic_name = App.Path & "\Graphics\#\PicNotAvail.jpg"
 Image1.Picture = LoadPicture(App.Path & "\Graphics\#\PicNotAvail.jpg")
End If
 lbl7.Text = Format((Date - DTPicker1.Value) / 365, "00")
End If
 DTPicker1.MinDate = Date - (50 * 365)
 DTPicker1.MaxDate = Date - (20 * 365)
End Sub


Private Sub lbl13_KeyPress(KeyAscii As Integer)
If Len(Trim(lbl13.Text)) = 0 Then
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 32 Then
Else
 MsgBox "Email Id Must start With Character!!", vbInformation + vbOKOnly, "Email"
 KeyAscii = 0
 lbl13.SetFocus
 Exit Sub
End If
End If
If InStr(lbl13.Text, "@") = False Then
 If KeyAscii = 95 Or KeyAscii = 46 Or KeyAscii = 64 Or (((KeyAscii >= 65) And (KeyAscii <= 90)) Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or (KeyAscii = 8) Or (KeyAscii = 32)) Or (KeyAscii >= 48 And KeyAscii <= 57) Then
   lbl13.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
     lbl16.SetFocus
  Else
   KeyAscii = 0
  End If
Else
  If KeyAscii = 95 Or KeyAscii = 46 Or (((KeyAscii >= 65) And (KeyAscii <= 90)) Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or (KeyAscii = 8) Or (KeyAscii = 32)) Or (KeyAscii >= 48 And KeyAscii <= 57) Then
   lbl13.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   lbl16.SetFocus
  Else
   KeyAscii = 0
  End If
End If
End Sub

Private Sub lbl13_LostFocus()
Dim domain As String
If Len(Trim(lbl13.Text)) <> 0 Then
 If Len(Trim(lbl13.Text)) <= 12 Then
 MsgBox "Invalid Email, Too Short Email", vbCritical + vbOKOnly, "Email"
lbl13.SetFocus
Exit Sub
End If
If InStr(lbl13.Text, "@") = False Then
 MsgBox "Invalid Email, It Must contain @..", vbCritical + vbOKOnly, "Email"
 lbl13.SetFocus
Exit Sub
End If
domain = Right(lbl13.Text, 4)
If UCase(domain) = UCase(".COM") Or UCase(domain) = UCase(".NET") Then
Exit Sub
Else
 MsgBox "Invalid Email", vbCritical + vbOKOnly, "Email"
 lbl13.SetFocus
 Exit Sub
 End If
domain = Right(lbl13.Text, 3)
If UCase(domain) = UCase(".TK") Or UCase(domain) = UCase(".IN") Then
Exit Sub
Else
 MsgBox "Invalid Email", vbCritical + vbOKOnly, "Email"
 lbl13.SetFocus
 Exit Sub
End If
 End If
End Sub
Private Sub lbl16_KeyPress(KeyAscii As Integer)

If Len(Trim(lbl16.Text)) = 0 Then
If KeyAscii >= 48 And KeyAscii <= 53 Then
 MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
KeyAscii = 0
lbl16.SetFocus
Exit Sub
End If
End If
If Len(Trim(lbl16.Text)) = 1 Then
 If lbl16.Text = 6 Then
  If KeyAscii <> 50 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   lbl16.SetFocus
  Exit Sub
 End If
End If
End If
If Len(Trim(lbl16.Text)) = 6 Then
 If Right(lbl16.Text, 4) = "0000" Then
  If KeyAscii = 48 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   lbl16.SetFocus
  Exit Sub
 End If
End If
End If
If Len(Trim(lbl16.Text)) = 7 Then
 If Right(lbl16.Text, 5) = "00000" Then
  If KeyAscii = 48 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   lbl16.SetFocus
  Exit Sub
 End If
End If
If Right(lbl16.Text, 5) = "11111" Then
  If KeyAscii = 49 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   lbl16.SetFocus
  Exit Sub
 End If
End If
If Right(lbl16.Text, 5) = "22222" Then
  If KeyAscii = 50 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   lbl16.SetFocus
  Exit Sub
 End If
End If
If Right(lbl16.Text, 5) = "55555" Then
  If KeyAscii = 53 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   lbl16.SetFocus
  Exit Sub
 End If
End If
If Right(lbl16.Text, 5) = "66666" Then
  If KeyAscii = 54 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   lbl16.SetFocus
  Exit Sub
 End If
End If
If Right(lbl16.Text, 5) = "77777" Then
  If KeyAscii = 55 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   lbl16.SetFocus
  Exit Sub
 End If
End If
If Right(lbl16.Text, 5) = "88888" Then
  If KeyAscii = 56 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   lbl16.SetFocus
  Exit Sub
 End If
End If
If Right(lbl16.Text, 5) = "99999" Then
  If KeyAscii = 57 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   lbl16.SetFocus
  Exit Sub
 End If
End If

End If
If Len(Trim(lbl16.Text)) = 8 Then
 If Right(lbl16.Text, 6) = "000000" Then
  If KeyAscii = 48 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   lbl16.SetFocus
  Exit Sub
 End If
End If
 If Right(lbl16.Text, 6) = "111111" Then
  If KeyAscii = 48 Then
   MsgBox "Invalid Mobile Number", vbInformation + vbOKOnly, "Mobile No"
   KeyAscii = 0
   lbl16.SetFocus
  Exit Sub
 End If
End If
End If
If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Then
        lbl16.SetFocus
  Else
   KeyAscii = 0
  End If
End Sub

Private Sub lbl16_LostFocus()
If (lbl16.Text <> "") Then
        If (Len(lbl16.Text) < 10) Then
            MsgBox "Invalid MOBILE NUMBER", vbExclamation + vbOKOnly, "Invalid  Mobile No"
            lbl16.Text = ""
            lbl16.SetFocus
        End If
End If
End Sub
Private Sub lbl2_KeyPress(KeyAscii As Integer)
    If (((KeyAscii >= 65) And (KeyAscii <= 90)) Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or (KeyAscii = 8) Or (KeyAscii = 32)) Then
       lbl2.SetFocus
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        lbl3.SetFocus
    Else
       KeyAscii = 0
    End If
End Sub

Private Sub lbl3_KeyPress(KeyAscii As Integer)
 If (((KeyAscii >= 65) And (KeyAscii <= 90)) Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or (KeyAscii = 8) Or (KeyAscii = 32)) Then
   lbl3.SetFocus
 ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   lbl4.SetFocus
 Else
   KeyAscii = 0
 End If
End Sub

Private Sub lbl4_KeyPress(KeyAscii As Integer)
 If (((KeyAscii >= 65) And (KeyAscii <= 90)) Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or (KeyAscii = 8) Or (KeyAscii = 32)) Then
       lbl4.SetFocus
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        lbl8.SetFocus
    Else
      KeyAscii = 0
    End If
End Sub

Private Sub lbl8_KeyPress(KeyAscii As Integer)
If (((KeyAscii >= 65) And (KeyAscii <= 90)) Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or (KeyAscii = 8) Or (KeyAscii = 32)) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 44 Then
   lbl8.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   lbl13.SetFocus
  Else
   KeyAscii = 0
  End If
End Sub

Private Sub modfy_Click() 'Modify Button
If modfy.Caption = "Modify" Then
lbl2.Locked = False
lbl3.Locked = False
lbl4.Locked = False
lbl8.Locked = False
lbl13.Locked = False
lbl16.Locked = False
  modfy.Caption = "Save"
ElseIf modfy.Caption = "Save" Then 'Need To Work Here
If lbl2.Text = "" Or lbl3.Text = "" Or lbl4.Text = "" Or lbl8.Text = "" Or lbl13.Text = "" Then
 MsgBox "Cannot left be blank", vbExclamation + vbOKOnly, "Empty Data"
 Exit Sub
End If
lbl2.Locked = True
lbl3.Locked = True
lbl4.Locked = True
lbl8.Locked = True
lbl13.Locked = True
lbl16.Locked = True
If Trim(lbl16.Text) <> "" Then
 c.Execute ("update emp set e_nm='" & lbl2.Text & "',e_father='" & lbl3.Text & "',e_add='" & lbl8.Text & "',e_mother='" & lbl4.Text & "',e_email='" & lbl13.Text & "',e_mob2=" & lbl16.Text & " where emp_id='" & EMP_login_reg_no & "' ")
Else
 c.Execute ("update emp set e_nm='" & lbl2.Text & "',e_father='" & lbl3.Text & "',e_add='" & lbl8.Text & "',e_mother='" & lbl4.Text & "',e_email='" & lbl13.Text & "',e_mob2=NULL where emp_id='" & EMP_login_reg_no & "' ")
End If
MsgBox "Successfully Updated", vbInformation + vbOKOnly, "Updated"
modfy.Caption = "Modify"
End If
End Sub

