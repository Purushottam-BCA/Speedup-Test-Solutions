VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form stu_profile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Profile Manager"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9600
   Icon            =   "stu_profile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   9600
   Begin VB.PictureBox Picture1 
      Height          =   8775
      Left            =   0
      ScaleHeight     =   8715
      ScaleWidth      =   9555
      TabIndex        =   0
      Top             =   0
      Width           =   9615
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
         Height          =   285
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   7260
         Width           =   1935
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
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   3435
         Width           =   1320
      End
      Begin VB.TextBox lbl8 
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
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   5835
         Width           =   3360
      End
      Begin VB.TextBox lbl7 
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
         Left            =   3300
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   8
         Top             =   5115
         Width           =   2880
      End
      Begin VB.TextBox lbl6 
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
         Height          =   375
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   4155
         Width           =   5160
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
         Left            =   3780
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   6
         Top             =   6555
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
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1875
         Width           =   2880
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
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1155
         Width           =   2880
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   735
         Left            =   0
         TabIndex        =   1
         Top             =   8040
         Width           =   9735
         Begin VB.CommandButton btnadd 
            BackColor       =   &H00E0E0E0&
            DisabledPicture =   "stu_profile.frx":038A
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   385
            Left            =   120
            MouseIcon       =   "stu_profile.frx":0BB0
            MousePointer    =   99  'Custom
            Picture         =   "stu_profile.frx":0D02
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Return Back."
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton modfy 
            Caption         =   "Modify"
            DisabledPicture =   "stu_profile.frx":13CB
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   385
            Left            =   8115
            MouseIcon       =   "stu_profile.frx":1BF1
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Click To Modify"
            Top             =   120
            Width           =   1335
         End
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
         Left            =   3240
         TabIndex        =   12
         Top             =   2520
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "24-Apr-2019"
         Format          =   93782019
         CurrentDate     =   40178
         MaxDate         =   40178
         MinDate         =   32874
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Only  *  value can be modified."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   0
         Width           =   6375
      End
      Begin VB.Label star4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         Enabled         =   0   'False
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
         Left            =   3050
         TabIndex        =   29
         Top             =   5040
         Width           =   105
      End
      Begin VB.Label star3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         Enabled         =   0   'False
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
         Left            =   3050
         TabIndex        =   28
         Top             =   4080
         Width           =   105
      End
      Begin VB.Label star5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         Enabled         =   0   'False
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
         Left            =   3050
         TabIndex        =   27
         Top             =   5760
         Width           =   105
      End
      Begin VB.Label star2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         Enabled         =   0   'False
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
         Left            =   3045
         TabIndex        =   26
         Top             =   1800
         Width           =   105
      End
      Begin VB.Label star1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         Enabled         =   0   'False
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
         Left            =   3050
         TabIndex        =   25
         Top             =   1080
         Width           =   105
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   2175
         Left            =   7455
         Stretch         =   -1  'True
         Top             =   225
         Width           =   1815
      End
      Begin VB.Image Image2 
         Height          =   195
         Left            =   3375
         Picture         =   "stu_profile.frx":1D43
         Stretch         =   -1  'True
         Top             =   7305
         Width           =   150
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
         Left            =   3375
         TabIndex        =   24
         Top             =   6555
         Width           =   495
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
         Left            =   1080
         TabIndex        =   23
         Top             =   2520
         Width           =   1515
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H80000004&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   405
         Index           =   9
         Left            =   3240
         Shape           =   4  'Rounded Rectangle
         Top             =   7200
         Width           =   2415
      End
      Begin VB.Label lbl1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "rs0001"
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
         Left            =   3240
         TabIndex        =   22
         Top             =   360
         Width           =   1455
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   405
         Index           =   6
         Left            =   3240
         Shape           =   4  'Rounded Rectangle
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H80000004&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   405
         Index           =   5
         Left            =   3240
         Shape           =   4  'Rounded Rectangle
         Top             =   5760
         Width           =   3495
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H80000004&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   405
         Index           =   4
         Left            =   3240
         Shape           =   4  'Rounded Rectangle
         Top             =   5040
         Width           =   3015
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H80000004&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   525
         Index           =   2
         Left            =   3240
         Shape           =   4  'Rounded Rectangle
         Top             =   4080
         Width           =   5295
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
         Left            =   1080
         TabIndex        =   21
         Top             =   5760
         Width           =   1140
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
         Left            =   1080
         TabIndex        =   20
         Top             =   4110
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
         Left            =   1080
         TabIndex        =   19
         Top             =   360
         Width           =   1740
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H80000004&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   405
         Index           =   1
         Left            =   3240
         Shape           =   4  'Rounded Rectangle
         Top             =   6480
         Width           =   3015
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   405
         Index           =   0
         Left            =   3240
         Shape           =   4  'Rounded Rectangle
         Top             =   1800
         Width           =   3015
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   405
         Index           =   3
         Left            =   3240
         Shape           =   4  'Rounded Rectangle
         Top             =   1080
         Width           =   3015
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
         Left            =   1080
         TabIndex        =   18
         Top             =   1080
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
         Left            =   1080
         TabIndex        =   17
         Top             =   1800
         Width           =   1605
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile No.  :"
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
         Left            =   1080
         TabIndex        =   16
         Top             =   6480
         Width           =   1290
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
         Left            =   1080
         TabIndex        =   15
         Top             =   3360
         Width           =   945
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount  :"
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
         Left            =   1080
         TabIndex        =   14
         Top             =   7200
         Width           =   1005
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
         Left            =   1080
         TabIndex        =   13
         Top             =   5040
         Width           =   1380
      End
   End
End
Attribute VB_Name = "stu_profile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pic_name As Variant

Private Sub btnadd_Click()
Unload Me
End Sub

Private Sub Form_Load()
conn
CenterForm Me
Set r = New ADODB.Recordset
Set r = c.Execute("select * from rstud where rstud_reg_no='" & Stu_login_reg_no & "' ")
If r.EOF = False Then
lbl1.Caption = r.Fields(0)
lbl1.Caption = UCase(lbl1.Caption)
lbl2.Text = r.Fields(1)
lbl3.Text = r.Fields(2)
DTPicker1.Value = Format(r.Fields(3), "dd-mmm-yyyy")
lbl5.Text = r.Fields(5)
lbl6.Text = r.Fields(6)
lbl7.Text = r.Fields(7)
lbl8.Text = r.Fields(8)
lbl9.Text = r.Fields(4)
lbl10.Text = r.Fields(16)
If IsNull(r.Fields(17)) = False Then
pic_name = r.Fields(17)
Image1.Picture = LoadPicture(pic_name)
Else
pic_name = App.Path & "\Graphics\#\PicNotAvail.jpg"
Image1.Picture = LoadPicture(App.Path & "\Graphics\#\PicNotAvail.jpg")
End If
End If
End Sub

Private Sub Form_Unload(cancel As Integer)
stu_profile.Hide
stu_dash.Enabled = True
End Sub

Private Sub lbl2_KeyPress(KeyAscii As Integer)
    If (((KeyAscii >= 65) And (KeyAscii <= 90)) Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or (KeyAscii = 8) Or (KeyAscii = 32)) Then
       lbl2.SetFocus
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub lbl3_KeyPress(KeyAscii As Integer)
 If (((KeyAscii >= 65) And (KeyAscii <= 90)) Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or (KeyAscii = 8) Or (KeyAscii = 32)) Then
       lbl3.SetFocus
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub lbl7_KeyPress(KeyAscii As Integer) 'Adhar Card
    If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Then
        lbl7.SetFocus
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub lbl7_LostFocus() 'Adhar Card
  If (lbl7.Text <> "") Then
        If (Len(lbl7.Text) < 12) Then
            MsgBox "Invalid AADHAR number", vbQuestion + vbOKOnly, "Invalid Adhar"
            lbl7.Text = ""
            lbl7.SetFocus
        End If
        Set r = New ADODB.Recordset
        Set r = c.Execute("select RSTUD_ADHR from rstud where RSTUD_REG_NO not in ( '" & Stu_login_reg_no & "') ")
        While r.EOF = False
         If r.Fields(0) = Trim(lbl7.Text) Then
          MsgBox "This Adhar No. Already Exist..", vbInformation + vbOKOnly, "Duplicate Adhar No"
          lbl7.SetFocus
         Exit Sub
         End If
        r.MoveNext
        Wend
   End If
End Sub

Private Sub lbl8_KeyPress(KeyAscii As Integer)
If Len(Trim(lbl8.Text)) = 0 Then
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 32 Then
Else
 MsgBox "Email Id Must start With Character!!", vbInformation + vbOKOnly, "Email"
 KeyAscii = 0
 lbl8.SetFocus
Exit Sub
End If
End If
If InStr(lbl8.Text, "@") = False Then
 If KeyAscii = 95 Or KeyAscii = 46 Or KeyAscii = 64 Or (((KeyAscii >= 65) And (KeyAscii <= 90)) Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or (KeyAscii = 8) Or (KeyAscii = 32)) Or (KeyAscii >= 48 And KeyAscii <= 57) Then
   lbl8.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   lbl9.SetFocus
  Else
   KeyAscii = 0
  End If
Else
  If KeyAscii = 95 Or KeyAscii = 46 Or (((KeyAscii >= 65) And (KeyAscii <= 90)) Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or (KeyAscii = 8) Or (KeyAscii = 32)) Or (KeyAscii >= 48 And KeyAscii <= 57) Then
   lbl8.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   lbl9.SetFocus
  Else
   KeyAscii = 0
  End If
End If
End Sub

Private Sub lbl8_LostFocus() 'check It Contains @ or not
Dim domain As String
If Len(Trim(lbl8.Text)) <> 0 Then
 If Len(Trim(lbl8.Text)) < 11 Then
  MsgBox "Invalid Email, Too Short Email", vbCritical + vbOKOnly, "Email"
  lbl8.SetFocus
  Exit Sub
 End If
 If InStr(lbl8.Text, "@") = False Then
  MsgBox "Invalid Email, It Must contain @..", vbCritical + vbOKOnly, "Email"
  lbl8.SetFocus
  Exit Sub
 End If
 domain = Right(lbl8.Text, 4)
 If UCase(domain) = UCase(".COM") Or UCase(domain) = UCase(".NET") Then
  Exit Sub
 Else
  MsgBox "Invalid Email", vbCritical + vbOKOnly, "Email"
  lbl8.SetFocus
 Exit Sub
 End If
domain = Right(lbl8.Text, 3)
 If UCase(domain) = UCase(".TK") Or UCase(domain) = UCase(".IN") Then
 Exit Sub
Else
 MsgBox "Invalid Email", vbCritical + vbOKOnly, "Email"
  lbl8.SetFocus
 Exit Sub
End If
 End If
End Sub

Private Sub modfy_Click() 'Modify Button
If modfy.Caption = "Modify" Then
 star1.Enabled = True
 star2.Enabled = True
 star3.Enabled = True
 star4.Enabled = True
 star5.Enabled = True

 lbl2.Locked = False
 lbl3.Locked = False
 lbl6.Locked = False
 lbl7.Locked = False
 lbl8.Locked = False
 DTPicker1.Enabled = True
   modfy.Caption = "Confirm"
 ElseIf modfy.Caption = "Confirm" Then 'Need To Work Here
   If lbl2.Text = "" Or lbl3.Text = "" Or lbl6.Text = "" Or lbl7.Text = "" Or lbl8.Text = "" Then
    MsgBox "Cannot left be blank", vbExclamation + vbOKOnly, "Empty Data"
   Exit Sub
 End If
star1.Enabled = False
star2.Enabled = False
star3.Enabled = False
star4.Enabled = False
star5.Enabled = False
lbl2.Locked = True
lbl3.Locked = True
lbl6.Locked = True
lbl7.Locked = True
lbl8.Locked = True
DTPicker1.Enabled = False
c.Execute ("update rstud set rstud_nm='" & lbl2.Text & "',rstud_father_nm='" & lbl3.Text & "',rstud_add='" & lbl6.Text & "',rstud_adhr='" & lbl7.Text & "',rstud_email='" & lbl8.Text & "' where RSTUD_REG_NO='" & Stu_login_reg_no & "' ")
MsgBox "Successfully Updated", vbInformation + vbOKOnly, "Updated"
modfy.Caption = "Modify"
End If
End Sub
