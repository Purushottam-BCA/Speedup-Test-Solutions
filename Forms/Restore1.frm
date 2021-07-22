VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FrmRestore 
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   6570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8820
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   5655
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   8055
      Begin VB.Frame Frame6 
         Caption         =   "Last  Restore"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   480
         TabIndex        =   14
         Top             =   720
         Width           =   7575
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "23/12/2019"
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
            Left            =   2790
            TabIndex        =   16
            Top             =   480
            Width           =   4620
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Last Restore Taken On : "
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
            Index           =   0
            Left            =   480
            TabIndex        =   15
            Top             =   480
            Width           =   2190
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Restore"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   480
         TabIndex        =   10
         Top             =   2400
         Width           =   7575
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Browse"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   415
            Left            =   6000
            MouseIcon       =   "Restore1.frx":0000
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select the Backup File from Folder to Restore Database : "
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
            Left            =   480
            TabIndex        =   13
            Top             =   480
            Width           =   5220
         End
         Begin VB.Label Label4 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   12
            Top             =   960
            Width           =   5295
         End
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   415
         Left            =   4320
         MouseIcon       =   "Restore1.frx":0152
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4800
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Restore"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   415
         Left            =   2160
         MouseIcon       =   "Restore1.frx":02A4
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4800
         Width           =   1575
      End
      Begin MSComDlg.CommonDialog cdb 
         Left            =   7320
         Top             =   1920
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox Text2 
         Height          =   975
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   3360
         Visible         =   0   'False
         Width           =   6735
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   600
         TabIndex        =   17
         Top             =   1320
         Visible         =   0   'False
         Width           =   6735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Restore Remainder"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   1800
      TabIndex        =   0
      Top             =   1200
      Width           =   5055
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   550
         Left            =   1560
         MouseIcon       =   "Restore1.frx":03F6
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   1440
         Width           =   1815
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Restore"
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
            Left            =   535
            TabIndex        =   4
            Top             =   120
            Width           =   720
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H8000000F&
            BackStyle       =   1  'Opaque
            Height          =   550
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   550
         Left            =   1560
         MouseIcon       =   "Restore1.frx":0548
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   2520
         Width           =   1815
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Skip"
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
            Left            =   720
            TabIndex        =   2
            Top             =   120
            Width           =   390
         End
         Begin VB.Shape Shape3 
            BackColor       =   &H8000000F&
            BackStyle       =   1  'Opaque
            Height          =   550
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Restore Taken On : "
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
         Left            =   480
         TabIndex        =   6
         Top             =   600
         Width           =   2190
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "23/12/2019"
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
         Left            =   2735
         TabIndex        =   5
         Top             =   600
         Width           =   4620
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "If No Backup File Then click Here"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   6000
      MouseIcon       =   "Restore1.frx":069A
      MousePointer    =   99  'Custom
      TabIndex        =   19
      ToolTipText     =   "Click To Restore Data From System Created Backup File."
      Top             =   6120
      Width           =   2670
   End
   Begin VB.Image Image1 
      Height          =   540
      Left            =   8400
      MouseIcon       =   "Restore1.frx":07EC
      MousePointer    =   99  'Custom
      Picture         =   "Restore1.frx":093E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   405
   End
End
Attribute VB_Name = "FrmRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim file_name As String, file_ext As String
Dim sd As Date

Private Sub Command1_Click()
On Error Resume Next
Dim filefilter As String
filefilter = "Backup File(*.dmp)|*.dmp|All Files(*.*)|*.*"
cdb.Filter = filefilter
cdb.ShowOpen
If cdb.FileName <> "" Then
 file_name = cdb.FileName
 file_ext = Right(cdb.FileTitle, 3)
 If UCase(Trim(file_ext)) = UCase("dmp") Then
  Label4.Caption = file_name
 Else
  MsgBox "Choose Correct Backup File ...."
  file_name = ""
  Label4.Caption = ""
 End If
Else
Exit Sub
End If
End Sub

Private Sub Command2_Click()
Frame1.Visible = True
Frame4.Visible = False
End Sub

Private Sub Command3_Click()
If Label4.Caption = "" Then
MsgBox "Select The Backup (*.dmp) File", vbInformation + vbOKOnly, "No file Selected"
Exit Sub
End If
Restore1
Shell "cmd.exe /c " & Text1.Text
MsgBox "Backup File SuccessFully Imported" & vbCrLf & "Wait For Some Time ....", vbInformation + vbOKOnly, "Imported Successfully"
c.Execute ("delete from restore1")
c.Execute ("insert into restore1 values('" & Format(Date, "dd-mmm-yyyy") & "') ")
Form_Load
End Sub

Private Sub Form_Load()
conn
Frame1.Visible = True
Frame4.Visible = False
Label4.Caption = ""
CenterForm Me
Set r = New ADODB.Recordset
Set r = c.Execute("select rdate from Restore1")
If r.EOF = False Then
sd = r.Fields(0)
Label2.Caption = Format(sd, "dd/mm/yyyy")
Label8.Caption = Label2.Caption
Else
Label2.Caption = "No Record Available"
Label8.Caption = Label2.Caption
End If
Shape1.BackColor = &H80000005
Shape3.BackColor = &H80000005
Label3.FontBold = False
Label5.FontBold = False
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape1.BackColor = &H80000005
Shape3.BackColor = &H80000005
Label3.FontBold = False
Label5.FontBold = False
End Sub

Private Sub Frame3_Click()
Unload Me
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape3.BackColor = &HD0C0C0
Label5.FontBold = True
End Sub

Private Sub Image1_Click()
MsgBox "All Catagory of Backup File, Whether It Is Of Student Backup, Questions Backup, Master Entry Backup, Accounts Backup Or Complete Backup , Can Be Restored In Simple Way . Just Select The backup File then Click On restore.", vbInformation + vbOKOnly, "Important Tips"
End Sub

Private Sub Label5_Click()
Unload Me
End Sub
Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape3.BackColor = &HD0C0C0
Label5.FontBold = True
End Sub
Private Sub Frame2_Click()
Frame1.Visible = False
Frame4.Visible = True
End Sub
Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape1.BackColor = &HD0C0C0
Label3.FontBold = True
End Sub

Private Sub Label3_Click()
Frame1.Visible = False
Frame4.Visible = True
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape1.BackColor = &HD0C0C0
Label3.FontBold = True
End Sub

Public Sub Restore1()
'Text1.Text = "imp.exe tem1/hello file=" & Label4.Caption & " full=y"
 Text1.Text = "imp.exe sts/sts file=" & Label4.Caption & " full=y"
End Sub

Private Sub Label7_Click(Index As Integer)
Label4.Caption = App.Path & "\Database\AutoBackupFile.DMP"
file_name = Label4.Caption
End Sub
