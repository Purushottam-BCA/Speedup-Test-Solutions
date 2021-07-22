VERSION 5.00
Begin VB.Form FrmClient3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client Payment"
   ClientHeight    =   9855
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13095
   Icon            =   "client_payment.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9855
   ScaleWidth      =   13095
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search Order No"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   12975
      Begin VB.CommandButton Command3 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3960
         MouseIcon       =   "client_payment.frx":0EE2
         MousePointer    =   99  'Custom
         TabIndex        =   38
         Top             =   555
         Width           =   1695
      End
      Begin VB.ComboBox Order 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1800
         MouseIcon       =   "client_payment.frx":1034
         MousePointer    =   99  'Custom
         TabIndex        =   37
         Text            =   "Combo1"
         Top             =   555
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   360
         TabIndex        =   42
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label62 
         BackStyle       =   0  'Transparent
         Caption         =   "Order No :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   41
         Top             =   555
         Width           =   1335
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Status :"
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
         Left            =   8760
         TabIndex        =   40
         Top             =   555
         Width           =   2655
      End
      Begin VB.Label PMT 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Caption         =   "Completed"
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
         Left            =   10680
         TabIndex        =   39
         Top             =   540
         Width           =   1860
      End
      Begin VB.Shape Shape1 
         Height          =   495
         Left            =   8640
         Top             =   480
         Width           =   3975
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H8000000B&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   8580
         Top             =   420
         Width           =   4095
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Payment Info"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   6600
      TabIndex        =   12
      Top             =   5280
      Width           =   6375
      Begin VB.TextBox Text1 
         BackColor       =   &H00B0FFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4200
         TabIndex        =   45
         Text            =   "0"
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "      Generate      Invoice"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4200
         MouseIcon       =   "client_payment.frx":1186
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Generate Invoice"
         Top             =   2520
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00808080&
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4200
         MouseIcon       =   "client_payment.frx":12D8
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Cancel"
         Top             =   3480
         Width           =   1815
      End
      Begin VB.CommandButton payment 
         BackColor       =   &H00E0E0E0&
         Caption         =   "<<<  Pay  >>> "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4200
         MouseIcon       =   "client_payment.frx":142A
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Pay Amount"
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox n14 
         BackColor       =   &H00B0FFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2760
         TabIndex        =   16
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Discount :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Left            =   4440
         TabIndex        =   44
         Top             =   560
         Width           =   1215
      End
      Begin VB.Image Image4 
         Height          =   255
         Left            =   2520
         Picture         =   "client_payment.frx":157C
         Stretch         =   -1  'True
         Top             =   3645
         Width           =   180
      End
      Begin VB.Image Image3 
         Height          =   255
         Left            =   2520
         Picture         =   "client_payment.frx":1979
         Stretch         =   -1  'True
         Top             =   2085
         Width           =   180
      End
      Begin VB.Image Image2 
         Height          =   255
         Left            =   2520
         Picture         =   "client_payment.frx":1D76
         Stretch         =   -1  'True
         Top             =   1365
         Width           =   180
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   2520
         Picture         =   "client_payment.frx":2173
         Stretch         =   -1  'True
         Top             =   660
         Width           =   180
      End
      Begin VB.Label n15 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label16"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2175
         TabIndex        =   34
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Due Date :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Left            =   480
         TabIndex        =   33
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label n11 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label16"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2760
         TabIndex        =   32
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label n12 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label16"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2760
         TabIndex        =   31
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label n13 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label16"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2760
         TabIndex        =   30
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Payable Amount : "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Left            =   465
         TabIndex        =   17
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Due Amount :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Left            =   465
         TabIndex        =   15
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Paid Amount :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Left            =   465
         TabIndex        =   14
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Left            =   465
         TabIndex        =   13
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Order Summary"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   0
      TabIndex        =   5
      Top             =   1560
      Width           =   6375
      Begin VB.Label n8 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label16"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2400
         TabIndex        =   29
         Top             =   3600
         Width           =   3375
      End
      Begin VB.Label n9 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label16"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   2400
         TabIndex        =   28
         Top             =   4440
         Width           =   3375
      End
      Begin VB.Label n10 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label16"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2400
         TabIndex        =   27
         Top             =   5400
         Width           =   3375
      End
      Begin VB.Label n5 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label16"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2400
         TabIndex        =   26
         Top             =   960
         Width           =   3375
      End
      Begin VB.Label n6 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label16"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2400
         TabIndex        =   25
         Top             =   1800
         Width           =   3375
      End
      Begin VB.Label n7 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label16"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2400
         TabIndex        =   24
         Top             =   2760
         Width           =   3375
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Institute Name :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   480
         TabIndex        =   11
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Address :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   480
         TabIndex        =   10
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Order date :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   480
         TabIndex        =   9
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Paper : "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   480
         TabIndex        =   8
         Top             =   5400
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Class :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   480
         TabIndex        =   7
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   2760
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Client Details"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   6600
      TabIndex        =   0
      Top             =   1560
      Width           =   6375
      Begin VB.Label n3 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label16"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2400
         TabIndex        =   23
         Top             =   2040
         Width           =   3375
      End
      Begin VB.Label n4 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label16"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2400
         TabIndex        =   22
         Top             =   2760
         Width           =   3375
      End
      Begin VB.Label n2 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label16"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2400
         TabIndex        =   21
         Top             =   1320
         Width           =   3375
      End
      Begin VB.Label n1 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label16"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2400
         TabIndex        =   20
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile No :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Gender :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Client name :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Client  ID : "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   1785
      Left            =   0
      Top             =   7920
      Width           =   6375
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   $"client_payment.frx":2570
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1815
      Left            =   120
      TabIndex        =   43
      Top             =   8160
      Width           =   6255
   End
End
Attribute VB_Name = "FrmClient3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
On Error Resume Next
Dim mp1 As String, Mp2 As String, Mp3 As Double, Mp4 As Double, Mp5 As Double
If PMT.Caption = "Completed" Then
 mp1 = "INVSTS" & Right(Order.Text, 3)
 Mp2 = "Question Paper For Class " & n6.Caption & " " & n7.Caption
 Mp3 = Val(n11.Caption) / Val(n10.Caption)
 Mp4 = Val(n11.Caption) - Val(Text1.Text)
 Mp5 = Mp4
 DV.CmdClientTotal
 RptInvoice.Sections("section4").Controls("InvoiceNo").Caption = mp1
 RptInvoice.Sections("section4").Controls("OrderDate").Caption = n5.Caption
 RptInvoice.Sections("section2").Controls("Label3").Caption = n2.Caption
 RptInvoice.Sections("section2").Controls("Label4").Caption = n8.Caption
 RptInvoice.Sections("section2").Controls("Label5").Caption = n9.Caption
 RptInvoice.Sections("section2").Controls("PaperInfo").Caption = Mp2
 RptInvoice.Sections("section2").Controls("PaperQty").Caption = n10.Caption
 RptInvoice.Sections("section2").Controls("PaperTotal").Caption = n11.Caption
 RptInvoice.Sections("section2").Controls("PaperDis").Caption = Text1.Text
 RptInvoice.Sections("section2").Controls("PaperNetTotal").Caption = Mp5
 RptInvoice.Sections("section2").Controls("FinalTotal").Caption = Mp4
 RptInvoice.Sections("section2").Controls("PaperRate").Caption = Mp3
 RptInvoice.Show 1, MDI
 DV.rsCmdClientTotal.Close
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
If Trim(Order.Text) = "" Then
 Exit Sub
End If
Set r = c.Execute("select * from clnt_ordr_chln where upper(ORD_NO)='" & Trim(UCase(Order.Text)) & "' ")
If r.EOF = False Then
 n5.Caption = Format(r.Fields(2), "dd-mmm-yyyy")
 n8.Caption = r.Fields(3)
 n9.Caption = r.Fields(4)
 n6.Caption = r.Fields(5)
 n7.Caption = r.Fields(6)
 n10.Caption = r.Fields(14)
 PMT.Caption = r.Fields(16)
 If UCase(PMT.Caption) = "COMPLETED" Then
  Command1.Enabled = False
  payment.Enabled = False
 Else
  payment.Enabled = True
 End If
 Set r1 = New ADODB.Recordset
 Set r1 = c.Execute("select * from  client where clnt_id='" & r.Fields(1) & "' ")
  If r1.EOF = False Then
   n1.Caption = UCase(r1.Fields(0))
   n2.Caption = UCase(r1.Fields(1))
   n4.Caption = UCase(r1.Fields(2))
   n3.Caption = UCase(r1.Fields(3))
  End If
 Set r3 = c.Execute("select * from CLIENT_PMT where upper(ORD_NO)='" & Trim(UCase(Order.Text)) & "' ")
 If r3.EOF = False Then
  n11.Caption = r3.Fields(3)
  n12.Caption = r3.Fields(4)
  n15.Caption = r3.Fields(5)
  n13.Caption = r3.Fields(6)
 End If
Else
 MsgBox "Order Id Not exist", vbCritical + vbOKOnly, ""
End If
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Top = 550
Me.Left = 3800
conn
Set r = c.Execute("select upper(ord_no) from CLNT_ORDR_CHLN ")
While r.EOF = False
 Order.AddItem r.Fields(0)
r.MoveNext
Wend
clkr
End Sub
Public Sub clkr()
n1.Caption = ""
n2.Caption = ""
n3.Caption = ""
n4.Caption = ""
n5.Caption = ""
n6.Caption = ""
n7.Caption = ""
n8.Caption = ""
n9.Caption = ""
n10.Caption = ""
n11.Caption = ""
n12.Caption = ""
n13.Caption = ""
n14.Text = ""
n15.Caption = ""
PMT.Caption = ""
Order.Text = ""
 Command1.Enabled = True
  payment.Enabled = True
End Sub
Private Sub n14_KeyPress(KeyAscii As Integer)
If (((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8)) Then
        n14.SetFocus
  ElseIf KeyAscii = 13 Then
   KeyAscii = 0
  Else
   KeyAscii = 0
   MsgBox "Only Positive Numeric Vlue Allowed..", vbInformation + vbOKOnly, ""
  End If
End Sub

Private Sub payment_Click()
If Trim(Order.Text) <> "" Then
 If Val(n14.Text) < Val(n13.Caption) Then
  MsgBox "Cannot Pay Less Amount Than Due Amount", vbInformation + vbOKOnly, ""
 Exit Sub
 End If
 c.Execute ("update CLNT_ORDR_CHLN set CSTATUS='Completed' where ord_no='" & Order.Text & "' ")
 c.Execute ("update CLIENT_PMT set CL_DAMT=0,CL_PAMT=" & Val(n11.Caption) & ",CL_PDATE='" & Format(Date, "dd-mmm-yyyy") & "' where ord_no='" & Order.Text & "'  ")
 'Inserting into Account
Dim statement As String
Set r = New ADODB.Recordset
statement = n2.Caption & " Has Paid Due Amount " & n14.Text & " for Order ID " & Order.Text
Set r = c.Execute("select count(*) from incm")
c.Execute ("insert into incm values (" & r.Fields(0) + 1 & ",'" & n2.Caption & "','" & statement & "'," & Val(n14.Text) & ",'" & Format(Date, "dd-mmm-yyyy") & "' )")
 MsgBox "Payment SuccessFully Done. Click On Generate Invoice To print Reciept", vbInformation + vbOKOnly, "Payment"
 PMT.Caption = "Completed"
 Command1.Enabled = True
 payment.Enabled = False
End If

End Sub

Private Sub Text1_Change()
n13.Caption = (Val(n11.Caption) - Val(n12.Caption)) - Val(Text1.Text)
End Sub
