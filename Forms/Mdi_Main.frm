VERSION 5.00
Object = "{F5E116E1-0563-11D8-AA80-000B6A0D10CB}#1.0#0"; "HookMenu.ocx"
Begin VB.MDIForm MDI 
   BackColor       =   &H00FFFFFF&
   Caption         =   "SPEEDUP TEST SOLUTIONS"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   20250
   Icon            =   "Mdi_Main.frx":0000
   LinkTopic       =   "MDIForm1"
   MouseIcon       =   "Mdi_Main.frx":0ECA
   Picture         =   "Mdi_Main.frx":11D4
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin HookMenu.XpMenu XpMenu1 
      Left            =   5880
      Top             =   1920
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   0
      CheckBorderColor=   7021576
      SelMenuBorder   =   7021576
      SelMenuBackColor=   14073525
      SelMenuForeColor=   16646297
      SelCheckBackColor=   14791828
      MenuBorderColor =   6956042
      SeparatorColor  =   -2147483632
      MenuBackColor   =   14609903
      MenuForeColor   =   0
      CheckBackColor  =   15326939
      CheckForeColor  =   10027263
      DisabledMenuBorderColor=   -2147483632
      DisabledMenuBackColor=   15660791
      DisabledMenuForeColor=   -2147483631
      MenuBarBackColor=   15790320
      MenuPopupBackColor=   16777215
      ShortCutNormalColor=   0
      ShortCutSelectColor=   16646297
      ArrowNormalColor=   10027263
      ArrowSelectColor=   12484864
      ShadowColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "MDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Unload(cancel As Integer)
'If MsgBox("Are You Sure To Exit Application...", vbYesNo + vbCritical, "Exit Application") = vbYes Then
'cancel = 0
End
'Else
'cancel = 1
'End If
End Sub
