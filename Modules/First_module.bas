Attribute VB_Name = "first_module"
'For Round Form (64 Bit & 32 Bit)
Private Declare Function CreateRoundRectRgn _
Lib "gdi32" (ByVal X1 As Long, _
ByVal Y1 As Long, _
ByVal X2 As Long, _
ByVal Y2 As Long, _
ByVal X3 As Long, _
ByVal Y3 As Long) As Long

Private Declare Function SetWindowRgn _
Lib "user32" (ByVal hWnd As Long, _
ByVal hRgn As Long, _
ByVal bRedraw As Boolean) As Long

'Student Current Background Image
Public CurrentDashPic As String

'++++++++++++ QUESTION PAPER +++++++++++++'
'Variable to Know Select Question Catogary
   
   Public Ques_Purpose As String  'Purpose
   
   Public Ques_Cat As String  'Catogary
   
   Public Ques_selection As String 'Selection
   
   Public Ques_Course As String  'Course
   Public Ques_Subject As String  'Subject
   Public Ques_Topic As String    'topic
   Public Ques_diff_leval As String  'Diff Leval
    
   Public tot_Ques_in_bank As Integer 'Total Question in database
   Public Tot_Ques As Integer 'Total Selected Question By User
   
   Public ppr_org_nm As String
   Public ppr_org_address As String
   Public ppr_tst_nm As String
   Public ppr_class As String
   Public ppr_sub As String
   Public ppr_maxMarks As Integer
   Public ppr_Time_hr As Integer
   Public ppr_time_min As Integer
   Public ppr_mrk_p_ques As Double
   
   Public Ques_Include_Ans As Integer  'Answer Key
 '+++++++++++++++++Question paper+++++++++++++++'
 
'Procedure to position the form in center
Public Sub CenterForm(frm As Form)
    Dim TopCorner As Integer
    Dim LeftCorner As Integer
    
    If frm.WindowState <> 0 Then Exit Sub
    TopCorner = (Screen.Height - frm.Height) \ 2
    LeftCorner = (Screen.Width - frm.Width) \ 2
    frm.Move LeftCorner, TopCorner
End Sub

'+++++++++++Log Out++++++++++++' Student
Public Function log_out_rstud()
On Error Resume Next
If r.State = adStateOpen Then r.Close
If r1.State = adStateOpen Then r1.Close
If r2.State = adStateOpen Then r2.Close
If r3.State = adStateOpen Then r3.Close
If r11.State = adStateOpen Then r11.Close
If rs_reg_stu.State = adStateOpen Then rs_reg_stu.Close
If rs_unreg_stu.State = adStateOpen Then rs_unreg_stu.Close
If rs_course.State = adStateOpen Then rs_course.Close
If rs_sub.State = adStateOpen Then rs_sub.Close
If rs_topic.State = adStateOpen Then rs_topic.Close
If rs_qtyp.State = adStateOpen Then rs_qtyp.Close
If rs_mcqMain.State = adStateOpen Then rs_mcqMain.Close
 Unload stu_profile
 Unload stu_pwsd
 Unload Stud_Ranking
 Unload Stu_Test_selection
 Unload rstud_pkg
 Unload rstud_Pkg_renew
 Unload stu_prev_record
 Unload stu_dash
 Unload login_new
 Unload FrmSelectUser
 FrmSelectUser.Show
' login_new.userID.Text = ""
' login_new.pswd.Text = ""
' login_new.Show
FrmSelectUser.Show
End Function

Public Sub autobackup()
Dim sl As String
sl = "exp sts/sts grants=y file=" & App.Path & "\Database\AutoBackupFile.DMP"
Shell "cmd.exe /c " & s1
End Sub

Public Function log_out_Admin()
On Error Resume Next
autobackup
admin_login_reg_no = ""
If r.State = adStateOpen Then r.Close
If r1.State = adStateOpen Then r1.Close
If r2.State = adStateOpen Then r2.Close
If r3.State = adStateOpen Then r3.Close
If r11.State = adStateOpen Then r11.Close
If rs_reg_stu.State = adStateOpen Then rs_reg_stu.Close
If rs_unreg_stu.State = adStateOpen Then rs_unreg_stu.Close
If rs_course.State = adStateOpen Then rs_course.Close
If rs_sub.State = adStateOpen Then rs_sub.Close
If rs_topic.State = adStateOpen Then rs_topic.Close
If rs_qtyp.State = adStateOpen Then rs_qtyp.Close
If rs_mcqMain.State = adStateOpen Then rs_mcqMain.Close
 Unload about_org
 Unload admin_dash
 Unload emp_id_pass
 Unload Form1
 Unload formConfirm
 Unload frmbackup
 Unload FrmClient1
 Unload FrmClient2
 Unload FrmClient3
 Unload FrmClient4
 Unload frmCourseMaster
 Unload FrmEmpMaster
 Unload FrmExpense
 Unload FrmExportQues
 Unload frmFrgtPswd
 Unload FrmImportQues
 Unload FrmExportQues
 Unload FrmIncmExpense
 Unload FrmPackage
 Unload FrmQuesType
 Unload FrmQuesUpdate
 Unload FrmReportMain
 Unload FrmRestore
 Unload FrmSchedule
 Unload frmSubMaster
 Unload FrmTestPrpt1
 Unload FrmTestPrpt2
 Unload FrmTopicMaster
 Unload LoginINFO
 Unload mcq_s
 Unload Paper_Preview
 Unload QpaprSetup
 Unload ques_entry_dash
 Unload QuesBank
 Unload Question_PPR
 Unload QuestionPPRdashboard
 Unload regstudnt
 Unload Search_registered
 Unload Security_Question
 Unload stud_id_pass
 Unload StuPendingReq
 Unload Stud_Ranking
 Unload FrmSelectUser
 FrmSelectUser.Show
' login_Admin.userID.Text = ""
' login_Admin.pswd.Text = ""
' login_Admin.vkCheck1.Value = vbUnchecked
' login_Admin.Show
End Function

Public Function log_out_Emp() 'LogOut User (Normal)
EMP_login_reg_no = ""
If r.State = adStateOpen Then r.Close
If r1.State = adStateOpen Then r1.Close
If r2.State = adStateOpen Then r2.Close
If r3.State = adStateOpen Then r3.Close
If r11.State = adStateOpen Then r11.Close
If rs_reg_stu.State = adStateOpen Then rs_reg_stu.Close
If rs_unreg_stu.State = adStateOpen Then rs_unreg_stu.Close
If rs_course.State = adStateOpen Then rs_course.Close
If rs_sub.State = adStateOpen Then rs_sub.Close
If rs_topic.State = adStateOpen Then rs_topic.Close
If rs_qtyp.State = adStateOpen Then rs_qtyp.Close
If rs_mcqMain.State = adStateOpen Then rs_mcqMain.Close
 Unload emp_dash
 Unload Stud_Ranking
 Unload Search_registered
 Unload regstudnt
 Unload QuestionPPRdashboard
 Unload Question_PPR
 Unload ques_entry_dash
 Unload QpaprSetup
 Unload Paper_Preview
 Unload mcq_s
 Unload FrmImportQues
 Unload FrmExportQues
 Unload frmFrgtPswd
 Unload FrmClient1
 Unload FrmClient2
 Unload FrmClient3
 Unload FrmClient4
 Unload about_org
 Unload Emp_Profil
 Unload EmpPSWRD
 Unload FrmSelectUser
 FrmSelectUser.Show
' login_EMP.userID.Text = ""
' login_EMP.pswd.Text = ""
' login_EMP.Show
End Function

Public Sub CreateRoundRectFromWindow(ByRef oWindow As Object)
Dim lRight As Long
Dim lBottom As Long
Dim hRgn As Long
With oWindow
lRight = .Width / Screen.TwipsPerPixelX
lBottom = .Height / Screen.TwipsPerPixelY
hRgn = CreateRoundRectRgn(0, 0, lRight, lBottom, 38, 38)
SetWindowRgn .hWnd, hRgn, True
End With
End Sub
