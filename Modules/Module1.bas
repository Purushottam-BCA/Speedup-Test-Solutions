Attribute VB_Name = "Module1"
'For Admin Login
Public admin_login_reg_no As String
'For Normal User or Employee
Public EMP_login_reg_no As String
'For Students
Public Stu_login_reg_no As String
Public Current_Logged_ID As String
'Test Properties
Public tst_course As String
Public tst_Type As String
Public tst_Tot_Ques As String
Public tst_Tot_Time As Integer
Public tst_Tot_Marks As Integer
Public tst_Pass_marks As Integer
Public tst_MrkPerQues As Double

'Student Test Selection Time Holding type of Test
Public Stu_typ_of_tst As Integer

'++++++++For Generating Student ID Card while Registration
Public stuPicPath As String
Public stuname As String
Public stufather As String
Public stuCourse As String
Public stuBatch As String
Public stuIddate As String

'+++++++For Enabling Option of Section tab++++++++++++'
Public IsFullLengthSelected As Byte
Public selectedType As String
Public selectedlvl As String
Public remainTIM As String
Public StuNam As String
Public GivenTESTCourse As String
Public ToTaTiMe As String

' Setting Tot ques and tot time 4 instruction page show during test
Public Total4InstructionPage As Integer
Public min4InstructionPage As String

'Set The Test Properties to display different properties as frame hiding and showing
Public choiceTST As Byte

'For Instruction general
Public GlobalPic As String

'For kNow Whether admin_Dash is open or not
Public adminOpen As Byte

'For Non Package LogIn
Public NonPackage As Integer
'For Student Ranking
Public Package2 As Byte

'For Full Length test All Atributes
Public FTOTSUB As Integer
Public FTOTQUESTION As Integer
Public FTOTTIMEMINUTE As Integer
Public FTOTTIMESECOND As Integer
Public FTOTMARKS As Integer
Public FPASSPERCENTG As Double
Public FMRKPERCOR As Double
Public FMRKPERWRONG As Double
Public Fsub1 As String
Public Fsub2 As String
Public Fsub3 As String
Public Fsub4 As String
Public Fsub5 As String
Public FNOQ1 As Integer
Public FNOQ2 As Integer
Public FNOQ3 As Integer
Public FNOQ4 As Integer
Public FNOQ5 As Integer
 'For  Setting caption & Behaviour Of Command Button Of Secotion tab during Full Mock Test
Public subName11 As String
Public Question11 As Integer

Public subName21 As String
Public Question21 As Integer
Public subName22 As String
Public Question22 As Integer

Public subName31 As String
Public Question31 As Integer
Public subName32 As String
Public Question32 As Integer
Public subName33 As String
Public Question33 As Integer

Public subName41 As String
Public Question41 As Integer
Public subName42 As String
Public Question42 As Integer
Public subName43 As String
Public Question43 As Integer
Public subName44 As String
Public Question44 As Integer

Public subName51 As String
Public Question51 As Integer
Public subName52 As String
Public Question52 As Integer
Public subName53 As String
Public Question53 As Integer
Public subName54 As String
Public Question54 As Integer
Public subName55 As String
Public Question55 As Integer

'Just For Reset  Password
Public Owner1 As String

'For Creating Admin at starting of form
Public Frm_Starting As Byte
 
' For Right Click on Ques Entry Page Side.....
Public QUESTIONRight As String

'For New Client Entry
Public CurrentClient As String

'For Instudction Available on paper or Not
Public WantInstruction As Integer
