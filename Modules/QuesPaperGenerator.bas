Attribute VB_Name = "QuesPaperGenerator"
'Declaring all Variables used while printing the question paper
Public school_nm As String 'Name Of School
Public school_pic As String 'Logo of School
Public school_add As String 'Address of school
Public Test_nm As String 'Type of test (semester/Unit/Annual)
Public testSUB_nm As String 'Subject Name
Public testclass_nm As String 'Class Name
Public testFULLmrk As Integer 'Total Mark
Public testTOTALtime As String 'Total time
Public testTotQues As Integer 'Total Questions
Public testCorrectMRK As Byte 'Number for each correct answer
Public testWrongMRK As Byte 'Number for each wrong answer
Public instructionSET As String 'Instructions on paper

'For Dashboard of question paper generator
Public ordrdt As Date
Public delivrdt As Date
Public subdash As String
Public classdash As String
Public purposedash As String 'client or general
Public tstTypDash As String
Public totqsdash As Integer
Public totmrkdash As Integer
Public ClintOrdrDate As String 'Special For Client
Public autoGenNum As Integer 'Special For Serial No

Public CurrentSub As String
Public CurrentTopic As String


