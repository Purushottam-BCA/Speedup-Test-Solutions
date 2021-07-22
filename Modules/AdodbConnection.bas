Attribute VB_Name = "AdodbConnection"
'DECLARATION AS USED IN THE PROJECT FOR CONNECTION
Option Explicit
Public c1 As New ADODB.Connection
Public c As New ADODB.Connection
Public r1 As New ADODB.Recordset
Public r11 As New ADODB.Recordset
Public r2 As New ADODB.Recordset
Public r3 As New ADODB.Recordset
Public r4 As New ADODB.Recordset
Public r As New ADODB.Recordset
Public rs_reg_stu As New ADODB.Recordset
Public rs_unreg_stu As New ADODB.Recordset
Public rs_course As New ADODB.Recordset
Public rs_sub As New ADODB.Recordset
Public rs_topic As New ADODB.Recordset
Public rs_qtyp As New ADODB.Recordset
Public rs_mcqMain As New ADODB.Recordset
Public sql As String
Public str As String

Public Function conn()
Set c = New ADODB.Connection
Set c1 = New ADODB.Connection
Set r = New ADODB.Recordset
Set r1 = New ADODB.Recordset
Set r11 = New ADODB.Recordset
Set r2 = New ADODB.Recordset
Set r3 = New ADODB.Recordset
Set rs_sub = New ADODB.Recordset
Set rs_topic = New ADODB.Recordset
Set rs_qtyp = New ADODB.Recordset
Set rs_mcqMain = New ADODB.Recordset
c.CursorLocation = adUseClient
c1.CursorLocation = adUseClient
c.ConnectionString = "Provider=MSDAORA.1;User ID=sts/sts;Persist Security Info=True"
c1.ConnectionString = "Provider=MSDAORA.1;User ID=sts/sts;Persist Security Info=True"
c1.Open
c.Open
 Set rs_course = c1.Execute("select c_nm from course")
 Set rs_qtyp = c1.Execute("select * from q_typ")
 Set rs_sub = c1.Execute("select * from sub")
 Set rs_mcqMain = c1.Execute("select * from quesms")
End Function
