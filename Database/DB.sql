drop user sts cascade
/
create user sts identified by sts
/
GRANT CONNECT,RESOURCE TO sts
/
conn sts/sts
/
drop table RESTORE1
/
drop table BACKUP1
/
drop table HOLDIMPORT
/
drop table EXP
/
drop table INCM
/
drop table CLIENT_PMT
/
drop table CLNT_ORDR_CHLN
/
drop table CLIENT
/
drop table PKG_RENEW
/
drop  table STURANK
/
drop table STUD_PREV_REC
/
drop table RPTSTUDENTS
/
drop table STUD_LOGIN
/
drop table RSTUD
/
drop table EMP_LOGIN
/
drop table EMP
/
drop table QTESTBANK
/
drop table mcqtest
/
drop table SECQUES
/
drop table ANSWERHOLD
/
drop table QPAPRDASH
/
drop table FULLMOCKTEST
/
drop table SUBWISETEST
/
drop table TOPICWISETEST
/
drop table QPAPERFINAL
/
drop table quesMS
/
drop table pkg
/
drop table schdl
/
drop table topic
/
drop table sub 
/
drop table course
/
drop table Q_TYP
/
drop table ADMIN_LOGIN
/
drop table ADMINTBL
/
drop table org
/
create table org(
 org_rg_no varchar(15) primary key,
 org_GST varchar(20) not null,
 org_name varchar(45) not null,
 org_add varchar(120) not null,
 org_mob varchar(13) not null,
 org_email varchar(45) null ,
 org_owner varchar(50) not null,
 org_ownr_mob varchar(13) not null)
/
create table AdminTbl(
a_id varchar(4) constraint admin_PK primary key,
a_nm varchar(50) not null,
a_father varchar(50) not null,
a_mother varchar(50) not null,
a_add varchar(100) not null,
a_state varchar(20) not null,
a_mob number(10) not null,
a_mob2 number(10) null,
a_dob date not null,
a_gndr varchar(7)  not null,
a_sal number constraint a_sal check(a_sal between 1000 and 100000),
a_adhr varchar(12) not null unique,
a_email varchar(40) null,
a_pinCD number(6) Not Null,
a_j_dt date not null,
a_qualif varchar(10) not null,
a_pic varchar(500) not null)
/
create table Admin_login(
a_log_id varchar(12) primary key,
a_id varchar(4) constraint Admin_login_FK_Admintbl references AdminTbl(a_id) on delete cascade, 
a_pswd varchar(25) not null,
a_hnt varchar2(150)  not null,
a_hnt_ans varchar(100) not null)
/
Create Table Qtestbank(
Q_No Number Not Null,
Q_Txt varchar(1000) Not Null,
Opt1 varchar(500),
Opt2 varchar(500),
Opt3 varchar(500),
Opt4 varchar(500),
Ans_Txt varchar(500),
Ans_no number Not Null)
/
Create Table QPaperFinal(
Q_No Number Not Null,
Q_Txt varchar(1000) Not Null,
Opt1 varchar(500),
Opt2 varchar(500),
Opt3 varchar(500),
Opt4 varchar(500),
Ans_Txt varchar(500),
Ans_no number Not Null)
/
Create table answerhold(
ID number Not Null,
Corr_Ans Number Not Null,
User_Ans number DEFAULT 0,
BookMrk Number DEFAULT 0)
/
Create Table Mcqtest(
Q_No Number Not Null,
Q_Txt Varchar(1000) not Null,
opt1 varchar(500) not Null,
opt2 varchar(500) not Null,
opt3 varchar(500) not Null,
opt4 varchar(500) not Null,
Ans_txt varchar(500) not Null,
Ans_No number Not Null,
Img varchar(500) Null)
/
Create Table Backup1(
Bdate date )
/
Create Table Restore1(
Rdate date )
/
Create Table SECQUES(
SNo Number,
Ques Varchar(200) Not Null)
/
create table course(
c_id varchar(5) constraint course_PK primary key,
c_nm varchar(25) not null unique,
c_full_nm varchar(50) not null)
/
create table sub(
sub_id varchar(5) constraint sub_PK primary key,
c_id varchar(5) constraint sub_Fk_course references course(c_id)on delete cascade,
sub_nm varchar(35) not null)
/
create table topic(
tp_id varchar(5) constraint topic_PK primary key,
tp_nm varchar(50) not null,
tp_dur number not null,
sub_id varchar(5) constraint topic_FK_sub references sub(sub_id)on delete cascade,
c_id varchar(5) constraint topic_FK_course references course(c_id)on delete cascade)
/
create table FullmockTest(
c_id varchar(5) references course(c_id) on delete cascade,
totSub number,
totQuestion number,
totTIMEminute number,
totTIMEsecond number,
totMarks number,
passpercentg number(5,2),
mrkPercor number,
mrkPerwrong number,
sub1 varchar(5) references sub(sub_id) on delete cascade,
sub2 varchar(5) references sub(sub_id) on delete cascade,
sub3 varchar(5) references sub(sub_id) on delete cascade,
sub4 varchar(5) references sub(sub_id) on delete cascade,
sub5 varchar(5) references sub(sub_id) on delete cascade,
NoQ1 number,
NoQ2 number,
NoQ3 number,
NoQ4 number,
NoQ5 number)
/
 create table SubWiseTest(
 c_id varchar(5) references course(c_id) on delete cascade,
 tot_ques number not null,
 time_min number(2) not null,
 time_sec number(2) not null,
 tot_mrks number(3) not null,
 pass_prcn number(3) not null,
 mrkCor number,
 mrkWrng number)
/
create table TopicWiseTest(
 c_id varchar(5) references course(c_id) on delete cascade,
 tot_ques number not null,
 time_min number(2) not null,
 time_sec number(2) not null,
 tot_mrks number(3) not null,
 pass_prcn number(3) not null,
 mrkCor number,
 mrkWrng number)
/
create table q_typ(
q_typ_id varchar(6) primary key,
q_typ_nm varchar(20) not null,
q_typ_mrk number not null)
/
create table quesMS(
q_id varchar(6) Not null,
q_no number not null,
c_id varchar(5) constraint quesMS_FK_course references course(c_id)on delete cascade,
sub_id varchar(5) constraint quesMS_FK_sub references sub(sub_id)on delete cascade,
tp_id varchar(5) constraint quesMS_FK_topic references topic(tp_id)on delete cascade,
q_typ_id varchar(6) constraint quesMS_FK_q_typ references q_typ(q_typ_id)on delete cascade,
q_txt varchar(1000) not null,
opt1 varchar(500) not null,
opt2 varchar(500) not null,
opt3 varchar(500) not null,
opt4 varchar(500) not null,
ans_txt varchar(500) not null,
ans_no number not null,
q_dif_lvl varchar(7) not null, 
q_expln varchar(900) null,
q_pic varchar(500) null)
/  
create table Pkg(
PKg_id varchar(6) constraint Pkg_PK primary key,
PKg_nm varchar(30) not null,
PKg_fee decimal(6,2) not null,
PKg_dur number not null ,
PKg_all_tst number not null,
c_id varchar(5) constraint Pkg_FK_Course references course(c_id)on delete cascade)
/
create table schdl(
sch_id varchar(6) constraint Schdl_PK primary key,
sch_strnth number not null check (sch_strnth between 1 and 100),
sch_timing varchar(15) not null,
strt_time varchar(10) not null,
end_time varchar(10) not null,
c_id varchar(5) constraint schdl_FK_Course references course(c_id)on delete cascade)
/
create table Holdimport(
hq_id varchar(6) null,
hq_no number null,
hc_id varchar(5) null,
hsub_id varchar(5) null,
htp_id varchar(5) null,
hq_typ_id varchar(6) null,
hq_txt varchar(1000) null,
hopt1 varchar(500) null,
hopt2 varchar(500) null,
hopt3 varchar(500) null,
hopt4 varchar(500) null,
hans_txt varchar(500) null,
hans_no number null,
hq_dif_lvl varchar(7) null, 
hq_expln varchar(900) null,
hq_pic varchar(500) null)
/  
create table emp(
emp_id varchar(4) constraint emp_PK primary key,
e_nm varchar(60) not null,
e_father varchar(60) not null,
e_mother varchar(60) not null,
e_add varchar(200) not null,
e_state varchar(20) not null,
e_mob number(10) not null,
e_mob2 number(10) null,
e_dob date not null,
e_gndr varchar(7)  not null,
e_sal number constraint e_sal check(e_sal between 1000 and 50000),
e_adhr varchar(12) not null unique,
e_email varchar(40) null,
e_pinCD number(6) Not Null,
e_j_dt date not null,
e_qualif varchar(30) not null,
e_pic varchar(800) not null)
/
create table emp_login (
e_id varchar(4) constraint emp_log_in_FK_emp references emp(emp_id) on delete cascade, 
e_log_id varchar(20) constraint emp_log_in_PK primary key,
e_pswd varchar(20) not null check(e_pswd like '%@%'),
e_hnt varchar2(150) not null,
e_hnt_ans varchar(100) not null)
/
create table rstud(
rstud_reg_no varchar(7) constraint rstud_PK primary key,
rstud_nm varchar(50) not null,
rstud_father_nm varchar(50) not null,
rstud_dob date not null,
rstud_mob varchar(13) not null,
rstud_gndr varchar(7) not null,
rstud_add varchar(150) not null,
rstud_adhr varchar(13) not null unique,
rstud_email varchar(50) not null,
rstud_status varchar(15) not null,
c_id varchar(5) constraint r_stud_FK_Course references course(c_id)on delete cascade,
PKg_id varchar(6) constraint rstud_FK_Pkg references Pkg(pkg_id)on delete cascade,
sch_id varchar(6) constraint stud_FK_schdl references schdl(sch_id)on delete cascade,
rstud_doj date not null,
rstud_doe date not null,
rstud_tot_test number not null,
rstud_amnt decimal(8,2)not null,
rstud_pic varchar(500) Null,
rstud_all_test number)
/
create table stud_login(
rstud_reg_no varchar(7) references rstud(rstud_reg_no)on delete cascade, 
rstud_log_id varchar(15)  primary key,
rstud_pswd varchar(25) not null,
rstud_hnt varchar2(200) not null,
rstud_hnt_ans varchar(100) not null)
/
create table stud_prev_rec(
sno number not null,
sdate date not null,
tst_typ varchar(30) not null,
tot_mrk number(3) not null,
obt_mrk number(3) not null,
dif_lvl varchar(13) not null,
q_status varchar(5) not null,
rstud_reg_no varchar(7) references rstud(rstud_reg_no)on delete cascade,
totqs number not null,
totcorr number not null,
totincorr number not null,
totunatampt number not null,
TOTTIME VARCHAR2(15) null,
ELAPSEDTIME  VARCHAR2(15)null)
/
create table PKG_RENEW(
sno number primary key,
Req_DT date not null,
pkg_id varchar(6) references pkg(pkg_id)on delete cascade,
strt_DT date,
expr_DT date,
Tot_Test number(2),
Tot_FEE decimal(6,1),
RSTUD_REG_NO varchar(7) references rstud(RSTUD_REG_NO)on delete cascade)
/
create table StuRank(
RegNo varchar(7) not null,
stuNm varchar(30) not null,
Course varchar(15) not null,
totMrk number(3) not null,
ObtMrk number(3) not null,
DiffLevel varchar(13) not null,
Mobile varchar(13) not null)
/
create table client(
clnt_id varchar(5) constraint clint_PK primary key,
clnt_nm varchar(35) not null,
clnt_mob number(11) not null,
clnt_gndr varchar(7) not null)
/
create table clnt_ordr_chln(
ord_no varchar(6) constraint clnt_ord_PK primary key,
clnt_id varchar(5) constraint clnt_ordr_chln_FK_client references client(clnt_id)on delete cascade not null,
ord_date date not null,
Sch_nm varchar(80) not null,
sch_add varchar(250) not null,
class varchar(15) not null,
subject varchar(40) not null,
TotQues number(2) not null,
tot_marks number(4) not null,
totTimeHr number(2) not null,
totTimeMin number(2) not null,
Exam_nm varchar(50) not null,
MrkCorrect number ,
mrkWrong number,
Total_Paper number,
Sch_Logo Varchar(500) null,
CSTATUS varchar(20) null)
/
create table client_pmt(
pmt_id number primary key,
ord_no varchar(6) references clnt_ordr_chln(ord_no)on delete cascade not null,
clnt_id varchar(5) constraint clnt_pmt_FK_clint references client(clnt_id)on delete cascade not null,
totamt number(9,2) not null,
cl_pamt number(9,2) not null,
cl_pdate date not null,
cl_damt number(9,2) null)
/
create table incm(
s_no number constraint incm_PK primary key,
inc_from varchar(100) not  null,
inc_reason varchar(200) not null,
inc_amt number(8,2) not null,
inc_date date not null)
/
create table exp(
ex_no number constraint exp_PK primary key,
ex_where varchar(100) not null,
ex_reason varchar(200) not null,
ex_amt number(8,2) not null,
ex_date date not null)
/
create table QpaprDash(
sno number not null,
oDate date ,
deliver date,
ordrTo varchar(30),
tstType varchar(20),
Class_nm varchar(15),
sub_nm varchar(25),
totQues number,
totmrk number )
/
Create Table RPTSTUDENTS as ( select R.rstud_reg_no, R.rstud_nm, R.rstud_father_nm , R.rstud_mob, R.rstud_gndr, R.RSTUD_ADHR,  R.RSTUD_DOJ, C.c_nm, P.Pkg_nm, S.SCH_TIMING, R.RSTUD_AMNT  from rstud R, pkg P, Course C, schdl S where R.c_id=c.c_id and R.pkg_id=P.pkg_id and  R.sch_id=S.sch_id )
/