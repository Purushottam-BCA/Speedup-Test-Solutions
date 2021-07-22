insert into course values('C001','BCA','Bachelor In Computer Applications')
/
insert into course values('C002','BBM','Bachelor In Business Management')
/
insert into course values('C003','11th','Intermediate 1st')
/
insert into course values('C004','12th','Intermediate 2nd')
/
insert into sub values('S001','C001','C')
/
insert into sub values('S002','C001','C++')
/
insert into sub values('S003','C001','DBMS')
/
insert into sub values('S004','C002','ECONOMICS')
/
insert into sub values('S005','C003','PHYSICS')
/
insert into sub values('S006','C003','CHEMISTRY')
/
insert into sub values('S007','C003','MATH')
/
insert into sub values('S008','C004','PHYSICS')
/
insert into sub values('S009','C004','CHEMISTRY')
/
insert into sub values('S010','C004','MATH')
/
insert into  topic values('T001','FUNCTIONS',10,'S001','C001')
/
insert into  topic values('T002','LOOPS',10,'S001','C001')
/
insert into  topic values('T003','POINTERS',10,'S001','C001')
/
insert into  topic values('T004','CLASS & OBJECTS',15,'S002','C001')
/
insert into  topic values('T005','FUNDAMENTALS IN C++',15,'S002','C001')
/
insert into  topic values('T006','Addition and Subtraction',10,'S001','C002')
/
insert into  topic values('T007','Multiplication and Division',15,'S001','C002')
/
insert into  topic values('T008','Square Root and Cube Root',10,'S001','C002')
/
insert into  topic values('T009','Simplification',10,'S001','C002')
/
insert into  topic values('T010','Number System',10,'S001','C002')
/
insert into  topic values('T011','Alphabates',10,'S003','C002')
/
insert into  topic values('T012','Coding-Decoding',10,'S003','C002')
/
insert into  topic values('T013','Input-Output',10,'S003','C002')
/
insert into  topic values('T014','Blood Relation',15,'S003','C002')
/
insert into  topic values('T015','Calender',15,'S003','C002')
/
insert into  topic values('T016','Clock',15,'S003','C002')
/
insert into  topic values('T017','Noun',10,'S004','C001')
/
insert into  topic values('T018','Pronoun',10,'S004','C001')
/
insert into  topic values('T019','Adjective',10,'S004','C001')
/
insert into q_typ values('qt001','MCQs',1)
/
insert into pkg values('P001','Package1',100,15,10,'C001')
/
insert into pkg values('P002','Package2',150,15,15,'C001')
/
insert into pkg values('P003','Package3',200,15,20,'C001')
/
insert into pkg values('P004','package4',250,20,25,'C002')
/
insert into pkg values('P005','package5',100,15,10,'C002')
/
insert into pkg values('P006','package6',100,15,10,'C002')
/
insert into pkg values('P007','package7',100,15,10,'C003')
/
insert into pkg values('P008','package8',100,15,10,'C003')
/
insert into secques values(1,'What Is Your Nick Name ?')
/
insert into secques values(2,'In Which Year Are You Born ?')
/
EXIT
/