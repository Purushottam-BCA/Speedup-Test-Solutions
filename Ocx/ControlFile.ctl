options(skip=1)
LOAD DATA 
INFILE 'H:\SPEEDUP_TEST\Ocx\mm.csv'
TRUNCATE
INTO TABLE Holdimport
fields terminated by ","
(
 hq_no,
 hq_txt,
 hopt1,
 hopt2,
 hopt3,
 hopt4,
 hans_txt,
 hans_no,
 hq_expln
)
