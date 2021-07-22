set feed off
set markup html on
spool on
spool C:\Users\Purushottam\Pictures\jhjhj.xls
select hq_no,hq_txt,hopt1,hopt2,hopt3,hopt4,hans_txt,hans_no,hq_expln from holdimport;
spool off
set markup html off
commit;
exit;
