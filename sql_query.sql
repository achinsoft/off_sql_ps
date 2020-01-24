if replace(left(cast(serverproperty('productversion') 
as varchar),2),'.','')<=8 
begin 
select 
serverproperty('computernamephysicalnetbios') as physicalname, 
dbid, 
convert(varchar(25), db.name) as dbname, 
convert(varchar(10), databasepropertyex(name, 'status')) as [status], 
(select count(1) from sysaltfiles where db_name(dbid) = db.name 
and groupid<>0) as datafiles, 
(select sum((size*8)/1024) from sysaltfiles where db_name(dbid) = 
db.name and groupid<>0) as [data mb], 
(select count(1) from sysaltfiles where db_name(dbid) = 
db.name and groupid=0) as logfiles, 
(select sum((size*8)/1024) from sysaltfiles where 
db_name(dbid) = db.name and groupid=0) as [log mb], 
databasepropertyex(name, 'recovery') as [recovery model], 
case cmptlevel 
when 60 then '60 (sql server 6.0)' 
when 65 then '65 (sql server 6.5)' 
when 70 then '70 (sql server 7.0)' 
when 80 then '80 (sql server 2000)' 
when 90 then '90 (sql server 2005)' 
when 100 then '100 (sql server 2008)' 
end as [compatibility level], 
convert(varchar(20), crdate, 103) + ' ' + 
convert(varchar(20), crdate, 108) as [creation date], 
'torn page detection' as [page verify option], 
case 
When( 
select top 1 
convert(varchar,datediff(day,bk.backup_finish_date, getdate())  ) 
from msdb..backupset bk 
where bk.database_name = db.name 
and bk.type='d'  
order by backup_set_id desc 
) <7 
then 'Backup Ok' 
else  
'Full Backup N/A or Older than 7 Days' 
end 
as [Full Backup], 
case 
when (db.dbid  in(1,2,3,4)) then 
'N/R' 
When ( 
select top 1 
convert(varchar,datediff(hh,bk.backup_finish_date, getdate())  ) 
from msdb..backupset bk 
where bk.database_name = db.name 
and bk.type='i' 
order by backup_set_id desc 
) <24 
then 'Backup Ok' 
else 
'Diff Backup N/A or Older than 24 hours' 
end 
as [Diff backup], 
case 
when (db.dbid  in(1,2,3,4)) 
then 'N/R' 
when 
(databasepropertyex(name, 'recovery')='simple') 
then 'N/R' 
When ( 
select top 1 
convert(varchar,datediff(mi,bk.backup_finish_date, getdate())  ) 
from msdb..backupset bk 
where bk.database_name = db.name 
and bk.type='l' 
order by backup_set_id desc 
) <=120 
then 'Backup Ok' 
else 
'Log Backup N/A or Older than 120 Mins ' 
end 
as [Transaction log] 
from sysdatabases db 
where db.name <>'tempdb' 
order by 2 
end 
if replace(left(cast(serverproperty('productversion') 
as varchar),2),'.','')>=9 
begin 
select 
serverproperty('computernamephysicalnetbios') as physicalname, 
database_id, 
convert(varchar(25), db.name) as dbname, 
convert(varchar(10), databasepropertyex(name, 'status')) as [status], 
(select count(1) from sys.master_files where 
db_name(database_id) = db.name and type_desc = 'rows') as datafiles, 
(select sum((size*8)/1024) from sys.master_files  where 
db_name(database_id) = db.name and type_desc = 'rows') as [data mb], 
(select count(1) from sys.master_files where 
db_name(database_id) = db.name and type_desc = 'log') as logfiles, 
(select sum((size*8)/1024) from sys.master_files  where 
db_name(database_id) = db.name and type_desc = 'log') as [log mb], 
recovery_model_desc as [recovery model], 
case compatibility_level 
when 60 then '60 (sql server 6.0)' 
when 65 then '65 (sql server 6.5)' 
when 70 then '70 (sql server 7.0)' 
when 80 then '80 (sql server 2000)' 
when 90 then '90 (sql server 2005)' 
when 100 then '100 (sql server 2008)' 
when 110 then '110 (sql server 2012)' 
end as [compatibility level], 
convert(varchar(20), create_date, 103) + ' ' + 
convert(varchar(20), create_date, 108) as [creation date], 
page_verify_option_desc as [page verify option], 
case 
When ( 
select top 1 
convert(varchar,datediff(day,bk.backup_finish_date, getdate())  ) 
from msdb..backupset bk 
where bk.database_name = db.name 
and bk.type='d' 
order by backup_set_id desc 
) <7 
then 'Backup Ok' 
else 
'Full Backup N/A or Older than 7 Days' 
end 
as [Full Backup], 
case 
when (db.database_id  in(1,2,3,4)) then 
'N/R' 
When ( 
select top 1 
convert(varchar,datediff(hh,bk.backup_finish_date, getdate())  ) 
from msdb..backupset bk 
where bk.database_name = db.name 
and bk.type='i' 
order by backup_set_id desc 
) <24 
then 'Backup Ok' 
else 
'Diff Backup N/A or Older than 24 hours' 
end 
as [Diff backup], 
case 
when (db.database_id  in(1,2,3,4)) 
then 'N/R' 
when 
(db.recovery_model_desc='simple') 
then 'N/R' 
When ( 
select top 1 
convert(varchar,datediff(mi,bk.backup_finish_date, getdate())  ) 
from msdb..backupset bk 
where bk.database_name = db.name 
and bk.type='l' 
order by backup_set_id desc 
) <=120 
then 'Backup Ok' 
else 
'Log Backup N/A or Older than 120 Mins ' 
end 
as [Transaction log] 
from sys.databases  db 
where state_desc ='online' and db.name <>'tempdb' 
order by 2 
end