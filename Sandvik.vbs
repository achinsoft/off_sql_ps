
Dim Connection
Dim Recordset
Dim SQL, sSrc, iSize
Dim fLog, flog1, oFso
Dim CompArray()
Set ofso = CreateObject("Scripting.FileSystemObject")

set flog1 = ofso.createtextfile("Not_Reachable_Servers.csv",true)

sSrc = "E:\Backup Validation Script\Servers.txt"
Call Main

Sub Main

'	If UCase(Right(Wscript.FullName, 11)) = "WSCRIPT.EXE" Then
 '   	WScript.Echo "This script must be run under CScript.  Run 'cscript from a command line."
  '  	WScript.Quit
'	End If

'	WScript.Echo ("Running checks, please wait...")

	Call Server_List						' Reads server list, opens log file

	For i = 0 To (iSize-1)
'wscript.echo isize
'		WScript.Echo "Now checking : " & i & "--" & CompArray(i)
		Call siddique(CompArray(i))
		call update		
	Next
	
'	wscript.echo "Updating Records to SQL Table..." 

	fLog.close
	flog1.close
'	wscript.echo "Sending email..." 
'	call send_email
'	wscript.echo "Email succesfully sent..." 
'	wscript.echo "Script successfully executed ! :)"
End Sub

sub siddique(sComputer)

oFso.deleteFile "Raw_data.csv"
Set fLog = oFso.CreateTextFile("Raw_data.csv",true) 

SQL = "if replace(left(cast(serverproperty('productversion') " & _
"as varchar),2),'.','')<=8 " & _
"begin " & _
"select " & _
"serverproperty('computernamephysicalnetbios') as physicalname, " & _
"dbid, " & _
"convert(varchar(25), db.name) as dbname, " & _
"convert(varchar(10), databasepropertyex(name, 'status')) as [status], " & _
"(select count(1) from sysaltfiles where db_name(dbid) = db.name " & _
"and groupid<>0) as datafiles, " & _
"(select sum((size*8)/1024) from sysaltfiles where db_name(dbid) = " & _
"db.name and groupid<>0) as [data mb], " & _
"(select count(1) from sysaltfiles where db_name(dbid) = " & _
"db.name and groupid=0) as logfiles, " & _
"(select sum((size*8)/1024) from sysaltfiles where " & _
"db_name(dbid) = db.name and groupid=0) as [log mb], " & _
"databasepropertyex(name, 'recovery') as [recovery model], " & _
"case cmptlevel " & _
"when 60 then '60 (sql server 6.0)' " & _
"when 65 then '65 (sql server 6.5)' " & _
"when 70 then '70 (sql server 7.0)' " & _
"when 80 then '80 (sql server 2000)' " & _
"when 90 then '90 (sql server 2005)' " & _
"when 100 then '100 (sql server 2008)' " & _
"end as [compatibility level], " & _
"convert(varchar(20), crdate, 103) + ' ' + " & _
"convert(varchar(20), crdate, 108) as [creation date], " & _
"'torn page detection' as [page verify option], " & _
"case " & _
"When( " & _
"select top 1 " & _
"convert(varchar,datediff(day,bk.backup_finish_date, getdate())  ) " & _
"from msdb..backupset bk " & _
"where bk.database_name = db.name " & _
"and bk.type='d' " & _ 
"order by backup_set_id desc " & _
") <7 " & _
"then 'Backup Ok' " & _
"else " & _ 
"'Full Backup N/A or Older than 7 Days' " & _
"end " & _
"as [Full Backup], " & _
"case " & _
"when (db.dbid  in(1,2,3,4)) then " & _
"'N/R' " & _
"When ( " & _
"select top 1 " & _
"convert(varchar,datediff(hh,bk.backup_finish_date, getdate())  ) " & _
"from msdb..backupset bk " & _
"where bk.database_name = db.name " & _
"and bk.type='i' " & _
"order by backup_set_id desc " & _
") <24 " & _
"then 'Backup Ok' " & _
"else " & _
"'Diff Backup N/A or Older than 24 hours' " & _
"end " & _
"as [Diff backup], " & _
"case " & _
"when (db.dbid  in(1,2,3,4)) " & _
"then 'N/R' " & _
"when " & _
"(databasepropertyex(name, 'recovery')='simple') " & _
"then 'N/R' " & _
"When ( " & _
"select top 1 " & _
"convert(varchar,datediff(mi,bk.backup_finish_date, getdate())  ) " & _
"from msdb..backupset bk " & _
"where bk.database_name = db.name " & _
"and bk.type='l' " & _
"order by backup_set_id desc " & _
") <=120 " & _
"then 'Backup Ok' " & _
"else " & _
"'Log Backup N/A or Older than 120 Mins ' " & _
"end " & _
"as [Transaction log] " & _
"from sysdatabases db " & _
"where db.name <>'tempdb' " & _
"order by 2 " & _
"end " & _
"if replace(left(cast(serverproperty('productversion') " & _
"as varchar),2),'.','')>=9 " & _
"begin " & _
"select " & _
"serverproperty('computernamephysicalnetbios') as physicalname, " & _
"database_id, " & _
"convert(varchar(25), db.name) as dbname, " & _
"convert(varchar(10), databasepropertyex(name, 'status')) as [status], " & _
"(select count(1) from sys.master_files where " & _
"db_name(database_id) = db.name and type_desc = 'rows') as datafiles, " & _
"(select sum((size*8)/1024) from sys.master_files  where " & _
"db_name(database_id) = db.name and type_desc = 'rows') as [data mb], " & _
"(select count(1) from sys.master_files where " & _
"db_name(database_id) = db.name and type_desc = 'log') as logfiles, " & _
"(select sum((size*8)/1024) from sys.master_files  where " & _
"db_name(database_id) = db.name and type_desc = 'log') as [log mb], " & _
"recovery_model_desc as [recovery model], " & _
"case compatibility_level " & _
"when 60 then '60 (sql server 6.0)' " & _
"when 65 then '65 (sql server 6.5)' " & _
"when 70 then '70 (sql server 7.0)' " & _
"when 80 then '80 (sql server 2000)' " & _
"when 90 then '90 (sql server 2005)' " & _
"when 100 then '100 (sql server 2008)' " & _
"when 110 then '110 (sql server 2012)' " & _
"end as [compatibility level], " & _
"convert(varchar(20), create_date, 103) + ' ' + " & _
"convert(varchar(20), create_date, 108) as [creation date], " & _
"page_verify_option_desc as [page verify option], " & _
"case " & _
"When ( " & _
"select top 1 " & _
"convert(varchar,datediff(day,bk.backup_finish_date, getdate())  ) " & _
"from msdb..backupset bk " & _
"where bk.database_name = db.name " & _
"and bk.type='d' " & _
"order by backup_set_id desc " & _
") <7 " & _
"then 'Backup Ok' " & _
"else " & _
"'Full Backup N/A or Older than 7 Days' " & _
"end " & _
"as [Full Backup], " & _
"case " & _
"when (db.database_id  in(1,2,3,4)) then " & _
"'N/R' " & _
"When ( " & _
"select top 1 " & _
"convert(varchar,datediff(hh,bk.backup_finish_date, getdate())  ) " & _
"from msdb..backupset bk " & _
"where bk.database_name = db.name " & _
"and bk.type='i' " & _
"order by backup_set_id desc " & _
") <24 " & _
"then 'Backup Ok' " & _
"else " & _
"'Diff Backup N/A or Older than 24 hours' " & _
"end " & _
"as [Diff backup], " & _
"case " & _
"when (db.database_id  in(1,2,3,4)) " & _
"then 'N/R' " & _
"when " & _
"(db.recovery_model_desc='simple') " & _
"then 'N/R' " & _
"When ( " & _
"select top 1 " & _
"convert(varchar,datediff(mi,bk.backup_finish_date, getdate())  ) " & _
"from msdb..backupset bk " & _
"where bk.database_name = db.name " & _
"and bk.type='l' " & _
"order by backup_set_id desc " & _
") <=120 " & _
"then 'Backup Ok' " & _
"else " & _
"'Log Backup N/A or Older than 120 Mins ' " & _
"end " & _
"as [Transaction log] " & _
"from sys.databases  db " & _
"where state_desc ='online' and db.name <>'tempdb' " & _
"order by 2 " & _
"end"

Const adopenstatic = 3
Const adlockoptimistic = 3
Set Connection = CreateObject("ADODB.Connection")
Set Recordset = CreateObject("ADODB.Recordset")
	on error resume next
	Connection.Open "Provider=SQLOLEDB;SERVER=" & sComputer & ";DATABASE=master;Trusted_Connection=Yes;Integrated_Security=SSPI;"
	if err.number <> 0 then
		flog1.writeline sComputer & " " & err.description
'		wscript.echo sComputer & " " & err.description
	else
    	

    Recordset.Open SQL, Connection

	'flog.writeline recordset.fields(0).name & "," & recordset.fields(1).name & "," & recordset.fields(2).name & "," & recordset.fields(3).name & "," & recordset.fields(4).name & "," & recordset.fields(5).name & "," & recordset.fields(6).name & "," & recordset.fields(7).name & "," & recordset.fields(8).name & "," & recordset.fields(9).name & "," & recordset.fields(10).name & "," & recordset.fields(11).name & "," & recordset.fields(12).name & "," & recordset.fields(13).name & "," & recordset.fields(14).name

do until recordset.eof
if not recordset.fields(0) = "" then
	flog.writeline sComputer & "," & recordset.fields(0) & "," & recordset.fields(1) & "," & recordset.fields(2) & "," & recordset.fields(3) & "," & recordset.fields(4) & "," & recordset.fields(5) & "," & recordset.fields(6) & "," & recordset.fields(7) & "," & recordset.fields(8) & "," & recordset.fields(9) & "," & recordset.fields(10) & "," & recordset.fields(11) & "," & recordset.fields(12) & "," & recordset.fields(13) & "," & recordset.fields(14)
end if
	recordset.movenext
loop
	recordset.close
	connection.close
end if
on error goto 0
flog.close
End Sub

Sub Server_List()
	Dim sTest
	Set oFso = CreateObject("Scripting.FileSystemObject")
	If not oFso.FileExists(sSrc) Then
		MsgBox "Please create a file named " & sSrc & " containing a list of computers to run script on, one per line", vbExclamation, "Input File Needed"
		wscript.quit
	end if
	Set ts = oFso.OpenTextFile(sSrc,1)

	iSize = 0
	Do Until ts.AtEndOfStream 
		ReDim Preserve CompArray(isize) ' resizes array based on each system name read from file  
		sTest = Trim(ts.ReadLine)
		If Not stest = "" Then
			CompArray(isize) =sTest
			isize = isize + 1
		End if
	Loop
End Sub

sub send_email()
strSMTPFrom = "Enter the email address here"
strSMTPTo = "Enter the email address here"

	body = "<p>Hi,</p>" & vbnewline & vbnewline & _
		"<p>Please find attached the Sandvik SQL backup rawdata</p>" & vbnewline & vbnewline & _
		"<p>Thank you !<br>" 
		

strTextBody = body
strSubject = "Sandvik -- Backup Report"
strAttachment = "\\ServerName\E$\Backup Validation Script\Raw_data.csv"
'C:\Users\UserID\Desktop\Scripts

Set oMessage = CreateObject("CDO.Message")
oMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
oMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "SMTP Server Name"
oMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 
oMessage.Configuration.Fields.Update

oMessage.Subject = strSubject
oMessage.From = strSMTPFrom
oMessage.To = strSMTPTo
oMessage.HTMLBody = strTextBody
oMessage.AddAttachment strAttachment

oMessage.Send
WScript.Sleep 3000 'Sleep for 3 seconds
end sub

sub update

if len(month(date())) = 1 then
	m = "0" & month(date())
else
	m = month(date())
end if
if len(day(date())) = 1 then
	d = "0" & day(date())
else
	d = day(date())
end if

if len(hour(time())) = 1 then
	h = "0" & hour(time())
else
	h = hour(time())
end if

if len(minute(time())) = 1 then
	m1 = "0" & minute(time())
else
	m1 = minute(time())
end if

if len(second(time())) = 1 then
	s = "0" & second(time())
else
	s = second(time())
end if

cur_date = year(date()) & "-" & m & "-" & d & " " & h & ":" & m1 & ":" & s


	Dim count
	output = "Raw_data.csv"
	Set out = oFso.OpenTextFile(output,1)

	Set Connection = CreateObject("ADODB.Connection")
	Set Recordset = CreateObject("ADODB.Recordset")
	
	Connection.Open "Provider=SQLOLEDB;SERVER=SERVERNAME;DATABASE=PowerSQL;Trusted_Connection=Yes;Integrated_Security=SSPI;"
    
	oSize = 0
	Do Until out.AtEndOfStream 
		ReDim Preserve outArray(osize) 
		count = Trim(out.ReadLine)
		If Not count = "" Then
			outArray(osize) = count
			osize = osize + 1
		End if
	Loop

	for i = 0 to (oSize -1)
		temp = outArray(i)
		temp1 = split(temp,",")
'		fulldate = format(temp1(11), "yyyy-mm-dd hh:mm:ss")
'14/10/2005 01:54:05
		fulldate = mid(temp1(11), 7,4) & "-" & mid(temp1(11), 4,2) & "-" & mid(temp1(11), 1,2) & " " & mid(temp1(11), 12,2) & ":" & mid(temp1(11), 15,2) & ":" & mid(temp1(11), 18,2)
		SQL = "INSERT INTO BackupReportDump ( Server_Name, Physicalname, database_id, dbname, status, datafiles, data_mb, logfiles, log_mb, recovery_model, compatibilityLevel, creationdate, page_verify_option, fullbackup, Diffbackup, Transaction_log, date)  " & _
				"VALUES " & _
		"('" & temp1(0) & "','" & temp1(1) & "'," & temp1(2) & ",'" & temp1(3) & "','" & temp1(4) & "'," & temp1(5) & "," & temp1(6) & "," & temp1(7) & "," & temp1(8) & ",'" & temp1(9) & "','" & temp1(10) & "','" & fulldate & "','" & temp1(12) & "','" & temp1(13) & "','" & temp1(14) & "','" & temp1(15) & "','" & cur_date & "');"
		
	    Recordset.Open SQL, Connection
	next
End Sub

