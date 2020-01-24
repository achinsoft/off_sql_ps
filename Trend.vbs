
Dim Connection
Dim Recordset, rst
Dim SQL, sSrc, iSize, strhtml, strhtml1, cur_date, cur_date1, SR_green, SR_red, SR_empty
Dim fLog, flog1, oFso
Dim CompArray()
Set oFso = CreateObject("Scripting.FileSystemObject")
Set fLog = oFso.CreateTextFile("Trend.htm",true) 

sSrc = "E:\Backup Validation Script\Servers.txt"
Call Main

Sub Main


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

cur_date = year(date()) & "-" & m & "-" & d

cur_date1 = date - 15

strhtml0 = "<table border=1><tr><td>Legend</td></tr>"
strhtml0 = strhtml0 & "<tr><td bgcolor=green><font color=white>Green = Success</font></td></tr>"
strhtml0 = strhtml0 & "<tr><td bgcolor=red><font color=white>Red = Failed</font></td></tr>"
strhtml0 = strhtml0 & "<tr><td bgcolor=white>White = Servers Not Reachable</td></tr>"
strhtml0 = strhtml0 & "<tr><td bgcolor=blue><font color=white>Blue = Report Not Ran</font></td></tr></table><br>"


strhtml = "<table border=1><tr><td>Server</td>"

for a = cur_date to cur_date1 step -1
	strhtml = strhtml & "<td> " & a & "</td>" 
next

strhtml = strhtml & "<td>Success Rate</td></tr>"


'	WScript.Echo ("Running checks, please wait...")

	Call Server_List						' Reads server list, opens log file
'	flog.writeline "Server_Name,Status,Date"
	For i = 0 To (iSize-1)
'		WScript.Echo "Now checking : " & i & " -- " & CompArray(i)
		Call siddique(CompArray(i))
	Next
'	wscript.echo strhtml & strhtml1
	flog.writeline strhtml0 & strhtml & strhtml1 & "</table>"
	fLog.close
'	flog1.close
'	wscript.echo "Sending email..." 
	call send_email

'	wscript.echo "Email succesfully sent..." 
'	wscript.echo "Script successfully executed ! :)"
End Sub

sub siddique(sComputer)
'if instr(sComputer, "\") then
'	asComputer = replace(sComputer, "\", "/\")
'else
'	asComputer = sComputer
'end if

strhtml1 = strhtml1 & "<tr><td>" & sComputer & "</td>"

Const adopenstatic = 3
Const adlockoptimistic = 3
Set Connection = CreateObject("ADODB.Connection")
Set Recordset = CreateObject("ADODB.Recordset")
Set rst = CreateObject("ADODB.Recordset")

	Connection.Open "Provider=SQLOLEDB;SERVER=ServerName;DATABASE=PowerSQL;Trusted_Connection=Yes;Integrated_Security=SSPI;"
SR_rate = 0
Sr_red = 0
sr_green = 0

for a = cur_date to cur_date1 step -1

if len(month(a)) = 1 then
	m = "0" & month(a)
else
	m = month(a)
end if

if len(day(a)) = 1 then
	d = "0" & day(a)
else
	d = day(a)
end if

dimdate = year(a) & "-" & m & "-" & d

SQL_records = "Select * from backupreportdump where Date >= '" & dimdate & " 00:00:00" & "' and Date <= '" & dimdate & " 23:59:59" & "'"
SQL = "select Server_Name, fullbackup, Date from BackupReportDump where Server_Name='" & sComputer & "' and Date >= '" & dimdate & " 00:00:00" & "' and Date <= '" & dimdate & " 23:59:59" & "'"


'wscript.echo sql
	Recordset.Open SQL, Connection
	rst.open SQL_records, Connection

	check = false
	acount = 0 
	counter = 0
	do until recordset.eof
		if recordset.fields(1) <> "Backup Ok" then
			check = true
		end if
		recordset.movenext
		counter = counter + 1
	loop

	do until rst.eof
		rst.movenext
		acount = acount + 1
	loop

	
	if counter > 0  and check = false then
		strhtml1 = strhtml1 & "<td bgcolor=green>&nbsp;</td>" & vbnewline
		sr_green = sr_green + 1
	elseif counter > 0  and check = true then
		strhtml1 = strhtml1 & "<td bgcolor=red>&nbsp;</td>" & vbnewline
		sr_red = sr_red + 1
	elseif acount = 0 then
		strhtml1 = strhtml1 & "<td bgcolor=blue>&nbsp;</td>" & vbnewline
	else
		strhtml1 = strhtml1 & "<td>&nbsp;</td>" & vbnewline
	end if

		recordset.close
		rst.close
next
'wscript.echo sr_green
'wscript.echo sr_red

on error resume next
if sr_green+sr_red > 0 then
	SR_rate = (sr_green/16) * 100
end if
on error goto 0
		strhtml1 = strhtml1 & "<td> " & SR_rate & "%</td>"

StrHtml1 = STrHTML1 & "</tr>"
	connection.close
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
strSMTPFrom = "Enter the email ID here"
strSMTPTo = "Enter the email ID here"
'strSMTPTo = "Enter the email ID here"

	Set ss = oFso.OpenTextFile("E:\Backup Validation Script\trend.htm",1)
	trend_data = ss.readall()


strSubject = "Sandvik -- 15 Days Backup Trend"
strAttachment = "\\ServerName\E$\Backup Validation Script\Trend.htm"

Set oMessage = CreateObject("CDO.Message")
oMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
oMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "Enter the SMTP Server Name here"
oMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 
oMessage.Configuration.Fields.Update

oMessage.Subject = strSubject
oMessage.From = strSMTPFrom
oMessage.To = strSMTPTo
'oMessage.HTMLBody = strhtml & strhtml1 & "</table>"
oMessage.HTMLBody = trend_data
'oMessage.AddAttachment strAttachment

oMessage.Send
WScript.Sleep 3000 'Sleep for 3 seconds
end sub
