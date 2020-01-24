
Dim Connection
Dim Recordset
Dim SQL, sSrc, iSize
Dim fLog, flog1, oFso
Dim CompArray()
Set ofso = CreateObject("Scripting.FileSystemObject")
Set fLog = oFso.CreateTextFile("Daily_Output.csv",true) 
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
	flog.writeline "Server_Name,Status"
	For i = 0 To (iSize-1)
'		WScript.Echo "Now checking : " & i & " -- " & CompArray(i)
		Call siddique(CompArray(i))
	Next
	

	fLog.close
	flog1.close
'	wscript.echo "Sending email..." 
	call send_email
'	wscript.echo "Email succesfully sent..." 
'	wscript.echo "Script successfully executed ! :)"
End Sub

sub siddique(sComputer)

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

cur_date = year(date()) & "-" & m & "-" & d & " " & "00:00:00"



SQL = "select Server_Name, fullbackup from BackupReportDump where Server_Name='" & sComputer & "' and Date >= '" & cur_date & "'"

Const adopenstatic = 3
Const adlockoptimistic = 3
Set Connection = CreateObject("ADODB.Connection")
Set Recordset = CreateObject("ADODB.Recordset")

	Connection.Open "Provider=SQLOLEDB;SERVER=ServerName;DATABASE=PowerSQL;Trusted_Connection=Yes;Integrated_Security=SSPI;"
	Recordset.Open SQL, Connection
	check = false
	do until recordset.eof
		if recordset.fields(1) <> "Backup Ok" then
			check = true
		end if
		recordset.movenext
	loop
	if check = false then
		flog.writeline sComputer & "," & "Success"
	else
		flog.writeline sComputer & "," & "Failed"
	end if

	recordset.close
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
strSMTPFrom = "Enter the email id here"
strSMTPTo = "Enter the email id here"
'strSMTPTo = "Enter the email id here"

	body = "<p>Hi,</p>" & vbnewline & vbnewline & _
		"<p>Please find attached the Sandvik SQL backup report</p>" & vbnewline & vbnewline & _
		"<p>Thank you !<br>" 
		

strTextBody = body
strSubject = "Sandvik -- Daily Backup Report"
strAttachment = "\\ServerName\E$\Backup Validation Script\Daily_Output.csv"
'C:\Users\UserID\Desktop\Scripts

Set oMessage = CreateObject("CDO.Message")
oMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
oMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "Mention the SMTP Server Name here"
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

