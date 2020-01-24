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
wscript.echo cur_date