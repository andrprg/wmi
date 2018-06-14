dim strCommand, strComputer
dim arr(2)

strComputer = "10.57.102."
'Логин админа
strUser = "ce\CETL_Andrey"
'Пароль
strPassword = "l.k.j987TY"


arr(0) = "10.57.102." & 6
arr(1) = "10.57.131.177"
'arr(2) = "10.57.102." & 9




For each i in arr

    comp = i

    RunScan(comp)
Next

Sub RunScan(comp)
	' Create WMI object
	Dim objLocator: Set objLocator = CreateObject("WbemScripting.SWbemLocator")
	Set objWMIService = objLocator.ConnectServer (comp, "root\cimv2", strUser, strPassword)  

	Set colOS = objWMIService.InstancesOf("Win32_OperatingSystem") 
	For each objOS in colOS 
		bd = objOS.LastBootUpTime 
		'WScript.Echo bd
		bd1 = Mid(bd,1,4) & "-" & Mid(bd,5,2) & "-" & Mid(bd,7,2) & " " & Mid(bd,9,2) & ":" & Mid(bd,11,2) & ":" & Mid(bd,13,2) 
		WScript.Echo bd1
		nw = Now
		WScript.Echo nw
		s = abs(DateDiff("s",bd1,nw)) 
		m = s \ 60 
		'WScript.Echo "m: " & m
		h = m \ 60 
		'WScript.Echo "h: " & h
		'm = m mod 60 
		's = s mod 60 
		su = h \ 24
		'su = right("00" & h, 2) \ 24
		hh = h - (su*24)

		WScript.Echo comp & " ==> Время работы системы: " & su & " дней  " & right("00" & hh, 2) & " часов " & right("00" & m, 2) & " минут " & right("00" & s, 2) & " секунд" '& vbCrLf 
	Next
End Sub