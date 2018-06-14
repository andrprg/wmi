On Error Resume Next

dim strCommand, strComputer
strComputer = "10.57.153.11"
strCommand = "cmd.exe /c md c:\hash" 
'Логин админа
'strUser = "ce\a_v_peregudov"
strUser = "ce\CETL_Andrey"
'Пароль
'strPassword = "tr3falTX5"
strPassword = "l.k.j987TY"

' Create WMI object
'Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Dim objLocator: Set objLocator = CreateObject("WbemScripting.SWbemLocator")
Set objWMIService = objLocator.ConnectServer (strComputer, "root\cimv2", strUser, strPassword)  


'strCommand = "cmd.exe /c md c:\Symantec" 
'RemoteExecute "Создаем"
'MsgErr()
'WScript.Sleep (3000)

'strCommand = "xcopy \\10.57.131.155\start\Symantec\SEP\*.* C:\Symantec /S/Y"
'RemoteExecute "Копируем"
'MsgErr()
'WScript.Sleep (3000)

'strCommand = "Msiexec /i C:\Symnatec\Sep.msi /quiet /norestart"
'RemoteExecute "Установка"
'MsgErr()
'WScript.Sleep (3000)

'strCommand = "xcopy \\10.57.131.155\start\Symantec\DNTUS26.exe C:\Windows\System32 /S/Y"
'RemoteExecute "Копируем"
'MsgErr()
'Wscript.Sleep (3000)

AddReg()


Function DelProcess()
	Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process")

	For Each objProcess in colProcessList
		colProperties = objProcess.GetOwner(strNameOfUser,strUserDomain)
		If strNameOfUser = "CETL_Andrey" Then
			objProcess.Terminate()
		End If	
	Next
End Function


Function RemoteExecute(msg)
	Set objNewProcess = objWMIService.Get("Win32_Process")
	MsgErr()
	WriteLog msg
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
		'Do While objWMIService.execquery("select * from Win32_process where ProcessId = " & intProcessID).count > 0
		'	Wscript.sleep 2000
		'Loop
		WriteLog  " Return value: " & intReturn
	End If
	
	
End Function 

Function MsgErr()
	If Err.Number <> 0 Then
		MsgBox "	" & Err.Number & " Источник: " & Err.Source & " Описание: " &  Err.Description
		Err.Clear
		WScript.Quit()
	End If
End Function

Function WriteLog(msg)
  WScript.StdOut.WriteLine msg
End Function

Function AddReg()
	Const HKEY_LOCAL_MACHINE = &H80000002
	Set oReg = objWMIService.Get("StdRegProv")
	
	strKeyPath = "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters"
	strValueName = "AutoShareWks"
	dwValue = 1
	oReg.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue
	
	

End Function
