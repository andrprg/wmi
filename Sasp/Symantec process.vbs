dim strCommand, strComputer
strComputer = "10.57.118.205"
'strComputer = "10.57.141.62"
'Логин админа
strUser = "ce\CETL_Andrey"
'Пароль
strPassword = "l.k.j987TY"

' Create WMI object
'Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Dim objLocator: Set objLocator = CreateObject("WbemScripting.SWbemLocator")
Set objWMIService = objLocator.ConnectServer (strComputer, "root\cimv2", strUser, strPassword)  
Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process")

For Each objProcess in colProcessList
    'colProperties = objProcess.GetOwner(strNameOfUser,strUserDomain)
    If objProcess.Name = "msiexec.exe" Then
    	Wscript.Echo "Process " & objProcess.Name & " ProcessId " & objProcess.ProcessID 
    End If 
Next