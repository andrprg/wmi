dim strCommand, strComputer
strComputer = "10.57.105.28"
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
Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process where name='msiexec.exe'")

For Each objProcess in colProcessList
    colProperties = objProcess.GetOwner(strNameOfUser,strUserDomain)
    Wscript.Echo "Process " & objProcess.Name & " ProcessId " _ 
        & objProcess.ProcessID 
    'objProcess.Terminate()    
Next