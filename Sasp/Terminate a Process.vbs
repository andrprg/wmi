' Terminate a Process


dim strCommand, strComputer
strComputer = "10.57.131.129"
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


Set colProcessList = objWMIService.ExecQuery _
    ("Select * from Win32_Process Where Name = 'xcopy.exe'")

For Each objProcess in colProcessList
   objProcess.Terminate()
Next