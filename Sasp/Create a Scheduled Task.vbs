On Error Resume Next

dim strCommand, strComputer
strComputer = "10.57.131.177"
strCommand = "cmd.exe /c shutdown -r -t 0" 
'Логин админа
'strUser = "ce\a_v_peregudov"
strUser = "ce\CETL_Andrey"
'Пароль
'strPassword = "tr3falTX5"
strPassword = "l.k.j987TY"

' Create WMI object
'Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Dim objLocator: Set objLocator = CreateObject("WbemScripting.SWbemLocator")
'Уровень аутентификации - уровень секретности пакетов:
objLocator.Security_.AuthenticationLevel = 6
'Уровень олицетворения - олицетворение:
objLocator.Security_.ImpersonationLevel = 3
Set objWMIService = objLocator.ConnectServer (strComputer, "root\cimv2", strUser, strPassword)  

p_StartTime = "20150402164912.546875+300"
p_RunRepeatedly = false
p_DaysOfWeek = 0
p_DaysOfMonth = 0
p_InteractWithDesktop = false
p_JobId = 0	' byReference, content may change! 

intResult = objWMIService.Create(strComputer, p_StartTime, p_RunRepeatedly, p_DaysOfWeek, p_DaysOfMonth, p_InteractWithDesktop, p_JobId)

Select case intResult
	Case 0 : WScript.Echo "Successful completion"
	Case 1 : WScript.Echo "Not supported"
	Case 2 : WScript.Echo "Access denied"
	Case 8 : WScript.Echo "Unknown failure"
	Case 9 : WScript.Echo "Path not found"
	Case 21 : WScript.Echo "Invalid parameter"
	Case 22 : WScript.Echo "Service not started"
End Select






Set objNewJob = objWMIService.Get("Win32_ScheduledJob",False) ',,,False)

errJobCreated = objNewJob.Create _
    ("shutdown -r -t 0", "20150403013000.000000+060",False ) 
Wscript.Echo errJobCreated