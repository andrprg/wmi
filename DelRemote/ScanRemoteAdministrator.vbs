On Error Resume Next
Const DATA_DIR = "\\10.57.131.155\start\"
'Сканируемая сеть
strNetwork = "10.57.154."
'С какого начинать IP адреса сканировать
minIp = 5
'Каким IP заканчивать
maxIp = 250
'Логин админа
strUser = "ce\CETL_Andrey"
'Пароль
strPassword = 

Set wshShell = WScript.CreateObject("WScript.Shell")
Dim fso, f1
Set fso = CreateObject("Scripting.FileSystemObject")
Set fl = fso.CreateTextFile(DATA_DIR & "Log\DelRemote\Scan_" & strNetwork & "log", True)
Dim objWMI  
Dim objLocator: Set objLocator = CreateObject("WbemScripting.SWbemLocator")
Const HKLM = &H80000002
strR_serverOld = "System\CurrentControlSet\Services\r_server" 'ключ реестра - путь к сервису radmin 2 


For j = minIp To maxIP Step 1
	strComputer = strNetwork & CStr(j) 
	WScript.StdOut.WriteLine ""
	WScript.StdOut.WriteLine "Обрабатывается: " & strComputer
	WScript.StdOut.WriteLine "============================================================="  

    If Avaible(strComputer) = True Then
		fl.WriteLine(strNetwork & CStr(j))
		fl.WriteLine("=============================================================")
		Set objWMI = objLocator.ConnectServer (strComputer, "root\cimv2", strUser, strPassword)  
		If Err.Number <> 0 Then
			WriteLog "	" & Err.Number & " Источник: " & Err.Source & " Описание: " &  Err.Description
			Err.Clear
		Else	
            isRemote() 
        End If
    Else
        WScript.StdOut.WriteLine vbTab & "Компьютер не доступен"
    End If
Next

'=====================================================================================================================================================
Function RemoteExecute(CmdLine,Msg)
    On Error Resume Next
    Const Impersonate = 3
    WScript.StdOut.WriteLine Msg
    objWMI.Security_.ImpersonationLevel = Impersonate 
    Set Process = objWMI.Get("Win32_Process")
    result = Process.Create(CmdLine, , , ProcessId)
	'WScript.StdOut.WriteLine CStr(result)
	'WScript.StdOut.WriteLine CmdLine
	If result = 0 Then
		Do While Service.execquery("select * from Win32_process where ProcessId = " & ProcessId).count > 0
			Wscript.sleep 2000
		Loop
	Else
		WScript.StdOut.WriteLine vbTab & "Не могу создать процесс на удаленной машине. Код ошибки: " & result
	End If
    If Err.Number <> 0 Then
        WriteLog "	" & Err.Number & " Источник: " & Err.Source & " Описание: " &  Err.Description
        Err.Clear
    End If
End Function





'пингом проверяет доступность компьютера в сети
Function Avaible(name) 
    On Error Resume Next
    Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("select * from Win32_PingStatus where address = '" & name & "'")
	
    For Each objStatus in objPing
        If IsNull(objStatus.StatusCode) Or objStatus.StatusCode <> 0 Then
            Avaible = False
        Else
            Avaible = True
        End If
    Next
End function


Function isRemote()
    msg = vbTab & "Remote Administrator версии 3 не установлен"
    isRemote = False
        If keyExist(strR_serverOld) = "TRUE" Then 
            WriteLog vbTab & "Remote Administrator версии 2 установлен"
            'cmd = "cmd " & Chr(34) & "%windir%\system32\r_server.exe" & Chr(34) & " /stop" 
            'RemoteExecute cmd, vbTab & "Остановка службы......."
            'cmd = "cmd " & Chr(34) & "r_server.exe" & Chr(34) & " /unregister" 
            'RemoteExecute cmd, vbTab & "Unregister......."
            'cmd = "cmd " & Chr(34) & "r_server.exe" & Chr(34) & " /uninstall /silence"                         
            'RemoteExecute cmd, vbTab & "Удаление службы"
            'cmd = "cmd " & Chr(34) & "del /f /s /q %windir%\system32\r_server.exe"
            'RemoteExecute cmd, vbTab & "Удаление r_server.exe из системного каталога"
            'cmd = "cmd " & Chr(34) & "del /f /s /q %windir%\system32\raddrv.dll"
            'RemoteExecute cmd, vbTab & "Удаление r_server.exe из системного каталога"
            'cmd = "cmd " & Chr(34) & "del /f /s /q %ProgramFiles%\Radmin\* "
            'RemoteExecute cmd, vbTab & "Удаляем всё из папки радмина"            
            'cmd = "cmd " & Chr(34) & "rmdir /s /q" & Chr(34) & "%ProgramFiles%\Radmin" & Chr(34) 
            'RemoteExecute cmd, vbTab & "Удаляем папку радмина"
 
            'cmd = "cmd " & Chr(34) & "del /f /s /q %ProgramFiles%\Radmin\* "
            'RemoteExecute cmd, vbTab & "Удаляем всё из папки радмина"
            
            
        Else
            WriteLog vbTab & "Remote Administrator версии 2 не установлен"
        End If   
		Dim cItems: Set cItems = objWMI.ExecQuery("Select * from Win32_Product where Name like '%Radmin Server%'",,48)
        'Dim cItems: Set cItems = objWMI.ExecQuery("Select * from Win32_Product",,48)
		If Err.Number = 0 Then
			For Each objItem in cItems
                msg = vbTab & "Remote Administrator версии 3 установлен"
                isRemote = True
                objItem.Uninstall()                
			Next
		End If
		WriteLog msg
        msg = ""
		Set cItems = Nothing
End Function

Function keyExist(strRegKey)
    On Error Resume Next 
    Dim RegKeyValue
    strValueName = "ImagePath"
    Set Srv = objLocator.ConnectServer(strComputer, "root\default", strUser, strPassword)
    Set objReg = Srv.Get("StdRegProv")
    
    objReg.GetStringValue HKLM, strRegKey, strValueName, strValue  
    If TypeName(strValue) <> "Null" Then 	
		keyExist = "TRUE" 
    Else 
        keyExist = "FALSE"
    End If 
End function  

Function WriteLog(msg)
  fl.WriteLine(msg)
  WScript.StdOut.WriteLine msg
End Function