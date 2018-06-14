On Error Resume Next
Const DATA_DIR = "\\10.57.131.155\start\"
'Сканируемая сеть
strNetwork = "10.57.131."
'С какого начинать IP адреса сканировать
minIp =29
'Каким IP заканчивать
maxIp = 29
'Логин админа
strUser = "ce\CETL_Andrey"
'Пароль
strPassword = "l.k.j987TY"
'Время на которое укладываем спать скрипт ожидая окончания установки
'Если время установки больше этого времени то можно получить неверную информацию о результате установки
timeSleep = 55000


Set wshShell = WScript.CreateObject("WScript.Shell")
Dim fso, f1
Set fso = CreateObject("Scripting.FileSystemObject")
Set fl = fso.CreateTextFile(DATA_DIR & "Printers\LogFullInstall_" & strNetwork & "log", True)
Dim objWMI  
Dim objLocator: Set objLocator = CreateObject("WbemScripting.SWbemLocator")

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
            if GetDelPrinter() = True Then
			   'GetDelPrinter()            
			End If	
        End If

    End If
Next
'=====================================================================================================================================================
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


'Проверяет наличие локальных принтеров
Function GetDelPrinter()
    WriteLog "	Получение локальных принтеров"
    GetPrinter = False
		msg = "	Локальных принтеров нет"
		Dim cItems: Set cItems = objWMI.ExecQuery("Select * from Win32_Printer",,48)
		If Err.Number = 0 Then
			For Each objItem in cItems
				'Проверка по имени порта
				If InStr(objItem.PortName,"USB")>0 or InStr(objItem.PortName,"LPT")>0 or InStr(objItem.PortName,"DOT")>0 Then
				'If objItem.Local =True and objItem.Network = False Then
					'fl.WriteLine("Компьютер:  " & strComputer  & " Принтер:  " & objItem.DeviceID & " Порт:" & objItem.PortName)  
					WriteLog "		Принтер: " & objItem.DeviceID & " Порт: " & objItem.PortName 
					GetPrinter = True
                    objItem.Delete_
					msg = ""
				End If	 
			Next
		End If
		WriteLog msg
		Set cItems = Nothing
End Function

Function WriteLog(msg)
  fl.WriteLine(msg)
  WScript.StdOut.WriteLine msg
End Function