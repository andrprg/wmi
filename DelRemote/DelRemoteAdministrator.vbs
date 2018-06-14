On Error Resume Next
Const DATA_DIR = "\\10.57.131.155\start\"
'Сканируемая сеть
strNetwork = "10.57.131."
'С какого начинать IP адреса сканировать
minIp = 5
'Каким IP заканчивать
maxIp = 250
'Логин админа
strUser = "ce\CETL_Andrey"
'Пароль
strPassword = "l.k.j987TY"

Set wshShell = WScript.CreateObject("WScript.Shell")
Dim fso, f1
Set fso = CreateObject("Scripting.FileSystemObject")
Set fl = fso.CreateTextFile(DATA_DIR & "Log\DelRemote\LogDel_" & strNetwork & "log", True)
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
            If isRemote() = True Then
                'DelRemote()
                isRemote()
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


Function isRemote()
    msg = vbTab & "Проверка: Remote Administrator не установлен"
    isRemote = False
		Dim cItems: Set cItems = objWMI.ExecQuery("Select * from Win32_Product where name like '%Radmin Server%'",,48)
		If Err.Number = 0 Then
			For Each objItem in cItems
				WriteLog vbTab & "Проверка: Remote Administrator установлен"
                isRemote = True
                WriteLog vbTab & "Удаление..."
                objItem.Uninstall()
			Next
		End If
		WriteLog msg
        msg = ""
		Set cItems = Nothing
End Function


Function DelRemote()
    WriteLog  "	Удаление Remote Administrator"
    DelRemote = False
		Dim cItems: Set cItems = objWMI.ExecQuery("Select * from Win32_Product where name like '%Radmin Server%'",,48)
		If Err.Number = 0 Then
			For Each objItem in cItems
				DelRemote = True
                'objItem.Uninstall()
			Next
		End If
		Set cItems = Nothing
End Function

Function WriteLog(msg)
  fl.WriteLine(msg)
  WScript.StdOut.WriteLine msg
End Function