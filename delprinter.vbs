On Error Resume Next
Const DATA_DIR = "\\10.57.131.155\start\"
'����������� ����
strNetwork = "10.57.131."
'� ������ �������� IP ������ �����������
minIp =29
'����� IP �����������
maxIp = 29
'����� ������
strUser = "ce\CETL_Andrey"
'������
strPassword = "l.k.j987TY"
'����� �� ������� ���������� ����� ������ ������ ��������� ���������
'���� ����� ��������� ������ ����� ������� �� ����� �������� �������� ���������� � ���������� ���������
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
	WScript.StdOut.WriteLine "��������������: " & strComputer
	WScript.StdOut.WriteLine "============================================================="  

    If Avaible(strComputer) = True Then
		fl.WriteLine(strNetwork & CStr(j))
		fl.WriteLine("=============================================================")
		Set objWMI = objLocator.ConnectServer (strComputer, "root\cimv2", strUser, strPassword)  
		If Err.Number <> 0 Then
			WriteLog "	" & Err.Number & " ��������: " & Err.Source & " ��������: " &  Err.Description
			Err.Clear
		Else	
            if GetDelPrinter() = True Then
			   'GetDelPrinter()            
			End If	
        End If

    End If
Next
'=====================================================================================================================================================
'������ ��������� ����������� ���������� � ����
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


'��������� ������� ��������� ���������
Function GetDelPrinter()
    WriteLog "	��������� ��������� ���������"
    GetPrinter = False
		msg = "	��������� ��������� ���"
		Dim cItems: Set cItems = objWMI.ExecQuery("Select * from Win32_Printer",,48)
		If Err.Number = 0 Then
			For Each objItem in cItems
				'�������� �� ����� �����
				If InStr(objItem.PortName,"USB")>0 or InStr(objItem.PortName,"LPT")>0 or InStr(objItem.PortName,"DOT")>0 Then
				'If objItem.Local =True and objItem.Network = False Then
					'fl.WriteLine("���������:  " & strComputer  & " �������:  " & objItem.DeviceID & " ����:" & objItem.PortName)  
					WriteLog "		�������: " & objItem.DeviceID & " ����: " & objItem.PortName 
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