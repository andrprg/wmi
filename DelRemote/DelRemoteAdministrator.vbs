On Error Resume Next
Const DATA_DIR = "\\10.57.131.155\start\"
'����������� ����
strNetwork = "10.57.131."
'� ������ �������� IP ������ �����������
minIp = 5
'����� IP �����������
maxIp = 250
'����� ������
strUser = "ce\CETL_Andrey"
'������
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
            If isRemote() = True Then
                'DelRemote()
                isRemote()
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


Function isRemote()
    msg = vbTab & "��������: Remote Administrator �� ����������"
    isRemote = False
		Dim cItems: Set cItems = objWMI.ExecQuery("Select * from Win32_Product where name like '%Radmin Server%'",,48)
		If Err.Number = 0 Then
			For Each objItem in cItems
				WriteLog vbTab & "��������: Remote Administrator ����������"
                isRemote = True
                WriteLog vbTab & "��������..."
                objItem.Uninstall()
			Next
		End If
		WriteLog msg
        msg = ""
		Set cItems = Nothing
End Function


Function DelRemote()
    WriteLog  "	�������� Remote Administrator"
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