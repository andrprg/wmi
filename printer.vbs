'On Error Resume Next

'���� ��� ����� ����������
Const DATA_DIR = "\\10.57.131.155\start\"
'����������� ����
strNetwork = "10.57.145."
'� ������ �������� IP ������ �����������
minIp = 5
'����� IP �����������
maxIp = 250
'����� ������
strUser = "ce\CETL_Andrey"
'������
strPassword = 
'����� �� ������� ���������� ����� ������ ������ ��������� ���������
'���� ����� ��������� ������ ����� ������� �� ����� �������� �������� ���������� � ���������� ���������
timeSleep = 55000
'================================================================================================================

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
			printer = GetPrinter()
			snmp = False
			if printer = True Then
				if CheckInstallSnmp() = False Then
					if InstallSnmp() = True Then
						snmp = True
					End If
				Else 
					snmp = True
				End If
			End If 'if printer = True Then
    
			If snmp = true Then
				'netframework2 = isInstallNetFramework2() 
				'CmdLine2 = DATA_DIR & "Printers\InstallNetFramework\insNet2_0.bat"
				'If  netframework2 = False Then
				'	RemoteExecute CmdLine2, "	��������� .Net Framework 2.0"  
				'	netframework2 = isInstallNetFramework2() 		   
				'End If
	  

				netframework3 = isInstallNetFramework3() 
				CmdLine3 = DATA_DIR & "Printers\InstallNetFramework\insNet3_0.bat"
				If  netframework3 = False Then
					RemoteExecute CmdLine3, "	��������� .Net Framework 3.5"  
					netframework3 = isInstallNetFramework3() 		   
				End If

				netframework4 = isInstallNetFramework4() 
				CmdLine4 = DATA_DIR & "Printers\InstallNetFramework\insNet4_0.bat"
				If  netframework4 = False Then
					RemoteExecute CmdLine4, "	��������� .Net Framework 4.0"  
					netframework4 = isInstallNetFramework4() 		   
				End If
				
				If  netframework3 = True and netframework4 = True Then
					If InstallHpProxyAgent() = True Then
						InstallHpWSProxyService()
					Else	
					End If
				End If 'If  netframework2 = True and netframework3 and netframework4 = True Then
			End If 'If snmp = true Then	

		End If		
	Else 
		WScript.StdOut.WriteLine "���� �� ��������"
	End If 'If Avaible(strComputer) = True Then
Next	

fl.Close()
MsgBox "������������ ���������."
WScript.Quit()

'�������� ������� ������ snmp
Function CheckInstallSnmp()
	msg = "	������ SNMP �� �����������"
	CheckInstallSnmp = False

	Dim cItems:Set cItems = objWMI.ExecQuery("Select * from Win32_Service",,48)
	'MsgBox (strComputer & "== Error # " & CStr(Err.Number) & " " & Err.Description)

	For Each objItem in cItems 
		If objItem.Name = "SNMP" Then
			msg = "	������ SNMP �����������"
			CheckInstallSnmp = True
		    If VersionOc() = 6 Then
			    Const HKEY_LOCAL_MACHINE = &H80000002
				Set oReg = objWMI.Get("StdRegProv")
				strKeyPath = "SYSTEM\CurrentControlSet\services\SNMP\Parameters\ValidCommunities"
				strValueName = "public"
				dwValue = 8
				oReg.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue
			End If
		End If
	Next
	WriteLog msg	
	
	
End function

'������ ��
Function VersionOc()
	Set colOperatingSystems = objWMI.ExecQuery ("Select * from Win32_OperatingSystem")
	For Each objOperatingSystem in colOperatingSystems
		VersionOc = Left(objOperatingSystem.Version,1)
	Next
End Function

'��������� ������
Function InstallSnmp()
    WriteLog "	��������� SNMP"
    InstallSnmp = False
	Const Impersonate = 3
    objWMI.Security_.ImpersonationLevel = Impersonate 
    Set Process = objWMI.Get("Win32_Process")
	if VersionOc() = 5 Then
		'����������� ���� � ������������
		Const HKEY_LOCAL_MACHINE = &H80000002    	
		Set Srv = objLocator.ConnectServer(strComputer, "root\default", strUser, strPassword)
		Set oReg = Srv.Get("StdRegProv")
		strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Setup"
		strValueName = "SourcePath"
		strValue = DATA_DIR & "WindowsXP"
		oReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
		strValueName = "ServicePackSourcePath"
		oReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
		'������� ���������
		CmdLine = "sysocmgr.exe /i:%WINDIR%\inf\sysoc.inf /u:\\10.57.131.155\start\Printers\snmp.txt /x /q /r"
	Elseif VersionOc() = 6 Then	
	     '������� ���������
		CmdLine = "ocsetup SNMP /unattendfile:\\10.57.131.155\start\Printers\snmp.txt /passive"
	End If
	'���������		
	result = Process.Create(CmdLine, , , ProcessId)
	Wscript.Sleep (timeSleep)
	If CheckInstallSnmp() = True Then
		InstallSnmp = True
	End If
End Function

'��������� ������� ��������� ���������
Function GetPrinter()
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
					msg = ""
				End If	 
			Next
		End If
		WriteLog msg
		Set cItems = Nothing
End Function


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

'��������� HP SNMP Proxy
Function InstallHpProxyAgent()
  InstallHpProxyAgent = False
  If isInstallSnmpProxy() = False Then  
	 If OcBit() = 32 Then
  		CmdLine = "Msiexec /i " & DATA_DIR & "Printers\InstallHpJetAgent\HPSNMPProxy_32_10_3_0010.msi /qn" 
	   	RemoteExecute CmdLine, "	��������� HP SNMP Proxy"
	 ElseIf OcBit() = 64 Then
	   	CmdLine =  "Msiexec /i " & DATA_DIR & "Printers\InstallHpJetAgent\HPSNMPProxy_64_10_3_0010.msi /qn" 
		RemoteExecute CmdLine, "	��������� HP SNMP Proxy"		
	 Else 
	 	 WriteLog "	�� ������� ���������� ����������� �������"
	 End If
   InstallHpProxyAgent = isInstallSnmpProxy()
  Else
	InstallHpProxyAgent = True
  End if 
End Function


Function isInstallSnmpProxy()	
	isInstallSnmpProxy = False
	msg = "	HP SNMP Proxy �� ����������"
	Const HKEY_LOCAL_MACHINE = &H80000002
	Set Srv = objLocator.ConnectServer(strComputer, "root\default", strUser, strPassword)
	Set oReg = Srv.Get("StdRegProv")
	strKeyPath = "SOFTWARE\Hewlett-Packard\HP SNMP Proxy"
	strValueName = "Name"
	oReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
    If TypeName(strValue) <> "Null" Then 	
		msg = "	HP SNMP Proxy ����������"
		isInstallSnmpProxy = True
    End If    
	WriteLog msg                 
End Function

Function isInstallNetFramework4()
  isInstallNetFramework4 = False
	msg = "	Net Framework 4.0 �� ����������"
	Const HKEY_LOCAL_MACHINE = &H80000002
	Set Srv = objLocator.ConnectServer(strComputer, "root\default", strUser, strPassword)
	Set oReg = Srv.Get("StdRegProv")
	strKeyPath = "SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Client"
	strValueName = "Install"
	oReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
	if strValue = 1 Then
		isInstallNetFramework4 = True
		msg = "	Net Framework 4.0 ����������"
	End If	
	WriteLog msg	
End Function

Function isInstallNetFramework3()
  isInstallNetFramework3 = False
	msg = "	Net Framework 3.5 �� ����������"
	Const HKEY_LOCAL_MACHINE = &H80000002
	Set Srv = objLocator.ConnectServer(strComputer, "root\default", strUser, strPassword)
	Set oReg = Srv.Get("StdRegProv")
	strKeyPath = "SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5"
	strValueName = "Install"
	oReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
	if strValue = 1 Then
		isInstallNetFramework3 = True
		msg = "	Net Framework 3.5 ����������"
	End If	
	WriteLog msg	
End Function


Function isInstallNetFramework2()
  isInstallNetFramework2 = False
	msg = "	Net Framework 2.0 �� ����������"
	Const HKEY_LOCAL_MACHINE = &H80000002
	Set Srv = objLocator.ConnectServer(strComputer, "root\default", strUser, strPassword)
	Set oReg = Srv.Get("StdRegProv")
	strKeyPath = "SOFTWARE\Microsoft\NET Framework Setup\NDP\v2.0.50727"
	strValueName = "Install"
	oReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
	if strValue = 1 Then
		isInstallNetFramework2 = True
		msg = "	Net Framework 2.0 ����������"
	End If	
	WriteLog msg	
End Function



Function RemoteExecute(CmdLine,Msg)
    Const Impersonate = 3
    WriteLog Msg
    objWMI.Security_.ImpersonationLevel = Impersonate 
    Set Process = objWMI.Get("Win32_Process")
    result = Process.Create(CmdLine, , , ProcessId)
	WriteLog "		������������� ��������: " & ProcessId
    If result = 0 Then
		Do While objWMI.execquery("select * from Win32_process where ProcessId = " & ProcessId).count > 0
			Wscript.sleep 2000
		Loop
	Else
		WScript.StdOut.WriteLine "�� ���� ������� ������� �� ��������� ������. ��� ������: " & result
	End If
End Function


Function InstallNetFramework(CmdLine, msg)
	'Const Impersonate = 3
	if isInstallNetFramework = False Then
	    
		'objWMI.Security_.ImpersonationLevel = Impersonate 
		'Set Process = objWMI.Get("Win32_Process")
		'CmdLine = DATA_DIR & "Printers\InstallNetFramework\insNet4_0.bat" 
		RemoteExecute CmdLine,msg		
		'result = Process.Create(CmdLine, , , ProcessId)		
		'WScript.StdOut.WriteLine "		������������� ��������: " & ProcessId
		'Do While objWMI.execquery("select * from Win32_process where ProcessId = " & ProcessId).count > 0
			'WScript.StdOut.Write "|"
		'	Wscript.sleep 2000
		'Loop
		isInstallNetFramework()
	End If	
End Function

Function isInstallHpWSProxyService()
	isInstallHpWSProxyService = False
	WriteLog "	�������� ������� HPWSProxyService"
	msg =  "	������ HPWSProxyService �� �����������"
    Const OpenAsDefault = -2
    Const FailIfNotExist = 0
    Const ForReading = 1	
 
    Set oShell = CreateObject("WScript.Shell")
    Set oFSO = CreateObject("Scripting.FileSystemObject")
 
    sTemp = oShell.ExpandEnvironmentStrings("%TEMP%")
    sTempFile = sTemp & "\" & oFSO.GetTempName
    
    oShell.Run "%comspec% /c sc \\" & strComputer & " query HPWSProxyService>" & sTempFile, 0, True
	WScript.Sleep 4000 
	
	Set fFile = oFSO.OpenTextFile(sTempFile, ForReading, FailIfNotExist, OpenAsDefault)
    sResults = fFile.ReadAll
    fFile.Close
    oFSO.DeleteFile (sTempFile)
 
    If InStr(lcase(sResults), "running") Then
        GetServiceStatus = "running"
        HPWSInstalled = "Yes"
    End If
    If InStr(lcase(sResults), "stopped") Then
        GetServiceStatus = "stopped"
        HPWSInstalled = "Yes"
    End If
    If InStr(lcase(sResults), "paused") Then
        GetServiceStatus = "paused"
        HPWSInstalled = "Yes"
    End If
    If InStr(lcase(sResults), "continue_pending") Then
        GetServiceStatus = "continue_pending"
        HPWSInstalled = "Yes"
    End If
    If InStr(lcase(sResults), "pause_pending") Then
        GetServiceStatus = "pause_pending"
        HPWSInstalled = "Yes"
    End If
    If InStr(lcase(sResults), "start_pending") Then
        GetServiceStatus = "start_pending"
        HPWSInstalled = "Yes"
    End If
    If InStr(lcase(sResults), "stop_pending") Then
        GetServiceStatus = "stop_pending"
        HPWSInstalled = "Yes"
    End If
    If Not Len(GetServiceStatus) > 0 Then
        GetServiceStatus = "unknown"
        HPWSInstalled = "No"
    End If
	If HPWSInstalled = "Yes" Then 
		msg = "	������ HPWSProxyService �����������"
		WriteLog msg
		WriteLog  "		��������� ������: " & GetServiceStatus
		isInstallHpWSProxyService = True
	Else
		WriteLog msg
	End If
End Function

Function InstallHpWSProxyService()
	InstallHpWSProxyService = False
	If isInstallHpWSProxyService() = False Then	
		CmdLine =  "Msiexec /i " & DATA_DIR & "Printers\InstallHpJetAgent\HPWSProxyService_10_3_1.msi /quiet" 
		RemoteExecute CmdLine, "	��������� HPWSProxyService"  
		InstallHpWSProxyService = isInstallHpWSProxyService()
	End if 
End Function


'������� ����������� �������� ��
Function OcBit()
	OcBit = 0
	Set colOperatingSystems = objWMI.ExecQuery ("Select * from Win32_ComputerSystem")
	For Each objItem in colOperatingSystems
	    If LCase(objItem.SystemType) = "x86-based pc" Then
			OcBit = 32
		End If
		If LCase(objItem.SystemType) = "x64-based pc" Then
			OcBit = 64
		End If
	Next
End Function

Function WriteLog(msg)
  fl.WriteLine(msg)
  WScript.StdOut.WriteLine msg
End Function