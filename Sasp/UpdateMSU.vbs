On Error Resume Next

Dim strCommand, strComputer
Dim WshShell
'===================================================================================================================

'���� ������
strComputer = "10.57.131.8"

'��������������� ����������
'���� /quiet - ����� ���������
'���� /norestart - �� �����������
softPath = "\\10.57.131.155\start\Msu\Windows7-KB2998527-x64.msu" ' /quiet /norestart"



'����� ������
strUser = "ce\CETL_Andrey"

'������
strPassword = "l.k.j987TY"

'===================================================================================================================


'Set WshShell = WScript.CreateObject("WScript.Shell")

'������� WMI ������
Dim objLocator: Set objLocator = CreateObject("WbemScripting.SWbemLocator")
Set objWMIService = objLocator.ConnectServer (strComputer, "root\cimv2", strUser, strPassword)  
MsgErr()

'cmmd = "wusa " & WshShell.CurrentDirectory & "\" & softPath
'WScript.StdOut.WriteLine cmmd

If isInstallUpdate() = False Then
	strCommand = "cmd.exe /c md c:\updatesystem" 
	RemoteExecute(strCommand)
	strCommand = "xcopy \\10.57.131.155\start\Msu\*.* c:\updatesystem\ /Y"	
	RemoteExecute(strCommand)
	strCommand = "cmd c:\updatesystem\runWusa.bat"
	RemoteExecute(strCommand)
	
	isInstallUpdate()
End If

WScript.StdOut.WriteLine "--------------------------------------------------------------------"
WScript.StdOut.WriteLine "������� ����� ������ ��� �������� ����"
WScript.StdIn.ReadLine


'===================================================================================================================

Function RemoteExecute(CmdLine)
    Const Impersonate = 3
    WScript.StdOut.WriteLine Msg
    objWMIService.Security_.ImpersonationLevel = Impersonate 
    Set Process = objWMIService.Get("Win32_Process")
    result = Process.Create(CmdLine, , , ProcessId)

	If result = 0 Then
		WScript.StdOut.WriteLine "���������: ������������� �������� " & ProcessId
		Do While objWMIService.execquery("select * from Win32_process where ProcessId = " & ProcessId).count > 0
			Wscript.sleep 2000
		Loop
	Else
		WScript.StdOut.WriteLine "�� ���� ������� ������� �� ��������� ������. ��� ������: " & result
	End If
End Function


Function isInstallUpdate()

    isInstallUpdate = False
	msg = "���������� Windows7-KB2998527-x64.msu �� �����������"

	Set colItems = objWMIService.ExecQuery("Select * from Win32_QuickFixEngineering where HotFixID ='KB2998527'")

	For Each objItem in colItems
		isInstallUpdate = True
		msg = "���������� Windows7-KB2998527-x64.msu �����������"

	Next
	WScript.StdOut.WriteLine Msg
End Function

Function MsgErr()
	If Err.Number <> 0 Then
		MsgBox "	" & Err.Number & " ��������: " & Err.Source & " ��������: " &  Err.Description
		Err.Clear
		WScript.Quit()
	End If
End Function