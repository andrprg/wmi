On Error Resume Next

Set WshShell = WScript.CreateObject("WScript.Shell")
Set Process = WshShell.Exec("%comspec% /c " & RunCmd)

Dim fso, f1
Set fso = CreateObject("Scripting.FileSystemObject")
Set f1 = fso.CreateTextFile("D:\IE8.log", True)
For j = 89 To 150 Step 1

	strComputer = "10.57.154." & CStr(j) 
	ServicePack = ""
	Version = ""
	'WScript.StdOut.WriteLine(strComputer)
	
	If Avaible(strComputer) = True Then

		Set objWMIService = GetObject("winmgmts:" _
    		& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")


		Set colOperatingSystem = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")

		For Each objOperatingSystem in colOperatingSystem
		ServicePack = objOperatingSystem.ServicePackMajorVersion
		Version = objOperatingSystem.Version

		Next

		If Mid(Version,1,3)="5.1" Then
			Set objWMIService = GetObject("winmgmts:\\" & strComputer & _
    		"\root\cimv2\Applications\MicrosoftIE")
			Set colIESettings = objWMIService.ExecQuery _
    			("Select * from MicrosoftIE_Summary")
			For Each strIESetting in colIESettings
                'WScript.Echo "���������:  " & strComputer  & "   ������:  " & strIESetting.Version 
                'f1.WriteLine("���������:  " & strComputer  & "   ������:  " & strIESetting.Version)  
			    If Mid(strIESetting.Version,1,3)<>"8.0" Then
    				WScript.Echo "���������:  " & strComputer  & "   ������:  " & strIESetting.Version
    				f1.WriteLine("���������:  " & strComputer  & "   ������:  " & strIESetting.Version)
    			Else
    				WScript.Echo "���������:  " & strComputer  & " ���������� �� ���������"
    			End If

			Next

		Else
    		WScript.Echo "���������:  " & strComputer  & " ���������� �� ���������"
			'WScript.Quit()
		End If
	Else
		WScript.Echo "���������:  " & strComputer  & " �� ��������"
	End If
Next	
f1.Close
MsgBox "������������ ���������. ���� � ������������ D:\Nmsk.log"
WScript.Quit()

Function Avaible(name) '������ ��������� ����������� ���������� name � ����
    On Error Resume Next
    Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}")._
        ExecQuery("select * from Win32_PingStatus where address = '"_
        & name & "'")
    For Each objStatus in objPing
        If IsNull(objStatus.StatusCode) Or objStatus.StatusCode <> 0 Then
            Avaible = False
        Else
            Avaible = True
    End If
    Next
End function