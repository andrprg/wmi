dim strCommand, strComputer
strComputer = "10.57.162.145"
'Логин админа
strUser = "ce\CETL_Andrey"
'Пароль
strPassword = "l.k.j987TY"

' Create WMI object
Dim objLocator: Set objLocator = CreateObject("WbemScripting.SWbemLocator")
Set objWMIService = objLocator.ConnectServer (strComputer, "root\cimv2", strUser, strPassword)  

Set colShares = objWMIService.ExecQuery("Select * from Win32_Share Where Name = 'Admin$'")


If colShares.Count = 0 Then
	MsgBox "Admin$ Does Not Exist On: " & UCase(strComputer)
	Set objInstance = objWMIService.Get("Win32_Share")

p_Path = "C:\WINDOWS"
p_Name = "ADMIN"
p_Type = 2147483648

intResult = objInstance.Create(p_Path, p_Name, p_Type)

Select case intResult
	Case 0 : WScript.Echo "Success"
	Case 2 : WScript.Echo "Access denied"
	Case 8 : WScript.Echo "Unknown failure"
	Case 9 : WScript.Echo "Invalid name"
	Case 10 : WScript.Echo "Invalid level"
	Case 21 : WScript.Echo "Invalid parameter"
	Case 22 : WScript.Echo "Duplicate share"
	Case 23 : WScript.Echo "Redirected path"
	Case 24 : WScript.Echo "Unknown device or directory"
	Case 25 : WScript.Echo "Net name not found"
End Select
Else
	'MsgBox "Admin$ Exist On: " & UCase(strComputer)
	For Each objItem In colShares
		WScript.Echo "AccessMask: " & objItem.AccessMask
    	WScript.Echo "AllowMaximum: " & objItem.AllowMaximum
    	WScript.Echo "Caption: " & objItem.Caption
    	WScript.Echo "Description: " & objItem.Description
    	WScript.Echo "InstallDate: " & objItem.InstallDate
    	WScript.Echo "MaximumAllowed: " & objItem.MaximumAllowed
    	WScript.Echo "Name: " & objItem.Name
    	WScript.Echo "Path: " & objItem.Path
    	WScript.Echo "Status: " & objItem.Status
    	Wscript.Echo "Type: " & objItem.Type	
    Next
End If

