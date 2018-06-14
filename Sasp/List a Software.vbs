dim strCommand, strComputer
strComputer = "10.57.253.88"
'Логин админа
strUser = "ce\CETL_Andrey"
'Пароль
strPassword = "l.k.j987TY"

' Create WMI object
Dim objLocator: Set objLocator = CreateObject("WbemScripting.SWbemLocator")
Set objWMIService = objLocator.ConnectServer (strComputer, "root\cimv2", strUser, strPassword)  
Set colProcessList = objWMIService.ExecQuery("SELECT * FROM Win32_Product where name like'Symantec%'","WQL",wbemFlagReturnImmediately)

'Set colShares = objWMIService.ExecQuery("Select * from Win32_Share Where Name = 'Admin$'")


'If colShares.Count = 0 Then
'	MsgBox "Admin$ Does Not Exist On: " & UCase(strComputer)
'Else
	'MsgBox "Admin$ Exist On: " & UCase(strComputer)
'	For Each objSoftware in colShares
' 		Wscript.Echo objSoftware.Path
' 		Wscript.Echo objSoftware.Name
' 		Wscript.Echo objSoftware.Type
	'Next
'End If


For Each objSoftware in colProcessList
 WScript.Echo objSoftware.Caption & vbtab & _
 objSoftware.Version
Next

