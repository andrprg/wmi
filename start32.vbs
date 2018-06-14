
'On Error Resume Next
Set WshShell = WScript.CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")
Set WshNetwork = CreateObject("WScript.Network")

LogFolder = "\\10.100.131.155\start\Log\"

Set LogFile = FSO.OpenTextFile(LogFolder & WshNetwork.ComputerName & "--" & WshNetwork.UserName & ".log", 2, True)




strStart = WshShell.SpecialFolders.Item("AllUsersStartUp") & "\MsStart.exe"
if  FSO.FileExists(strStart)  Then
	FSO.DeleteFile strStart
end if
LogFile.WriteLine Now & ": MsStart.exe удален"