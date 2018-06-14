Dim fso, logFile, scriptDir
scriptDir = left(WScript.ScriptFullName,(Len(WScript.ScriptFullName))-(len(WScript.ScriptName)))
Set fso = CreateObject("Scripting.FileSystemObject")
Set logFile = fso.OpenTextFile(scriptDir & "MicrosoftUpdatesInstall.log", 8, True)
logFile.WriteLine "==============================================================================="
logFile.WriteLine Now()
'=======================================================================
' Please place the updates in a folder and modify the below folder path
spPatchFolder = "c:\updatesystem" 
'=======================================================================
InstallPatches(spPatchFolder)
logFile.WriteLine "==============================================================================="
MsgBox "Microsoft updates installation completed successfully!",64, "Microsoft updates"

Set fso = Nothing
Set logFile = Nothing
WScript.Quit

Function InstallPatches(sFolder)
	Set objfso = CreateObject("Scripting.FileSystemObject")
	Set objShell = CreateObject("Wscript.Shell")
	Set folder = objfso.GetFolder(sFolder)
	Set files = folder.Files
	For each sFile In files
		If Ucase(Right(sFile.name,3)) = "CAB" Then
			'pkgmgr /ip /m:<path><file name>.cab /quiet
			i=objShell.Run ("pkgmgr /ip /m:" & Chr(34) & sfolder & "\" & sFile.name & chr(34) & " /quiet")
			If i = 0 Or i = 3010 Then 
				logFile.WriteLine chr(34) & folder & "\" & sFile.name & chr(34) & " installation completed Successfully" 
			Else
				logFile.WriteLine chr(34) & folder & "\" & sFile.name & chr(34) & " installation failed. Exit Code: " & i 
			End If
		End If 

	Next
	For Each Subfolder in folder.SubFolders
    	InstallPatches(Subfolder)
	Next
	Set objfso = Nothing 
	Set objShell = Nothing 
	Set folder = Nothing 
	Set files = Nothing 
End Function
