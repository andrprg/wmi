Set FSO = CreateObject("Scripting.FileSystemObject")
Set WshNetwork = CreateObject("WScript.Network")
Set WshShell = CreateObject("WScript.Shell")
Set objEnv = WshShell.Environment

'===================================================================
SourceFolder = "C:\Проекты\Тележка вспомогательная 89.01.001.03.00.00.00-0"
DestFolder = "C:\1\"
LogFolder = "C:\1"
'===================================================================


iReadFile=0


'2  перезаписывается файл   8- добавляется к концу   
Set LogFile = FSO.OpenTextFile(LogFolder & WshNetwork.ComputerName & "--" & WshNetwork.UserName & ".log", 2, True)

If Not FSO.FolderExists(SourceFolder) Then
	LogFile.WriteLine Now & ", " & WshNetwork.ComputerName & ", " & WshNetwork.UserName & " : Исходный каталог " & SourceFolder & " не существует."
End If

If Not FSO.FolderExists(DestFolder) Then
	LogFile.WriteLine Now & ", " & WshNetwork.ComputerName & ", " & WshNetwork.UserName & " : Каталог назначения " & DestFolder & " не существует."
End If
AllFolders SourceFolder,DestFolder
LogFile.WriteLine Now & ": Найдено файлов для копирования:       " & iReadFile
LogFile.Close

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function AllFolders(WSourceFolder,WDestFolder)
    Set SubF = FSO.GetFolder(WSourceFolder).SubFolders
    CopyFile WSourceFolder,WDestFolder
    For Each Folder In SubF
  	    
		If Not FSO.FolderExists(WDestFolder & Folder.Name) Then
			FSO.CreateFolder(WDestFolder & Folder.Name)
		End If

		AllFolders WSourceFolder & "\" & Folder.Name,WDestFolder & Folder.Name & "\" 

		
    Next

    AllFolders = Rezult
End Function

Function CopyFile(src,dest)
    'MsgBox src&"--" &dest
	For Each File In FSO.GetFolder(src).Files
		If Right(File.Name, 12) = ". CATDrawing" or Right(File.Name, 4) = ".doc" or Right(File.Name, 11) = ".CATDrawing" Then
			iReadFile=iReadFile+1
 			' TRUE-перезаписать файл
			FSO.CopyFile File.Path, dest & File.Name,True
			If Err.Number<>0 Then
				LogFile.WriteLine Now & ": " & Err.Description & " " & File.Name
				Err.Clear
			End If
		End If
	Next

End Function
