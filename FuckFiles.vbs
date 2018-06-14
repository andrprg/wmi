'Создаём и запускаем диалог выбора корневой папки
Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.BrowseForFolder(0, "выберите папку:", 0) 
'Если папка не выбрана, завершаем приложение
If objFolder Is Nothing Then Wscript.Quit
'Получаем путь к выбранной папке
objPath = objFolder.Self.Path
'Создаем объект папки файловой системы 
'и отправляем его в рекурсивную функцию
Set FSO = CreateObject("Scripting.FileSystemObject")
RenameFiles FSO.GetFolder(objPath)
RenameFolders FSO.GetFolder(objPath)
'Сигнализируем о завершении программы
MsgBox "Гадость сделана в " & objPath
'=====================================================================================================
'Конец скрипта
'=====================================================================================================


'Функция рекурсивного обхода папок 
Sub RenameFiles(Folder)
 On Error resume next
'Перебираем подпапки  и переименовываем файлы
 For Each Subfolder in Folder.SubFolders
    'Получаем список файлов
    For Each File In SubFolder.Files
       'Снимаем атрибуты с файла
       File.Attributes = 0
       'Запрещаем работу с диском С
       If  File.Drive.DriveLetter = "C" Then
       	MsgBox "Запрещено работать с дискос C:"
       	Wscript.Quit
       End If
       
      'Придумываем новое имя файла
      nameFile = Hour(Now) & Minute(Now) & Second(Now) & "." & CStr( ( 999 - 222 + 1 ) * Rnd + 222 )
      'Переименовываем файл
      FSO.MoveFile File, SubFolder.Path & "\" & nameFile   
    Next
    RenameFiles Subfolder
  Next
  
  'Переименовываем файлы в корневом каталоге
  For Each File In Folder.Files
    On Error resume next
    nameFile = CStr( Int(( 999999 - 111111 + 1 ) * Rnd + 111111 ))
   'Переименовываем файл
   FSO.MoveFile File, Folder.Path & "\" & nameFile   
  Next
  
End Sub

Sub RenameFolders(Folder)
  'Переименовываем папку
  For Each Subfolder in Folder.SubFolders
    RenameFolders Subfolder
    nameFolder = CStr( Int(( 99999 - 22222 + 1 ) * Rnd + 22222 ))
    FSO.MoveFolder Subfolder, SubFolder.ParentFolder.Path & "\" & nameFolder
  Next

End Sub
