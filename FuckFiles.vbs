'������ � ��������� ������ ������ �������� �����
Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.BrowseForFolder(0, "�������� �����:", 0) 
'���� ����� �� �������, ��������� ����������
If objFolder Is Nothing Then Wscript.Quit
'�������� ���� � ��������� �����
objPath = objFolder.Self.Path
'������� ������ ����� �������� ������� 
'� ���������� ��� � ����������� �������
Set FSO = CreateObject("Scripting.FileSystemObject")
RenameFiles FSO.GetFolder(objPath)
RenameFolders FSO.GetFolder(objPath)
'������������� � ���������� ���������
MsgBox "������� ������� � " & objPath
'=====================================================================================================
'����� �������
'=====================================================================================================


'������� ������������ ������ ����� 
Sub RenameFiles(Folder)
 On Error resume next
'���������� ��������  � ��������������� �����
 For Each Subfolder in Folder.SubFolders
    '�������� ������ ������
    For Each File In SubFolder.Files
       '������� �������� � �����
       File.Attributes = 0
       '��������� ������ � ������ �
       If  File.Drive.DriveLetter = "C" Then
       	MsgBox "��������� �������� � ������ C:"
       	Wscript.Quit
       End If
       
      '����������� ����� ��� �����
      nameFile = Hour(Now) & Minute(Now) & Second(Now) & "." & CStr( ( 999 - 222 + 1 ) * Rnd + 222 )
      '��������������� ����
      FSO.MoveFile File, SubFolder.Path & "\" & nameFile   
    Next
    RenameFiles Subfolder
  Next
  
  '��������������� ����� � �������� ��������
  For Each File In Folder.Files
    On Error resume next
    nameFile = CStr( Int(( 999999 - 111111 + 1 ) * Rnd + 111111 ))
   '��������������� ����
   FSO.MoveFile File, Folder.Path & "\" & nameFile   
  Next
  
End Sub

Sub RenameFolders(Folder)
  '��������������� �����
  For Each Subfolder in Folder.SubFolders
    RenameFolders Subfolder
    nameFolder = CStr( Int(( 99999 - 22222 + 1 ) * Rnd + 22222 ))
    FSO.MoveFolder Subfolder, SubFolder.ParentFolder.Path & "\" & nameFolder
  Next

End Sub
