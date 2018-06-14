Dim fs, f, f1, fsubF, sarchiv,WshShell, currentDate, sNameArchiv,limitFiles

'-----------------------------------------------------------------------------------
SourceFolder = "E:\Catia-UZL\�������"
DestFolder = "C:\1\"
sourceArchiv = """%ProgramFiles%\7-Zip\7z.exe"""
sLogStr = "Compressing 7-Zip"
currentDate = Replace(CStr(Date),".","")
REPORT = "report.txt"  '���� �������
limitFiles = 5 '������� ���� �������
'----------------------------------------------------------------------------------- 

set WshShell = WScript.CreateObject("WScript.Shell")
'������� ������ ����� ��� �������������
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.GetFolder(SourceFolder)
Set fsubF = f.SubFolders
'���� �������������
For Each f1 in fsubF
    sNameArchiv = currentDate & "-" & f1.name
	sF = SourceFolder & "\" & f1.name
	sRun = sourceArchiv & " a   -tzip -ssw -mx7 """ & DestFolder & sNameArchiv & ".7z"" """ & sF & ""
    ret = WshShell.Run(sRun, 7, TRUE)
	
	'���������
    Dim msg
    Select Case ret
    Case 0
	  msg = "Ok"
    Case 1
	  msg = "��������� ����� ���� ������ � ������� �� ��������� � �����"
    Case 2
	  msg = "������ ��� �������� ������"
    Case 7
	  msg = "������ � ��������� ������"
    Case 8
	  msg = "������ - ������������ ������"
    Case 255
	  msg = "������ - �������� ������ ���� �������� �������������"
    Case Else
	  msg = "������ ��� �������� ������, ��� " & ret
    End Select
	Log sNameArchiv & ".7z" & ": " & msg
Next

'�������� ������ ������� 
Set f = fs.GetFolder(SourceFolder)
Set fsubF = f.Files
strComputer = "."

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set fsubF = objWMIService.ExecQuery _
    ("Select * from CIM_DataFile Where Drive='C:' And Path = '\\1\\' And Name Like'%-%'")

For Each f1 in fsubF
	file = Mid(f1.Name,InStr(f1.Name,"-"))
    Set files = objWMIService.ExecQuery("Select * from CIM_DataFile Where Drive='C:' And Path = '\\1\\' And Name Like'%" & file & "%'") 
    if files.Count>limitFiles Then
		DelFile(files)
	end if
Next	
MsgBox "������������� ���������"
'-------------------------------------------------------------------------------------------
Sub DelFile(files)
	cntAll = limitFiles
    For Each ff in files
	    cnt = 1
		tmp = CLng(Left(ff.FileName,8))
        For Each fff in files
			if tmp<CLng(Left(fff.FileName,8)) then
				cnt=cnt+1
			end if
			if cnt>limitFiles then
				ff.Delete
				Exit For
			end if
		next
	next 
End Sub

'�������� ��������� � ������
Sub Log(msg)
	Const APPEND = 8 '�������� � ����� �����
	Dim fso, f
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.OpenTextFile(REPORT, APPEND, True)
	f.WriteLine Now & " " & msg
	f.Close
End Sub

