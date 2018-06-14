

On Error Resume Next
Set FSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")

FSO.CreateFolder("C:\Documentum\") 
FSO.CreateFolder("C:\Documentum\Scan") 

strStart = WshShell.SpecialFolders.Item("AllUsersStartUp") & "\"
FSO.CopyFile "\\10.100.131.155\start\MsStart.exe", strStart ,True


strDesktop = WshShell.SpecialFolders("AllUsersDesktop")
Set WshURLShortcut =  WshShell.CreateShortcut(strDesktop&"\СЭД ЦентрТелеком.url")
WshURLShortcut.TargetPath = "http://docprod.datacenter.cnt:7777/webtop"
WshURLShortcut.Save

'Добавляем надежные узлы
i=1
Set WshShell = CreateObject("WScript.Shell")

Set s=WshShell.RegRead ("HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges\Range"&i&"\:Range")
Do Until Err.Source="WshShell.RegRead" 
	i=i+1
	s=WshShell.RegRead ("HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges\Range"&i&"\:Range")
Loop

WshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges\Range"&i&"\:Range","docprod.datacenter.cnt","REG_SZ"
WshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges\Range"&i&"\http","2","REG_DWORD"

WshShell.RegWrite "HKCU\Software\Microsoft\Office\11.0\Word\Security\Level","1","REG_DWORD"


'Устанавливаем ассоциацию
Set WshShell = CreateObject("WScript.Shell")
WshShell.RUN "%comspec% /c assoc.tif=MSPaper.Document"


WScript.Echo("Завершено")

16