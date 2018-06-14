strComputer = "10.57.161.50"
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colSoftware = objWMIService.ExecQuery ("Select * from Win32_Product where name='Check Point VPN-1 SecuRemote/SecureClient NGX R60 HFA3'")
 MsgBox colSoftware.Count
For Each objSoftware in colSoftware
    MsgBox objSoftware.Caption
    'objSoftware.Uninstall()
Next