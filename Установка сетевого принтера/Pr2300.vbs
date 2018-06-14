'
'IP адрес принтера
ip = "10.57.131.55"

ipComp = "10.57.131.29"
'Логин админа
strUser = "ce\CETL_Andrey"

'Пароль
strPassword = "l.k.j987TY"

'Имя драйвера
driverName = "HP LaserJet 500 MFP M525 PCL 6"

'Путь к inf файлу
inf = "U:\Soft\Drivers\Принтер\HP 500\hpzid4vp.inf"

'Имя принтера (любое)
namePrinter = "HP LaserJet 500 MFP M525 PCL 6"

Dim objLocator: Set objLocator = CreateObject("WbemScripting.SWbemLocator")


Set objWMIService = objLocator.ConnectServer (ipComp, "root\cimv2", strUser, strPassword)  

'Set objWMIService = GetObject("winmgmts:" _
'    & "{impersonationLevel=impersonate}!\\" & ip & "\root\cimv2")
Set objNewPort = objWMIService.Get _
    ("Win32_TCPIPPrinterPort").SpawnInstance_

objNewPort.Name = "IP_" & ip
objNewPort.Protocol = 1
objNewPort.HostAddress = ip
objNewPort.PortNumber = "9100"
objNewPort.SNMPEnabled = False
objNewPort.Put_

Set objDriver = objWMIService.Get("Win32_PrinterDriver")
objWMIService.Security_.Privileges.AddAsString "SeLoadDriverPrivilege", True

objDriver.Name = driverName
objDriver.Infname = inf
errResult = objDriver.AddPrinterDriver(objDriver)

Set objPrinter = objWMIService.Get("Win32_Printer").SpawnInstance_

objPrinter.DriverName = driverName
objPrinter.PortName   = "IP_" & ip
objPrinter.Location = "office" 
objPrinter.DeviceID   = namePrinter
objPrinter.Network = True
objPrinter.Put_