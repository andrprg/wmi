'где сохранять отчет
Const DATA_DIR = "\\10.100.131.155\start\invent\"

Const DATA_EXT = ".ini" 'расширение файла отчета

'объект для доступа к файловой системе
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

'объект WMI
Dim wmio

'узнать имя локального компьютера
Dim nwo, comp
Set nwo = CreateObject("WScript.Network")
comp = LCase(nwo.ComputerName)

'файл отчета
Set tf = fso.CreateTextFile(DATA_DIR & comp & DATA_EXT, True)


'провести инвентаризацию
If Len(comp) > 0 Then InventComp(comp)

'== ПОДПРОГРАММЫ

'инвентаризация компьютера, заданного сетевым именем или IP-адресом
'сохранение отчета с указанным именем
Sub InventComp(compname)

	Set wmio = GetObject("WinMgmts:{impersonationLevel=impersonate}!\\" & compname & "\Root\CIMV2")



	tf.WriteLine "[Инвентаризация]"

	'дата проверки
	tf.WriteLine "     Дата проверки=" & Now
	
	tf.WriteLine "[Информация о сети]"
	tf.WriteLine "     Имя компьютера=" & compname		
	Log "Win32_ComputerSystem", _
		  "Domain,PrimaryOwnerName", _
		  "", _
		  "Компьютер", _
		  "Рабочая группа/домен,Владелец"
		  
    Log "Win32_NetworkAdapterConfiguration", _
		  "IPAddress,MACAddress", "IPEnabled=TRUE", _
		  "Сетевой адаптер", _
		  "IP-адрес,MAC-адрес"
		  
	tf.WriteLine "[Операционная система]"
    Log "Win32_OperatingSystem", _
		"Caption,Version,CSDVersion,Description,SerialNumber,InstallDate", "", _
		"Операционная система", _
		"Наименование,Версия,Пакет обновления,Описание,Серийный номер,Дата установки"
		
	tf.WriteLine "[Сводка по оборудованию]"		
    Log "Win32_Processor", _
		  "Name", "", _
		  "Процессор", _
		  "Процессор"
	LogE "Win32_BaseBoard", _
		  "Manufacturer,Product", "", _
		  "Материнская плата", _
		  "Материнская плата, ","Y"
    Log "Win32_ComputerSystem", _
		  "TotalPhysicalMemory", "", _
		  "Компьютер", _
		  "Объем памяти (Мб)"
    Log "Win32_VideoController", _
		  "Name", "", _
		  "Видеоконтроллер", _
		  "Видеоадаптер"
    Log "Win32_DiskDrive", _
		  "Model,Size","InterfaceType <> 'USB'", _
	 	  "Диск", _
		  "Жесткий диск,Размер(Гб)"
    tf.WriteLine "[Установленное програмное обеспечение]"				   
		  Log "Win32_Product", _
		  "Caption", "", _
		  "Программа", _
		  "Программа"

		  

		  


	
	'закрыть файл 
	tf.Close
	

End Sub

'from - класс WMI
'sel - свойства WMI, через запятую
'where - условие отбора или пустая строка
'sect - соответствующая секция отчета
'param - соответствующие параметры внутри секции отчета, через запятую
'для отображения в кратных единицах, нужно их указать в скобках
Sub Log(from, sel, where, sect, param)

	Const RETURN_IMMEDIATELY = 16
	Const FORWARD_ONLY = 32

	Dim query, cls, item, prop
	query = "Select " & sel & " From " & from

	If Len(where) > 0 Then query = query & " Where " & where
	Set cls = wmio.ExecQuery(query,, RETURN_IMMEDIATELY + FORWARD_ONLY)

	Dim props, names, num, value
	props = Split(sel, ",")
	names = Split(param, ",")

	num = 1 'номер экземпляра
	For Each item In cls
		For i = 0 To UBound(props)

			'взять значение
			Set prop = item.Properties_(props(i))
			value = prop.Value

			'без проверки на Null возможнен вылет с ошибкой
			If IsNull(value) Then
				value = ""

			'если тип данных - массив, собрать в строку
			ElseIf IsArray(value) Then
				value = Join(value,",")

			'если указана кратная единица измерения, перевести значение
			ElseIf Right(names(i), 4) = "(Мб)" Then
				value = CStr(Round(value / 1024 ^ 2))
			ElseIf Right(names(i), 4) = "(Гб)" Then
				value = CStr(Round(value / 1024 ^ 3))

			'если тип данных - дата, преобразовать в читаемый вид
			ElseIf prop.CIMType = 101 Then
				value = ReadableDate(value)
			End If

			'вывести в файл непустое значение, заменить спецсимвол ";"
			value = Trim(Replace(value, ";", "_"))
			If Len(value) > 0 Then tf.WriteLine "     " & names(i) & "="  & value

		Next 'i

		'перейти к следующему экземпляру
		num = num + 1
	Next 'item

End Sub

'from - класс WMI
'sel - свойства WMI, через запятую
'where - условие отбора или пустая строка
'sect - соответствующая секция отчета
'param - соответствующие параметры внутри секции отчета, через запятую
'для отображения в кратных единицах, нужно их указать в скобках
'join-объединять поля в одну строку
Sub LogE(from, sel, where, sect, param,join)

	Const RETURN_IMMEDIATELY = 16
	Const FORWARD_ONLY = 32

	Dim query, cls, item, prop,jn 
	query = "Select " & sel & " From " & from

	If Len(where) > 0 Then query = query & " Where " & where
	Set cls = wmio.ExecQuery(query,, RETURN_IMMEDIATELY + FORWARD_ONLY)

	Dim props, names, num, value
	props = Split(sel, ",")
	names = Split(param, ",")

	num = 1 'номер экземпляра
	For Each item In cls
		For i = 0 To UBound(props)

			'взять значение
			Set prop = item.Properties_(props(i))
			value = prop.Value

			'без проверки на Null возможнен вылет с ошибкой
			If IsNull(value) Then
				value = ""

			'если тип данных - массив, собрать в строку
			ElseIf IsArray(value) Then
				value = Join(value,",")

			'если указана кратная единица измерения, перевести значение
			ElseIf Right(names(i), 4) = "(Мб)" Then
				value = CStr(Round(value / 1024 ^ 2)) & "Мб"
			ElseIf Right(names(i), 4) = "(Гб)" Then
				value = CStr(Round(value / 1024 ^ 3)) & "Гб"

			'если тип данных - дата, преобразовать в читаемый вид
			ElseIf prop.CIMType = 101 Then
				value = ReadableDate(value)
			End If

			'вывести в файл непустое значение, заменить спецсимвол ";"
			value = Trim(Replace(value, ";", "_"))
			If Len(value) > 0 Then
				If join="N" Then 
					tf.WriteLine "     " & names(i) & "="  & value  
				ElseIf  join="Y" Then 
					jn=jn & " " & value 
				End if
            End If

		Next 'i
        If join="Y" Then    tf.WriteLine "     " & param & "="  & jn  
		'перейти к следующему экземпляру
		num = num + 1
	Next 'item

End Sub

'преобразование даты формата DMTF в читаемый вид (ДД.ММ.ГГГГ)
'http://msdn.microsoft.com/en-us/library/aa389802.aspx
Function ReadableDate(str)
'объект недоступен в Windows 2000, поэтому см. далее
'	Dim dto
'	Set dto = CreateObject("WbemScripting.SWbemDateTime")
'	dto.Value = str
'	ReadableDate = dto.GetVarDate(True)
	ReadableDate = Mid(str, 7, 2) & "." & Mid(str, 5, 2) & "." & Left(str, 4)
End Function

