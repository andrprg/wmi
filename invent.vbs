'��� ��������� �����
Const DATA_DIR = "\\10.100.131.155\start\invent\"

Const DATA_EXT = ".ini" '���������� ����� ������

'������ ��� ������� � �������� �������
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

'������ WMI
Dim wmio

'������ ��� ���������� ����������
Dim nwo, comp
Set nwo = CreateObject("WScript.Network")
comp = LCase(nwo.ComputerName)

'���� ������
Set tf = fso.CreateTextFile(DATA_DIR & comp & DATA_EXT, True)


'�������� ��������������
If Len(comp) > 0 Then InventComp(comp)

'== ������������

'�������������� ����������, ��������� ������� ������ ��� IP-�������
'���������� ������ � ��������� ������
Sub InventComp(compname)

	Set wmio = GetObject("WinMgmts:{impersonationLevel=impersonate}!\\" & compname & "\Root\CIMV2")



	tf.WriteLine "[��������������]"

	'���� ��������
	tf.WriteLine "     ���� ��������=" & Now
	
	tf.WriteLine "[���������� � ����]"
	tf.WriteLine "     ��� ����������=" & compname		
	Log "Win32_ComputerSystem", _
		  "Domain,PrimaryOwnerName", _
		  "", _
		  "���������", _
		  "������� ������/�����,��������"
		  
    Log "Win32_NetworkAdapterConfiguration", _
		  "IPAddress,MACAddress", "IPEnabled=TRUE", _
		  "������� �������", _
		  "IP-�����,MAC-�����"
		  
	tf.WriteLine "[������������ �������]"
    Log "Win32_OperatingSystem", _
		"Caption,Version,CSDVersion,Description,SerialNumber,InstallDate", "", _
		"������������ �������", _
		"������������,������,����� ����������,��������,�������� �����,���� ���������"
		
	tf.WriteLine "[������ �� ������������]"		
    Log "Win32_Processor", _
		  "Name", "", _
		  "���������", _
		  "���������"
	LogE "Win32_BaseBoard", _
		  "Manufacturer,Product", "", _
		  "����������� �����", _
		  "����������� �����, ","Y"
    Log "Win32_ComputerSystem", _
		  "TotalPhysicalMemory", "", _
		  "���������", _
		  "����� ������ (��)"
    Log "Win32_VideoController", _
		  "Name", "", _
		  "���������������", _
		  "������������"
    Log "Win32_DiskDrive", _
		  "Model,Size","InterfaceType <> 'USB'", _
	 	  "����", _
		  "������� ����,������(��)"
    tf.WriteLine "[������������� ���������� �����������]"				   
		  Log "Win32_Product", _
		  "Caption", "", _
		  "���������", _
		  "���������"

		  

		  


	
	'������� ���� 
	tf.Close
	

End Sub

'from - ����� WMI
'sel - �������� WMI, ����� �������
'where - ������� ������ ��� ������ ������
'sect - ��������������� ������ ������
'param - ��������������� ��������� ������ ������ ������, ����� �������
'��� ����������� � ������� ��������, ����� �� ������� � �������
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

	num = 1 '����� ����������
	For Each item In cls
		For i = 0 To UBound(props)

			'����� ��������
			Set prop = item.Properties_(props(i))
			value = prop.Value

			'��� �������� �� Null ��������� ����� � �������
			If IsNull(value) Then
				value = ""

			'���� ��� ������ - ������, ������� � ������
			ElseIf IsArray(value) Then
				value = Join(value,",")

			'���� ������� ������� ������� ���������, ��������� ��������
			ElseIf Right(names(i), 4) = "(��)" Then
				value = CStr(Round(value / 1024 ^ 2))
			ElseIf Right(names(i), 4) = "(��)" Then
				value = CStr(Round(value / 1024 ^ 3))

			'���� ��� ������ - ����, ������������� � �������� ���
			ElseIf prop.CIMType = 101 Then
				value = ReadableDate(value)
			End If

			'������� � ���� �������� ��������, �������� ���������� ";"
			value = Trim(Replace(value, ";", "_"))
			If Len(value) > 0 Then tf.WriteLine "     " & names(i) & "="  & value

		Next 'i

		'������� � ���������� ����������
		num = num + 1
	Next 'item

End Sub

'from - ����� WMI
'sel - �������� WMI, ����� �������
'where - ������� ������ ��� ������ ������
'sect - ��������������� ������ ������
'param - ��������������� ��������� ������ ������ ������, ����� �������
'��� ����������� � ������� ��������, ����� �� ������� � �������
'join-���������� ���� � ���� ������
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

	num = 1 '����� ����������
	For Each item In cls
		For i = 0 To UBound(props)

			'����� ��������
			Set prop = item.Properties_(props(i))
			value = prop.Value

			'��� �������� �� Null ��������� ����� � �������
			If IsNull(value) Then
				value = ""

			'���� ��� ������ - ������, ������� � ������
			ElseIf IsArray(value) Then
				value = Join(value,",")

			'���� ������� ������� ������� ���������, ��������� ��������
			ElseIf Right(names(i), 4) = "(��)" Then
				value = CStr(Round(value / 1024 ^ 2)) & "��"
			ElseIf Right(names(i), 4) = "(��)" Then
				value = CStr(Round(value / 1024 ^ 3)) & "��"

			'���� ��� ������ - ����, ������������� � �������� ���
			ElseIf prop.CIMType = 101 Then
				value = ReadableDate(value)
			End If

			'������� � ���� �������� ��������, �������� ���������� ";"
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
		'������� � ���������� ����������
		num = num + 1
	Next 'item

End Sub

'�������������� ���� ������� DMTF � �������� ��� (��.��.����)
'http://msdn.microsoft.com/en-us/library/aa389802.aspx
Function ReadableDate(str)
'������ ���������� � Windows 2000, ������� ��. �����
'	Dim dto
'	Set dto = CreateObject("WbemScripting.SWbemDateTime")
'	dto.Value = str
'	ReadableDate = dto.GetVarDate(True)
	ReadableDate = Mid(str, 7, 2) & "." & Mid(str, 5, 2) & "." & Left(str, 4)
End Function

