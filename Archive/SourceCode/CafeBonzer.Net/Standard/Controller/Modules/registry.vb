Option Strict Off
Option Explicit On
Module mRegistry
	Public Const HKEY_CLASSES_ROOT As Integer = &H80000000
	Public Const HKEY_CURRENT_CONFIG As Integer = &H80000005
	Public Const HKEY_CURRENT_USER As Integer = &H80000001
	Public Const HKEY_DYN_DATA As Integer = &H80000006
	Public Const HKEY_LOCAL_MACHINE As Integer = &H80000002
	Public Const HKEY_PERFORMANCE_DATA As Integer = &H80000004
	Public Const HKEY_USERS As Integer = &H80000003
	Public Const ERROR_SUCCESS As Short = 0
	
	' Registry API prototypes
	Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Integer) As Integer
	Declare Function RegCreateKey Lib "advapi32.dll"  Alias "RegCreateKeyA"(ByVal hKey As Integer, ByVal lpSubKey As String, ByRef phkResult As Integer) As Integer
	Declare Function RegDeleteKey Lib "advapi32.dll"  Alias "RegDeleteKeyA"(ByVal hKey As Integer, ByVal lpSubKey As String) As Integer
	Declare Function RegDeleteValue Lib "advapi32.dll"  Alias "RegDeleteValueA"(ByVal hKey As Integer, ByVal lpValueName As String) As Integer
	Declare Function RegOpenKey Lib "advapi32.dll"  Alias "RegOpenKeyA"(ByVal hKey As Integer, ByVal lpSubKey As String, ByRef phkResult As Integer) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1016"'
	Declare Function RegQueryValueEx Lib "advapi32.dll"  Alias "RegQueryValueExA"(ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByRef lpData As Any, ByRef lpcbData As Integer) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1016"'
	Declare Function RegSetValueEx Lib "advapi32.dll"  Alias "RegSetValueExA"(ByVal hKey As Integer, ByVal lpValueName As String, ByVal Reserved As Integer, ByVal dwType As Integer, ByRef lpData As Any, ByVal cbData As Integer) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1016"'
	Declare Function RegEnumKeyEx Lib "advapi32.dll"  Alias "RegEnumKeyExA"(ByVal hKey As Integer, ByVal dwIndex As Integer, ByVal lpName As String, ByRef lpcbName As Integer, ByVal lpReserved As Integer, ByVal lpClass As String, ByRef lpcbClass As Integer, ByRef lpftLastWriteTime As Any) As Integer
	Declare Function RegEnumValue Lib "advapi32.dll"  Alias "RegEnumValueA"(ByVal hKey As Integer, ByVal dwIndex As Integer, ByVal lpValueName As String, ByRef lpcbValueName As Integer, ByVal lpReserved As Integer, ByRef lpType As Integer, ByRef lpData As Byte, ByRef lpcbData As Integer) As Integer
	Public Const REG_SZ As Short = 1 ' Unicode nul terminated string
	Public Const REG_DWORD As Short = 4 ' 32-bit number
	Public Const REG_BINARY As Short = 3 ' Free form binary
	Public Sub SaveKey(ByRef hKey As Integer, ByRef strPath As String)
		Dim r As Object
		Dim Keyhand As Integer
		'UPGRADE_WARNING: Couldn't resolve default property of object r. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		r = RegCreateKey(hKey, strPath, Keyhand)
		'UPGRADE_WARNING: Couldn't resolve default property of object r. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		r = RegCloseKey(Keyhand)
	End Sub
	Public Function GetString(ByRef hKey As Integer, ByRef strPath As String, ByRef strValue As String) As Object
		Dim lValueType As Object
		Dim r As Object
		Dim datatype, Keyhand, lResult As Integer
		Dim strBuf As String
		Dim lDataBufSize As Integer
		Dim intZeroPos As Short
		'UPGRADE_WARNING: Couldn't resolve default property of object r. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		r = RegOpenKey(hKey, strPath, Keyhand)
		'UPGRADE_WARNING: Couldn't resolve default property of object lValueType. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		lResult = RegQueryValueEx(Keyhand, strValue, 0, lValueType, 0, lDataBufSize)
		'UPGRADE_WARNING: Couldn't resolve default property of object lValueType. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If lValueType = REG_SZ Then
			strBuf = New String(" ", lDataBufSize)
			lResult = RegQueryValueEx(Keyhand, strValue, 0, 0, strBuf, lDataBufSize)
			If lResult = ERROR_SUCCESS Then
				intZeroPos = InStr(strBuf, Chr(0))
				If intZeroPos > 0 Then
					GetString = Left(strBuf, intZeroPos - 1)
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object GetString. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					GetString = strBuf
				End If
			End If
		End If
	End Function
	Public Sub SaveString(ByRef hKey As Integer, ByRef strPath As String, ByRef strValue As String, ByRef strdata As String)
		Dim Keyhand, r As Integer
		r = RegCreateKey(hKey, strPath, Keyhand)
		r = RegSetValueEx(Keyhand, strValue, 0, REG_SZ, strdata, Len(strdata))
		r = RegCloseKey(Keyhand)
	End Sub
	Function GetDWord(ByVal hKey As Integer, ByVal strPath As String, ByVal strValueName As String) As Integer
		Dim lValueType, lResult, lBuf As Integer
		Dim r, lDataBufSize, Keyhand As Integer
		r = RegOpenKey(hKey, strPath, Keyhand)
		' Get length/data type
		lDataBufSize = 4
		lResult = RegQueryValueEx(Keyhand, strValueName, 0, lValueType, lBuf, lDataBufSize)
		If lResult = ERROR_SUCCESS Then
			If lValueType = REG_DWORD Then
				GetDWord = lBuf
			End If
		End If
		r = RegCloseKey(Keyhand)
	End Function
	Function SaveDword(ByVal hKey As Integer, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Integer) As Object
		Dim lResult As Integer
		Dim Keyhand As Integer
		Dim r As Integer
		r = RegCreateKey(hKey, strPath, Keyhand)
		lResult = RegSetValueEx(Keyhand, strValueName, 0, REG_DWORD, lData, 4)
		r = RegCloseKey(Keyhand)
	End Function
	Public Function DeleteKey(ByVal hKey As Integer, ByVal strKey As String) As Object
		Dim r As Integer
		r = RegDeleteKey(hKey, strKey)
	End Function
	Public Function DeleteValue(ByVal hKey As Integer, ByVal strPath As String, ByVal strValue As String) As Object
		Dim r As Object
		Dim Keyhand As Integer
		'UPGRADE_WARNING: Couldn't resolve default property of object r. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		r = RegOpenKey(hKey, strPath, Keyhand)
		'UPGRADE_WARNING: Couldn't resolve default property of object r. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		r = RegDeleteValue(Keyhand, strValue)
		'UPGRADE_WARNING: Couldn't resolve default property of object r. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		r = RegCloseKey(Keyhand)
	End Function
	Public Sub EnumKey(ByVal hKey As Integer, ByVal strPath As String, ByRef cResult As Collection)
		Dim Cnt, Keyhand As Integer
		Dim sName As String
		RegOpenKey(hKey, strPath, Keyhand)
		Do 
			sName = New String(vbNullChar, 255)
			If RegEnumKeyEx(Keyhand, Cnt, sName, 255, 0, vbNullString, 0, 0) <> 0 Then Exit Do
			cResult.Add(StripTerminator(sName))
			Cnt = Cnt + 1
		Loop 
		RegCloseKey(Keyhand)
	End Sub
	Public Sub EnumValue(ByVal hKey As Integer, ByVal strPath As String, ByRef cResult As Collection)
		Dim Cnt, Keyhand As Integer
		Dim sName As String
		RegOpenKey(hKey, strPath, Keyhand)
		Do 
			sName = New String(vbNullChar, 255)
			If RegEnumValue(Keyhand, Cnt, sName, 255, 0, 0, 0, 0) <> 0 Then Exit Do
			cResult.Add(StripTerminator(sName))
			Cnt = Cnt + 1
		Loop 
		RegCloseKey(Keyhand)
	End Sub
	Public Function StripTerminator(ByRef sInput As String) As String
		Dim ZeroPos As Short
		ZeroPos = InStr(1, sInput, vbNullChar)
		If ZeroPos > 0 Then
			StripTerminator = Left(sInput, ZeroPos - 1)
		Else
			StripTerminator = sInput
		End If
	End Function
	Public Function GetBinary(ByVal hKey As Integer, ByVal strPath As String, ByVal strValueName As String, ByRef bArray() As Byte) As Boolean
		'Dim bArray() As Byte
		'If GetBinary(KEY, PATH, VALUE, bArray()) = True Then
		'   MsgBox StrConv(bArray, vbUnicode)
		'End If
		Dim lValueType, lResult, lBuf As Integer
		Dim r, lDataBufSize, Keyhand As Integer
		r = RegOpenKey(hKey, strPath, Keyhand)
		' Get length/data type
		lDataBufSize = 0
		'UPGRADE_WARNING: Lower bound of array bArray was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1033"'
		ReDim bArray(1)
		lResult = RegQueryValueEx(Keyhand, strValueName, 0, lValueType, bArray(1), lDataBufSize)
		If lResult > 0 And lValueType = REG_BINARY Then
			'UPGRADE_WARNING: Lower bound of array bArray was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1033"'
			ReDim bArray(lDataBufSize)
			lResult = RegQueryValueEx(Keyhand, strValueName, 0, lValueType, bArray(1), lDataBufSize)
			If lResult = ERROR_SUCCESS Then GetBinary = True
		End If
		r = RegCloseKey(Keyhand)
	End Function
	Public Function SaveBinary(ByVal hKey As Integer, ByVal strPath As String, ByVal strValueName As String, ByRef bStart As Byte, ByRef bLen As Integer) As Boolean
		'Dim bArray(1 To 3) As Byte
		'SaveBinary Key, Path, Value, bArray(1), 3
		Dim lResult As Integer
		Dim Keyhand As Integer
		Dim r As Integer
		r = RegCreateKey(hKey, strPath, Keyhand)
		lResult = RegSetValueEx(Keyhand, strValueName, 0, REG_BINARY, bStart, bLen)
		If lResult = ERROR_SUCCESS Then SaveBinary = True
		r = RegCloseKey(Keyhand)
	End Function
End Module