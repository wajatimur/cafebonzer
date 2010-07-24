Option Strict Off
Option Explicit On
Module mDiskKey
	Private Declare Function GetVolumeInformation Lib "kernel32.dll"  Alias "GetVolumeInformationA"(ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Short, ByRef lpVolumeSerialNumber As Integer, ByRef lpMaximumComponentLength As Integer, ByRef lpFileSystemFlags As Integer, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Integer) As Integer
	
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Membina DiskKey
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Function CreateDiskKey(ByRef Name As String, ByRef Key As String, ByRef CbDrvStr As String) As Boolean
		On Error GoTo ErrInt
		Dim Buffer As String
		Dim Sn As Integer
		Sn = GetSerial(CbDrvStr)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Crypt(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Buffer = Crypt(Sn & "/:" & Name & "/:" & Key)
		
		FileOpen(1, CbDrvStr & "\nmtdkey.k", OpenMode.Output)
		PrintLine(1, Buffer)
		FileClose(1)
		SetAttr(CbDrvStr & "\nmtdkey.k", FileAttribute.Hidden)
		CreateDiskKey = True
		Exit Function
		
ErrInt: 
		CreateDiskKey = False
	End Function
	
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Mengesahkan DiskKey
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Function ValidateDisk(ByRef CbDrvStr As String) As Boolean
		Dim x As Object
		On Error GoTo ErrInt
		Dim Sn As Integer
		Dim SnF As String
		Dim sKeyFile As String
		
		ValidateDisk = False
		sKeyFile = "\nmtdkey.k"
		
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
		If Len(Dir(CbDrvStr & sKeyFile, FileAttribute.Hidden)) = 0 Then Exit Function
		
		'get the disk volume
		Sn = GetSerial(CbDrvStr)
		'get the internal disk volume
		FileOpen(1, CbDrvStr & sKeyFile, OpenMode.Input)
		SnF = LineInput(1)
		FileClose(1)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Crypt(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		SnF = Crypt(SnF)
		'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		x = InStr(1, SnF, "/:")
		'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		SnF = Left(SnF, x - 1)
		
		If Sn = CDbl(SnF) Then ValidateDisk = True
		Exit Function
		
ErrInt: 
		ValidateDisk = False
	End Function
	
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Mendapatkan Nama Pendaftaran Dari DiskKey
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Function GetName(ByRef CbDrvStr As Object) As String
		Dim y As Object
		Dim x As Object
		Dim Buffer As String
		Dim TmpName As String
		
		'UPGRADE_WARNING: Couldn't resolve default property of object CbDrvStr. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		FileOpen(1, CbDrvStr & "\nmtdkey.k", OpenMode.Input)
		Buffer = LineInput(1)
		'UPGRADE_WARNING: Couldn't resolve default property of object Crypt(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Buffer = Crypt(Buffer)
		'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		x = InStr(1, Buffer, "/:")
		'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		y = InStr(x + 2, Buffer, "/:")
		'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		TmpName = Mid(Buffer, x + 2, y - (x + 2))
		GetName = TmpName
		FileClose(1)
	End Function
	
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Mendapatkan Key Pendaftaran dari DiskKey
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Function GetKey(ByRef CbDrvStr As Object) As String
		Dim x As Object
		Dim Buffer As String
		Dim TmpKey As String
		
		'UPGRADE_WARNING: Couldn't resolve default property of object CbDrvStr. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		FileOpen(1, CbDrvStr & "\nmtdkey.k", OpenMode.Input)
		Buffer = LineInput(1)
		'UPGRADE_WARNING: Couldn't resolve default property of object Crypt(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Buffer = Crypt(Buffer)
		'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		x = InStrRev(Buffer, "/:", -1)
		'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		TmpKey = Right(Buffer, Len(Buffer) - x - 1)
		GetKey = TmpKey
		FileClose(1)
	End Function
	
	
	
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Internal Function - not for outdoor use.
	' Please caution
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Private Function GetSerial(ByRef CbDrvStr As String) As Integer
		Dim SerialNum As Integer
		Dim Res As Integer
		Dim Temp1 As String
		Dim Temp2 As String
		Temp1 = New String(Chr(0), 255)
		Temp2 = New String(Chr(0), 255)
		Res = GetVolumeInformation(CbDrvStr, Temp1, Len(Temp1), SerialNum, 0, 0, Temp2, Len(Temp2))
		GetSerial = SerialNum
	End Function
	Private Function Crypt(ByRef Text As String) As Object
		Dim Tmp As String
		Dim Itg As Short
		For Itg = 1 To Len(Text)
			Tmp = Tmp & Chr(Asc(Mid(Text, Itg, 1)) Xor 4)
		Next Itg
		'UPGRADE_WARNING: Couldn't resolve default property of object Crypt. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Crypt = Tmp
	End Function
End Module