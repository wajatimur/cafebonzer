Option Strict Off
Option Explicit On
Module mNetwork
	'==================================================================
	' Aplication codename : CafeBonzer
	' Programmer          : Azri Jamil a.k.a wajatimur
	' Module Name         : Network
	' Description         : Network Engine
	'==================================================================
	Public StackNetData As New Collection
	Public DataArray() As dNetData
	Public Structure dNetData
		Dim SockIndex As Integer
		Dim Data As String
	End Structure
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' hidupkan server dan tunggu setiap sambungan
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Sub NetUp()
		On Error GoTo ErrTrap
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		FrmSysHost.DefInstance.Socket(0).LocalPort = SetAmbil("porttempatan")
		FrmSysHost.DefInstance.Socket(0).Listen()
		FrmSysHost.DefInstance.NetTimer.Enabled = True
		ReDim DataArray(0)
		Exit Sub
		
ErrTrap: 
		ErrLog(Err, "mNet | NetUp")
	End Sub
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' terminate all network
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Sub NetClose()
		Dim Sck As Object
		Dim w As Short
		On Error GoTo ErrTrap
		FrmSysHost.DefInstance.NetTimer.Enabled = False
		For w = 0 To FrmSysHost.DefInstance.Socket.Count - 1
			Sck = FrmSysHost.DefInstance.Socket(w)
			'UPGRADE_WARNING: Couldn't resolve default property of object Sck.Cleanup. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			Sck.Cleanup()
			'UPGRADE_NOTE: Object Sck may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
			Sck = Nothing
		Next w
		Exit Sub
		
ErrTrap: 
		ErrLog(Err, "mNet | NetClose")
		Resume Next
	End Sub
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Ping client
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Sub Ping(ByRef SockIndex As Integer)
		Send(SockIndex, "//hey")
	End Sub
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Hantar data (masukkan kedalam "query list")
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Sub Send(ByRef SockIndex As Integer, ByRef Data As String)
		Dim TmpDst As New clsDataStore
		
		'Dim lUbnd As Long
		'lUbnd = UBound(DataArray) + 1
		'ReDim Preserve DataArray(lUbnd)
		TmpDst.Name = CStr(SockIndex)
		TmpDst.Add(SockIndex, "sockindex")
		TmpDst.Add(Data, "data")
		StackNetData.Add(TmpDst)
		'DataArray(lUbnd).SockIndex = SockIndex
		'DataArray(lUbnd).Data = Data
		'UPGRADE_NOTE: Object TmpDst may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
		TmpDst = Nothing
	End Sub
End Module