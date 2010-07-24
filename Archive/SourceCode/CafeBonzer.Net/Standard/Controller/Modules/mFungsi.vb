Option Strict Off
Option Explicit On
Module mCommand
	Public Echo As Boolean
	
	'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
	' NETWORK FUNCTIONS
	'
	'
	'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Fungsi arahan yang diterima dari agent
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Sub FnMesej(ByRef arahan As String)
		Dim strNama, strMsg As String
		strNama = Mid(arahan, 8, InStr(8, arahan, ":") - 8)
		strMsg = Mid(arahan, InStr(8, arahan, ":") + 1)
		If CbConsole = True Then
			If FrmSysConsole.DefInstance.Text1.Text <> "" Then FrmSysConsole.DefInstance.List1.Items.Add(FrmSysConsole.DefInstance.Text1.Text)
			FrmSysConsole.DefInstance.wr(strNama & ">" & strMsg)
			If Echo = True Then FrmSysConsole.DefInstance.wr((FrmSysConsole.DefInstance.Text2.Text)) Else FrmSysConsole.DefInstance.Text1.Text = ""
			Exit Sub
		End If
		FrmAgnMsg.DefInstance.server.Text = strNama
		FrmAgnMsg.DefInstance.rcv.Text = strMsg
		If CbMsgRcv <> True And CbConsole = False Then CbMsgRcv = True : FrmAgnMsg.DefInstance.Show()
	End Sub
	
	
	'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
	' COMMANDS
	'
	'
	'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
	'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
	' [CmdName] - Return Command Name
	'
	'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
	Function CmdName(ByRef DataString As String) As String
		Dim l_dataLen, l As Integer
		
		l_dataLen = Len(DataString)
		If l_dataLen = 0 Then Exit Function
		
		For l = 1 To l_dataLen
			If Mid(DataString, l, 1) = ":" Then
				CmdName = Mid(DataString, 2, l - 2)
				Exit Function
			End If
		Next l
		
		CmdName = Mid(DataString, 2)
	End Function
	
	'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
	' [CmdValue] - Return Command Value
	'
	'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
	Function CmdValue(ByRef DataString As String) As String
		Dim l_dataLen, a As Integer
		
		l_dataLen = Len(DataString)
		If l_dataLen = 0 Then Exit Function
		
		For a = 1 To l_dataLen
			If Mid(DataString, a, 1) = ":" Then
				CmdValue = Mid(DataString, a + 1)
				Exit Function
			End If
		Next a
	End Function
	
	'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
	' [SubVal] - Return Sub Command Value
	'
	'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
	'FIXIT: Declare 'Val' and 'Default' with an early-bound data type                          FixIT90210ae-R1672-R1B8ZE
	'UPGRADE_NOTE: Default was upgraded to Default_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
	Function SubVal(ByRef DataString As String, ByRef SubCmdName As String, Optional ByRef Default_Renamed As Object = "", Optional ByRef Special As Integer = 0) As Object
		On Error GoTo ErrInt
		Dim s_CmdVal As String
		Dim l_PosCmdName As Integer
		Dim l_PosCmdDataA, l_PosCmdDataB As Integer
		' Note
		'   format bagi sub command
		'   {subcmdname|'data'}{subcmdname2|'data'}
		'
		
		l_PosCmdName = InStr(1, DataString, "{" & LCase(SubCmdName))
		If l_PosCmdName = 0 Then GoTo ErrInt
		l_PosCmdDataA = InStr(l_PosCmdName, DataString, "|'") + 2
		l_PosCmdDataB = InStr(l_PosCmdDataA, DataString, "'}")
		SubVal = Mid(DataString, l_PosCmdDataA, l_PosCmdDataB - l_PosCmdDataA)
		'UPGRADE_WARNING: Couldn't resolve default property of object SubVal. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Default_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object SubVal. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If SubVal = "" Then SubVal = Default_Renamed
		Exit Function
ErrInt: 
		'UPGRADE_WARNING: Couldn't resolve default property of object Default_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object SubVal. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		SubVal = Default_Renamed
	End Function
	
	Function SubBuild(ByRef DataName As String, ByRef DataValue As String) As String
		SubBuild = "{" & DataName & "|'" & DataValue & "'}"
	End Function
	
	
	
	'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
	' CONSOLE
	'
	'
	'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
	Public Sub CurrentHook()
		Dim idx As Integer
		'UPGRADE_WARNING: Couldn't resolve default property of object AgentGetIndex(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		idx = AgentGetIndex((FrmSysConsole.DefInstance.CurSocket))
		If idx = 0 Then
			FrmSysConsole.DefInstance.wr("No socket currently hook !")
		Else
			FrmSysConsole.DefInstance.wr("Socket currently hook to " & FrmMain.DefInstance.Lv1.ListItems(idx).Text)
		End If
	End Sub
	
	Public Sub Hook(ByRef StationName As Object)
		Dim idx As Integer
		'UPGRADE_WARNING: Couldn't resolve default property of object AgentGetIndexB(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		idx = AgentGetIndexB(StationName)
		If idx = 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object StationName. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			FrmSysConsole.DefInstance.wr("Station not exist ! - " & StationName)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object FrmMain.Lv1.ListItems().Tag. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			FrmSysConsole.DefInstance.CurSocket = FrmMain.DefInstance.Lv1.ListItems(idx).Tag
			'UPGRADE_WARNING: Couldn't resolve default property of object StationName. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			FrmSysConsole.DefInstance.wr("Hooking socket success > " & StationName)
		End If
	End Sub
	
	Public Sub DisEcho(ByRef Param As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object Param. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If Param = "" Then Exit Sub
		'UPGRADE_WARNING: Couldn't resolve default property of object Param. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If Left(Param, 1) = "1" Then
			Echo = True
			FrmSysConsole.DefInstance.wr("Echo enable")
		Else
			Echo = False
			FrmSysConsole.DefInstance.wr("Echo disable")
		End If
	End Sub
	
	Public Sub DkeyVar(ByRef Param As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object Param. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If Param = "" Then Exit Sub
		'UPGRADE_WARNING: Couldn't resolve default property of object Param. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		CbDrvStr = Param & ":"
		FrmSysConsole.DefInstance.wr("DiskKey drive set to " & CbDrvStr)
	End Sub
	
	Public Sub SendMesej(ByRef Param As String, ByRef Sck As Integer)
		Dim Nama, Mesej As String
		If InStr(1, Param, ":") <> 0 Then
			Nama = Mid(Param, 1, InStr(1, Param, ":") - 1)
			Mesej = Mid(Param, InStr(1, Param, ":") + 1)
			Send(Sck, "//mesej:" & Nama & ":" & Mesej)
		Else
			Nama = "Server"
			Mesej = Param
			Send(Sck, "//mesej:Server:" & Param)
		End If
		FrmSysConsole.DefInstance.wr(Nama & ">" & Mesej)
	End Sub
	
	' kita selalu tidak mahu terjadi sesuatu yang tidak kita suka
	' dan kita mesti selalu bersedia menghadapi sesuatu yang tidak kita suka
	' azri jamil - oct,2002
End Module