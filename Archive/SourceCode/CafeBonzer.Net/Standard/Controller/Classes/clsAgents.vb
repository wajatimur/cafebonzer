Option Strict Off
Option Explicit On
Friend Class clsAgents
	Private StackTmpAgent As New Collection
	Private StackAgent As New Collection
	Private StackSocket As New Collection
	
	Public Enum EnCountStatus
		Used = 0
		UnUsed = 1
		Locked = 2
		Unlocked = 3
	End Enum
	
	Public Enum EnEvent
		Agent_Added = 1
		Agent_Removed = 2
		Agent_Disconnect = 3
		Agent_Added_Offline_ = 4
		Info_Updated = 5
	End Enum
	
	Public Event AgentAdded(ByRef Agent As clsAgent)
	Public Event AgentRemove(ByRef Agent As clsAgent)
	Public Event InfoUpdated(ByRef Agent As clsAgent, ByRef InfoType As Integer)
	
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Members | Agents
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Function Agents(ByRef Name As Object) As clsAgent
		Agents = StackAgent.Item(Name)
	End Function
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Members | Agents by Index
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Function AgentsByIndex(ByRef SockIndex As Object) As clsAgent
		Dim CtmpAgn As clsAgent
		
		For	Each CtmpAgn In StackAgent
			'UPGRADE_WARNING: Couldn't resolve default property of object SockIndex. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			If CtmpAgn.AgentSockIndex = SockIndex Then
				AgentsByIndex = CtmpAgn
				Exit Function
			End If
		Next CtmpAgn
		
		'UPGRADE_NOTE: Object CtmpAgn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
		CtmpAgn = Nothing
	End Function
	
	
	
	'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
	' Agent Members
	'
	'
	'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Count Agent
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Function Count() As Integer
		Count = StackAgent.Count()
	End Function
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	'
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Function CountStatus(ByRef Status As EnCountStatus) As Integer
		Dim TmpAgn As clsAgent
		Dim lAgnCnt As Integer
		Select Case Status
			Case 0
				For	Each TmpAgn In StackAgent
					If TmpAgn.AgentStatus = VS(3) Then lAgnCnt = lAgnCnt + 1
				Next TmpAgn
				CountStatus = lAgnCnt
			Case 1
				For	Each TmpAgn In StackAgent
					If TmpAgn.AgentStatus = VS(4) Then lAgnCnt = lAgnCnt + 1
				Next TmpAgn
				CountStatus = lAgnCnt
		End Select
	End Function
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	'
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Function AgentAdd(ByRef Sock As AxSocketWrenchCtrl.AxSocket, ByRef SocketID As Integer) As clsAgent
		Dim CtmpAgn As New clsAgent
		
		Sock.Accept = SocketID
		CtmpAgn.AgnInit(True, Sock)
		
		'UPGRADE_ISSUE: VBControlExtender property Sock.Index was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		StackTmpAgent.Add(CtmpAgn, "#" & CStr(Sock.Index))
		'UPGRADE_ISSUE: VBControlExtender property Sock.Index was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		StackSocket.Add(Sock, "#" & CStr(Sock.Index))
		'UPGRADE_NOTE: Object CtmpAgn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
		CtmpAgn = Nothing
	End Function
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	'
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Function AgentCert(ByRef SockIndex As Object, ByRef CertCommand As Object) As Object
		Dim CtmpAgn As clsAgent
		Dim BoolRet As Boolean
		Dim StrTmpAgnName As Object
		Dim BoolOfflineSignal As Boolean
		
		'[Fetch agent name to check if agent is offline type ]'
		'UPGRADE_WARNING: Couldn't resolve default property of object CertCommand. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object SubVal(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object StrTmpAgnName. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		StrTmpAgnName = SubVal(CStr(CertCommand), "name")
		For	Each CtmpAgn In StackAgent
			'UPGRADE_WARNING: Couldn't resolve default property of object StrTmpAgnName. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			If StrTmpAgnName = CtmpAgn.AgentName Then
				BoolOfflineSignal = True
			End If
		Next CtmpAgn
		
		
		If BoolOfflineSignal = True Then
			CtmpAgn = StackAgent.Item(StrTmpAgnName)
			'UPGRADE_WARNING: Couldn't resolve default property of object SockIndex. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Couldn't resolve default property of object StackSocket(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			CtmpAgn.AgnInit(True, StackSocket.Item("#" & SockIndex))
			'UPGRADE_WARNING: Couldn't resolve default property of object CertCommand. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			BoolRet = CtmpAgn.AgnInitCert(CStr(CertCommand), True)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object SockIndex. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			CtmpAgn = StackTmpAgent.Item("#" & SockIndex)
			'UPGRADE_WARNING: Couldn't resolve default property of object CertCommand. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			BoolRet = CtmpAgn.AgnInitCert(CStr(CertCommand))
			If BoolRet = True Then
				'UPGRADE_WARNING: Couldn't resolve default property of object SockIndex. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				StackTmpAgent.Remove("#" & SockIndex)
				StackAgent.Add(CtmpAgn, CtmpAgn.AgentName)
				TriggerEvent(CtmpAgn, EnEvent.Agent_Added)
				FrmSysHost.DefInstance.Pinger.Enabled = True
			End If
		End If
		
		If BoolRet = False Then CtmpAgn.AgnInitReject()
		'UPGRADE_NOTE: Object CtmpAgn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
		CtmpAgn = Nothing
	End Function
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	'
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Sub AgentDisconnect(ByRef SockIndex As Object)
		Dim CtmpAgn As clsAgent
		
		For	Each CtmpAgn In StackAgent
			'UPGRADE_WARNING: Couldn't resolve default property of object SockIndex. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			If CtmpAgn.AgentSockIndex = SockIndex Then
				CtmpAgn.AgnOffline()
				'UPGRADE_WARNING: Couldn't resolve default property of object SockIndex. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				StackSocket.Remove("#" & SockIndex)
				TriggerEvent(CtmpAgn, EnEvent.Agent_Disconnect)
				Exit Sub
			End If
		Next CtmpAgn
		
		'UPGRADE_NOTE: Object CtmpAgn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
		CtmpAgn = Nothing
	End Sub
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	'
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	'UPGRADE_NOTE: AgentRemove was upgraded to AgentRemove_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
	Public Sub AgentRemove_Renamed(ByRef SockIndex As Object)
		Dim CtmpAgn As clsAgent
		
		For	Each CtmpAgn In StackAgent
			'UPGRADE_WARNING: Couldn't resolve default property of object SockIndex. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			If CtmpAgn.AgentSockIndex = SockIndex Then
				CtmpAgn.AgnRemove()
				StackAgent.Remove(CtmpAgn.AgentName)
				'UPGRADE_WARNING: Couldn't resolve default property of object SockIndex. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				StackSocket.Remove("#" & SockIndex)
				TriggerEvent(CtmpAgn, EnEvent.Agent_Removed)
				Exit Sub
			End If
		Next CtmpAgn
		
		'UPGRADE_NOTE: Object CtmpAgn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
		CtmpAgn = Nothing
	End Sub
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	'
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Sub AgentRecoverUsed()
		Dim CtmpAgn As clsAgent
		
		For	Each CtmpAgn In StackAgent
			If CtmpAgn.AgentStatus <> VS(4) Then
				CtmpAgn.AgnRecover()
			End If
		Next CtmpAgn
		
		'UPGRADE_NOTE: Object CtmpAgn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
		CtmpAgn = Nothing
	End Sub
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	'
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Sub AgentRecoverUp()
		Dim DbR As DAO.Database
		Dim RsR As DAO.Recordset
		
		DbR = DAODBEngine_definst.OpenDatabase(CurIDBPath, False, False, ";pwd=nsb2003")
		RsR = DbR.OpenRecordset("TerminalRecover", DAO.RecordsetTypeEnum.dbOpenDynaset)
		
		Dim CtmpAgn As New clsAgent
		With RsR
			If .BOF = True Then Exit Sub
			.MoveLast()
			.MoveFirst()
			Do Until .EOF = True
				CtmpAgn.AgnInit(False, Nothing, .Fields("AgentName").Value)
				StackAgent.Add(CtmpAgn, CtmpAgn.AgentName)
				TriggerEvent(CtmpAgn, EnEvent.Agent_Added_Offline_)
				'UPGRADE_NOTE: Object CtmpAgn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
				CtmpAgn = Nothing
				.MoveNext()
			Loop 
		End With
		RsR.Close()
		DbR.Close()
	End Sub
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	'
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Sub AgentCheckUsed()
		Dim CtmpAgn As clsAgent
		
		For	Each CtmpAgn In StackAgent
			If CtmpAgn.AgentStatus <> VS(4) Then
				CtmpAgn.CusCheckExpired()
				CtmpAgn.CusStatusUpdate()
			End If
		Next CtmpAgn
		
		'UPGRADE_NOTE: Object CtmpAgn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
		CtmpAgn = Nothing
	End Sub
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	'
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
	Public Sub SendCommand(ByRef Command_Renamed As Object)
		Dim CtmpAgn As clsAgent
		If StackAgent.Count() = 0 Then Exit Sub
		For	Each CtmpAgn In StackAgent
			'UPGRADE_WARNING: Couldn't resolve default property of object Command_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			CtmpAgn.NetSend("//" & Command_Renamed)
		Next CtmpAgn
		
		'UPGRADE_NOTE: Object CtmpAgn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
		CtmpAgn = Nothing
	End Sub
	
	
	
	'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
	' SOCKET MEMBERS
	'
	'
	'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Socks Member
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Function Socks(ByRef Index As Object) As AxSocketWrenchCtrl.AxSocket
		Socks = StackSocket.Item(Index)
	End Function
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Socket Count
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Function SockCount() As Integer
		SockCount = StackSocket.Count()
	End Function
	
	
	
	'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
	' EVENTS
	'
	'
	'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
	Public Sub TriggerEvent(ByRef Agent As clsAgent, ByRef EventNum As EnEvent, Optional ByRef Param1 As Object = "", Optional ByRef Param2 As Object = "")
		If EventNum = 0 Then Exit Sub
		Select Case EventNum
			Case 1
				RaiseEvent AgentAdded(Agent)
			Case 2
				RaiseEvent AgentRemove(Agent)
		End Select
	End Sub
End Class