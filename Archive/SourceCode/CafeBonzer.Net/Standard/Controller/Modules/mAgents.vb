Option Strict Off
Option Explicit On
Module mAgents
	Public lSock As Integer
	
	Public Enum EnAgentSymbol
		User_Online = 0
		User_Offline = 1
		User_Prepaid = 2
		User_Prepaid_Ended = 3
		Terminal_Online = 10
		Terminal_Offline = 11
		Terminal_Lock = 12
		Terminal_Cleaning = 13
	End Enum
	
	
	
	''=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	'' Parse Network Command
	''=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Sub ParseCmd(ByRef DataRcv As String, ByRef SockIndex As Integer)
		Dim s_CmdName, s_CmdVal As String
		Dim DataString As String
		Dim uA As clsAgent
		Dim xPos As Short
		Dim yPos As Short
		
		xPos = 0
		Do Until xPos = Len(DataRcv)
			xPos = InStr(xPos + 1, DataRcv, "/")
			yPos = InStr(xPos + 1, DataRcv, "/")
			
			If xPos = 0 Then Exit Do
			If yPos = 0 Then
				DataString = Mid(DataRcv, xPos)
				xPos = Len(DataRcv)
			Else
				DataString = Mid(DataRcv, xPos, yPos - xPos)
			End If
		Loop 
		
		s_CmdName = CmdName(DataString)
		s_CmdVal = CmdValue(DataString)
		
		If s_CmdName = "cert" Then
			UniAgents.AgentCert(SockIndex, s_CmdVal)
		Else
			uA = UniAgents.AgentsByIndex(SockIndex)
			Select Case s_CmdName
				Case Is = "hoi"
					uA.NetPingReset()
				Case Is = "mesej"
					Call FnMesej(DataString)
			End Select
			If Left(s_CmdName, 4) = "info" Then InfoParse(DataString, uA)
		End If
	End Sub
	
	''=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	'' Update Main Panel
	''=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Sub UpdatePanel(ByRef KeyName As Object)
		Dim uA As clsAgent
		
		FrmMain.DefInstance.SpgInfoLblB(0).Text = CStr(UniAgents.Count)
		FrmMain.DefInstance.SpgInfoLblB(1).Text = CStr(UniAgents.CountStatus(clsAgents.EnCountStatus.UnUsed))
		
		'UPGRADE_WARNING: Couldn't resolve default property of object KeyName. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If UniAgents.Count > 0 And KeyName <> "" Then
			uA = UniAgents.Agents(KeyName)
			FrmMain.DefInstance.SpgInfoLblD(0).Text = uA.AgentConnected
			FrmMain.DefInstance.SpgInfoLblD(1).Text = uA.AgentIPAdd
			FrmMain.DefInstance.SpgInfoLblD(2).Text = uA.AgentMAC
			If uA.AgentStatus = VS(3) Then
				FrmMain.DefInstance.SpgInfoLblD(3).Text = Crnc & VB6.Format(uA.CusGetPrice, "#0.00")
			Else
				FrmMain.DefInstance.SpgInfoLblD(3).Text = ""
			End If
			'UPGRADE_NOTE: Object uA may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
			uA = Nothing
		Else
			FrmMain.DefInstance.SpgInfoLblD(0).Text = ""
			FrmMain.DefInstance.SpgInfoLblD(1).Text = ""
			FrmMain.DefInstance.SpgInfoLblD(2).Text = ""
			FrmMain.DefInstance.SpgInfoLblD(3).Text = ""
		End If
	End Sub
	
	''=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	'' Update Stat
	''=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Sub UpdateStat(ByRef sItem As MSComctlLib.ListItem)
		Dim sCustomerTimeIn, sCustomerName, sCustomerTimeOut As String
		Dim sAgentStatus As String
		Dim uA As clsAgent
		
		If sItem Is Nothing Then
			StatText(3)
			Exit Sub
		End If
		
		uA = UniAgents.Agents((sItem.Text))
		sAgentStatus = uA.AgentStatus
		sCustomerName = uA.CustomerName
		sCustomerTimeIn = uA.CustomerTimeIn
		sCustomerTimeOut = uA.CustomerTimeOut
		
		If sAgentStatus = VS(4) Then
			StatText(3, "")
			Exit Sub
		End If
		
		'untuk prepaid
		If sCustomerTimeOut <> VS(1) And sAgentStatus = VS(3) Then
			StatText(3, "Time Left : " & uA.CusGetTimeLeft & " - " & sCustomerName)
			Exit Sub
			'UPGRADE_NOTE: Object uA may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
			uA = Nothing
		End If
		
		'jika prepaid telah tamat..
		If sCustomerTimeOut <> VS(1) And sAgentStatus = VS(5) Then
			StatText(3, "Time End - " & sCustomerName)
			Exit Sub
			'UPGRADE_NOTE: Object uA may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
			uA = Nothing
		End If
		
		'untuk pay as you go
		If sCustomerTimeOut = VS(1) And sAgentStatus = VS(3) Then
			StatText(3, Crnc & " " & uA.CusGetPrice & "  (" & uA.CusGetTimeUse & ")" & " - " & sCustomerName)
			Exit Sub
			'UPGRADE_NOTE: Object uA may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
			uA = Nothing
		End If
	End Sub
	
	''=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	'' Selected Agent
	''=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Function AgentSel() As clsAgent
		AgentSel = UniAgents.Agents(SelText)
	End Function
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Dapatkan index dalam listview melalui SckID
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Function AgentGetIndex(ByRef SockIndex As Object) As Object
		Dim g As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object AgentGetIndex. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		AgentGetIndex = 0
		For g = 1 To FrmMain.DefInstance.Lv1.ListItems.Count
			'UPGRADE_WARNING: Couldn't resolve default property of object SockIndex. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			If FrmMain.DefInstance.Lv1.ListItems(g).Tag = CStr(SockIndex) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object g. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				'UPGRADE_WARNING: Couldn't resolve default property of object AgentGetIndex. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				AgentGetIndex = g : Exit Function
			End If
		Next g
	End Function
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Dapatkan index dalam listview melalui Nama Terminal
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Function AgentGetIndexB(ByRef Nama As Object) As Object
		Dim g As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object AgentGetIndexB. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		AgentGetIndexB = 0
		For g = 1 To FrmMain.DefInstance.Lv1.ListItems.Count
			'UPGRADE_WARNING: Couldn't resolve default property of object Nama. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			If FrmMain.DefInstance.Lv1.ListItems(g).Text = CStr(Nama) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object g. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				'UPGRADE_WARNING: Couldn't resolve default property of object AgentGetIndexB. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				AgentGetIndexB = g : Exit Function
			End If
		Next g
	End Function
End Module