Option Strict Off
Option Explicit On
Module mDataInfo
	Public Sub InfoParse(ByRef DataCommand As String, ByRef Agent As clsAgent)
		Dim sCmdName, sCmdVal As String
		sCmdName = CmdName(DataCommand)
		sCmdVal = CmdValue(DataCommand)
		
		Dim BoolUlock As Boolean
		If sCmdName = "info.me" Then
			BoolUlock = (sCmdVal = "unlock") Or (sCmdVal = "cleanok")
			
			If BoolUlock Then
				If Agent.AgentStatus = VS(4) Then Agent.AgentIcon = "TerminalOnline"
				If Agent.AgentStatus = VS(3) Then Agent.AgentIcon = "UserOnline"
				If Agent.AgentStatus = VS(5) Then Agent.AgentIcon = "UserEnded"
			End If
			
			If sCmdVal = "lock" Then Agent.AgentIcon = "TerminalLock"
			If sCmdVal = "block" Then Agent.AgentIcon = "mouse"
		ElseIf sCmdName = "info.net" Then 
			'Agent.AgentMAC = SubVal(sCmdVal, "mac")
		End If
	End Sub
End Module