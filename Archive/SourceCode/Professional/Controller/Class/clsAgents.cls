VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAgents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private StackTmpAgent As New Collection
Private StackAgent As New Collection
Private StackSocket As New Collection

Public Enum eCountStatus
    [Used] = 0
    [UnUsed] = 1
    [Locked] = 2
    [Unlocked] = 3
End Enum

Public Enum eEventNum
    [Agent Added] = 1
    [Agent Removed] = 2
    [Info Updated] = 3
End Enum

Public Event AgentAdded(ByRef Agent As clsAgent)
Public Event AgentRemove(ByRef Agent As clsAgent)
Public Event InfoUpdated(ByRef Agent As clsAgent, InfoType As Long)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Agent Member
''
''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Count() As Long
    Count = StackAgent.Count
End Function

Public Function CountStatus(Status As eCountStatus) As Long
    Dim TmpAgn As clsAgent, lAgnCnt As Long
    Select Case Status
        Case 0
            For Each TmpAgn In StackAgent
                If TmpAgn.AgentStatus = VS(3) Then lAgnCnt = lAgnCnt + 1
            Next
            CountStatus = lAgnCnt
        Case 1
            For Each TmpAgn In StackAgent
                If TmpAgn.AgentStatus = VS(4) Then lAgnCnt = lAgnCnt + 1
            Next
            CountStatus = lAgnCnt
    End Select
End Function

Public Function Agents(Name) As clsAgent
Attribute Agents.VB_UserMemId = 0
    Set Agents = StackAgent.Item(Name)
End Function

Public Function AgentsByIndex(SockIndex) As clsAgent
    Dim cTmpAgn As clsAgent
    
    For Each cTmpAgn In StackAgent
        If cTmpAgn.AgentSockIndex = SockIndex Then
            Set AgentsByIndex = cTmpAgn
            Exit Function
        End If
    Next
    
    Set cTmpAgn = Nothing
End Function

Public Function AddAgent(Sock As Socket, SocketID As Long) As clsAgent
    Dim cTmpAgn As New clsAgent
    
    Sock.Accept = SocketID
    cTmpAgn.AgnInit Sock
    
    StackTmpAgent.Add cTmpAgn, "#" & CStr(Sock.Index)
    StackSocket.Add Sock, "#" & CStr(Sock.Index)
    Set cTmpAgn = Nothing
End Function

Public Function AgentCert(SockIndex, CertCommand)
    Dim cTmpAgn As clsAgent, b_Ret As Boolean
    
    Set cTmpAgn = StackTmpAgent("#" & SockIndex)
    b_Ret = cTmpAgn.AgnInitCert(CStr(CertCommand), Me)
    
    If b_Ret = True Then
        StackTmpAgent.Remove "#" & SockIndex

        StackAgent.Add cTmpAgn, cTmpAgn.AgentName
        TriggerEvent cTmpAgn, [Agent Added]
        Call Recovery(cTmpAgn.AgentName)
        FrmHost.pinger.Enabled = True
    Else
        cTmpAgn.AgnInitReject
    End If
    
    Set cTmpAgn = Nothing
End Function

Public Sub RemoveAgent(SockIndex)
    Dim cTmpAgn As clsAgent
    
    For Each cTmpAgn In StackAgent
        If cTmpAgn.AgentSockIndex = SockIndex Then
            cTmpAgn.AgnRemove
            StackAgent.Remove cTmpAgn.AgentName
            StackSocket.Remove "#" & SockIndex
            TriggerEvent cTmpAgn, [Agent Removed]
            Exit Sub
        End If
    Next
    
    Set cTmpAgn = Nothing
End Sub

Public Sub SendCommand(Command)
    Dim cTmpAgn As clsAgent
    If StackAgent.Count = 0 Then Exit Sub
    For Each cTmpAgn In StackAgent
        cTmpAgn.NetSend "//" & Command
    Next
    
    Set cTmpAgn = Nothing
End Sub

Public Sub TriggerEvent(Agent As clsAgent, EventNum As eEventNum, Optional Param1 = "", Optional Param2 = "")
    If EventNum = 0 Then Exit Sub
    Select Case EventNum
        Case 1
            RaiseEvent AgentAdded(Agent)
        Case 2
            RaiseEvent AgentRemove(Agent)
        Case 3
            RaiseEvent InfoUpdated(Agent, CLng(Param1))
    End Select
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Socket Member
''
''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SockCount() As Long
    SockCount = StackSocket.Count
End Function

Public Function Socks(Index) As Socket
    Set Socks = StackSocket(Index)
End Function
