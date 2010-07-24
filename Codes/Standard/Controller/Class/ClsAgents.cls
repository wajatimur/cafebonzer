VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsAgents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : clsAgents
'    Project    : CafeBonzer
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Private StackTmpAgent As New Collection
Private StackAgent As New Collection
Private StackSocket As New Collection

Public Enum EnControlAll
    LockAll = 1010
    LockAllUnused = 1020
    UnlockAll = 1030
    UnlockAllUsed = 1040
    ShutdownAll = 2010
    ShutdownAllUnused = 2020
    ShutdownAllForced = 2030
    ShutdownAllForcedUnused = 2040
    RebootAll = 2050
    RebootUnused = 2060
    CleanAll = 3010
    CleanAllTemp = 3020
    CleanAllRecycleBin = 3030
    CleanAllHistory = 3040
    CleanAllRecentDocs = 3050
End Enum

Public Enum EnCountStatus
    [Used] = 0
    [UnUsed] = 1
    [Locked] = 2
    [Unlocked] = 3
    [Connected] = 4
End Enum

Public Enum EnEvent
    [Agent Added] = 1
    [Agent Removed] = 2
    [Agent Disconnect] = 3
    [Agent Added Offline ] = 4
    [Info Updated] = 5
End Enum

Public Event AgentAdded(ByRef Agent As ClsAgent)
Public Event AgentRemove(ByRef Agent As ClsAgent)
Public Event InfoUpdated(ByRef Agent As ClsAgent, InfoType As Long)

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Members | Agents
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    Public Function Agents(Name) As ClsAgent
Attribute Agents.VB_UserMemId = 0
        Set Agents = StackAgent.Item(Name)
    End Function
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Members | Agents by Index
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    Public Function AgentsByIndex(SockIndex) As ClsAgent
        Dim CtmpAgn As ClsAgent
        
        For Each CtmpAgn In StackAgent
            If CtmpAgn.AgentSockIndex = SockIndex Then
                Set AgentsByIndex = CtmpAgn
                Exit Function
            End If
        Next
        
        Set CtmpAgn = Nothing
    End Function



'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' Agent Members
'
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Count Agent
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    Public Function Count() As Long
        Count = StackAgent.Count
    End Function
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    Public Function CountStatus(Status As EnCountStatus) As Long
        Dim TmpAgn As ClsAgent, lAgnCnt As Long
        Select Case Status
            Case 0
                For Each TmpAgn In StackAgent
                    If TmpAgn.AgentStatus = VS(1, 1) Then lAgnCnt = lAgnCnt + 1
                Next
                CountStatus = lAgnCnt
            Case 1
                For Each TmpAgn In StackAgent
                    If TmpAgn.AgentStatus = VS(1, 2) Then lAgnCnt = lAgnCnt + 1
                Next
                CountStatus = lAgnCnt
            Case 4
                For Each TmpAgn In StackAgent
                    If TmpAgn.AgentCertified = True Then lAgnCnt = lAgnCnt + 1
                Next
                CountStatus = lAgnCnt
        End Select
    End Function
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    Public Function AgentAdd(Sock As Socket, SocketID As Long) As ClsAgent
        Dim CtmpAgn As New ClsAgent
        
        Sock.Accept = SocketID
        CtmpAgn.AgnInit True, Sock
        
        StackTmpAgent.Add CtmpAgn, "#" & CStr(Sock.Index)
        StackSocket.Add Sock, "#" & CStr(Sock.Index)
        Set CtmpAgn = Nothing
    End Function
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    Public Function AgentCert(SockIndex As Long, CertCommand As String)
        Dim CtmpAgn As ClsAgent, BoolRet As Boolean
        Dim StrTmpAgnName, BoolOfflineSignal As Boolean
        
      '[Fetch agent name to check if agent is offline type ]'
        StrTmpAgnName = CmdSubGet(CertCommand, "NAME")
        For Each CtmpAgn In StackAgent
            If StrTmpAgnName = CtmpAgn.AgentName Then
                BoolOfflineSignal = True
            End If
        Next
        
        If BoolOfflineSignal = True Then
            Set CtmpAgn = StackAgent(StrTmpAgnName)
            CtmpAgn.AgnInit True, StackSocket("#" & SockIndex)
            '{ Certifying Agent }'
            BoolRet = CtmpAgn.AgnInitCert(CStr(CertCommand), True)
        Else
            Set CtmpAgn = StackTmpAgent("#" & SockIndex)
            '{ Certifying Agent }'
            BoolRet = CtmpAgn.AgnInitCert(CStr(CertCommand))
            If BoolRet = True Then
                StackTmpAgent.Remove "#" & SockIndex
                StackAgent.Add CtmpAgn, CtmpAgn.AgentName
                TriggerEvent CtmpAgn, [Agent Added]
                FrmSysHost.TmrPing.Enabled = True
            End If
        End If
        
        If BoolRet = False Then CtmpAgn.AgnInitReject
        Set CtmpAgn = Nothing
    End Function
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    Public Sub AgentDisconnect(SockIndex)
        Dim CtmpAgn As ClsAgent
        
        For Each CtmpAgn In StackAgent
            If CtmpAgn.AgentSockIndex = SockIndex Then
                CtmpAgn.AgnOffline
                StackSocket.Remove "#" & SockIndex
                TriggerEvent CtmpAgn, [Agent Disconnect]
                Exit Sub
            End If
        Next
        
        Set CtmpAgn = Nothing
    End Sub
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    Public Sub AgentRemove(SockIndex)
        Dim CtmpAgn As ClsAgent
        
        For Each CtmpAgn In StackAgent
            If CtmpAgn.AgentSockIndex = SockIndex Then
                CtmpAgn.AgnRemove
                StackAgent.Remove CtmpAgn.AgentName
                StackSocket.Remove "#" & SockIndex
                TriggerEvent CtmpAgn, [Agent Removed]
                Exit Sub
            End If
        Next
        
        Set CtmpAgn = Nothing
    End Sub
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    Public Sub AgentRecoverUsed()
        Dim CtmpAgn As ClsAgent
        
        For Each CtmpAgn In StackAgent
            If CtmpAgn.AgentStatus <> VS(1, 2) Then
                CtmpAgn.AgnRecover
            End If
        Next
        
        Set CtmpAgn = Nothing
    End Sub
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    Public Sub AgentRecoverUp()
        Dim DbR As Database, RsR As Recordset
    
        Set DbR = OpenDatabase(CurIDBPath, False, False, ";pwd=nsb2003")
        Set RsR = DbR.OpenRecordset("TerminalRecover", dbOpenDynaset)
        
        With RsR
            If .BOF = True Then Exit Sub
            .MoveLast
            .MoveFirst
            Do Until .EOF = True
                Dim CtmpAgn As New ClsAgent
                CtmpAgn.AgnInit False, Nothing, !AgentName
                StackAgent.Add CtmpAgn, CtmpAgn.AgentName
                TriggerEvent CtmpAgn, [Agent Added Offline ]
                Set CtmpAgn = Nothing
                .MoveNext
            Loop
        End With
        RsR.Close
        DbR.Close
    End Sub
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    Public Sub AgentCheckUsed()
        Dim CtmpAgn As ClsAgent
        
        For Each CtmpAgn In StackAgent
            If CtmpAgn.AgentStatus <> VS(1, 2) Then
                CtmpAgn.CusCheckExpired
                CtmpAgn.CusStatusUpdate
            End If
        Next
        
        Set CtmpAgn = Nothing
    End Sub
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    Public Sub SendCommand(Command)
        Dim CtmpAgn As ClsAgent
        If StackAgent.Count = 0 Then Exit Sub
        For Each CtmpAgn In StackAgent
            CtmpAgn.NetSend StrCmdSep + Command
        Next
        
        Set CtmpAgn = Nothing
    End Sub
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    Public Sub AgentControl(Action As EnControlAll)
        Dim CtmpAgn As ClsAgent, StrAgnStatus As String
        If StackAgent.Count = 0 Then Exit Sub
        
        For Each CtmpAgn In StackAgent
            StrAgnStatus = CtmpAgn.AgentStatus
            Select Case Action
            Case 1010
                CtmpAgn.Commands.ConLock (TerminalLock)
            Case 1020
                If StrAgnStatus = VS(1, 2) Then CtmpAgn.Commands.ConLock (TerminalLock)
            Case 1030
                CtmpAgn.Commands.ConLock (TerminalUnlock)
            Case 1040
                If StrAgnStatus = VS(1, 1) Then CtmpAgn.Commands.ConLock (TerminalUnlock)
            Case 2010
                CtmpAgn.Commands.ConExitWin (Shutdown)
            Case 2020
                If StrAgnStatus = VS(1, 2) Then CtmpAgn.Commands.ConExitWin (Shutdown)
            Case 2030
                CtmpAgn.Commands.ConExitWin (ShutdownForce)
            Case 2040
                If StrAgnStatus = VS(1, 2) Then CtmpAgn.Commands.ConExitWin (ShutdownForce)
            Case 2050
                CtmpAgn.Commands.ConExitWin (Reboot)
            Case 2060
                If StrAgnStatus = VS(1, 2) Then CtmpAgn.Commands.ConExitWin (Reboot)
            Case 3010
                CtmpAgn.Commands.ConCleaning (Clean)
            Case 3020
                CtmpAgn.Commands.ConCleaning (CleanTemp)
            Case 3030
                CtmpAgn.Commands.ConCleaning (CleanRecycleBin)
            Case 3040
                CtmpAgn.Commands.ConCleaning (CleanHistory)
            Case 3050
                CtmpAgn.Commands.ConCleaning (CleanRecentDocs)
            End Select
        Next
        
        Set CtmpAgn = Nothing
    End Sub


'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' SOCKET MEMBERS
'
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Socks Member
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    Public Function Socks(Index) As Socket
        Set Socks = StackSocket(Index)
    End Function
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Socket Count
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    Public Function SockCount() As Long
        SockCount = StackSocket.Count
    End Function



'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' EVENTS
'
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
    Public Sub TriggerEvent(Agent As ClsAgent, EventNum As EnEvent, Optional Param1 = "", Optional Param2 = "")
        If EventNum = 0 Then Exit Sub
        Select Case EventNum
            Case 1
                RaiseEvent AgentAdded(Agent)
            Case 2
                RaiseEvent AgentRemove(Agent)
        End Select
    End Sub
