Attribute VB_Name = "MdlCommand"
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : MdlCommand
'    Project    : CafeBonzerAG
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>

Public Function CmdCount(Value As String) As Long
    Dim LngIdxA As Long, LngCmd As Long
    For LngIdxA = 1 To Len(Value)
        If Mid$(Value, LngIdxA, 1) = StrCmdSep Then LngCmd = LngCmd + 1
    Next
    CmdCount = LngCmd
End Function


Public Function CmdSubCount(Value As String) As Long
    Dim LngIdxA As Long, LngCmd As Long
    For LngIdxA = 1 To Len(Value)
        If Mid$(Value, LngIdxA, 1) = StrCmdSubSep1 Then LngCmd = LngCmd + 1
    Next
    CmdSubCount = LngCmd
End Function


Public Function CmdSubPut(Data As String, Value As String) As String
    CmdSubPut = StrCmdSubSep1 & Data & StrCmdSubSep2 & Value
End Function


Public Function CmdSubGet(Value As String, SubName As String) As String
    Dim StrTmp As String, StrTmp2 As String
    Dim LngIdxA As Long, LngCnt As Long
    
    LngCnt = CmdSubCount(Value)
    For LngIdxA = 1 To LngCnt
        StrTmp2 = Split(Value, StrCmdSubSep1)(LngIdxA)
        StrTmp = Split(StrTmp2, StrCmdSubSep2)(1)
        If LCase$(StrTmp) = LCase$(SubName) Then
            CmdSubGet = Split(StrTmp2, StrCmdSubSep2)(2)
            Exit Function
        End If
    Next
End Function


Public Sub CmdParse(CmdData As String)
    Dim LngCmdCount As Long, LngIdxA As Long, StrCmdData As String
    Dim StrCmdMain As String, StrCmdSub As String, StrCmdSubData As String

    LngCmdCount = CmdCount(CmdData)
    If LngCmdCount = 0 Then Exit Sub
    
    For LngIdxA = 1 To LngCmdCount
        StrCmdData = Split(CmdData, StrCmdSep)(LngIdxA)
        StrCmdMain = Mid$(StrCmdData, 1, 2)
        StrCmdSub = Mid$(StrCmdData, 3, 4)
        StrCmdSubData = Mid$(StrCmdData, 7)
        
        If StrCmdMain = "01" Then
            Select Case StrCmdSub
            Case Is = "0010"
                Call NetSend("010020")
            Case Is = "0020"
                'NetPingReset
            Case Is = "0030"
                Call AgentCertified(StrCmdSubData)
            End Select
            
        ElseIf StrCmdMain = "02" Then
            Select Case StrCmdSub
            Case Is = "0010"
                Call SysWindowsExit(StrCmdSubData)
            Case Is = "0020"
                Call SysShellLock(StrCmdSubData)
            Case Is = "0030"
                Call SysDevDiskClean(StrCmdSubData)
            Case Is = "0040"
                Call AgentLogin(StrCmdSubData)
            Case Is = "0050"
                Call AgentUsage(StrCmdData)
            End Select
            
        ElseIf StrCmdMain = "03" Then
            Select Case StrCmdSub
            Case Is = "0010"
                Call AgentMsgReceive(StrCmdSubData)
            Case Is = "0020"
                Call AgentMsgReceive(StrCmdSubData, 1)
            End Select
            
        ElseIf StrCmdMain = "04" Then
            Select Case StrCmdSub
            Case Is = "0010"
                Call AgentInfoStatus
            Case Is = "0020"
                Call AgentConfiguration(StrCmdSubData)
            Case Is = "0100"
                Call AppExit
            End Select
            
        End If
    Next
End Sub

