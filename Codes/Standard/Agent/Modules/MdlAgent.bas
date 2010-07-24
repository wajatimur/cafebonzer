Attribute VB_Name = "MdlAgent"
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : MdlAgent
'    Project    : CafeBonzerAG
'
'    Description: Module Agent
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Private LngTickMsgCd As Long


Public Sub AgentUsage(SubCommand As String)
    If LngTickMsgCd > 0 Then
        'Sequence Lag for Ticker Message.
        LngTickMsgCd = LngTickMsgCd - 1
    Else
        If CmdSubGet(SubCommand, "ACTION") = 0 Then
            MdlTicker.StrTickerText = CmdSubGet(SubCommand, "PRICEUSE")
        Else
            MdlTicker.StrTickerText = CmdSubGet(SubCommand, "TIMELEFT")
        End If
    End If
End Sub


Public Sub AgentLogin(SubCommand As String)
    If CmdSubGet(SubCommand, "ACTION") = "0" Then
        SettingSave "LOGIN", False
        Call SysShellLock(1)
    Else
        SettingSave "LOGIN", True
        Call SysShellLock(0)
    End If
End Sub


Public Sub AgentCertified(SubCommand As String)
    If CmdSubGet(SubCommand, "ACTION") = 1 Then
        Call NetSessionStart
        Call AgentInfoStatus
    Else
        NetClose
        NetConnect
    End If
End Sub

Public Sub AgentConfiguration(SubCommand As String)
    '
End Sub


Public Sub AgentInfoStatus()
    ' ShellLock
    '   Lock    = 1
    '   Unlock  = 0
    
    NetSend "040010" & CmdSubPut("LOCK", CStr(LngStatusLock))
End Sub


Public Sub AgentMsgReceive(SubCommand As String, Optional MsgType As Long = 0)
    Dim StrMessage As String
    
    Select Case MsgType
    Case 0
        If LngStatusLock = 0 Then
            StrMessage = "[Server] " & SubCommand
            FrmMessaging.LsvMessage.ListItems.Add , , StrMessage, , "MSGIN"
            FrmMessaging.Show
        End If
    Case 1
        If MsgType = 1 Then
            MdlTicker.StrTickerText = SubCommand
            LngTickMsgCd = 1
        End If
    End Select
End Sub

