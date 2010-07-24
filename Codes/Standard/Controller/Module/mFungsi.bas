Attribute VB_Name = "MdlCommand"
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : MdlCommand
'    Project    : CafeBonzer
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Private StrCmdDataPool As String


Public Function CmdCount(Value) As Long
    Dim lngCmd As Long
    For d = 1 To Len(Value)
        If Mid(Value, d, 1) = StrCmdSep Then lngCmd = lngCmd + 1
    Next d
    CmdCount = lngCmd
End Function

Public Function CmdSubCount(Value) As Long
    Dim lngCmd As Long
    For d = 1 To Len(Value)
        If Mid(Value, d, 1) = StrCmdSubSep1 Then lngCmd = lngCmd + 1
    Next d
    CmdSubCount = lngCmd
End Function


Public Function CmdSubPut(Data, Value) As String
    CmdSubPut = StrCmdSubSep1 & Data & StrCmdSubSep2 & Value
End Function


Public Function CmdSubGet(Value, SubName) As String
    Dim strTmp As String, strTmp2 As String
    Dim lngCnt As Long
    
    lngCnt = CmdSubCount(Value)
    For g = 1 To lngCnt
        strTmp2 = Split(Value, StrCmdSubSep1)(g)
        strTmp = Split(strTmp2, StrCmdSubSep2)(0)
        If LCase$(strTmp) = LCase$(SubName) Then
            CmdSubGet = Split(strTmp2, StrCmdSubSep2)(1)
            Exit Function
        End If
    Next
End Function


''=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'' Parse Network Command
''=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub CmdParse(CmdData As String, SockIndex As Long)
    Dim LngCmdCount As Long, LngIdxA As Long, StrCmdData As String
    Dim StrCmdMain As String, StrCmdSub As String, ObjAg As ClsAgent

    'StrCmdDataPool = StrCmdDataPool + CmdData
    LngCmdCount = CmdCount(CmdData)
    Set ObjAg = UniAgents.AgentsByIndex(SockIndex)
    
    For LngIdxA = 1 To LngCmdCount
        StrCmdData = Split(CmdData, StrCmdSep)(LngIdxA)
        StrCmdMain = Mid(StrCmdData, 1, 2)
        StrCmdSub = Mid(StrCmdData, 3, 4)
        
        If StrCmdMain = "01" Then
            Select Case StrCmdSub
                Case Is = "0010"
                    ObjAg.Commands.NetPing (Pong)
                Case Is = "0020"
                    ObjAg.NetPingReset
                Case Is = "0030"
                    UniAgents.AgentCert SockIndex, StrCmdData
            End Select
        ElseIf StrCmdMain = "02" Then
            Select Case StrCmdSub
                Case Is = "0040"
                    ObjAg.NetSetVariable "NETMAC", CmdSubGet(StrCmdData, "NETMAC")
            End Select
        ElseIf StrCmdMain = "03" Then
            Select Case StrCmdSub
                Case Is = "0010"
                    FrmAgnMsg.AddMessage ObjAg, Mid(StrCmdData, 7)
            End Select
        ElseIf StrCmdMain = "04" Then
            Select Case StrCmdSub
            Case Is = "0010"
                ObjAg.AgStatusRequested StrCmdData
            Case Is = "0020"
                'Sent Result
            End Select
        End If
    Next
End Sub


' kita selalu tidak mahu terjadi sesuatu yang tidak kita suka
' tetapi kita mesti selalu bersedia menghadapi sesuatu yang tidak kita suka
' azri jamil - oct,2002
