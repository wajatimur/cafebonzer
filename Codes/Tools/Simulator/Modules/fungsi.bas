Attribute VB_Name = "mdlCommand"
Public NoReply As Boolean
Public colVC As New Collection

Public bConLock As Long
Public bConBlock As Long
Public bConMsg As Boolean
Public bConTick As Boolean

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [CmdName] - Return Command Name
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Function CmdName(DataString As String) As String
    Dim l_dataLen As Long, l_dataFnd As Long
    
    l_dataLen = Len(DataString)
    If l_dataLen = 0 Then Exit Function
    
    For a = 1 To l_dataLen
        If Mid(DataString, a, 1) = ":" Then
            CmdName = Mid(DataString, 3, a - 3)
            Exit Function
        End If
    Next a
    
    CmdName = Mid(DataString, 3)
End Function

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [CmdValue] - Return Command Value
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Function CmdValue(DataString As String) As String
    Dim l_dataLen As Long
    
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
' [Command - Sub Value] - Get Value for Sub Command
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Function SubVal(DataString As String, SubCmdName As String, Optional Default As Variant = "") As Variant
    Dim s_CmdVal As String, l_PosCmdName As Long
    Dim l_PosCmdDataA As Long, l_PosCmdDataB As Long
    ' Note
    '   format bagi sub command
    '   {subcmdname|'data'}{subcmdname2|'data'}
    '
    
    s_CmdVal = CmdValue(DataString)
    l_PosCmdName = InStr(1, DataString, "{" & LCase(SubCmdName))
    l_PosCmdDataA = InStr(l_PosCmdName, DataString, "|'") + 2
    l_PosCmdDataB = InStr(l_PosCmdDataA, DataString, "'}")
    SubVal = Mid(DataString, l_PosCmdDataA, l_PosCmdDataB - l_PosCmdDataA)
    If SubVal = "" Then SubVal = Default
End Function

Function SubBuild(DataName As String, DataValue As String) As String
    SubBuild = "{" & DataName & "|'" & DataValue & "'}"
End Function

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [CmdParse] - Parse Net Command
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub CmdParse(Index, DataString As String)
    Dim s_CmdName As String, s_CmdVal As String
    
    s_CmdName = CmdName(DataString)
    s_CmdVal = CmdValue(DataString)
    
    Select Case s_CmdName
        Case Is = "cert"
            If SubVal(s_CmdVal, "granted", 0) = 1 Then
                NetStart Index
            Else
                MsgBox " [ Cannot connect | " & SubVal(s_CmdVal, "info") & "]"
                NetClose
                NetConnect
            End If
            
        Case Is = "hey"
            If NoReply = False Then NetSend Index, "/hoi"
        
        Case Is = "login"
            'SetSave "login", True
            FrmHost.List.AddItem MyName & Index & "> Login"
            
        Case Is = "logout"
            'SetSave "login", False
            FrmHost.List.AddItem MyName & Index & "> Logout"
            
        Case Is = "tutup"
            FrmHost.List.AddItem MyName & Index & "> Tutup = " & s_CmdVal
            
        Case Is = "mesej"
            FrmHost.List.AddItem MyName & Index & "> Mesej = " & s_CmdVal
            
        Case Is = "harga"
            FrmHost.List.AddItem MyName & Index & "> Harga = " & s_CmdVal
            
        Case Is = "tiker"
            FrmHost.List.AddItem MyName & Index & "> Tiker = " & s_CmdVal
            
        Case Is = "shell"
            FrmHost.List.AddItem MyName & Index & "> Shell = " & s_CmdVal
            
        Case Is = "kunci"
            FrmHost.List.AddItem MyName & Index & "> Kunci = " & s_CmdVal
            bConLock = c_cmdval
            
        Case Is = "sdown"
            FrmHost.List.AddItem MyName & Index & "> Sdown = " & s_CmdVal
            
        Case Is = "block"
            FrmHost.List.AddItem MyName & Index & "> Block = " & s_CmdVal
            NetSend Index, "/info.me:block"
            
        Case Is = "sleep"
            FrmHost.List.AddItem MyName & Index & "> Sleep = " & s_CmdVal
            
        Case Is = "cleand"
            FrmHost.List.AddItem MyName & Index & "> CleanD = " & s_CmdVal
            NetSend Index, "/info.me:cleanok"
            
        Case Is = "status"
            FrmHost.List.AddItem MyName & Index & "> Status = " & s_CmdVal
            If bConLock = 1 Then NetSend Index, "/info.me:lock" Else NetSend Index, "/info.me:unlock"
            If bConBlock = 1 Then NetSend Index, "/info.me:block" Else NetSend Index, "/info.me:unblock"
    
        Case Is = "screen"
            FrmHost.List.AddItem MyName & Index & "> Screen = " & s_CmdVal
            
        Case Is = "mon.switch"
            FrmHost.List.AddItem MyName & Index & "> Mon.Switch = " & s_CmdVal
    End Select
End Sub
