Attribute VB_Name = "mCommand"
Option Explicit

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Fungsi arahan yang diterima dari agent
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Sub FnMesej(arahan As String)
    Dim strNama As String, strMsg As String
    strNama = Mid$(arahan, 8, InStr(8, arahan, ":") - 8)
    strMsg = Mid$(arahan, InStr(8, arahan, ":") + 1)
    If CbConsole = True Then
        If FrmTerminal.Text1 <> "" Then FrmTerminal.List1.AddItem FrmTerminal.Text1
        FrmTerminal.wr strNama & ">" & strMsg
        If Echo = True Then FrmTerminal.wr FrmTerminal.Text2.Text Else FrmTerminal.Text1 = ""
        Exit Sub
    End If
    FrmMesej.server.Caption = strNama
    FrmMesej.rcv.Caption = strMsg
    If CbMsgRcv <> True And CbConsole = False Then CbMsgRcv = True: FrmMesej.Show
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [CmdName] - Return Command Name
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Function CmdName(DataString As String) As String
    Dim l_dataLen As Long, l As Long
    
    l_dataLen = Len(DataString)
    If l_dataLen = 0 Then Exit Function
    
    For l = 1 To l_dataLen
        If Mid$(DataString, l, 1) = ":" Then
            CmdName = Mid$(DataString, 2, l - 2)
            Exit Function
        End If
    Next l
    
    CmdName = Mid$(DataString, 2)
End Function

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [CmdValue] - Return Command Value
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Function CmdValue(DataString As String) As String
    Dim l_dataLen As Long, a As Long
    
    l_dataLen = Len(DataString)
    If l_dataLen = 0 Then Exit Function
    
    For a = 1 To l_dataLen
        If Mid$(DataString, a, 1) = ":" Then
            CmdValue = Mid$(DataString, a + 1)
            Exit Function
        End If
    Next a
End Function

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [SubVal] - Return Sub Command Value
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
'FIXIT: Declare 'Val' and 'Default' with an early-bound data type                          FixIT90210ae-R1672-R1B8ZE
Function SubVal(DataString As String, SubCmdName As String, Optional Default As Variant = "", Optional Special As Long = 0) As Variant
On Error GoTo ErrInt
    Dim s_CmdVal As String, l_PosCmdName As Long
    Dim l_PosCmdDataA As Long, l_PosCmdDataB As Long
    ' Note
    '   format bagi sub command
    '   {subcmdname|'data'}{subcmdname2|'data'}
    '
    
    l_PosCmdName = InStr(1, DataString, "{" & LCase$(SubCmdName))
    If l_PosCmdName = 0 Then GoTo ErrInt
    l_PosCmdDataA = InStr(l_PosCmdName, DataString, "|'") + 2
    l_PosCmdDataB = InStr(l_PosCmdDataA, DataString, "'}")
    SubVal = Mid$(DataString, l_PosCmdDataA, l_PosCmdDataB - l_PosCmdDataA)
    If SubVal = "" Then SubVal = Default
Exit Function
ErrInt:
    SubVal = Default
End Function

Function SubBuild(DataName As String, DataValue As String) As String
    SubBuild = "{" & DataName & "|'" & DataValue & "'}"
End Function

