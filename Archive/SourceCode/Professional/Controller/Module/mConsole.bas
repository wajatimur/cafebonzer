Attribute VB_Name = "mConsole"
Public Echo As Boolean

Public Sub CurrentHook()
    idx = GetAgentIndex(FrmTerminal.CurSocket)
    If idx = 0 Then
        FrmTerminal.wr "No socket currently hook !"
    Else
        FrmTerminal.wr "Socket currently hook to " & FrmMain.Lv1.ListItems(idx).Text
    End If
End Sub

Public Sub Hook(StationName)
    idx = GetAgentIndexB(StationName)
    If idx = 0 Then
        FrmTerminal.wr "Station not exist ! - " & StationName
    Else
        FrmTerminal.CurSocket = FrmMain.Lv1.ListItems(idx).Tag
        FrmTerminal.wr "Hooking socket success > " & StationName
    End If
End Sub

Public Sub DisEcho(Param)
    If Param = "" Then Exit Sub
    If Left(Param, 1) = "1" Then
        Echo = True
        FrmTerminal.wr "Echo enable"
    Else
        Echo = False
        FrmTerminal.wr "Echo disable"
    End If
End Sub

Public Sub DkeyVar(Param)
    If Param = "" Then Exit Sub
    CbDrvStr = Param & ":"
    FrmTerminal.wr "DiskKey drive set to " & CbDrvStr
End Sub

Public Sub SendMesej(Param As String, Sck As Long)
    Dim Nama As String, Mesej As String
    If InStr(1, Param, ":") <> 0 Then
        Nama = Mid(Param, 1, InStr(1, Param, ":") - 1)
        Mesej = Mid(Param, InStr(1, Param, ":") + 1)
        Send Sck, "//mesej:" & Nama & ":" & Mesej
    Else
        Nama = "Server"
        Mesej = Param
        Send Sck, "//mesej:Server:" & Param
    End If
    FrmTerminal.wr Nama & ">" & Mesej
End Sub

' kita selalu tidak mahu terjadi sesuatu yang tidak kita suka
' dan kita mesti selalu bersedia menghadapi sesuatu yang tidak kita suka
' azri jamil - oct,2002
