Attribute VB_Name = "MdlNetwork"
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : MdlNetwork
'    Project    : CafeBonzerAG
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Private Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type

Private Const WM_COPYDATA = &H4A
Private Const WM_DESTROY = &H2

Public Const GWL_WNDPROC = (-4)

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lngParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private HwndHost As Long
Private LngPrevWndProc As Long


'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [NetConnect] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub NetConnect()
    FrmHost.Connecter.Enabled = True
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [NetSessionStart] - Start a Session With  Server
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub NetSessionStart()
On Error GoTo ErrInt
    NetSend "020040" & CmdSubPut("NETMAC", SysNetGetMac)
    NetSend "020030" & SysDevPrintersGet
    NetSend "040020" & CmdSubPut("DB", "NetDefaultPass")
    BlnConnected = True
    FrmHost.Pinger = True
    
    If LngStatusLock = 1 Then FrmKey.StatIcon (Connected)
    If LngStatusTray = 1 Then Call TrayEdit(FrmHost.IconStatOn.Picture.Handle)
Exit Sub

ErrInt:
    AppErrorLog Err, "MdlNetwork | NetSessionStart"
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [NetClose] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub NetClose()
On Error GoTo ErrInt
    BlnConnected = False
    FrmHost.Pinger = False
    FrmHost.Socket.Disconnect

    If LngStatusLock = 1 Then FrmKey.StatIcon (Discconnet)
    If LngStatusTray = 1 Then Call TrayEdit(FrmHost.IconStatOff.Picture.Handle)
Exit Sub

ErrInt:
    AppErrorLog Err, "NetClose | MdlNetwork"
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [NetPing] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub NetPing()
    Call NetSend("010010")
End Sub


'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [NetSend] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub NetSend(Data)
On Error GoTo ErrInt
    Dim StrData As String
    
    If Trim$(Data) = "" Then Exit Sub
    StrData = StrCmdSep + CStr(Data)
    
    If FrmHost.Socket.IsWritable = True Then
        FrmHost.Socket.SendLen = Len(StrData)
        FrmHost.Socket.SendData = StrData
    End If
Exit Sub

ErrInt:
    If Err.Number = 24054 Or Err.Number = 24022 Then
        Call NetClose
        Call NetConnect
    Else
        AppErrorLog Err, "Module Net | NetSend"
    End If
End Sub
