Attribute VB_Name = "MdlApplication"
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : MdlApplication
'    Project    : CafeBonzerAG
'
'    Description: Main Module
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>

Public StrNetHost As String
Public StrNetPort As String

Public lMonSwitch As Long
Public BlnAppFirstTime As Boolean
Public BlnConnected As Boolean
Public LngStatusLock As Long
Public LngStatusTray As Long

Public StrTickMsgWelcome As String
Public StrAppVersion As String
Public StrAppBuild As String
Public LngErrorType As Long
Public LngEnvPlatformId As Long
Public LngEnvRegistryRoot As Long

Public StrCmdSep As String
Public StrCmdSubSep1 As String
Public StrCmdSubSep2 As String

Public Const StrDesktopName As String = "CbDaemonShell"
Public Const StrShellViewClass As String = "Progman"
Public Const StrShellTaskBarClass As String = "Shell_TrayWnd"

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Program Entry Points] - Di mana semuanya bermula.....
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Sub Main()
    If App.PrevInstance = True Then End

    Call SettingFirstLoad
    Call SettingEnv
    Call NetConnect
    Call SettingProtect
End Sub

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' AppExit
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub AppExit()
    'Call IPCSendData("TFrmMain", "ACTIONLETSGO")
    Call NetClose
    Call TickerStop
    Call TrayRemove
    FrmHost.Socket.Cleanup
    CDesktop.ClearUp
    
    For Each Form In Forms
        Unload Form
    Next
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Change Priority] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub AppSetPriority(Optional pNormal As Boolean = False)
On Error GoTo ErrInt
    Dim pId As Long, hProcess As Long
    
    pId = GetCurrentProcessId
    hProcess = OpenProcess(PROCESS_DUP_HANDLE, True, pId)
    If pNormal = False Then
        SetPriorityClass hProcess, REALTIME_PRIORITY_CLASS
    Else
        SetPriorityClass hProcess, IDLE_PRIORITY_CLASS
    End If
    Call CloseHandle(hProcess)
Exit Sub

ErrInt:
    AppErrorLog Err, "Module System | AppSetPriority"
End Sub

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Error handler
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub AppErrorLog(ErrObj As ErrObject, ProcName As String)
On Error GoTo ErrInt
    Dim IntErrNum As Integer, StrErrDesc As String, StrErrSource As String
    'Dim ErrDesc As String
    
    IntErrNum = ErrObj.Number
    StrErrSource = ErrObj.Source
    StrErrDesc = ErrObj.Description
    
    Select Case LngErrorType
    Case 1
        MsgBox IntErrNum & " / " & StrErrSource & vbNewLine & StrErrDesc, vbExclamation, ProcName
    End Select
    
    'ErrDesc = Now & " - " & s_errDesc & " - " & s_errSource & " - " & i_errNum
    'Open "AgErrLog.txt" For Append As #1
    'Write #1, ErrDesc
    'Close #1
Exit Sub
ErrInt:
    MsgBox Err.Number & " / " & Err.Source & vbNewLine & Err.Description, vbExclamation, "Error Handler"
End Sub

Public Function ProcLowLevelKeyboard(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Dim p As KBDLLHOOKSTRUCT, BlnKeyEat As Boolean
   Dim BlnKeyAltTab As Boolean, BlnKeyAltEsc As Boolean
   Dim BlnKeyCtlEsc As Boolean, BlnKeyWin As Boolean
   
   If (nCode = HC_ACTION) Then
        If wParam = WM_KEYDOWN Or wParam = WM_SYSKEYDOWN Or wParam = WM_KEYUP Or wParam = WM_SYSKEYUP Then
           '[ Retrive Key Data Structure ]'
            CopyMemory p, ByVal lParam, Len(p)
           '[ Definining Key Set ]'
            BlnKeyAltTab = (p.vkCode = VK_TAB) And ((p.flags And LLKHF_ALTDOWN) <> 0)
            BlnKeyAltEsc = (p.vkCode = VK_ESCAPE) And ((p.flags And LLKHF_ALTDOWN) <> 0)
            BlnKeyCtlEsc = (p.vkCode = VK_ESCAPE) And ((GetKeyState(VK_CONTROL) And &H8000) <> 0)
            BlnKeyWin = p.vkCode = VK_LWIN Or p.vkCode = VK_RWIN
           '[ Key Set To Eat ]'
            BlnKeyEat = BlnKeyAltTab Or BlnKeyAltEsc Or BlnKeyCtlEsc Or BlnKeyWin
        End If
    End If
    
    If BlnKeyEat Then
    '[ Eat KeyStroke ]'
        ProcLowLevelKeyboard = -1
    Else
    '[ Continue Hook ]'
        ProcLowLevelKeyboard = CallNextHookEx(0, nCode, wParam, ByVal lParam)
    End If
End Function
