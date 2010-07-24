Attribute VB_Name = "MdlSystem"
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : MdlSystem
'    Project    : CafeBonzerAG
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Enum EnuInfoOsVer
    GetPlatformid = 1
    GetMajorVersion = 2
    GetMinorVersion = 3
    GetBuild = 4
    GetCSDVersion = 5
End Enum
Public LngHookLLKey As Long
Public CDesktop As New ClsDesktop

Public Function SysNetGetMac() As String
   Dim StrMAC As String, lASTAT As Long
   Dim NCB As NET_CONTROL_BLOCK
   Dim AST As ASTAT

  'The IBM NetBIOS 3.0 specifications defines four basic
  'NetBIOS environments under the NCBRESET command. Win32
  'follows the OS/2 Dynamic Link Routine (DLR) environment.
  'This means that the first NCB issued by an application
  'must be a NCBRESET, with the exception of NCBENUM.
  'The Windows NT implementation differs from the IBM
  'NetBIOS 3.0 specifications in the NCB_CALLNAME field.
   NCB.ncb_command = NCBRESET
   NCB.ncb_lana_num = 0
   Call Netbios(NCB)
   
  'To get the Media Access Control (MAC) address for an
  'ethernet adapter programmatically, use the Netbios()
  'NCBASTAT command and provide a "*" as the name in the
  'NCB.ncb_CallName field (in a 16-chr string).
   NCB.ncb_callname = "*               "
   NCB.ncb_command = NCBASTAT
   
  'For machines with multiple network adapters you need to
  'enumerate the LANA numbers and perform the NCBASTAT
  'command on each. Even when you have a single network
  'adapter, it is a good idea to enumerate valid LANA numbers
  'first and perform the NCBASTAT on one of the valid LANA
  'numbers. It is considered bad programming to hardcode the
  'LANA number to 0 (see the comments section below).
   NCB.ncb_lana_num = 0
   NCB.ncb_length = Len(AST)
   lASTAT = HeapAlloc(GetProcessHeap(), HEAP_GENERATE_EXCEPTIONS Or HEAP_ZERO_MEMORY, NCB.ncb_length)
   If lASTAT = 0 Then Exit Function
   NCB.ncb_buffer = lASTAT
   Call Netbios(NCB)
   
   CopyMemory AST, NCB.ncb_buffer, Len(AST)
   StrMAC = Format$(Hex$(AST.adapt.adapter_address(0)), "00") & _
         Format$(Hex$(AST.adapt.adapter_address(1)), "00") & _
         Format$(Hex$(AST.adapt.adapter_address(2)), "00") & _
         Format$(Hex$(AST.adapt.adapter_address(3)), "00") & _
         Format$(Hex$(AST.adapt.adapter_address(4)), "00") & _
         Format$(Hex$(AST.adapt.adapter_address(5)), "00")

   HeapFree GetProcessHeap(), 0, lASTAT
   SysNetGetMac = StrMAC
End Function

Public Sub SysDisCtlAltDel(Opt As Boolean)
    Dim LngResult As Long
    LngResult = SystemParametersInfo(SPI_SCREENSAVERRUNNING, Opt, vbNull, 0)
End Sub

Public Function SysInfoGetOs(Optional InfoOsVersionNumber As EnuInfoOsVer) As Variant
    Dim DstOsVer As OSVERSIONINFO, LngIovn As Long
        
    LngIovn = InfoOsVersionNumber
    DstOsVer.dwOSVersionInfoSize = Len(DstOsVer)
    Call GetVersionEx(DstOsVer)
    
    If LngIovn > 0 Then
        If LngIovn = 1 Then SysInfoGetOs = DstOsVer.dwPlatformId
        If LngIovn = 2 Then SysInfoGetOs = DstOsVer.dwMajorVersion
        If LngIovn = 3 Then SysInfoGetOs = DstOsVer.dwMinorVersion
        If LngIovn = 4 Then SysInfoGetOs = DstOsVer.dwBuildNumber
        If LngIovn = 5 Then SysInfoGetOs = DstOsVer.szCSDVersion
        Exit Function
    End If
    
    Select Case DstOsVer.dwMajorVersion
    Case 4
        Select Case DstOsVer.dwMinorVersion
        Case 0
            Select Case DstOsVer.dwBuildNumber
            Case 950
            
            Case 1111
            
            Case 1381
            
            End Select
        Case 10
            Select Case DstOsVer.dwBuildNumber
            Case 1998
            
            Case 2222
            
            End Select
        Case 90
        
        End Select
    Case 5
        Select Case DstOsVer.dwMinorVersion
        Case 0
        
        Case 1
        
        End Select
    End Select
End Function

Public Function SysInfoGetName() As String
    Dim StrNama As String
    StrNama = String$(255, Chr$(0))
    GetComputerName StrNama, 255
    StrNama = Left$(StrNama, InStr(1, StrNama, Chr$(0)) - 1)
    SysInfoGetName = StrNama
End Function

Public Sub SysInfoSetName(NetName As String)
    If NetName = "" Then Exit Sub
    SetComputerName NetName
End Sub

Public Function SysDevPrintersGet() As String
    Dim StrPrinter As String, UtPrinter As Printer
    ' Format
    '  total|name|default|port|drivername|... and so on
    
    StrPrinter = CmdSubPut("TOTAL", Printers.Count)
    For Each UtPrinter In Printers
        StrPrinter = StrPrinter & CmdSubPut("NAME", UtPrinter.DeviceName)
        StrPrinter = StrPrinter & CmdSubPut("DEFAULT", UtPrinter.TrackDefault)
        StrPrinter = StrPrinter & CmdSubPut("PORT", UtPrinter.Port)
        StrPrinter = StrPrinter & CmdSubPut("DRIVERNAME", UtPrinter.DriverName)
        StrPrinter = StrPrinter & CmdSubPut("PAPERSIZE", UtPrinter.PaperSize)
        StrPrinter = StrPrinter & CmdSubPut("ORIENTATION", UtPrinter.Orientation)
    Next
    SysDevPrintersGet = StrPrinter
End Function

Public Sub SysDevMonitorOff(OpCode As String)
On Error GoTo ErrInt
    Dim LngRet As Long
    ' Reference
    ' 1 = On
    ' 2 = Off
    ' 3 = Suspend
    
    LngRet = Choose(OpCode, -1, 2, 1)
    LngRet = SendMessage(FrmMain.hWnd, &H112, &HF170, LngRet)
Exit Sub

ErrInt:
    AppErrorLog Err, "Module Command | ScreenOff"
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Cleaning Disk] - Harddisk Cleaning Command
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub SysDevDiskClean(OpCode As String)
On Error GoTo ErrInt
    Dim CWinPath As New ClsWinPath
    Dim StrOpCode As String
    '0 = clean all
    '1 = clean temp. dir
    '2 = clean recycle
    '3 = clean history
    '4 = clean opened doc
    '
    ' Send Status
    '  1 = Clean Ok
    
    StrOpCode = OpCode
    
    If InStr(1, StrOpCode, "0") Or InStr(1, StrOpCode, "1") Then
        Call PathDelTree(CWinPath.Temp)
    ElseIf InStr(1, StrOpCode, "0") Or InStr(1, StrOpCode, "2") Then
        SHEmptyRecycleBin FrmHost.hWnd, vbNullString, SHERB_NOCONFIRMATION + SHERB_NOSOUND
    ElseIf InStr(1, StrOpCode, "0") Or InStr(1, StrOpCode, "3") Then
        Dim uRl As New UrlHistory
        uRl.ClearHistory
    ElseIf InStr(1, StrOpCode, "0") Or InStr(1, StrOpCode, "4") Then
        SHAddToRecentDocs 2, vbNullString
    End If
    
    NetSend "040010" & CmdSubPut("CLEAN", 1)
    Set CWinPath = Nothing
Exit Sub

ErrInt:
    AppErrorLog Err, "Module Command | CleanDisk"
End Sub

Public Sub SysWindowsExit(OpCode As String)
    Dim LngResult As Long
    '0 = shutdown
    '1 = force shutdown
    '2 = Reboot
    '3 = force reboot
    If SysInfoGetOs(GetPlatformid) = 2 Then Call SecAdjustToken
    Select Case Mid$(OpCode, 1, 1)
    Case 0
        LngResult = ExitWindowsEx(EWX_SHUTDOWN, 0)
    Case 1
        LngResult = ExitWindowsEx(EWX_SHUTDOWN Or EWX_FORCE, 0)
    Case 2
        LngResult = ExitWindowsEx(EWX_REBOOT, 0)
    Case 3
        LngResult = ExitWindowsEx(EWX_REBOOT Or EWX_FORCE, 0)
    End Select
End Sub

Public Sub SysWindowsSleep(OpCode As String)
On Error GoTo ErrInt
    If OpCode = "1" Then
        SetSystemPowerState 1, 1
        NetSend "040010" & CmdSubPut("SLEEP", 1)
    Else
        SetSystemPowerState 0, 0
        NetSend "040010" & CmdSubPut("SLEEP", 0)
    End If
Exit Sub

ErrInt:
    AppErrorLog Err, "Module Command | Tidur"
End Sub

Public Sub SysWindowsHook(Install As Boolean)
    If Install = True Then
        LngHookLLKey = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf ProcLowLevelKeyboard, App.hInstance, 0)
    Else
        UnhookWindowsHookEx LngHookLLKey
    End If
End Sub

Public Sub SysWindowsBlockInput(OpCode As String)
On Error GoTo ErrInt
    Dim StrOpCode As String
    StrOpCode = OpCode
    
    Select Case StrOpCode
    Case 1
        BlockInput True
        NetSend "040010" & CmdSubPut("BLOCKINPUT", 1)
    Case 2
        BlockInput False
        NetSend "040010" & CmdSubPut("BLOCKINPUT", 0)
    End Select
Exit Sub

ErrInt:
    AppErrorLog Err, "Module Command | BlockInput"
End Sub

Public Sub SysShellHide(OpCode As String)
    hWndtsk = FindShellTaskBar
    hwnddsk = FindShellWindow
    
    Select Case OpCode
    Case 0
        HideShowWindow hWndtsk, True
        HideShowWindow hwnddsk, True
    Case 1
        HideShowWindow hWndtsk
        HideShowWindow hwnddsk
    End Select
End Sub

Public Sub SysShellLock(OpCode As String)
    Dim LngVal As Long
    
    Select Case CmdSubGet(OpCode, "ACTION")
    Case 0
        If LngStatusLock = 0 Then Exit Sub
        If LngEnvPlatformId = 2 Then
            'Call SysWindowsHook(False)
            CDesktop.Switch True
        Else
            SystemParametersInfo SPI_SCREENSAVERRUNNING, 0, LngVal, 0&
            MinAllWindow False
            Call HideDesktop
            Unload FrmKey
        End If
    Case 1
        If LngStatusLock = 1 Then Exit Sub
        If LngEnvPlatformId = 2 Then
            'Call SysWindowsHook(True)
            CDesktop.Switch
        Else
            SystemParametersInfo SPI_SCREENSAVERRUNNING, 1, LngVal, 0&
            MinAllWindow True
            Call HideDesktop(True)
            FrmKey.Show
        End If

    End Select
    
    LngStatusLock = OpCode
    Call AgentInfoStatus
End Sub
