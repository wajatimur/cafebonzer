Attribute VB_Name = "MdlConfiguration"
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : MdlConfiguration
'    Project    : CafeBonzerAG
'
'    Description: Configuration Module
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
' New Setting Redefined
'
'   SysDisCad       = Disable Ctl+Alt+Del
'   AppFirstTime    = First Time Flag
'   TickGuiDisable  = Disable Ticker
'   TickMsgWelcome  = Ticker Welcome Message
'
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
    Public Const REG_SZ = 1                         ' Unicode null terminated string
    Public Const HKEY_CLASSES_ROOT = &H80000000
    Public Const HKEY_CURRENT_CONFIG = &H80000005
    Public Const HKEY_CURRENT_USER = &H80000001
    Public Const HKEY_DYN_DATA = &H80000006
    Public Const HKEY_LOCAL_MACHINE = &H80000002
    Public Const HKEY_PERFORMANCE_DATA = &H80000004
    Public Const HKEY_USERS = &H80000003
    Public Const ERROR_SUCCESS = 0&

Public Const CStrTickMsgWelcome = "[ CafeBonzer System - Welcome ]"
Public Const CStrSettingPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Security"
Public Const CStrAutoStartPath = "Software\Microsoft\Windows\CurrentVersion\Run"

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub SettingEnv()
  '[ Global Variable ]
    StrTickMsgWelcome = SettingGet("TickMsgWelcome", CStrTickMsgWelcome)
    StrAppVersion = "CafeBonzer v" & App.Major & "." & App.Minor
    StrAppBuild = App.Minor & "." & App.Revision
    LngErrorType = 1
    LngEnvPlatformId = SysInfoGetOs(GetPlatformid)
    
    StrNetHost = SettingGet("NetServerIp")
    StrNetPort = SettingGet("NetServerPort")
    
  '[ Command Separator ]
    StrCmdSep = Chr$(20)
    StrCmdSubSep1 = Chr$(210)
    StrCmdSubSep2 = Chr$(220)

    If LngEnvPlatformId = 1 Then
        LngEnvRegistryRoot = HKEY_LOCAL_MACHINE
    Else
        'CDesktop.Create StrDesktopName
        'CDesktop.StartProcess App.Path & "\CbDaemon.exe"
        LngEnvRegistryRoot = HKEY_CURRENT_USER
    End If
    If SettingGet("AppAutoStart") = 1 Then
        SaveString LngEnvRegistryRoot, "Software\Microsoft\Windows\CurrentVersion\Run", "Component", "CbAg.exe"
    Else
        DeleteValue LngEnvRegistryRoot, "Software\Microsoft\Windows\CurrentVersion\Run", "Component"
    End If
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub SettingFirstLoad()
    If SettingGet("AppFirstTime") = "" Then
        BlnAppFirstTime = True
        FrmMain.Show vbModal
        End
    End If
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub SettingProtect()
  ' + PLATFORM DEPENDENT SETTINGS ------------------------------------------
    If LngEnvPlatformId = 1 Then
        RegisterServiceProcess GetCurrentProcessId, 1
        If SettingGet("SysDisCad", 1) = 1 Then SysDisCtlAltDel True
        Call DeskWallProtect
    End If

  ' + GENERAL SETTINGS -----------------------------------------------------
    If SettingGet("TickGuiDisable", 1) = 1 Then TrayStart Else TickerStart FrmTicker
    If SettingGet("SysAutoLock", 1) = 1 Then
        If SettingGet("LOGIN", False) = False Then SysShellLock 1
    End If
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub SettingSave(SettingName As String, SettingValue As String)
    Dim LngHandle As Long
    
    RegCreateKey HKEY_LOCAL_MACHINE, CStrSettingPath, LngHandle
    RegSetValueEx LngHandle, SettingName, 0, REG_SZ, ByVal SettingValue, Len(SettingValue)
    RegCloseKey LngHandle
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Function SettingGet(SettingName As String, Optional Default As Variant = "") As Variant
    Dim LngHandle As Long
    Dim LngDataBufSize As Long, LngDataType As Long
    Dim LngResult As Long, StrDataBuf As String
    SettingGet = Default
    
    RegOpenKey HKEY_LOCAL_MACHINE, CStrSettingPath, LngHandle
    RegQueryValueEx LngHandle, SettingName, 0&, LngDataType, ByVal 0&, LngDataBufSize
    If LngDataType = REG_SZ Then
        StrDataBuf = String$(LngDataBufSize, " ")
        LngResult = RegQueryValueEx(LngHandle, SettingName, 0&, 0&, ByVal StrDataBuf, LngDataBufSize)
        If LngResult = ERROR_SUCCESS Then
            If LngDataBufSize = 0 Then
                SettingGet = ""
            Else
                SettingGet = Left$(StrDataBuf, InStr(StrDataBuf, Chr$(0)) - 1)
            End If
        End If
    End If
    RegCloseKey LngHandle
End Function
