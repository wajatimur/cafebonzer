VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPolicies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private lFileNumDP As Long

Private PolCategory(1 To 4) As ut_PolCategory
Private PolInfo_Exp(1 To 14) As ut_PolKey
Private PolInfo_Net(1 To 6) As ut_PolKey
Private PolInfo_Sys(1 To 13) As ut_PolKey
Private PolInfo_Dos(1 To 2) As ut_PolKey


Private Type ut_DisApp
    ID As Long
    Enable As Boolean
    ProgramName As String * 32
End Type

Private Type ut_PolKey
    Key As String
    Description As String
End Type

Private Type ut_PolCategory
    Category As String
    Path As String
End Type


Public Enum en_PolCategory
    pcExplorer = 1
    pcNetwork = 2
    pcSystem = 3
    pcDos = 4
End Enum

Public Enum en_PolExp
    NoRun = 1
    NoFind = 2
    NoSetFolders = 3
    NoSetTaskbar = 4
    NoRecentDocsHistory = 5
    ClearRecentDocsHistory = 6
    NoDeletePrinter = 7
    NoAddPrinter = 8
    NoClose = 9
    NoSaveSettings = 10
    NoDesktop = 11
    NoNetHood = 12
    NoInternetIcon = 13
    NoPrinterTabs = 14
End Enum

Public Enum en_PolNet
    NoNetSetup = 1
    NoNetSetupIDPage = 2
    NoNetSetupSecurityPage = 3
    NoFileSharingControl = 4
    NoFileSharing = 5
    NoPrintSharing = 6
End Enum

Public Enum en_PolSys
    NoDispCPL = 1
    NoDispBackgroundPage = 2
    NoDispScrSavPage = 3
    NoDispAppearancePage = 4
    NoDispSettingPage = 5
    NoConfigPage = 6
    NoDevMgrPage = 7
    NoFileSysPage = 8
    NoVirtMemPage = 9
    NoAdminPage = 10
    NoProfilePage = 11
    NoPwdPage = 12
    NoSecCPL = 13
End Enum

Public Enum en_PolDos
    Disabled = 1
    NoRealMode = 2
End Enum

Private Const DisAppPath = "\NsCbDapp.set"

Private Const pSys = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
Private Const pExp = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
Private Const pNet = "Software\Microsoft\Windows\CurrentVersion\Policies\Network"
Private Const pDos = "Software\Microsoft\Windows\CurrentVersion\Policies\Dos"

Private Const REG_DWORD As Long = 4
Private Const KEY_ALL_ACCESS As Long = &H3F
Private Const HKEY_CURRENT_USER As Long = &H80000001
Private Const ERROR_SUCCESS As Long = 0

Public Event EnumerateDisProg(sKeyPath As String, l_ID As Long, sProgName As String, sEnable As Boolean)


Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Sub DisProgAdd(ProgName As String, Optional ProgID As Long, Optional RegistryOnly As Boolean = False)
    Dim RecordMax As Long
    Dim utDA As ut_DisApp
    
    utDA.Enable = True
    utDA.ProgramName = ProgName
    
   ' dapatkan jumlah record
    RecordMax = LOF(lFileNumDP) \ Len(utDA)
    
   ' simpan ke file
    If RegistryOnly = False Then
        utDA.ID = RecordMax + 1
    Else
        utDA.ID = ProgID
    End If
    
    Put #lFileNumDP, utDA.ID, utDA
   ' simpan ke registry
    regCreate_LongValue HKEY_CURRENT_USER, pExp, "RestrictRun", 1
    regCreate_StringValue HKEY_CURRENT_USER, pExp & "\RestrictRun", utDA.ID, utDA.ProgramName
End Sub

Public Sub DisProgEnum(Optional NoEvent As Boolean = False)
    Dim lRecordMax As Long, lRecordNum As Long, sProgName As String
    Dim utDA As ut_DisApp, l_Cnt As Long
    
    lRecordMax = LOF(lFileNumDP) \ Len(utDA)
    regDelete_SubKey HKEY_CURRENT_USER, pExp, "RestrictRun"
    RegDeleteKey HKEY_CURRENT_USER, pExp & "\RestrictRun"
    
    For lRecordNum = lRecordMax To 1 Step -1
        Get #lFileNumDP, lRecordNum, utDA
        sProgName = Trim$(utDA.ProgramName)
        If utDA.Enable = True Then
            l_Cnt = l_Cnt + 1
            regCreate_LongValue HKEY_CURRENT_USER, pExp, "RestrictRun", 1
            regCreate_StringValue HKEY_CURRENT_USER, pExp & "\RestrictRun", l_Cnt, sProgName
        End If
        If NoEvent = False Then
            RaiseEvent EnumerateDisProg(pExp & "\RestrictRun", utDA.ID, sProgName, utDA.Enable)
        End If
    Next lRecordNum
    If l_Cnt > 0 Then regCreate_StringValue HKEY_CURRENT_USER, pExp & "\RestrictRun", l_Cnt + 1, App.EXEName & ".exe"
  '( pening kepala aku buat benda nih - 7,Nov,2002 )'
End Sub

Public Sub DisProgRemove(ProgID As Long, Optional RegistryOnly As Boolean = False)
    Dim arrDA() As ut_DisApp, l_Cnt As Long, sDisProgPath As String
    Dim lRecordMax As Long, lRecordNum As Long, utDA As ut_DisApp
    
    lRecordMax = LOF(lFileNumDP) \ Len(utDA)
    If ProgID > lRecordMax Then Exit Sub
    ReDim arrDA(lRecordMax)
    
    For lRecordNum = lRecordMax To 1 Step -1
        Get #lFileNumDP, lRecordNum, utDA
        If RegistryOnly = False Then
            If utDA.ID <> ProgID Then
                l_Cnt = l_Cnt + 1
                arrDA(l_Cnt) = utDA
            End If
        Else
            l_Cnt = l_Cnt + 1
            If utDA.ID = ProgID Then utDA.Enable = False
            arrDA(l_Cnt) = utDA
        End If
    Next lRecordNum
    
    If RegistryOnly = False Then
        sDisProgPath = App.Path & DisAppPath
        Close #lFileNumDP
        Kill sDisProgPath
        
        lFileNumDP = FreeFile
        Open sDisProgPath For Random Lock Read Write As #lFileNumDP Len = Len(utDA)
        ReDim Preserve arrDA(l_Cnt)
    End If
    
    For lRecordNum = 1 To UBound(arrDA)
        arrDA(lRecordNum).ID = lRecordNum
        Put #lFileNumDP, lRecordNum, arrDA(lRecordNum)
    Next lRecordNum
End Sub


'=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~
' [ Policies ]
'
'=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~
'>> Get Policies by Bool
Public Function GetPolicies(PolCat As en_PolCategory, PolIndex As Long) As Boolean
    Dim polKey As String
    Select Case PolCat
    Case 1
        polKey = PolInfo_Exp(PolIndex).Key
    Case 2
        polKey = PolInfo_Net(PolIndex).Key
    Case 3
        polKey = PolInfo_Sys(PolIndex).Key
    Case 4
        polKey = PolInfo_Dos(PolIndex).Key
    End Select
    
    GetPolicies = regValue_Exist(HKEY_CURRENT_USER, PolCategory(PolCat).Path, polKey)
End Function
'>> Get Policies by Long
Public Function GetPolicies2(PolCat As en_PolCategory, PolIndex As Long) As Long
    Dim polKey As String
    Select Case PolCat
    Case 1
        polKey = PolInfo_Exp(PolIndex).Key
    Case 2
        polKey = PolInfo_Net(PolIndex).Key
    Case 3
        polKey = PolInfo_Sys(PolIndex).Key
    Case 4
        polKey = PolInfo_Dos(PolIndex).Key
    End Select
    
    GetPolicies2 = regGet_LongValue(HKEY_CURRENT_USER, PolCategory(PolCat).Path, polKey)
End Function

'>> Get Policies Long
Public Function GetPoliciesLong(hKey As Long, PolCat As en_PolCategory, PolicyKey As String) As Long
    Dim s_KeyPath As String
    s_KeyPath = Choose(PolCat, pExp, pNet, pSys, pDos)
    
    GetPoliciesLong = regGet_LongValue(HKEY_CURRENT_USER, s_KeyPath, PolicyKey)
End Function

'>> Get Policies Count
Public Function GetPoliciesCount(PolCat As en_PolCategory) As Long
    Select Case PolCat
    Case 1
        GetPoliciesCount = UBound(PolInfo_Exp)
    Case 2
        GetPoliciesCount = UBound(PolInfo_Net)
    Case 3
        GetPoliciesCount = UBound(PolInfo_Sys)
    Case 4
        GetPoliciesCount = UBound(PolInfo_Dos)
    End Select
End Function

'>> Set Policies
Public Sub SetPolicies(PolCat As en_PolCategory, PolIndex As Long, Optional Enabled As Boolean = True)
    Dim polKey As String
    Select Case PolCat
    Case 1
        polKey = PolInfo_Exp(PolIndex).Key
    Case 2
        polKey = PolInfo_Net(PolIndex).Key
    Case 3
        polKey = PolInfo_Sys(PolIndex).Key
    Case 4
        polKey = PolInfo_Dos(PolIndex).Key
    End Select
    
    If Enabled = True Then
        regCreate_LongValue HKEY_CURRENT_USER, PolCategory(PolCat).Path, polKey, 1
    Else
        regDelete_SubKey HKEY_CURRENT_USER, PolCategory(PolCat).Path, polKey
    End If
End Sub

Public Function GetPoliciesDesc(PolCat As en_PolCategory, PolIndex As Long)
    Select Case PolCat
    Case 1
        GetPoliciesDesc = PolInfo_Exp(PolIndex).Description
    Case 2
        GetPoliciesDesc = PolInfo_Net(PolIndex).Description
    Case 3
        GetPoliciesDesc = PolInfo_Sys(PolIndex).Description
    Case 4
        GetPoliciesDesc = PolInfo_Dos(PolIndex).Description
    End Select
End Function



'=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~
' [ Properties ]
'
'
'=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~
Public Property Let Exp_NoRun(nVal As Boolean)

End Property
Public Property Get Exp_NoRun() As Boolean

End Property


Public Property Let Exp_NoFind(nVal As Boolean)

End Property
Public Property Get Exp_NoFind() As Boolean

End Property


Public Property Let Exp_NoSetFolders(nVal As Boolean)

End Property
Public Property Get Exp_NoSetFolders() As Boolean

End Property


Public Property Let Exp_NoTaskbar(nVal As Boolean)

End Property
Public Property Get Exp_NoTaskbar() As Boolean

End Property


Public Property Let Exp_NoRecentDocsHistory(nVal As Boolean)

End Property
Public Property Get Exp_NoRecentDocsHistory() As Boolean

End Property


Public Property Let Exp_ClearRecentDocsHistory(nVal As Boolean)

End Property
Public Property Get Exp_ClearRecentDocsHistory() As Boolean

End Property


Public Property Let Dsk_NoDeletePrinter(nVal As Boolean)

End Property
Public Property Get Dsk_NoDeletePrinter() As Boolean

End Property


Public Property Let Dsk_NoAddPrinter(nVal As Boolean)

End Property
Public Property Get Dsk_NoAddPrinter() As Boolean

End Property


Public Property Let Dsk_NoClose(nVal As Boolean)

End Property
Public Property Get Dsk_NoClose() As Boolean

End Property


Public Property Let Dsk_NoSaveSettings(nVal As Boolean)

End Property
Public Property Get Dsk_NoSaveSettings() As Boolean

End Property


Public Property Let Dsk_NoDesktop(nVal As Boolean)

End Property
Public Property Get Dsk_NoDesktop() As Boolean

End Property


Public Property Let Dsk_NoDrives(nVal As Boolean)

End Property
Public Property Get Dsk_NoDrives() As Boolean

End Property


Public Property Let Dsk_NoNetHood(nVal As Boolean)

End Property
Public Property Get Dsk_NoNetHood() As Boolean

End Property


Public Property Let Dsk_NoInternetIcon(nVal As Boolean)

End Property
Public Property Get Dsk_NoInternetIcon() As Boolean

End Property


Public Property Let Dsk_NoPrinterTabs(nVal As Boolean)

End Property
Public Property Get Dsk_NoPrinterTabs() As Boolean

End Property


Public Property Let Net_NoNetSetup(nVal As Boolean)

End Property
Public Property Get Net_NoNetSetup() As Boolean

End Property


Public Property Let Net_NoNetSetupIDPage(nVal As Boolean)

End Property
Public Property Get Net_NoNetSetupIDPage() As Boolean

End Property


Public Property Let Net_NoNetSetupSecurityPage(nVal As Boolean)

End Property
Public Property Get Net_NoNetSetupSecurityPage() As Boolean

End Property


Public Property Let Net_NoFileSharingControl(nVal As Boolean)

End Property
Public Property Get Net_NoFileSharingControl() As Boolean

End Property


Public Property Let Net_NoFileSharing(nVal As Boolean)

End Property
Public Property Get Net_NoFileSharing() As Boolean

End Property


Public Property Let Net_NoPrintSharing(nVal As Boolean)

End Property
Public Property Get Net_NoPrintSharing() As Boolean

End Property


Public Property Let Sys_NoDispCPL(nVal As Boolean)

End Property
Public Property Get Sys_NoDispCPL() As Boolean

End Property


Public Property Let Sys_NoDispBackgroundPage(nVal As Boolean)

End Property
Public Property Get Sys_NoDispBackgroundPage() As Boolean

End Property


Public Property Let Sys_NoDispScrSavPage(nVal As Boolean)

End Property
Public Property Get Sys_NoDispScrSavPage() As Boolean

End Property


Public Property Let Sys_NoDispAppearancePage(nVal As Boolean)

End Property
Public Property Get Sys_NoDispAppearancePage() As Boolean

End Property


Public Property Let Sys_NoDispSettingPage(nVal As Boolean)

End Property
Public Property Get Sys_NoDispSettingPage() As Boolean

End Property


Public Property Let Sys_NoConfigPage(nVal As Boolean)

End Property
Public Property Get Sys_NoConfigPage() As Boolean

End Property


Public Property Let Sys_NoDevMgrPage(nVal As Boolean)

End Property
Public Property Get Sys_NoDevMgrPage() As Boolean

End Property


Public Property Let Sys_NoFileSysPage(nVal As Boolean)

End Property
Public Property Get Sys_NoFileSysPage() As Boolean

End Property


Public Property Let Sys_NoVirtMemPage(nVal As Boolean)

End Property
Public Property Get Sys_NoVirtMemPage() As Boolean

End Property


Public Property Let Sys_NoAdminPage(nVal As Boolean)

End Property
Public Property Get Sys_NoAdminPage() As Boolean

End Property


Public Property Let Sys_NoProfilePage(nVal As Boolean)

End Property
Public Property Get Sys_NoProfilePage() As Boolean

End Property


Public Property Let Sys_NoPwdPage(nVal As Boolean)

End Property
Public Property Get Sys_NoPwdPage() As Boolean

End Property


Public Property Let Sys_NoSecCPL(nVal As Boolean)

End Property
Public Property Get Sys_NoSecCPL() As Boolean

End Property


Public Property Let Dos_Disabled(nVal As Boolean)

End Property
Public Property Get Dos_Disabled() As Boolean

End Property


Public Property Let Dos_NoRealMode(nVal As Boolean)

End Property
Public Property Get Dos_NoRealMode() As Boolean

End Property


Public Property Let Misc_DriveHide(nVal As Long)
    regCreate_LongValue HKEY_CURRENT_USER, pExp, "NoDrives", nVal
End Property
Public Property Get Misc_DriveHide() As Long
    Misc_DriveHide = regGet_LongValue(HKEY_CURRENT_USER, pExp, "NoDrives")
End Property


Public Property Let Misc_DisableRegedit(nVal As Long)
    regCreate_LongValue HKEY_CURRENT_USER, pSys, "DisableRegistryTools", nVal
End Property
Public Property Get Misc_DisableRegedit() As Long
    Misc_DisableRegedit = regGet_LongValue(HKEY_CURRENT_USER, pSys, "DisableRegistryTools")
End Property


Private Sub Class_Initialize()
    Dim utDA As ut_DisApp
    
    lFileNumDP = FreeFile
    Open App.Path & DisAppPath For Random Lock Read Write As #lFileNumDP Len = Len(utDA)

    'Policies category
    PolCategory(1).Category = "Explorer"
    PolCategory(1).Path = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    PolCategory(2).Category = "Network"
    PolCategory(2).Path = "Software\Microsoft\Windows\CurrentVersion\Policies\Network"
    PolCategory(3).Category = "System"
    PolCategory(3).Path = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
    PolCategory(4).Category = "Dos"
    PolCategory(4).Path = "Software\Microsoft\Windows\CurrentVersion\Policies\Dos"


    ' Menu staff
    PolInfo_Exp(1).Key = "NoRun"
    PolInfo_Exp(1).Description = "Hide Run Menu"
    PolInfo_Exp(2).Key = "NoFind"
    PolInfo_Exp(2).Description = "Hide Find Menu"
    PolInfo_Exp(3).Key = "NoSetFolders"
    PolInfo_Exp(3).Description = "Hide Folders in Settings Menu"
    PolInfo_Exp(4).Key = "NoSetTaskbar"
    PolInfo_Exp(4).Description = "Hide Taskbar in Settings Menu"
    PolInfo_Exp(5).Key = "NoRecentDocsHistory"
    PolInfo_Exp(5).Description = "Hide Recent Documents Menu"
    PolInfo_Exp(6).Key = "ClearRecentDocsHistory"
    PolInfo_Exp(6).Description = "Clear Recent Documents Menu"
    
    
    'desktop staff
    PolInfo_Exp(7).Key = "NoDeletePrinter"
    PolInfo_Exp(7).Description = "Disable printer deletion"
    PolInfo_Exp(8).Key = "NoAddPrinter"
    PolInfo_Exp(8).Description = "Disable printer addition"
    PolInfo_Exp(9).Key = "NoClose"
    PolInfo_Exp(9).Description = "Shutdown Command"
    PolInfo_Exp(10).Key = "NoSaveSettings"
    PolInfo_Exp(10).Description = "Don't save Settings"
    PolInfo_Exp(11).Key = "NoDesktop"
    PolInfo_Exp(11).Description = "Hide all Items on the Desktop"
    PolInfo_Exp(12).Key = "NoNetHood"
    PolInfo_Exp(12).Description = "Network icon from Desktop"
    PolInfo_Exp(13).Key = "NoInternetIcon"
    PolInfo_Exp(13).Description = "Internet icon from Desktop"
    PolInfo_Exp(14).Key = "NoPrinterTabs"
    PolInfo_Exp(14).Description = "Details and General Pages"
    
    
    'Network
    PolInfo_Net(1).Key = "NoNetSetup"
    PolInfo_Net(1).Description = "Disable Network Control Panel"
    PolInfo_Net(2).Key = "NoNetSetupIDPage"
    PolInfo_Net(2).Description = "Disable Identification Page"
    PolInfo_Net(3).Key = "NoNetSetupSecurityPage"
    PolInfo_Net(3).Description = "Disable Access Control Page"
    PolInfo_Net(4).Key = "NoFileSharingControl"
    PolInfo_Net(4).Description = "Disable File and Print Sharing Controls"
    PolInfo_Net(5).Key = "NoFileSharing"
    PolInfo_Net(5).Description = "Disable File Sharing"
    PolInfo_Net(6).Key = "NoPrintSharing"
    PolInfo_Net(6).Description = "Disable Print Sharing"
    
    
    'System
    PolInfo_Sys(1).Key = "NoDispCPL"
    PolInfo_Sys(1).Description = "Disable Display System Control Panel"
    PolInfo_Sys(2).Key = "NoDispBackgroundPage"
    PolInfo_Sys(2).Description = "Disable Display Control Background Page"
    PolInfo_Sys(3).Key = "NoDispScrSavPage"
    PolInfo_Sys(3).Description = "Disable Display Control Screen Saver Page"
    PolInfo_Sys(4).Key = "NoDispAppearancePage"
    PolInfo_Sys(4).Description = "Disable Display Control Appearance Page"
    PolInfo_Sys(5).Key = "NoDispSettingPage"
    PolInfo_Sys(5).Description = "Disable Display Control Setting Page"
    PolInfo_Sys(6).Key = "NoConfigPage"
    PolInfo_Sys(6).Description = "Disable Configuration Page"
    PolInfo_Sys(7).Key = "NoDevMgrPage"
    PolInfo_Sys(7).Description = "Disable Device Manager Page"
    PolInfo_Sys(8).Key = "NoFileSysPage"
    PolInfo_Sys(8).Description = "Disable File System Page"
    PolInfo_Sys(9).Key = "NoVirtMemPage"
    PolInfo_Sys(9).Description = "Disable Virtual Memory Page"
    PolInfo_Sys(10).Key = "NoAdminPage"
    PolInfo_Sys(10).Description = "Disable Remote Administration Page"
    PolInfo_Sys(11).Key = "NoProfilePage"
    PolInfo_Sys(11).Description = "Disable User Profiles Page"
    PolInfo_Sys(12).Key = "NoPwdPage"
    PolInfo_Sys(12).Description = "Disable Change Passwords Page"
    PolInfo_Sys(13).Key = "NoSecCPL"
    PolInfo_Sys(13).Description = "Disable Password Control Panel"
    
    
    ' Dos
    PolInfo_Dos(1).Key = "Disabled"
    PolInfo_Dos(1).Description = "Disable MS-DOS prompt"
    PolInfo_Dos(2).Key = "NoRealMode"
    PolInfo_Dos(2).Description = "Disable MS-DOS app at real mode"
End Sub

Private Sub Class_Terminate()
    Close #lFileNumDP
End Sub

Private Function regValue_Exist(ByVal hKey As Long, ByVal sRegKeyPath As String, ByVal sRegSubKey As String) As Boolean
    Dim lKeyHandle As Long
    Dim lRet As Long
    Dim lDataType As Long
    Dim lBufferSize As Long
    lKeyHandle = 0
    lRet = RegOpenKey(hKey, sRegKeyPath, lKeyHandle)
    If lKeyHandle <> 0 Then lRet = RegQueryValueEx(lKeyHandle, sRegSubKey, 0&, lDataType, ByVal 0&, lBufferSize)
    If lRet = ERROR_SUCCESS Then
       regValue_Exist = True
       lRet = RegCloseKey(lKeyHandle)
    Else
       regValue_Exist = False
    End If
End Function

Private Sub regDelete_SubKey(ByVal hKey As Long, ByVal sRegKeyPath As String, ByVal sRegSubKey As String)
    Dim lKeyHandle As Long
    Dim lRet As Long
    If regValue_Exist(hKey, sRegKeyPath, sRegSubKey) Then
        lRet = RegOpenKey(hKey, sRegKeyPath, lKeyHandle)
        lRet = RegDeleteValue(lKeyHandle, sRegSubKey)
        lRet = RegCloseKey(lKeyHandle)
    End If
End Sub

Private Sub regCreate_LongValue(ByVal hKey As Long, ByVal sRegKeyPath As String, ByVal sRegSubKey As String, lKeyValue As Long)
    Dim lKeyHandle As Long
    Dim lRet As Long
    Dim lDataType As Long
    lRet = RegCreateKey(hKey, sRegKeyPath, lKeyHandle)
    lRet = RegSetValueEx(lKeyHandle, sRegSubKey, 0&, REG_DWORD, lKeyValue, 4&)
    lRet = RegCloseKey(lKeyHandle)
End Sub

Private Function regGet_LongValue(ByVal hKey As Long, ByVal sRegKeyPath As String, ByVal sRegSubKey As String) As Long
    Dim lKeyHandle As Long
    Dim lRet As Long
    Dim lDataType As Long
    Dim lResult As Long
    lRet = RegOpenKey(hKey, sRegKeyPath, lKeyHandle)
    lRet = RegQueryValueEx(lKeyHandle, sRegSubKey, 0&, lDataType, lResult, 4&)
    If lRet = ERROR_SUCCESS Then
        If lDataType = REG_DWORD Then
            regGet_LongValue = lResult
        End If
    End If
    lRet = RegCloseKey(lKeyHandle)
End Function

Private Sub regCreate_StringValue(ByVal hKey As Long, ByVal sRegKeyPath As String, ByVal sRegSubKey As String, sKeyValue As String)
    Dim Keyhand As Long, r As Long
    r = RegCreateKey(hKey, sRegKeyPath, Keyhand)
    r = RegSetValueEx(Keyhand, sRegSubKey, 0, REG_SZ, ByVal sKeyValue, Len(sKeyValue))
    r = RegCloseKey(Keyhand)
End Sub


