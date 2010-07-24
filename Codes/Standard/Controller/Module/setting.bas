Attribute VB_Name = "MdlConfiguration"
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : MdlConfiguration
'    Project    : CafeBonzer
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
'==================================================================
' Aplication codename : CafeBonzer
' Programmer          : Azri Jamil a.k.a wajatimur
' Module Name         : Setting
' Description         :
'==================================================================
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
    Public Const REG_SZ = 1                         ' Unicode nul terminated string
    Public Const HKEY_CLASSES_ROOT = &H80000000
    Public Const HKEY_CURRENT_CONFIG = &H80000005
    Public Const HKEY_CURRENT_USER = &H80000001
    Public Const HKEY_DYN_DATA = &H80000006
    Public Const HKEY_LOCAL_MACHINE = &H80000002
    Public Const HKEY_PERFORMANCE_DATA = &H80000004
    Public Const HKEY_USERS = &H80000003
    Public Const ERROR_SUCCESS = 0&

Public EnumArray() As String
Public LngGap As Long



'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Cek jika program buka untuk pertama kali
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub ConfigEnv()
    Dim StrAppFirstTime As String
    'General Setting - 18/08/2002 my 2nd year annivessary with my love one
    
    'Global variables
    CbAppVersion = "CafeBonzer v" & App.Major & "." & App.Minor & " Beta"
    CbAppBuild = App.Minor & "." & App.Revision
    CbAppLatestAgn = "2.0.00"
    
    'Command Seperator
    StrCmdSep = Chr(20)
    StrCmdSubSep1 = Chr(210)
    StrCmdSubSep2 = Chr(220)
    
    'Path untuk database utama semasa
    CurSDBPath = App.Path & "\data\sdata.mdb"
    CurIDBPath = App.Path & "\data\idata.mdb"
    
    'Initialize universal akses database.. untuk kegunaan umum
    CDataSe.InitDb = CurSDBPath
    CDataIe.InitDb = CurIDBPath
    
    Set CDataS = OpenDatabase(CurSDBPath, False, False, ";pwd=nsb2003")
    Set CDataI = OpenDatabase(CurIDBPath, False, False, ";pwd=nsb2003")
    
    StrAppFirstTime = SettingGet("AppFirstTime")
    If StrAppFirstTime = "" Or StrAppFirstTime = "YesFirst" Then
        Unload FrmAppSplash
        FrmSysSet.Show vbModal
    End If
    
    'Loading and checking current session
    OpenSessionCur = SetGetDb("FinSessionClose")
    OpenSessionDay = SetGetDb("FinSessionDay", 0)
    
    'Log user activity
    CbLogUser = SettingGet("SecLogUser", True)
End Sub


Public Sub ConfigForm()
On Error GoTo ErrInt
    Dim CInfoListView As ListView, CInfoListView1 As ListView
    Dim BlnToolBox As Boolean
    
    Set CInfoListView = FrmMain.InfoListView
    Set CInfoListView1 = FrmMain.InfoListView1
       
    Call MetricFrmLoad(FrmMain)
    Call MenuIconAttach
    
    FrmMain.Caption = "CafeBonzer v" & App.Major & "." & CbAppBuild & " Beta - " & SetGetDb("GenOrgMoto", "Actually We Are The Best CyberCafe")
    If CbDemoMode = True Then FrmMain.Caption = FrmMain.Caption & " [Evaluation]"
    FrmMain.MainNote = SetGetDb("AppMainNote")
    
    CInfoListView.ListItems.Add , "CONNECTION", "Connection", , "CHAIN"
    CInfoListView.ListItems.Add , "LOCK", "Lock", , "TerminalLock"
    CInfoListView.ListItems.Add , "STATUS", "Status", , "UserOnline"
    CInfoListView.ListItems.Add , "CURRENTUSAGE", "Usage", , "TIME"
    CInfoListView.ListItems.Add , "TERMCONNECTED", "Connected", , "CHAIN"
    CInfoListView.ListItems.Add , "IPADDRESS", "IP Address", , "NET"
    CInfoListView.ListItems.Add , "MACADDRESS", "MAC Address", , "PCI"
    
    CInfoListView1.ListItems.Add , "UNUSED", "Unused", , 1
    CInfoListView1.ListItems.Add , "TOTALCOUNT", "Total", , 1
    CInfoListView1.ListItems.Add , "CONNECTEDCOUNT", "Connected", , 1
      
    If FileExisted(App.Path & "\" & "CafeSmMgr.exe") Then
        Load FrmMain.MnuToolsModules(1)
        FrmMain.MnuToolsModules(1).Caption = "Services Manager"
        FrmMain.MnuToolsModules(1).Visible = True
    End If
    
    If CbDemoMode = True Then
        FrmMain.MenuInfoLiscenseSub(0).Enabled = True
        FrmMain.MenuInfoLiscenseSub(1).Enabled = False
    Else
        FrmMain.MenuInfoLiscenseSub(0).Enabled = False
        FrmMain.MenuInfoLiscenseSub(1).Enabled = True
    End If
    
Exit Sub

ErrInt:
    AppErrorLog Err, "ConfigForm"
End Sub


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' simpan setting dalam registry
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub SettingSave(SettingName As String, SettingValue As String)
    Dim LngHandle As Long, StrPath As String
    StrPath = "OsSecurity\Data"
    
    RegCreateKey HKEY_CLASSES_ROOT, StrPath, LngHandle
    RegSetValueEx LngHandle, SettingName, 0, REG_SZ, ByVal SettingValue, Len(SettingValue)
    RegCloseKey LngHandle
    
    'SaveString HKEY_CLASSES_ROOT, "OsSecurity\Data", Namasetting, Nilai
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ambil setting dari registry
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Function SettingGet(SettingName As String, Optional Default As Variant = "") As Variant
    Dim LngHandle As Long, StrPath As String
    Dim LngDataBufSize As Long, LngDataType As Long
    Dim LngResult, StrDataBuf As String
    StrPath = "OsSecurity\Data"
    SettingGet = Default
    
    RegOpenKey HKEY_CLASSES_ROOT, StrPath, LngHandle
    RegQueryValueEx LngHandle, SettingName, 0&, LngDataType, ByVal 0&, LngDataBufSize
    If LngDataType = REG_SZ Then
        StrDataBuf = String(LngDataBufSize, " ")
        LngResult = RegQueryValueEx(LngHandle, SettingName, 0&, 0&, ByVal StrDataBuf, LngDataBufSize)
        If LngResult = ERROR_SUCCESS Then
            SettingGet = Left$(StrDataBuf, InStr(StrDataBuf, Chr$(0)) - 1)
        End If
    End If
    RegCloseKey LngHandle
    'SettingGet = GetString(HKEY_CLASSES_ROOT, "OsSecurity\Data", SettingName)
    'If SettingGet = "" Then SettingGet = Default
End Function


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Setting Save | Database
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub SetSaveDb(Setting As String, Value As Variant)
    Dim Db As Database, CRset As Recordset
    Set Db = OpenDatabase(App.Path & "\data\sdata.mdb", False, False, ";pwd=nsb2003")
    Set CRset = Db.OpenRecordset(":setting", dbOpenTable)
    
    With CRset
        .Index = "setting"
        .Seek "=", Setting
        If .NoMatch = True Then
            .AddNew
            !Setting = Setting
            !Value = Value
            .Update
        Else
            .Edit
            !Value = Value
            .Update
        End If
    End With
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Setting Get | Database
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Function SetGetDb(Setting As String, Optional Default As Variant = "") As Variant
    Dim Db As Database, CRset As Recordset
    Set Db = OpenDatabase(App.Path & "\data\sdata.mdb", False, False, ";pwd=nsb2003")
    Set CRset = Db.OpenRecordset(":setting", dbOpenSnapshot)
    
    SetGetDb = Default
    With CRset
        .FindFirst "setting = '" & Setting & "'"
        If .NoMatch = False Then SetGetDb = !Value
    End With
End Function
