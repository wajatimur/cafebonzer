Attribute VB_Name = "mSetting"
'==================================================================
' Aplication codename : CafeBonzer
' Programmer          : Azri Jamil a.k.a wajatimur
' Module Name         : Setting
' Description         :
'==================================================================

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Cek jika program buka untuk pertama kali
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub PreSetting()

    'global variables
    CbAppVersion = "CafeBonzer v" & App.Major & "." & App.Minor
    CbAppBuild = App.Minor & "." & App.Revision
    CbAppLatestAgn = "1.7.55"
    
    'data path
    CbPathDatRecv = App.Path & "\data\recostate.dat"
    
    'path untuk database utama semasa
    CurSDBPath = App.Path & "\data\sdata.mdb"
    CurIDBPath = App.Path & "\data\idata.mdb"
    
    'initialize universal akses database.. untuk kegunaan umum
    uSDBe.InitDb = CurSDBPath
    uIDBe.InitDb = CurIDBPath
    
    Set uSDB = OpenDatabase(CurSDBPath)
    Set uIDB = OpenDatabase(CurIDBPath)
    
    If SetAmbil("pertamakali") = "" Or SetAmbil("pertamakali") = "ya" Then
        Unload FrmSplash
        FrmSet.Show vbModal
    End If
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' "Load" kan semua setting
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub SettingUp()
On Error GoTo ErrInt
    Dim autoCloseSession As String
    'General Setting - 18/08/2002 my 2nd year annivessary with my love one
    
    'Loading and checking current session
    OpenSessionCur = uSDBe.DbGetSetting("opensession")
    OpenSessionLast = uSDBe.DbGetSetting("lastsession")
    autoCloseSession = SetAmbil("autocloses")
    If Trim(OpenSessionCur) = "" Then
        OpenSessionCur = Date
        uSDBe.DbSaveSetting "opensession", OpenSessionCur
    End If
    If (OpenSessionCur & " " & Time) < (Date & " " & autoCloseSession) Then
        uSDBe.DbSaveSetting "lastsession", OpenSessionCur
        uSDBe.DbSaveSetting "opensession", Date
        OpenSessionCur = Date
    End If
    
    'log user activity
    CbLogUser = SetAmbil("logaktiviti", True)
    'loading POS category
    Call LoadPosCatCB(FrmMain.SerImgCb1, FrmMain.Iml)
Exit Sub

ErrInt:
    ErrLog Err, "SettingUp"
End Sub

Public Sub SettingFrm()
On Error GoTo ErrInt
    Dim s_tTab As Boolean
    Dim s_rBar As Boolean
    
    FrmMaster.Caption = "CafeBonzer v" & App.Major & "." & CbAppBuild & " - " & SetAmbil("tajukatas")
    If CbDemoMode = True Then FrmMain.Caption = FrmMain.Caption & " UNREGISTERED"
    FrmMain.MainNote = SetAmbil("mainnote")
    
    s_tTab = SetAmbil("tooltab", True)
    s_rBar = SetAmbil("dockbar", True)
    If s_tTab = False Then
        'FrmMain.Menu4EnvSub(0).Checked = s_tTab
        'FrmMain.MainPhold.PageCollapse = False
    End If
    If s_rBar = False Then
        'FrmMain.Menu4EnvSub(1).Checked = s_rBar
        'FrmMain.MainPdock.PageFlip = True
    End If
Exit Sub

ErrInt:
    ErrLog Err, "SettingFrm"
End Sub


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' simpan setting dalam registry
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Sub SetSimpan(Namasetting As String, Nilai As String)
    SaveString HKEY_CLASSES_ROOT, "externalthread\shell", Crypt(Namasetting, EcKey1), Crypt(Nilai, EcKey2)
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ambil setting dari registry
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function SetAmbil(Namasetting As String, Optional Default As Variant = "") As Variant
    SetAmbil = Crypt(GetString(HKEY_CLASSES_ROOT, "externalthread\shell", Crypt(Namasetting, EcKey1)), EcKey2)
    If SetAmbil = "" Then SetAmbil = Default
End Function



