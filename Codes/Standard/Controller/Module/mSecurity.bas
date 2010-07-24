Attribute VB_Name = "MdlSecurity"
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : MdlSecurity
'    Project    : CafeBonzer
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
' Access Policy Code Redesign
'
'   Load External Module    = 1
'   Configuration           = 2
'   Statistic               = 4
'   Console                 = 8
'   Agent Manager           = 16
'   Security Log            = 32

'   Unlock Client           = 64
'   Cancel Client           = 128
'   Change Price            = 256

Public Const LngTotalAccessCode = 9 ' = 512 (0 to 512)
Enum EnAccessCode
    ModExternal = 1
    ModConfiguration = 2
    ModStatistic = 4
    ModConsole = 8
    ModAgentManager = 16
    ModSecurityLog = 32
    TaskUnlock = 64
    TaskCancel = 128
    TaskChangePrice = 256
End Enum


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Check Passport
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Function SecPasswordCheck(NamePass, NeedPass) As Boolean
    SecPasswordCheck = False
    Cond1 = NamePass = SetGetDb("GenAdminName", "admin") And NeedPass = SetGetDb("GenAdminPass")
    If Cond1 Then SecPasswordCheck = True: CbUserName = "Admin": CbUserAccess = 1023: Exit Function
    
    For d = 0 To CDataSe.DataCount("ListEmployee") - 1
        Cond2 = NamePass = CDataSe.DataGet("ListEmployee", "Username", d) And NeedPass = CDataSe.DataGet("ListEmployee", "Password", d)
        If Cond2 Then
            CbUserName = CDataSe.DataGet("ListEmployee", "Username", d)
            CbUserAccess = CDataSe.DataGet("ListEmployee", "Access", d)
            SecPasswordCheck = True
            If SetGetDb("SecLogUser", True) = True Then CbLogUser = True
            Exit Function
        End If
    Next d
End Function
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Request a Security Policy Access on User
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Function SecAccessRequest(AccessModules As EnAccessCode, Optional UserAccess As Long) As Boolean
    Dim LngIdxA As Long, StrTmpAccCode As String, LngTmpUserAccess As Long
    
    SecAccessRequest = False
    If UserAccess = 0 Then
        LngTmpUserAccess = CbUserAccess
    Else
        LngTmpUserAccess = UserAccess
    End If
    For LngIdxA = LngTotalAccessCode To 0 Step -1
        StrTmpAccCode = 2 ^ LngIdxA
        If LngTmpUserAccess >= StrTmpAccCode Then
            If AccessModules = StrTmpAccCode Then
                SecAccessRequest = True
                Exit Function
            End If
            LngTmpUserAccess = LngTmpUserAccess - StrTmpAccCode
        End If
    Next
End Function
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Security Access Check
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub SecOpenModules(AccessModules As EnAccessCode, Optional Extended As String, Optional Extended2 As String)
    If SecAccessRequest(AccessModules) = False Then
        MsgBox ST(1, 1), vbOKOnly, CbMsgWarn
    Else
        Select Case AccessModules
            Case 1
                LngRet = ShellExecute(FrmMain.Hwnd, "open", Extended, Extended2, vbNullString, SW_NORMAL)
                If LngRet <= 32 Then MsgBox ST(1, 3), vbCritical, CbMsgWarn
            Case 2
                FrmSysSet.Show
                SecUserLog SL(5)
            Case 4
                FrmStat.Show
            Case 8
                'FrmSysConsole.Show
            Case 16
                FrmAgnMgr.Show
        End Select
    End If
End Sub
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Log User Activity
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub SecUserLog(Activity As String, ParamArray SeqParam())
    'If UBound(SeqParam) > -1 Then Activity = Parameter(Activity, SeqParam)
    Activity = LangPrcs(Activity)
    
    If CbLogUser = True Then
        CDataIe.DataSave "SecurityEmployeeLog", "Date", Date, True, False
        CDataIe.DataSave "SecurityEmployeeLog", "Time", Time, False, False
        CDataIe.DataSave "SecurityEmployeeLog", "UserName", CbUserName, False, False
        CDataIe.DataSave "SecurityEmployeeLog", "Access", CbUserAccess, False, False
        CDataIe.DataSave "SecurityEmployeeLog", "SecurityLog", Activity, False, True
    End If
End Sub
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Log To Main Form
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub SecAppMainLog(Log As String)
    'FrmMain.MainLog.AddItem Log
    'FrmMain.MainLog.Selected(FrmMain.MainLog.NewIndex) = True
    'FrmMain.MainLog.Refresh
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Demo Mode Check
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub DemoCheck()
    Dim LngDemoDay As Long, DteDemoDate As Date
    Dim StrRegName As String
    
    StrRegName = SettingGet("RegName", "Demo")
    
    If StrRegName = "" Then
        SetSaveDb "AppDemoDate", Date
        SetSaveDb "AppDemoDay", 1
        CbDemoMode = True
    ElseIf StrRegName = "Demo" Then
        CbDemoMode = True
        LngDemoDay = SetGetDb("AppDemoDay", CbDemoMaxDay)
        DteDemoDate = SetGetDb("AppDemoDate", Date)
        If DteDemoDate <> Date Then SetSaveDb "AppDemoDay", (LngDemoDay + 1)
        SetSaveDb "AppDemoDate", Date
        FrmSysDemo.Show vbModal
    Else
        SetSaveDb "Demo", False
        CbDemoMode = False
    End If
End Sub


Public Function SecLiscenseActivate() As Boolean
    Dim StrRegName As String, StrRegKey As String

    If DiskKeyValidate(CbDrvStr) = True Then
    '{ Disk is valid. Activation proceed }'
        StrRegName = DiskKeyName(CbDrvStr)
        StrRegKey = DiskKeyNum(CbDrvStr)
        SettingSave "RegName", StrRegName
        SettingSave "RegNumber", StrRegKey
        CbDemoMode = False
        DeleteFile CbDrvStr & ":\Boot"
        MsgBox ST(3, 1), vbOKOnly, CbMsgWarn
    Else
    '{ Unable to activate. Exit }'
        MsgBox ST(3, 3), vbInformation, CbMsgWarn
        End
    End If
End Function


Public Function SecLiscenseTransfer() As Boolean
    Dim LngRet As Long, StrRegName As String, StrRegKey As String
    
    LngRet = MsgBox(ST(3, 2), vbOKCancel + vbInformation, CbMsgWarn)
    
    If LngRet = vbOK Then
        StrRegName = SettingGet("RegName")
        StrRegKey = SettingGet("RegNumber")
        If DiskKeyCreate(StrRegName, StrRegKey, CbDrvStr) = True Then
            SetSaveDb "AppDemoDay", 1
            SettingSave "RegName", "Demo"
            SettingSave "RegNumber", "Demo"
            MsgBox "Liscense transfer complete. CafeBonzer will be close !", vbOKOnly + vbInformation, CbMsgWarn
            AppExit False
        End If
    End If
End Function
