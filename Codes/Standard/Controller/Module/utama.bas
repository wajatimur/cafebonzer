Attribute VB_Name = "MdlApplication"
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : MdlApplication
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
' Module Name         : Main
' Description         : Main Module
'==================================================================

'Public declaration
Public UniAgents As New ClsAgents

Public ModuleAgentInfo As Long
Public ModuleAgentManager As Long
' ModAgentInfo
'   0 = Inactive
'   1 = Resource
'   2 = Application\Process
'   3 = Printer Installed\Printing
'   4 = Network Traffic
'   5 = Hardware Enum\Os Info
'   6 = Drive Information
'
' ModAgentManager
'   0 = Inactive
'   1 = Active

Public Const CbDemoMaxDay = 14
Public Const EcKey2 = 6
Public Const CbDrvStr = "a"

Public StrCmdSep As String
Public StrCmdSubSep1 As String
Public StrCmdSubSep2 As String

Public CbAppVersion As String
Public CbAppBuild As String
Public CbAppLatestAgn As String

Public CurSDBPath As String
Public CurIDBPath As String

Public CRset As Recordset
Public CDataS As Database
Public CDataI As Database
Public CDataSe As New ClsData
Public CDataIe As New ClsData

Public CbUserName As String
Public CbUserAccess As Long
Public CbDemoMode As Boolean
Public CbLogUser As Boolean
Public CbViewMode As Long
Public CbMsgRcv As Boolean
Public CbConsole As Boolean

Public OpenSessionCur As String
Public OpenSessionDay As Long

Public Enum EnuModule
    [CafeReport] = 0
    [CafeSnmMgr] = 1
    [Help] = 100
End Enum

'//cbViewMode Constant//
' 0 = normal mode
' 1 = map mode

'//Form Sequence//
' Initalize -> Load -> Activate -> Paint

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Program Entry Point
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Sub Main()
 '[ Avoid multiple instance ]'
    If App.PrevInstance = True Then Exit Sub

    FrmAppSplash.Show           '[ Splash screen ]'
    DoEvents: Sleep 500

    Call LangLoad
    Call DemoCheck
    Call ConfigEnv
    Call ConfigForm
    
    Unload FrmAppSplash         '[ Terminate form FrmAppSplash ]'
    
    FrmAppPass.Show             '[ Authorization ]'
    'FrmMain.Show: CbUserName = "admin": CbUserAccess = 1023

    UniAgents.AgentRecoverUp    '[ Recover Offline Agent ]'
    Call NetUp
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' AppExit dari Program
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Function AppExit(Optional Ask As Boolean = True) As Integer
    Dim Frm As Form, LngRet As Long
    
    If CbDemoMode = True Then FrmSysDemo.Show vbModal
    Call MetricFrmSave(FrmMain)
    
    If Ask = True And FrmMain.ListView.ListItems.Count = 0 Then LngRet = MsgBox(ST(1, 0), vbOKCancel, CbMsgApp)
    If Ask = True And FrmMain.ListView.ListItems.Count <> 0 Then LngRet = MsgBox(ST(1, 2), vbOKCancel, CbMsgApp)
    
    'User decision
    If LngRet = vbCancel Then
        AppExit = 1
        Exit Function
    Else:
        Call MenuIconDetach
        Call NetClose
        For Each Frm In Forms
            'If Frm.Name <> "FrmMain" Then Unload Frm
            Unload Frm
        Next Frm
        AppExit = 0
    End If
End Function

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Load Module
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub AppLoadInfo(Module As EnuModule)
    Dim LngRet As Long

    Select Case Module
    Case 100
        LngRet = ShellExecute(FrmMain.Hwnd, "open", App.Path & "\help.chm", vbNullString, vbNullString, SW_NORMAL)
    End Select

    If LngRet <= 32 Then MsgBox ST(1, 3), vbCritical, CbMsgWarn
End Sub


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Date | US Format
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function DateGetUS() As Date
    hari = Day(Date)
    Bulan = Month(Date)
    Tahun = Year(Date)
    DateGetUS = Bulan & "/" & hari & "/" & Tahun
End Function

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Date | System
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Function DateGetSystem(tDay As Integer, tMonth As Integer, tYear As Integer) As String
    Dim ST As SYSTEMTIME
    Dim tBuffer As String

    ST.wDay = tDay
    ST.wMonth = tMonth
    ST.wYear = tYear
    tBuffer = String(255, 0)
        
    GetDateFormat ByVal 0&, 0, ST, vbNullString, tBuffer, Len(tBuffer)
    DateGetSystem = Left(tBuffer, InStr(1, tBuffer, Chr$(0)) - 1)
End Function


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Simpan error
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub AppErrorLog(ErrType As ErrObject, ProcName As String, Optional DisplayMsg As Boolean = True)
    Dim StrError As String
    Dim IntErrNum As Integer, StrErrDesc As String, StrErrSource As String
        
    IntErrNum = ErrType.Number
    StrErrSource = ErrType.Source
    StrErrDesc = ErrType.Description
    
    If DisplayMsg = True Then
        MsgBox IntErrNum & " / " & StrErrSource & vbNewLine & StrErrDesc, vbExclamation, ProcName
    End If
    
    StrError = Now & " - " & StrErrSource & " - " & StrErrDesc & " - " & IntErrNum
    
    Open "ErrLog.txt" For Append As #1
    Write #1, StrError
    Close #1
End Sub


Public Function RsFilter(RsTmp As Recordset, FilterStr As String) As Recordset
    If FilterStr = "" Then Exit Function
    RsTmp.Filter = FilterStr
    Set RsFilter = RsTmp.OpenRecordset
End Function


Public Function FileExisted(Filename As String) As Boolean
    If Dir(Filename) <> "" Then FileExisted = True
End Function
