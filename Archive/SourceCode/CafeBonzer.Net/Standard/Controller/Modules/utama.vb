Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Module mApplication
	'==================================================================
	' Aplication codename : CafeBonzer
	' Programmer          : Azri Jamil a.k.a wajatimur
	' Module Name         : Main
	' Description         : Main Module
	'==================================================================
	
	'Public declaration
	Public UniAgents As New clsAgents
	
	Public ModAgentInfo As Integer
	Public ModAgentManager As Integer
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
	
	
	Public Const EcKey1 As Short = 8
	Public Const EcKey2 As Short = 6
	
	Public CbAppVersion As String
	Public CbAppBuild As String
	Public CbAppLatestAgn As String
	
	Public CbPathDatRecv As String
	Public CurSDBPath As String
	Public CurIDBPath As String
	
	Public Rs As DAO.Recordset
	Public uSDB As DAO.Database
	Public uIDB As DAO.Database
	Public uSDBe As New clsData
	Public uIDBe As New clsData
	
	Public CbUserName As String
	Public CbUserAccess As String
	Public CbDemoMode As Boolean
	Public CbDrvStr As String
	Public CbLogUser As Boolean
	Public CbViewMode As Integer
	Public CbMsgRcv As Boolean
	Public CbConsole As Boolean
	
	Public OpenSessionCur As String
	Public OpenSessionLast As String
	
	Public Enum EnuModule
		CafeReport = 0
		CafeSnmMgr = 1
	End Enum
	
	'//cbViewMode Constant//
	' 0 = normal mode
	' 1 = map mode
	
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Program Entry Point
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	'UPGRADE_WARNING: Application will terminate when Sub Main() finishes. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1047"'
	Public Sub Main()
		'[ Avoid multiple instance ]'
		If (UBound(Diagnostics.Process.GetProcessesByName(Diagnostics.Process.GetCurrentProcess.ProcessName)) > 0) = True Then Exit Sub
		
		'[ Splash screen ]'
		FrmSplash.DefInstance.Show()
		System.Windows.Forms.Application.DoEvents() : Sleep(500)
		
		'[ Load language ]'
		Call LangLoad()
		'[ Demo check ]'
		Call DemoCheck()
		
		'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1039"'
		Load(FrmMain)
		
		'[ Startin rutine ]'
		Call PreSetting()
		'[ Load semua setting ]'
		Call SettingUp()
		'[ Setup mainform ]'
		Call SettingFrm()
		
		'[ Terminate form frmSplash ]'
		FrmSplash.DefInstance.Close()
		
		'[ Authorization ]'
		FrmPass.DefInstance.Show()
		'FrmMain.Show: CbUserName = "admin": CbUserAccess = "111"
		
		'[ Recover Offline Agent ]'
		UniAgents.AgentRecoverUp()
		
		'[ Hidupkan network ]'
		Call NetUp()
	End Sub
	
	Public Function RsFilter(ByRef RsTmp As DAO.Recordset, ByRef FilterStr As String) As Object
		If FilterStr = "" Then Exit Function
		RsTmp.Filter = FilterStr
		RsFilter = RsTmp.OpenRecordset
	End Function
	
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Keluar dari Program
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Sub Keluar(Optional ByRef Ask As Boolean = True)
		Dim oSpc As Object
		Dim msgret As Object
		Dim Frm As System.Windows.Forms.Form
		
		'save form position and size
		Call CbFrmMetricSave(FrmMain.DefInstance)
		
		'check routine
		'UPGRADE_WARNING: Couldn't resolve default property of object msgret. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If Ask = True And FrmMain.DefInstance.Lv1.ListItems.Count = 0 Then msgret = MsgBox(MB(11), MsgBoxStyle.OKCancel, CbMsgApp)
		'UPGRADE_WARNING: Couldn't resolve default property of object msgret. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If Ask = True And FrmMain.DefInstance.Lv1.ListItems.Count <> 0 Then msgret = MsgBox(MB(12), MsgBoxStyle.OKCancel, CbMsgApp)
		
		'user decision
		If msgret = MsgBoxResult.Cancel Then
			Exit Sub
		Else
			'UPGRADE_NOTE: Object oSpc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
			oSpc = Nothing
			Call NetClose() : FrmMain.DefInstance.Hide()
			'UPGRADE_ISSUE: Forms collection was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2068"'
			For	Each Frm In Forms
				Frm.Close()
			Next Frm
		End If
	End Sub
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Tarikh dalam format US
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Function Tarikh() As Date
		Dim Tahun As Object
		Dim Bulan As Object
		Dim Hari As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object Hari. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Hari = VB.Day(Today)
		'UPGRADE_WARNING: Couldn't resolve default property of object Bulan. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Bulan = Month(Today)
		'UPGRADE_WARNING: Couldn't resolve default property of object Tahun. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Tahun = Year(Today)
		'UPGRADE_WARNING: Couldn't resolve default property of object Tahun. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Hari. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Bulan. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Tarikh = CDate(Bulan & "/" & Hari & "/" & Tahun)
	End Function
	
	
	Public Function GetSystemDate(ByRef tDay As Short, ByRef tMonth As Short, ByRef tYear As Short) As String
		Dim St As SYSTEMTIME
		Dim tBuffer As String
		
		St.wDay = tDay
		St.wMonth = tMonth
		St.wYear = tYear
		tBuffer = New String(Chr(0), 255)
		
		GetDateFormat(0, 0, St, vbNullString, tBuffer, Len(tBuffer))
		GetSystemDate = Left(tBuffer, InStr(1, tBuffer, Chr(0)) - 1)
	End Function
	
	
	'UPGRADE_NOTE: Module was upgraded to Module_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
	Public Sub LoadModule(ByRef Module_Renamed As EnuModule)
		Dim lngRet As Integer
		
		Select Case Module_Renamed
			Case 0
				lngRet = ShellExecute(FrmMain.DefInstance.Handle.ToInt32, "open", VB6.GetPath & "\CafeReport.exe", "pc-usage", vbNullString, SW_NORMAL)
			Case 1
				lngRet = ShellExecute(FrmMain.DefInstance.Handle.ToInt32, "open", VB6.GetPath & "\CafeSmMgr.exe", vbNullString, vbNullString, SW_NORMAL)
		End Select
		
		If lngRet <= 32 Then MsgBox(MB(20), MsgBoxStyle.Critical, CbMsgWarn)
	End Sub
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Cek kewujudan File
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Function FileExist(ByVal PathName As String) As Boolean
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
		FileExist = IIf(Dir(PathName) = "", False, True)
	End Function
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Simpan error
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Sub ErrLog(ByRef errType As ErrObject, ByRef procName As String, Optional ByRef DisplayMsg As Boolean = True)
		Dim ErrDesc As String
		Dim i_errNum As Short
		Dim s_errDesc, s_errSource As String
		
		i_errNum = Err.Number
		s_errSource = errType.Source
		s_errDesc = Err.Description
		
		MsgBox(i_errNum & " / " & s_errSource & vbNewLine & s_errDesc, MsgBoxStyle.Exclamation, procName)
		ErrDesc = Now & " - " & s_errDesc & " - " & s_errSource & " - " & i_errNum
		
		FileOpen(1, "ErrLog.txt", OpenMode.Append)
		WriteLine(1, ErrDesc)
		FileClose(1)
	End Sub
End Module