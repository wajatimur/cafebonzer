Option Strict Off
Option Explicit On
Module mSecurity
	Enum eCbAccessAllow
		Allow_Setting = 1
		Allow_Statistic = 2
		Allow_Unlock = 3
	End Enum
	
	Enum eCbAccessTo
		Configuration = 1
		Statistic = 2
	End Enum
	
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Log Activity Pekerja
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	'UPGRADE_WARNING: ParamArray SeqParam was changed from ByRef to ByVal. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1003"'
	Public Sub LogWorker(ByRef Activity As String, ParamArray ByVal SeqParam() As Object)
		If UBound(SeqParam) > -1 Then Activity = LangParam(Activity, New Object(){SeqParam})
		Activity = LangPrcs(Activity)
		
		If CbLogUser = True Then
			uIDBe.DataSave("pekerja-log", "tarikh", Today, True, False)
			uIDBe.DataSave("pekerja-log", "masa", TimeOfDay, False, False)
			uIDBe.DataSave("pekerja-log", "nick", CbUserName, False, False)
			uIDBe.DataSave("pekerja-log", "akses", CbUserAccess, False, False)
			uIDBe.DataSave("pekerja-log", "perkara", Activity, False, True)
		End If
	End Sub
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Cek pasport
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	'kena buat encryption untuk ini.. kalau boleh untuk semua setting
	Public Function CekPass(ByRef NamePass As Object, ByRef NeedPass As Object) As Boolean
		Dim Cond2 As Object
		Dim d As Object
		Dim Cond1 As Object
		CekPass = False
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(mp). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object NeedPass. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object NamePass. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Cond1. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Cond1 = NamePass = SetAmbil("mu") And NeedPass = SetAmbil("mp")
		'UPGRADE_WARNING: Couldn't resolve default property of object Cond1. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If Cond1 Then CekPass = True : CbUserName = "Admin" : CbUserAccess = "111" : Exit Function
		
		'UPGRADE_WARNING: Couldn't resolve default property of object uSDBe.DataCount(pekerja-list). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		For d = 0 To uSDBe.DataCount("pekerja-list") - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object uSDBe.DataGet(pekerja-list, password, d). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Couldn't resolve default property of object NeedPass. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Couldn't resolve default property of object uSDBe.DataGet(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Couldn't resolve default property of object NamePass. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Cond2. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			Cond2 = NamePass = uSDBe.DataGet("pekerja-list", "nick", d) And NeedPass = uSDBe.DataGet("pekerja-list", "password", d)
			'UPGRADE_WARNING: Couldn't resolve default property of object Cond2. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			If Cond2 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object uSDBe.DataGet(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				CbUserName = uSDBe.DataGet("pekerja-list", "nick", d)
				'UPGRADE_WARNING: Couldn't resolve default property of object uSDBe.DataGet(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				CbUserAccess = uSDBe.DataGet("pekerja-list", "akses", d)
				CekPass = True
				'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(logaktiviti). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				If SetAmbil("logaktiviti") = True Then CbLogUser = True
				Exit Function
			End If
		Next d
	End Function
	
	
	Public Function CekAkses(ByRef AksesPart As eCbAccessAllow) As Boolean
		CekAkses = True
		If Mid(CbUserAccess, AksesPart, 1) = "0" Then
			MsgBox(MB(10), MsgBoxStyle.OKOnly, CbMsgWarn)
			CekAkses = False
		End If
	End Function
	
	Public Sub Accessing(ByRef WhichPart As eCbAccessTo)
		If CekAkses(WhichPart) = False Then
			MsgBox(MB(10), MsgBoxStyle.OKOnly, CbMsgWarn)
		Else
			Select Case WhichPart
				Case 1
					FrmSet.DefInstance.Show()
				Case 2
					FrmStat.DefInstance.Show()
			End Select
		End If
	End Sub
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Tulis log ke list1 dalam frmmain
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Sub MainLog(ByRef Log As String)
		FrmMain.DefInstance.MainLog.Items.Add(Log)
		'UPGRADE_ISSUE: ListBox property MainLog.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2059"'
		FrmMain.DefInstance.MainLog.SetSelected(FrmMain.DefInstance.MainLog.NewIndex, True)
		FrmMain.DefInstance.MainLog.Refresh()
	End Sub
	
	Public Sub DemoCheck()
		Dim lngDemoday As Integer
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(namadaftar). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If SetAmbil("namadaftar") = "" Then
			SetSaveDb("demodate", Tarikh)
			SetSaveDb("demoday", 1)
			CbDemoMode = True
			'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(namadaftar, demo). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ElseIf SetAmbil("namadaftar", "demo") = "demo" Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object SetGetDb(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			lngDemoday = SetGetDb("demoday", 0)
			' kira hari
			'UPGRADE_WARNING: Couldn't resolve default property of object SetGetDb(demodate, Tarikh). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			If SetGetDb("demodate", Tarikh) <> Tarikh Then SetSaveDb("demoday", (lngDemoday + 1))
			' simpan tarikh demo terakhir dibuka
			SetSaveDb("demodate", Tarikh)
			FrmSysDemo.DefInstance.ShowDialog()
			' set variable global cbDemo = True
			CbDemoMode = True
		Else
			SetSaveDb("demo", False)
			CbDemoMode = False
		End If
	End Sub
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Untuk mendapatkan jumlah hari yang telah digunakan
	' dalam demomode
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Function GetDayUse() As Object
		Dim buf As String
		FileOpen(1, VB6.GetPath & "\" & VB6.GetExeName(), OpenMode.Binary)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
		FileGet(1, buf, LOF(1) - 1)
		'UPGRADE_WARNING: Couldn't resolve default property of object GetDayUse. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		GetDayUse = buf
	End Function
End Module