Option Strict Off
Option Explicit On
Module mSetting
	'==================================================================
	' Aplication codename : CafeBonzer
	' Programmer          : Azri Jamil a.k.a wajatimur
	' Module Name         : Setting
	' Description         :
	'==================================================================
	Public EnumArray() As String
	
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Cek jika program buka untuk pertama kali
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Sub PreSetting()
		
		'global variables
		CbAppVersion = "CafeBonzer v" & System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).FileMajorPart & "." & System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).FileMinorPart & " Beta"
		'UPGRADE_ISSUE: App property App.Revision was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2069"'
		CbAppBuild = System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).FileMinorPart & "." & App.Revision
		CbAppLatestAgn = "1.7.55"
		
		'data path
		CbPathDatRecv = VB6.GetPath & "\data\recostate.dat"
		
		'path untuk database utama semasa
		CurSDBPath = VB6.GetPath & "\data\sdata.mdb"
		CurIDBPath = VB6.GetPath & "\data\idata.mdb"
		
		'initialize universal akses database.. untuk kegunaan umum
		uSDBe.InitDb = CurSDBPath
		uIDBe.InitDb = CurIDBPath
		
		uSDB = DAODBEngine_definst.OpenDatabase(CurSDBPath, False, False, ";pwd=nsb2003")
		uIDB = DAODBEngine_definst.OpenDatabase(CurIDBPath, False, False, ";pwd=nsb2003")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(pertamakali). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If SetAmbil("pertamakali") = "" Or SetAmbil("pertamakali") = "ya" Then
			FrmSplash.DefInstance.Close()
			FrmSet.DefInstance.ShowDialog()
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
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		autoCloseSession = SetAmbil("autocloses")
		If Trim(OpenSessionCur) = "" Then
			OpenSessionCur = CStr(Today)
			uSDBe.DbSaveSetting("opensession", OpenSessionCur)
		End If
		If (OpenSessionCur & " " & TimeOfDay) < (Today & " " & autoCloseSession) Then
			uSDBe.DbSaveSetting("lastsession", OpenSessionCur)
			uSDBe.DbSaveSetting("opensession", Today)
			OpenSessionCur = CStr(Today)
		End If
		
		'log user activity
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		CbLogUser = SetAmbil("logaktiviti", True)
		'loading POS category
		Call LoadPosCatCB((FrmMain.DefInstance.SerImgCb1), (FrmMain.DefInstance.ImgListSnm))
		Exit Sub
		
ErrInt: 
		ErrLog(Err, "SettingUp")
	End Sub
	
	Public Sub SettingFrm()
		On Error GoTo ErrInt
		Dim s_tTab As Boolean
		Dim s_rBar As Boolean
		
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		FrmMain.DefInstance.Text = "CafeBonzer v" & System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).FileMajorPart & "." & CbAppBuild & " - " & SetAmbil("tajukatas")
		If CbDemoMode = True Then FrmMain.DefInstance.Text = FrmMain.DefInstance.Text & " UNREGISTERED"
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		FrmMain.DefInstance.MainNote.Text = SetAmbil("mainnote")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		s_tTab = SetAmbil("tooltab", True)
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		s_rBar = SetAmbil("dockbar", True)
		If s_tTab = False Then
			FrmMain.DefInstance.menu4envsub(0).Checked = s_tTab
			FrmMain.DefInstance.MainPhold.PageCollapse = False
		End If
		If s_rBar = False Then
			FrmMain.DefInstance.menu4envsub(1).Checked = s_rBar
			FrmMain.DefInstance.MainPdock.PageFlip = True
		End If
		Exit Sub
		
ErrInt: 
		ErrLog(Err, "SettingFrm")
	End Sub
	
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' simpan setting dalam registry
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Sub SetSimpan(ByRef Namasetting As String, ByRef Nilai As String)
		'SaveString HKEY_CLASSES_ROOT, "externalthread\shell", Crypt(Namasetting, EcKey1), Crypt(Nilai, EcKey2)
		SaveString(HKEY_CLASSES_ROOT, "externalthread\shell", Namasetting, Nilai)
	End Sub
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' ambil setting dari registry
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	'UPGRADE_NOTE: Default was upgraded to Default_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
	Function SetAmbil(ByRef Namasetting As String, Optional ByRef Default_Renamed As Object = "") As Object
		'SetAmbil = Crypt(GetString(HKEY_CLASSES_ROOT, "externalthread\shell", Crypt(Namasetting, EcKey1)), EcKey2)
		'UPGRADE_WARNING: Couldn't resolve default property of object GetString(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		SetAmbil = GetString(HKEY_CLASSES_ROOT, "externalthread\shell", Namasetting)
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Default_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If SetAmbil = "" Then SetAmbil = Default_Renamed
	End Function
	
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Setting Save | Database
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Sub SetSaveDb(ByRef Setting As Object, ByRef Value As Object)
		Dim Db As DAO.Database
		Dim Rs As DAO.Recordset
		Db = DAODBEngine_definst.OpenDatabase(VB6.GetPath & "\data\sdata.mdb", False, False, ";pwd=nsb2003")
		Rs = Db.OpenRecordset(":setting", DAO.RecordsetTypeEnum.dbOpenTable)
		
		With Rs
			.Index = "setting"
			.Seek("=", Setting)
			If .NoMatch = True Then
				.AddNew()
				'UPGRADE_WARNING: Couldn't resolve default property of object Setting. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				.Fields("Setting").Value = Setting
				'UPGRADE_WARNING: Couldn't resolve default property of object Value. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				.Fields("Value").Value = Value
				.Update()
			Else
				.Edit()
				'UPGRADE_WARNING: Couldn't resolve default property of object Value. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				.Fields("Value").Value = Value
				.Update()
			End If
		End With
	End Sub
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Setting Get | Database
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	'UPGRADE_NOTE: Default was upgraded to Default_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
	Function SetGetDb(ByRef Setting As Object, Optional ByRef Default_Renamed As Object = "") As Object
		Dim Db As DAO.Database
		Dim Rs As DAO.Recordset
		Db = DAODBEngine_definst.OpenDatabase(VB6.GetPath & "\data\sdata.mdb", False, False, ";pwd=nsb2003")
		Rs = Db.OpenRecordset(":setting", DAO.RecordsetTypeEnum.dbOpenSnapshot)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Default_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object SetGetDb. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		SetGetDb = Default_Renamed
		With Rs
			'UPGRADE_WARNING: Couldn't resolve default property of object Setting. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			.FindFirst("setting = '" & Setting & "'")
			'UPGRADE_WARNING: Couldn't resolve default property of object SetGetDb. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			If .NoMatch = False Then SetGetDb = .Fields("Value").Value
		End With
	End Function
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Simpan string ke File *.ini
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Sub INIsimpan(ByRef NamaFail As Object, ByRef Bahagian As Object, ByRef Kunci As Object, ByRef Nilai As Object, Optional ByRef NoEncrypt As Boolean = False)
		NoEncrypt = True
		'UPGRADE_WARNING: Couldn't resolve default property of object Bahagian. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Bahagian = CStr(Bahagian)
		'UPGRADE_WARNING: Couldn't resolve default property of object Kunci. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Kunci = CStr(Kunci)
		'UPGRADE_WARNING: Couldn't resolve default property of object Nilai. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Crypt(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If NoEncrypt = False Then Nilai = Crypt(CStr(Nilai), EcKey2)
		'UPGRADE_WARNING: Couldn't resolve default property of object NamaFail. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		NamaFail = CStr(NamaFail)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object NamaFail. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Nilai. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Kunci. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Bahagian. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		WritePrivateProfileString(Bahagian, CStr(Kunci), CStr(Nilai), NamaFail)
	End Sub
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Ambil string dari file *.ini
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Function INIambil(ByRef NamaFail As Object, ByRef Bahagian As Object, ByRef Kunci As Object, Optional ByRef NoEncrypt As Boolean = False) As String
		Dim ostr As Object
		Dim retval As New VB6.FixedLengthString(255)
		NoEncrypt = True
		'UPGRADE_WARNING: Couldn't resolve default property of object Bahagian. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Bahagian = CStr(Bahagian)
		'UPGRADE_WARNING: Couldn't resolve default property of object Kunci. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Kunci = CStr(Kunci)
		'UPGRADE_WARNING: Couldn't resolve default property of object NamaFail. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		NamaFail = CStr(NamaFail)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object NamaFail. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Kunci. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Bahagian. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		GetPrivateProfileString(Bahagian, CStr(Kunci), "", retval.Value, Len(retval.Value), NamaFail)
		'UPGRADE_WARNING: Couldn't resolve default property of object ostr. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ostr = retval.Value & Chr(0)
		'UPGRADE_WARNING: Couldn't resolve default property of object ostr. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		INIambil = Left(ostr, InStr(1, ostr, Chr(0)) - 1)
		'UPGRADE_WARNING: Couldn't resolve default property of object Crypt(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If NoEncrypt = False Then INIambil = Crypt(INIambil, EcKey2)
	End Function
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Enumerate Section dan Key dalam *.ini file
	'  jika enumkey dinyatakan.. enumerator akan
	'  mengambil nilai bagi setiap key.. jika tidak
	'  ia akan enumerate bahagian dan bukannyer kunci
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Function INIenumSection(ByRef NamaFail As Object, Optional ByRef EnumKey As String = "") As String
		Dim Buff As String
		Dim idx As Short
		ReDim EnumArray(10)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object NamaFail. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
		If Len(Dir(NamaFail)) = 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object NamaFail. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			FileOpen(1, NamaFail, OpenMode.Output) : FileClose(1)
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object NamaFail. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		FileOpen(1, NamaFail, OpenMode.Input)
		
		If EnumKey = "" Then
			Do Until EOF(1)
				Buff = LineInput(1)
				If Left(Buff, 1) = "[" Then
					If idx = UBound(EnumArray) Then ReDim Preserve EnumArray(UBound(EnumArray) + 10)
					EnumArray(idx) = StrReverse(Mid(Buff, 2))
					EnumArray(idx) = StrReverse(Mid(EnumArray(idx), 2))
					idx = idx + 1
				End If
			Loop 
		Else
			Do Until EOF(1)
				Buff = LineInput(1)
				If LCase(Left(Buff, Len(EnumKey))) = LCase(EnumKey) Then
					If idx = UBound(EnumArray) Then ReDim Preserve EnumArray(UBound(EnumArray) + 10)
					'EnumArray(idx) = Crypt(Mid(Buff, InStr(1, Buff, "=") + 1), EcKey2)
					EnumArray(idx) = Mid(Buff, InStr(1, Buff, "=") + 1)
					idx = idx + 1
				End If
			Loop 
		End If
		
		FileClose(1)
		ReDim Preserve EnumArray(idx)
	End Function
	
	
	
	Function CryptX(ByRef Text As String, ByRef Codekey As Short) As String
		Dim Tmp, Tmp2 As String
		Dim Itg As Short
		Tmp2 = "*nematix*" & Text & "*seal*"
		For Itg = 1 To Len(Tmp2)
			Tmp = Tmp & Chr(Asc(Mid(Tmp2, Itg, 1)) Xor Codekey)
		Next Itg
		CryptX = Tmp
	End Function
	
	Function DecryptX(ByRef Text As String, ByRef Codekey As Short) As String
		Dim a002 As Object
		Dim a001 As Object
		Dim Tmp As String
		Dim Itg As Short
		For Itg = 1 To Len(Text)
			Tmp = Tmp & Chr(Asc(Mid(Text, Itg, 1)) Xor Codekey)
		Next Itg
		'UPGRADE_WARNING: Couldn't resolve default property of object a001. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		a001 = Right(Tmp, Len(Tmp) - 9)
		'UPGRADE_WARNING: Couldn't resolve default property of object a001. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object a002. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		a002 = Left(a001, Len(a001) - 6)
		'UPGRADE_WARNING: Couldn't resolve default property of object a002. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		DecryptX = a002
	End Function
	
	Function Crypt(ByRef Text As String, ByRef Codekey As Short) As Object
		Dim Tmp As String
		Dim Itg As Short
		For Itg = 1 To Len(Text)
			Tmp = Tmp & Chr(Asc(Mid(Text, Itg, 1)) Xor Codekey)
		Next Itg
		'UPGRADE_WARNING: Couldn't resolve default property of object Crypt. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Crypt = Tmp
	End Function
End Module