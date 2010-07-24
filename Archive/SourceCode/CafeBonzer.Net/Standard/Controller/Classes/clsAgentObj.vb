Option Strict Off
Option Explicit On
Friend Class clsAgent
	'==================================================================
	'  Aplication codename : CafeBonzer
	'  Programmer          : Azri Jamil a.k.a wajatimur
	'  Class Name          : Agent Object
	'  Description         : Class for total control of agent
	'==================================================================
	'Private ObjRoot As clsAgents
	'Private ObjInfo As New clsAgInfo
	Private ObjSck As AxSocketWrenchCtrl.AxSocket
	
	Private b_netMode As Boolean
	Private i_sockIndex As Integer
	Private i_NetPingCount As Short
	
	Private b_agCert As Boolean
	Private s_agVersion As String
	Private s_agName As String
	Private s_agMAC As String
	Private s_agIP As String
	Private s_agStatus As String
	Private s_agConnected As String
	
	Private s_CusName As String
	Private s_CusType As String
	Private s_CusFlag As String
	Private s_CusTimeIn As String
	Private s_CusTimeOut As String
	Private s_CusUsage As String
	Private l_CusPrinted As Integer
	
	Private ClistView As AxMSComctlLib.AxListView
	Private ClvItem As MSComctlLib.ListItem
	
	
	
	'=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*==*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=
	'  MAIN PROCEDURE
	'
	'
	'=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*==*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
	Private Sub Class_Initialize_Renamed()
		ClistView = FrmMain.DefInstance.Lv1
		b_agCert = False
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	
	'=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*==*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=
	'  PUBLIC SUBS & FUNCTIONS
	'
	'
	'=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*==*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	' Send Data
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	Public Sub NetSend(ByRef Data As String)
		Send(i_sockIndex, Data)
	End Sub
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	' Ping Client
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	Public Function NetPing() As Short
		Send(i_sockIndex, "//hey")
		i_NetPingCount = i_NetPingCount + 1
		NetPing = i_NetPingCount
	End Function
	Public Sub NetPingReset()
		i_NetPingCount = 0
	End Sub
	
	
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	'
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	Public Function ItemMain() As MSComctlLib.ListItem
		ItemMain = ClvItem
	End Function
	
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	' Initialize Agent
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	Public Sub AgnInit(ByRef Online As Boolean, Optional ByRef NewSocket As AxSocketWrenchCtrl.AxSocket = Nothing, Optional ByRef AgentName As String = "")
		If Online = True Then
			ObjSck = NewSocket
			b_netMode = True
			'UPGRADE_ISSUE: VBControlExtender property ObjSck.Index was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			i_sockIndex = ObjSck.Index
			s_agIP = ObjSck.PeerAddress
		Else
			s_agName = AgentName
			Call zAgentAdd(s_agName)
		End If
	End Sub
	
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	' Reject Agent
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	Public Sub AgnInitReject()
		ObjSck.Cleanup()
		'UPGRADE_NOTE: Object clsAgent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
		clsAgent = Nothing
	End Sub
	
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	' Initialize Certification
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	Public Function AgnInitCert(ByRef sCommand As String, Optional ByRef OfflineSignal As Boolean = False) As Boolean ', Root As clsAgents) As Boolean
		'UPGRADE_WARNING: Couldn't resolve default property of object SubVal(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		s_agVersion = SubVal(sCommand, "version")
		
		
		If s_agVersion >= "1.7.75" Then
			Send(i_sockIndex, "//cert:" & SubBuild("granted", CStr(1)))
			b_agCert = True
			AgnInitCert = True
			If OfflineSignal = False Then
				'UPGRADE_WARNING: Couldn't resolve default property of object SubVal(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				s_agName = SubVal(sCommand, "name")
				s_agName = zGetFreeName(s_agName)
				Call zAgentAdd(s_agName)
			Else
				If Me.AgentStatus = VS(4) Then
					Call zAgentSetSymbol(mAgents.EnAgentSymbol.Terminal_Online)
				Else
					Call zAgentSetSymbol(mAgents.EnAgentSymbol.User_Online)
				End If
			End If
		Else
			Send(i_sockIndex, "//cert:" & SubBuild("granted", CStr(0)) & SubBuild("info", "Version too old, please update !"))
			AgnInitCert = False
		End If
	End Function
	
	Public Function AgnOffline() As Object
		b_agCert = False
		b_netMode = False
		i_sockIndex = 0
		
		If Me.AgentStatus = VS(4) Then
			Call zAgentSetSymbol(mAgents.EnAgentSymbol.Terminal_Offline)
		Else
			Call zAgentSetSymbol(mAgents.EnAgentSymbol.User_Offline)
		End If
		
		ObjSck.Cleanup()
		'UPGRADE_NOTE: Object ObjSck may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
		ObjSck = Nothing
	End Function
	
	Public Sub AgnRecover()
		Dim DbR As DAO.Database
		Dim RsR As DAO.Recordset
		
		DbR = DAODBEngine_definst.OpenDatabase(CurIDBPath, False, False, ";pwd=nsb2003")
		RsR = DbR.OpenRecordset("TerminalRecover", DAO.RecordsetTypeEnum.dbOpenDynaset)
		
		With RsR
			.FindFirst("AgentName = '" & s_agName & "'")
			If .NoMatch = False Then
				.Edit()
			Else
				.AddNew()
			End If
			.Fields("AgentName").Value = s_agName
			.Fields("AgentStatus").Value = s_agStatus
			.Fields("CustomerName").Value = s_CusName
			.Fields("CustomerFlag").Value = s_CusFlag
			.Fields("CustomerType").Value = s_CusType
			.Fields("CustomerTimeIn").Value = s_CusTimeIn
			.Fields("CustomerTimeOut").Value = s_CusTimeOut
			.Fields("CustomerUsage").Value = s_CusUsage
			.Fields("CustomerPrinted").Value = l_CusPrinted
			.Update()
		End With
		RsR.Close()
		DbR.Close()
	End Sub
	
	Public Sub AgnRecoverRemove()
		Dim DbR As DAO.Database
		Dim RsR As DAO.Recordset
		
		DbR = DAODBEngine_definst.OpenDatabase(CurIDBPath, False, False, ";pwd=nsb2003")
		RsR = DbR.OpenRecordset("TerminalRecover", DAO.RecordsetTypeEnum.dbOpenDynaset)
		
		With RsR
			.FindFirst("AgentName = '" & s_agName & "'")
			If .NoMatch = False Then .Delete()
		End With
		RsR.Close()
		DbR.Close()
	End Sub
	
	Public Sub AgnTransfer(ByRef NewLvIndex As Short)
		Dim nLv As MSComctlLib.ListItem
		Dim uA As clsAgent
		
		nLv = ClistView.ListItems(NewLvIndex)
		uA = UniAgents.Agents((nLv.Text))
		Send(i_sockIndex, "//logout")
		
		'copy data
		uA.CustomerName = s_CusName
		uA.CustomerTimeIn = s_CusTimeIn
		uA.CustomerTimeOut = s_CusTimeOut
		uA.AgentStatus = s_agStatus
		uA.CustomerType = s_CusType
		uA.CustomerFlag = s_CusFlag
		uA.CustomerUsage = s_CusUsage
		uA.CustomerPrinted = l_CusPrinted
		
		'copy appearance
		uA.AgentForeColor = System.Drawing.ColorTranslator.FromOle(System.Convert.ToInt32(ClvItem.ForeColor))
		'UPGRADE_WARNING: Couldn't resolve default property of object ClvItem.SmallIcon. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		uA.AgentSmallIcon = ClvItem.SmallIcon
		'UPGRADE_WARNING: Couldn't resolve default property of object ClvItem.Icon. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		uA.AgentIcon = ClvItem.Icon
		uA.NetSend("//login")
		
		'reset current agent
		Call CusStop()
		Call AgnRecoverRemove()
	End Sub
	
	
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	' Remove Agent
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	Public Sub AgnRemove()
		On Error GoTo ErrInt
		ClistView.ListItems.Remove(ClvItem.Index)
		ObjSck.Cleanup()
		'UPGRADE_NOTE: Object clsAgent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
		clsAgent = Nothing
		Exit Sub
ErrInt: 
		ErrLog(Err, "Class Agent | AgnRemove")
		Resume Next
	End Sub
	
	
	Public Sub AgnAddPage(ByRef InfoType As Object)
		Dim AgentInfo As Object
		Dim LVdItem2 As Object
		Dim LVd2 As Object
		Dim ObjInfo As Object
		Dim a As Object
		Dim LVdItem1 As Object
		Dim tItm As MSComctlLib.ListItem
		Select Case InfoType
			Case 1
				tItm = MgoSlv1.ItemAdd(s_agName, s_agName, "dump")
				tItm.let_Tag(i_sockIndex)
				LVdItem1 = tItm
				'UPGRADE_WARNING: Couldn't resolve default property of object ObjInfo.PrintersCount. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				If ObjInfo.PrintersCount > 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object ObjInfo.PrintersCount. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					For a = 1 To ObjInfo.PrintersCount
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjInfo.Printers. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
						MgoSlv1.MatrixAdd(tItm, ObjInfo.Printers(a).DeviceName)
					Next a
					tItm.SubItems(1) = "All Printer"
				Else
					tItm.SubItems(1) = "No Printer"
				End If
				tItm.ListSubItems(1).ReportIcon = "printer"
				
				NetSend("//mon.switch:1")
			Case 2
				'UPGRADE_WARNING: Couldn't resolve default property of object LVd2.ListItems. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				tItm = LVd2.ListItems.Add( , s_agName, s_agName,  , "dump")
				tItm.let_Tag(i_sockIndex)
				LVdItem2 = tItm
				NetSend("//mon.switch:2")
				'UPGRADE_WARNING: Couldn't resolve default property of object AgentInfo.MemLoad. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				tItm.SubItems(1) = AgentInfo.MemLoad & " %"
				For a = 2 To 7
					'UPGRADE_WARNING: Couldn't resolve default property of object a. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					tItm.SubItems(a) = 0 & " Mb"
				Next a
		End Select
	End Sub
	
	
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	' Check Expires Customer for Prepaid & Fixed Time
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	Public Function CusCheckExpired() As Boolean
		CusCheckExpired = False
		If s_agStatus = VS(3) And CDate(s_CusTimeOut) <= Now Then
			'set ikon ke tamat
			Me.AgentSmallIcon = "tamat"
			'hantar ke agent supaya kunci
			Call NetSend("//kunci:1")
			'masukkan ke dalam lv utama
			Me.AgentStatus = VS(5)
			'tukar warna
			Me.AgentForeColor = System.Drawing.Color.Blue
			CusCheckExpired = True
		End If
	End Function
	
	Public Sub CusStatusUpdate()
		Dim mt As Object
		Dim hrg As Object
		Dim Cond3 As Object
		Dim Cond2 As Object
		Dim Cond1 As Object
		'{ updatekan harga dan hantar ke setiap terminal }'
		'UPGRADE_WARNING: Couldn't resolve default property of object Cond1. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Cond1 = s_CusTimeOut = VS(1) And s_agStatus = VS(3) 'kira jumlah harga
		'UPGRADE_WARNING: Couldn't resolve default property of object Cond2. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Cond2 = s_CusTimeOut <> VS(1) And s_agStatus = VS(3) 'kira masa tinggal
		'UPGRADE_WARNING: Couldn't resolve default property of object Cond3. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Cond3 = s_CusTimeOut <> VS(1) And s_agStatus = VS(5) 'tamat digunakan
		'UPGRADE_WARNING: Couldn't resolve default property of object Cond1. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If Cond1 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object hrg. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			hrg = CusGetPrice
			'UPGRADE_WARNING: Couldn't resolve default property of object hrg. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			NetSend("//harga:" & hrg)
			'UPGRADE_WARNING: Couldn't resolve default property of object Cond2. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ElseIf Cond2 Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object CusGetTimeLeft. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mt. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			mt = CusGetTimeLeft
			'UPGRADE_WARNING: Couldn't resolve default property of object mt. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			NetSend("//tiker: [ Time Left : " & mt & " ] ")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cond3. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ElseIf Cond3 Then 
			NetSend("//tiker: [ Thank You ] ")
		End If
	End Sub
	
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	' Get Price for Customer
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	Public Function CusGetPrice() As String
		Dim Ai As Short
		Dim gi01, pRate, gi02 As String
		
		'ambil harga mengikut skema
		If LCase(s_CusType) = LCase(VS(2)) Then
			'default pricing
			'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			pRate = SetAmbil("harga")
		Else
			'pricing for current scheme
			'UPGRADE_WARNING: Couldn't resolve default property of object uSDBe.DataFind(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			pRate = uSDBe.DataFind("skema", "skema", "harga", s_CusType)
		End If
		
		'ambil nilai biasa jika harga skema tidak dapat diterima
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If CDbl(pRate) < 0 Then pRate = SetAmbil("harga")
		
		'pengiraan harga..
		gi01 = CStr(DateDiff(Microsoft.VisualBasic.DateInterval.Minute, CDate(s_CusTimeIn), Now)) '+ Ai
		gi02 = CStr(CDbl(gi01) * CDbl(pRate)) ' + SetAmbil("hargaex")
		
		'return value.. easy heh..
		gi02 = CStr(GetRoundUpVal(CSng(gi02)))
		CusGetPrice = VB6.Format(gi02, "#0.00")
	End Function
	
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	' Get Time Usage for Customer
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	Public Function CusGetTimeUse(Optional ByRef InMinute As Boolean = False) As String
		Dim Jam, Minit As String
		
		'dapatkan jumlah minit
		Minit = CStr(DateDiff(Microsoft.VisualBasic.DateInterval.Minute, CDate(s_CusTimeIn), Now))
		If InMinute = True Then CusGetTimeUse = Minit : Exit Function
		'dapatkan jumlah jam
		Jam = CStr(CDbl(Minit) \ 60)
		'dapatkan minit selepas jam
		Minit = CStr(CDbl(Minit) - (CDbl(Jam) * 60))
		'gabungkan kesemua
		CusGetTimeUse = Jam & " " & VS(7) & ", " & Minit & " " & VS(8)
	End Function
	
	Public Function CusGetTimeUseEx() As Integer
		Dim Minit, Saat As Integer
		' dapatkan jumlah minit
		Minit = DateDiff(Microsoft.VisualBasic.DateInterval.Minute, CDate(s_CusTimeIn), Now)
		Saat = DateDiff(Microsoft.VisualBasic.DateInterval.Second, CDate(s_CusTimeIn), Now)
		CusGetTimeUseEx = Saat
	End Function
	
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	' Get Time Left for Customer
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	Public Function CusGetTimeLeft() As Object
		Dim s_Jam, s_Minit As String
		
		s_Minit = CStr(DateDiff(Microsoft.VisualBasic.DateInterval.Minute, Now, CDate(s_CusTimeOut)))
		s_Jam = CStr(CDbl(s_Minit) \ 60)
		s_Minit = CStr(CDbl(s_Minit) - (CDbl(s_Jam) * 60))
		'UPGRADE_WARNING: Couldn't resolve default property of object CusGetTimeLeft. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		CusGetTimeLeft = s_Jam & " " & VS(7) & ", " & s_Minit & " " & VS(8)
	End Function
	
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	' Start Pay As Your Go
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	Public Sub CusStartPAYG(ByRef CusName As String, ByRef CusType As String)
		Me.AgentStatus = VS(3)
		Me.CustomerName = CusName
		Me.CustomerType = CusType
		Me.CustomerFlag = "g"
		Me.CustomerTimeIn = CStr(Now)
		Me.CustomerTimeOut = VS(1)
		Me.CustomerUsage = VS(1)
		Me.CustomerPrinted = 0
		
		If b_agCert = True Then Call zAgentSetSymbol(mAgents.EnAgentSymbol.User_Online)
		If b_agCert = False Then Call zAgentSetSymbol(mAgents.EnAgentSymbol.User_Offline)
		Call NetSend("//login:" & SubBuild("name", CusName))
	End Sub
	
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	' Start Prepaid
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	Public Sub CusStartPPAID(ByRef CusName As String, ByRef CusType As String, ByRef PaidPrice As Double, Optional ByRef CusSambung As Boolean = False)
		Dim cRateEx, cRate, nVal As Double
		Dim gMin, gJam As Short
		Dim MinTmp As String
		
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		cRateEx = CDbl(SetAmbil("hargaex")) 'ambil harga extra
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		cRate = CDbl(SetAmbil("harga")) 'ambil harga
		
		Me.AgentStatus = VS(3)
		Me.CustomerName = CusName
		Me.CustomerType = CusType
		Me.CustomerTimeIn = CStr(Now)
		Me.CustomerPrinted = 0
		
		If CusSambung = True Then
			Select Case Left(s_CusFlag, 1)
				Case Is = "p"
					nVal = CDbl(Mid(s_CusFlag, 2)) + PaidPrice
					Me.CustomerFlag = "p" & nVal
				Case Is = "f"
					nVal = CDbl(Mid(s_CusFlag, 2)) * cRate
					Me.CustomerFlag = "p" & (nVal + PaidPrice)
			End Select
		Else
			Me.CustomerFlag = "p" & PaidPrice
		End If
		
		gMin = PaidPrice / cRate 'kira jumlah minit..
		Me.CustomerTimeOut = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Minute, gMin, Now)) 'bila akan berhenti..
		
		If gMin >= 60 Then
			gJam = gMin \ 60 : gMin = gMin - (gJam * 60)
			If gMin = 0 Then MinTmp = "" Else MinTmp = " " & gMin & " " & LCase(VS(8))
			Me.CustomerUsage = gJam & " " & LCase(VS(7)) & MinTmp
		Else
			Me.CustomerUsage = gMin & " " & LCase(VS(8))
		End If
		
		If b_agCert = True Then Call zAgentSetSymbol(mAgents.EnAgentSymbol.User_Online)
		If b_agCert = False Then Call zAgentSetSymbol(mAgents.EnAgentSymbol.User_Offline)
		Call NetSend("//kunci:0")
	End Sub
	
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	' Start Fixed Time
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	Public Sub CusStartTIME(ByRef CusName As String, ByRef CusType As String, ByRef UseHour As Short, ByRef UseMinute As Short, Optional ByRef CusSambung As Boolean = False)
		Dim sMin As Object
		Dim sJam As Object
		Dim id As Object
		Dim gMin, TmpMin As Short
		Dim cRateEx, cRate, cPaid As Double
		
		gMin = (UseHour * 60) + UseMinute 'jumlah masa yang ingin di gunakan dalam minit
		
		Me.AgentStatus = VS(3)
		Me.CustomerName = CusName
		Me.CustomerType = CusType
		Me.CustomerTimeIn = CStr(Now)
		Me.CustomerPrinted = 0
		
		If CusSambung = True Then
			'UPGRADE_WARNING: Couldn't resolve default property of object id. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			Select Case Left(s_CusFlag, 1)
				Case Is = "p"
					cPaid = CDbl(Mid(s_CusFlag, 2))
					'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					cRate = CDbl(SetAmbil("harga"))
					'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					cRateEx = CDbl(SetAmbil("hargaex"))
					TmpMin = cPaid / cRate
					Me.CustomerFlag = "f" & (TmpMin + gMin)
				Case CStr(id = "f")
					Me.CustomerFlag = "f" & (CDbl(Mid(s_CusFlag, 2)) + gMin)
			End Select
		Else
			Me.CustomerFlag = "f" & gMin
		End If
		
		Me.CustomerTimeOut = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Minute, gMin, Now)) 'masukkan masa berhenti ke dalam subitems(3)
		'UPGRADE_WARNING: Couldn't resolve default property of object sJam. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		sJam = IIf(UseHour <> 0, UseHour & " " & LCase(VS(7)) & " ", "")
		'UPGRADE_WARNING: Couldn't resolve default property of object sMin. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		sMin = IIf(UseMinute <> 0, UseMinute & " " & LCase(VS(8)), "")
		'UPGRADE_WARNING: Couldn't resolve default property of object sMin. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object sJam. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Me.CustomerUsage = sJam & sMin
		
		If b_agCert = True Then Call zAgentSetSymbol(mAgents.EnAgentSymbol.User_Online)
		If b_agCert = False Then Call zAgentSetSymbol(mAgents.EnAgentSymbol.User_Offline)
		Call NetSend("//kunci:0")
	End Sub
	
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	' Stop \ Terminate Customer
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	Public Sub CusStop()
		Me.AgentStatus = VS(4)
		Me.CustomerFlag = ""
		Me.CustomerName = ""
		Me.CustomerTimeIn = ""
		Me.CustomerTimeOut = ""
		Me.CustomerType = ""
		Me.CustomerUsage = ""
		Me.CustomerPrinted = 0
		
		If b_agCert = True Then Call zAgentSetSymbol(mAgents.EnAgentSymbol.Terminal_Online)
		If b_agCert = False Then Call zAgentSetSymbol(mAgents.EnAgentSymbol.Terminal_Offline)
		
		Me.AgentForeColor = System.Drawing.Color.Black
		Me.AgentFontBold = False
		
		FrmMain.DefInstance.MainSbar.Panels(2)._ObjectDefault = ""
		FrmMain.DefInstance.MainSbar.Panels(3)._ObjectDefault = ""
	End Sub
	
	
	'=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*==*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=
	'  PROPERTY SECTION
	'
	'
	'=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*==*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	' Agent | Connected Time
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	Public ReadOnly Property AgentConnected() As String
		Get
			AgentConnected = s_agConnected
		End Get
	End Property
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	' Agent | Name
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	Public Property AgentName() As String
		Get
			AgentName = s_agName
		End Get
		Set(ByVal Value As String)
			If Value <> "" Then
				s_agName = Value
				ClvItem.Text = s_agName
			End If
		End Set
	End Property
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	' Certification for Agent
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	Public ReadOnly Property AgentCertified() As Boolean
		Get
			AgentCertified = b_agCert
		End Get
	End Property
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	'  Agent Socket Index
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	Public ReadOnly Property AgentSockIndex() As Short
		Get
			AgentSockIndex = i_sockIndex
		End Get
	End Property
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	' Agent IP Address
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	Public ReadOnly Property AgentIPAdd() As String
		Get
			AgentIPAdd = s_agIP
		End Get
	End Property
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	' Agent MAC Address
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	Public ReadOnly Property AgentMAC() As String
		Get
			AgentMAC = s_agMAC
		End Get
	End Property
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	' Agent Small Icon
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	Public Property AgentSmallIcon() As String
		Get
			'UPGRADE_WARNING: Couldn't resolve default property of object ClvItem.SmallIcon. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			AgentSmallIcon = ClvItem.SmallIcon
		End Get
		Set(ByVal Value As String)
			If Value <> "" Then ClvItem.SmallIcon = Value
		End Set
	End Property
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	' Agent Icon
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	Public Property AgentIcon() As String
		Get
			'UPGRADE_WARNING: Couldn't resolve default property of object ClvItem.Icon. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			AgentIcon = ClvItem.Icon
		End Get
		Set(ByVal Value As String)
			If Value <> "" Then ClvItem.Icon = Value
		End Set
	End Property
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	' Agent Fore Colour
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	Public WriteOnly Property AgentForeColor() As System.Drawing.Color
		Set(ByVal Value As System.Drawing.Color)
			Dim B As Object
			On Error Resume Next
			ClvItem.ForeColor = System.Convert.ToUInt32(Value)
			For B = 1 To 7
				ClvItem.ListSubItems(B).ForeColor = System.Convert.ToUInt32(Value)
			Next B
		End Set
	End Property
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	' Agent Font Bold
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	Public WriteOnly Property AgentFontBold() As Boolean
		Set(ByVal Value As Boolean)
			Dim B As Object
			On Error Resume Next
			ClvItem.Bold = Value
			For B = 1 To 7
				ClvItem.ListSubItems(B).Bold = Value
			Next B
		End Set
	End Property
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	' Agent Status
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	Public Property AgentStatus() As String
		Get
			AgentStatus = s_agStatus
		End Get
		Set(ByVal Value As String)
			If s_agStatus = Value Then Exit Property
			s_agStatus = Value
			ClvItem.SubItems(1) = s_agStatus
		End Set
	End Property
	
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	' Customer Name
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	Public Property CustomerName() As String
		Get
			CustomerName = s_CusName
		End Get
		Set(ByVal Value As String)
			If s_CusName = Value Then Exit Property
			s_CusName = Value
			ClvItem.SubItems(2) = s_CusName
		End Set
	End Property
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	' Customer Time In
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	Public Property CustomerTimeIn() As String
		Get
			CustomerTimeIn = s_CusTimeIn
		End Get
		Set(ByVal Value As String)
			If s_CusTimeIn = Value Then Exit Property
			s_CusTimeIn = Value
			If s_CusTimeIn = VS(1) Or s_CusTimeIn = "" Then
				ClvItem.SubItems(4) = s_CusTimeIn
			Else
				ClvItem.SubItems(4) = CStr(TimeValue(s_CusTimeIn))
			End If
		End Set
	End Property
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	' Customer Time Out
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	Public Property CustomerTimeOut() As String
		Get
			CustomerTimeOut = s_CusTimeOut
		End Get
		Set(ByVal Value As String)
			If s_CusTimeOut = Value Then Exit Property
			s_CusTimeOut = Value
			If s_CusTimeOut = VS(1) Or s_CusTimeOut = "" Then
				ClvItem.SubItems(5) = s_CusTimeOut
			Else
				ClvItem.SubItems(5) = CStr(TimeValue(s_CusTimeOut))
			End If
		End Set
	End Property
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	' Customer Scheme
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	Public Property CustomerType() As String
		Get
			CustomerType = s_CusType
		End Get
		Set(ByVal Value As String)
			If s_CusType = Value Then Exit Property
			s_CusType = Value
			ClvItem.SubItems(3) = s_CusType
		End Set
	End Property
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	' Customer Flag
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	Public Property CustomerFlag() As String
		Get
			CustomerFlag = s_CusFlag
		End Get
		Set(ByVal Value As String)
			If s_CusFlag = Value Then Exit Property
			s_CusFlag = Value
			ClvItem.ListSubItems(1).let_Tag(Value)
		End Set
	End Property
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	' Customer Usage
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	Public Property CustomerUsage() As String
		Get
			CustomerUsage = s_CusUsage
		End Get
		Set(ByVal Value As String)
			If s_CusUsage = Value Then Exit Property
			s_CusUsage = Value
			ClvItem.SubItems(6) = Value
		End Set
	End Property
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	' Customer Printed
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	Public Property CustomerPrinted() As Integer
		Get
			CustomerPrinted = l_CusPrinted
		End Get
		Set(ByVal Value As Integer)
			If l_CusPrinted = Value Then Exit Property
			l_CusPrinted = Value
			ClvItem.SubItems(7) = CStr(Value)
		End Set
	End Property
	'=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*==*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=
	
	
	
	'=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*==*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=
	'  PRIVATE SUBS & FUNCTIONS
	'
	'
	'=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*==*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	' Add Agent to Listview, plus load recover data
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	Private Sub zAgentAdd(ByRef s_Nama As String)
		Dim DbR As DAO.Database
		Dim RsR As DAO.Recordset
		
		DbR = DAODBEngine_definst.OpenDatabase(CurIDBPath, False, False, ";pwd=nsb2003")
		RsR = DbR.OpenRecordset("TerminalRecover", DAO.RecordsetTypeEnum.dbOpenDynaset)
		ClvItem = ClistView.ListItems.Add( , s_Nama, s_Nama, "TerminalOnline", "TerminalOnline")
		
		ClvItem.let_Tag(i_sockIndex)
		s_agStatus = VS(4)
		s_agConnected = CStr(TimeOfDay)
		ClvItem.SubItems(1) = s_agStatus
		ClistView.Sorted = True : ClistView.Sorted = False
		
		With RsR
			.FindFirst("AgentName = '" & s_agName & "'")
			If .NoMatch = False Then
				Me.AgentStatus = .Fields("AgentStatus").Value
				Me.CustomerName = .Fields("CustomerName").Value
				Me.CustomerFlag = .Fields("CustomerFlag").Value
				Me.CustomerType = .Fields("CustomerType").Value
				Me.CustomerTimeIn = .Fields("CustomerTimeIn").Value
				Me.CustomerTimeOut = .Fields("CustomerTimeOut").Value
				Me.CustomerUsage = .Fields("CustomerUsage").Value
				Me.CustomerPrinted = .Fields("CustomerPrinted").Value
				If Me.AgentStatus = VS(5) Then
					Call zAgentSetSymbol(mAgents.EnAgentSymbol.User_Prepaid_Ended)
				Else
					If b_agCert = True Then Call zAgentSetSymbol(mAgents.EnAgentSymbol.User_Online)
					If b_agCert = False Then Call zAgentSetSymbol(mAgents.EnAgentSymbol.User_Offline)
				End If
			End If
		End With
		
		RsR.Close()
		DbR.Close()
	End Sub
	
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	' Set Agent Symbol
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	Private Sub zAgentSetSymbol(ByRef Symbol As mAgents.EnAgentSymbol)
		Select Case Symbol
			Case 0
				ClvItem.SmallIcon = "UserOnline"
				ClvItem.Icon = "UserOnline"
				Me.AgentForeColor = System.Drawing.ColorTranslator.FromOle(&H8000)
			Case 1
				ClvItem.SmallIcon = "UserOffline"
				ClvItem.Icon = "UserOnline" '!!!
				Me.AgentForeColor = System.Drawing.ColorTranslator.FromOle(&H8000)
			Case 2
				ClvItem.SmallIcon = "UserPrepaid"
				ClvItem.Icon = "UserPrepaid"
				Me.AgentForeColor = System.Drawing.ColorTranslator.FromOle(&H8000)
			Case 3
				ClvItem.SmallIcon = "UserEnded" ' = UserPrepaidEnded
				ClvItem.Icon = "UserEnded"
				Me.AgentForeColor = System.Drawing.Color.Blue
			Case 10
				ClvItem.SmallIcon = "TerminalOnline"
				ClvItem.Icon = "TerminalOnline"
			Case 11
				ClvItem.SmallIcon = "TerminalOffline"
				ClvItem.Icon = "TerminalOffline"
			Case 12
				ClvItem.SmallIcon = "TerminalLock"
				ClvItem.Icon = "TerminalLock"
			Case 13
				ClvItem.SmallIcon = "TerminalClean"
				ClvItem.Icon = "TerminalClean"
		End Select
	End Sub
	
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	' Get Free No Conflict Name
	'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
	Private Function zGetFreeName(ByRef Name As String) As String
Recheck: 
		zGetFreeName = Name
		If CekDuplicate(Name) = True Then
			zGetFreeName = CArrName(Name)
			GoTo Recheck
		End If
	End Function
	'=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*==*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=
End Class