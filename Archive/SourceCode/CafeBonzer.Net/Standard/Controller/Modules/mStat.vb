Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Module mStatistic
	
	'=-=-====-==-===-===-=-==-=-=--==-=-=-=-===-==-=-==-=-=-==---=--
	' Simpan Senarai Penggunaan PC
	'=-=-====-==-===-===-=-==-=-=--==-=-=-=-===-==-=-==-=-=-==---=--
	Public Sub SavePcUsage(ByRef PcName As Object, ByRef CusName As Object, ByRef TimeIn As Object, ByRef TimeOut As Object, ByRef JumlahBayar As Object)
		uIDBe.DataSave("pc-usage", "tahun", Year(Today), True, False)
		uIDBe.DataSave("pc-usage", "bulan", Month(Today), False, False)
		uIDBe.DataSave("pc-usage", "hari", VB.Day(Today), False, False)
		uIDBe.DataSave("pc-usage", "pcname", PcName, False, False)
		uIDBe.DataSave("pc-usage", "nama", CusName, False, False)
		uIDBe.DataSave("pc-usage", "masuk", TimeIn, False, False)
		uIDBe.DataSave("pc-usage", "keluar", TimeOut, False, False)
		uIDBe.DataSave("pc-usage", "harga", JumlahBayar, False, True)
	End Sub
	
	
	'=-=-====-==-===-===-=-==-=-=--==-=-=-=-===-==-=-==-=-=-==---=--
	' Jualan Sebulan Bagi Setiap PC
	'=-=-====-==-===-===-=-==-=-=--==-=-=-=-===-==-=-==-=-=-==---=--
	Public Sub SavePcBulanan(ByRef PcName As Object, ByRef JumlahMasa As Object, ByRef JumlahBayar As Object)
		Dim mBef As Integer
		Dim bBef As Double
		Dim RsSS As DAO.Recordset
		RsSS = uIDB.OpenRecordset("pc-bulanan", DAO.RecordsetTypeEnum.dbOpenDynaset)
		
		RsSS.Filter = "tahun = " & Year(Today) & " AND bulan = " & Month(Today) '& "'"
		Rs = RsSS.OpenRecordset
		
		With Rs
			'UPGRADE_WARNING: Couldn't resolve default property of object PcName. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			.FindFirst("namapc = '" & PcName & "'")
			If .NoMatch = False Then
				mBef = .Fields("JumlahMasa").Value
				bBef = .Fields("JumlahBayar").Value
				.Edit()
				'UPGRADE_WARNING: Couldn't resolve default property of object JumlahMasa. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				.Fields("JumlahMasa").Value = JumlahMasa + mBef
				'UPGRADE_WARNING: Couldn't resolve default property of object JumlahBayar. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				.Fields("JumlahBayar").Value = VB6.Format(JumlahBayar + bBef, "#0.00")
				.Update()
			Else
				.AddNew()
				.Fields("Tahun").Value = Year(Today)
				.Fields("Bulan").Value = Month(Today)
				'UPGRADE_WARNING: Couldn't resolve default property of object PcName. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				.Fields("NamaPc").Value = PcName
				'UPGRADE_WARNING: Couldn't resolve default property of object JumlahMasa. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				.Fields("JumlahMasa").Value = JumlahMasa
				'UPGRADE_WARNING: Couldn't resolve default property of object JumlahBayar. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				.Fields("JumlahBayar").Value = JumlahBayar
				.Update()
			End If
		End With
	End Sub
	
	
	'=-=-====-==-===-===-=-==-=-=--==-=-=-=-===-==-=-==-=-=-==---=--
	' Jualan Harian
	'=-=-====-==-===-===-=-==-=-=--==-=-=-=-===-==-=-==-=-=-==---=--
	Public Sub SavePcHarian(ByRef JumBayar As Object)
		Dim sMonth, sYear, sDay As String
		Dim vstBef As Short
		Dim jBef As Double
		Rs = uIDB.OpenRecordset("pc-harian", DAO.RecordsetTypeEnum.dbOpenDynaset)
		
		
		
		With Rs
			'for current date
			sYear = CStr(Year(Today))
			sMonth = CStr(Month(Today))
			sDay = CStr(VB.Day(Today))
			.FindFirst("tahun = " & sYear & " AND bulan = " & sMonth & " AND hari = " & sDay) ' & "'"
			If .NoMatch = False Then
				vstBef = .Fields("pelanggan").Value
				jBef = .Fields("pungutan").Value
				.Edit()
				.Fields("pelanggan").Value = vstBef + 1
				'UPGRADE_WARNING: Couldn't resolve default property of object JumBayar. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				.Fields("pungutan").Value = VB6.Format(jBef + JumBayar, "#0.00")
				.Update()
			Else
				.AddNew()
				.Fields("Tahun").Value = sYear
				.Fields("Bulan").Value = sMonth
				.Fields("Hari").Value = sDay
				.Fields("pelanggan").Value = 1
				'UPGRADE_WARNING: Couldn't resolve default property of object JumBayar. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				.Fields("pungutan").Value = JumBayar
				.Update()
			End If
			
			'for session date
			sYear = CStr(Year(CDate(OpenSessionCur)))
			sMonth = CStr(Month(CDate(OpenSessionCur)))
			sDay = CStr(VB.Day(CDate(OpenSessionCur)))
			.FindFirst("tahun = " & sYear & " AND bulan = " & sMonth & " AND hari = " & sDay) '& "'"
			If .NoMatch = False Then
				jBef = .Fields("closing").Value
				.Edit()
				'UPGRADE_WARNING: Couldn't resolve default property of object JumBayar. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				.Fields("closing").Value = VB6.Format(jBef + JumBayar, "#0.00")
				.Update()
			End If
		End With
	End Sub
	
	
	'=-=-====-==-===-===-=-==-=-=--==-=-=-=-===-==-=-==-=-=-==---=--
	' Jualan Harian
	'=-=-====-==-===-===-=-==-=-=--==-=-=-=-===-==-=-==-=-=-==---=--
	Public Sub SavePcMingguan(ByRef Hari As Object, ByRef JumJualan As Object)
		Dim jBef As Double
		Rs = uIDB.OpenRecordset("pc-grafminggu", DAO.RecordsetTypeEnum.dbOpenDynaset)
		
		With Rs
			.FindFirst("tahun = " & Year(Today) & " AND bulan = " & Month(Today)) '& "'"
			If .NoMatch = False Then
				jBef = .Fields(Hari).Value
				.Edit()
				'UPGRADE_WARNING: Couldn't resolve default property of object JumJualan. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				.Fields(Hari).Value = VB6.Format(jBef + JumJualan, "#0.00")
				.Update()
			Else
				.AddNew()
				.Fields("Tahun").Value = Year(Today)
				.Fields("Bulan").Value = Month(Today)
				'UPGRADE_WARNING: Couldn't resolve default property of object JumJualan. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				.Fields(Hari).Value = JumJualan
				.Update()
			End If
		End With
	End Sub
	
	
	'=-=-====-==-===-===-=-==-=-=--==-=-=-=-===-==-=-==-=-=-==---=--
	' Simpan Senarai Pelanggan
	'=-=-====-==-===-===-=-==-=-=--==-=-=-=-===-==-=-==-=-=-==---=--
	Public Sub SavePelanggan(ByRef CusName As Object, ByRef CusID As Object, ByRef JumlahMasa As Object, ByRef JumlahBayar As Object)
		Dim mBef As Integer
		Dim bBef As Double
		Dim vstBef As Short
		Rs = uSDB.OpenRecordset("pelanggan-list", DAO.RecordsetTypeEnum.dbOpenTable)
		
		With Rs
			'UPGRADE_WARNING: Couldn't resolve default property of object CusID. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			If CusID <> "" Then
				.Index = "noahli"
				.Seek("=", CusID)
			Else
				.Index = "nama"
				.Seek("=", CusName)
			End If
			
			If .NoMatch = False Then
				mBef = .Fields("JumlahMasa").Value
				bBef = .Fields("JumlahBayar").Value
				vstBef = .Fields("lawat").Value
				.Edit()
				'UPGRADE_WARNING: Couldn't resolve default property of object JumlahMasa. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				.Fields("JumlahMasa").Value = mBef + JumlahMasa
				'UPGRADE_WARNING: Couldn't resolve default property of object JumlahBayar. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				.Fields("JumlahBayar").Value = VB6.Format(bBef + JumlahBayar, "#0.00")
				.Fields("lawat").Value = vstBef + 1
				.Fields("tarikhakhir").Value = Now
				.Update()
			Else
				.AddNew()
				'UPGRADE_WARNING: Couldn't resolve default property of object CusID. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				.Fields("noahli").Value = CusID
				'UPGRADE_WARNING: Couldn't resolve default property of object CusName. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				.Fields("Nama").Value = CusName
				'UPGRADE_WARNING: Couldn't resolve default property of object JumlahMasa. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				.Fields("JumlahMasa").Value = JumlahMasa
				'UPGRADE_WARNING: Couldn't resolve default property of object JumlahBayar. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				.Fields("JumlahBayar").Value = JumlahBayar
				.Fields("lawat").Value = 1
				.Fields("tarikhakhir").Value = Now
				.Update()
			End If
		End With
	End Sub
	
	
	'=-=-====-==-===-===-=-==-=-=--==-=-=-=-===-==-=-==-=-=-==---=--
	' Simpan Transaksi POS
	'=-=-====-==-===-===-=-==-=-=--==-=-=-=-===-==-=-==-=-=-==---=--
	Public Sub SavePosTrans(ByRef Items As MSComctlLib.ListItems)
		Dim GroupId As Object
		Dim g As Object
		Dim RsF As DAO.Recordset
		If Items.Count = 0 Then Exit Sub
		
		Rs = uIDB.OpenRecordset("pos-usage", DAO.RecordsetTypeEnum.dbOpenDynaset)
		Rs.Filter = "tahun = " & Year(Today) & " AND bulan = " & Month(Today) & " AND hari = " & VB.Day(Today) '& "'"
		RsF = Rs.OpenRecordset
		
		With RsF
			'.MoveLast
			'.MoveFirst
			For g = 1 To Items.Count
				.AddNew()
				.Fields("Tahun").Value = Year(Today)
				.Fields("Bulan").Value = Month(Today)
				.Fields("Hari").Value = VB.Day(Today)
				.Fields("GroupId").Value = Mid(Items(g).Key, 2, 2)
				.Fields("id").Value = Items(g).Key
				'UPGRADE_WARNING: Couldn't resolve default property of object GroupId. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				.Fields("transid").Value = .Fields("Tahun").Value & .Fields("Bulan").Value & .Fields("Hari").Value & GroupId & VB6.Format(.RecordCount, "#0000")
				.Fields("Item").Value = Items(g).Text
				.Fields("qty").Value = Items(g).SubItems(2)
				.Fields("Harga").Value = VB6.Format(CDbl(Items(g).SubItems(1)) * CShort(.Fields("qty").Value), "#0.00")
				.Update()
			Next g
		End With
	End Sub
End Module