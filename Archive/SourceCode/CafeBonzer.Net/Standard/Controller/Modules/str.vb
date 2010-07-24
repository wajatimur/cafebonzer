Option Strict Off
Option Explicit On
Module mStringMath
	'==================================================================
	' Aplication codename : CafeBonzer
	' Programmer          : Azri Jamil a.k.a wajatimur
	' Module Name         : String Control
	' Description         :
	'==================================================================
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Bandingkan kewujudan huruf dalam ayat yang di berikan
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Function CharCompare(ByRef Ayat As String, ByRef SetAyat As String) As Boolean
		Dim BandingSetHuruf As Object
		Dim H As Object
		Dim g As Object
		Dim pjset As Object
		Dim Pjayat As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object Pjayat. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Pjayat = Len(Ayat)
		'UPGRADE_WARNING: Couldn't resolve default property of object pjset. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		pjset = Len(SetAyat)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Pjayat. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		For g = 1 To Pjayat
			'UPGRADE_WARNING: Couldn't resolve default property of object pjset. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			For H = 1 To pjset
				'UPGRADE_WARNING: Couldn't resolve default property of object H. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				'UPGRADE_WARNING: Couldn't resolve default property of object g. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				If Mid(Ayat, g, 1) = Mid(SetAyat, H, 1) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object BandingSetHuruf. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					BandingSetHuruf = True : Exit For
				End If
			Next H
		Next g
	End Function
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Bilang jumlah huruf dalam ayat yang diberikan
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Function CharCount(ByRef Ayat As String, ByRef Huruf As String) As Short
		Dim H As Object
		Dim BilangJumlahHuruf As Object
		Dim Pjayat As Object
		Dim Jhrf As Short
		'UPGRADE_WARNING: Couldn't resolve default property of object Pjayat. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Pjayat = Len(Ayat)
		
		If Len(Huruf) > 1 Or Len(Huruf) = 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object BilangJumlahHuruf. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			BilangJumlahHuruf = 0 : Exit Function
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Pjayat. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		For H = 1 To Pjayat
			'UPGRADE_WARNING: Couldn't resolve default property of object H. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			If Huruf = Mid(Ayat, H, 1) Then Jhrf = Jhrf + 1
		Next H
		
		'UPGRADE_WARNING: Couldn't resolve default property of object BilangJumlahHuruf. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		BilangJumlahHuruf = Jhrf
	End Function
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Nilai yang di roundup (depend on setting)
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Function GetRoundUpVal(ByRef dblValue As Double) As Double
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(roundup). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If SetAmbil("roundup") = 1 Then
			GetRoundUpVal = System.Math.Round(dblValue, 1)
		Else
			GetRoundUpVal = dblValue
		End If
	End Function
	
	Public Function GetBulan(ByRef MonthNum As Object) As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object MonthNum. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Choose(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		GetBulan = Choose(MonthNum, "Januari", "Februari", "Mac", "April", "Mei", "Jun", "Julai", "Ogos", "September", "Oktober", "November", "Disember")
	End Function
	
	Public Function Text2Num(ByRef Text As String) As Double
		On Error Resume Next
		If Text = "" Then
			Text2Num = 0
		Else
			Text2Num = CDbl(Text)
		End If
	End Function
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Tukarkan nama ke bentuk array
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Function CArrName(ByRef Nama As String) As String
		Dim arrnum As Object
		Dim K As Object
		Dim j As Object
		If Right(Nama, 1) <> ")" Then
			CArrName = Nama & "(1)"
			Exit Function
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			j = InStrRev(Nama, "(", -1)
			'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			If j <> 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				'UPGRADE_WARNING: Couldn't resolve default property of object K. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				K = Mid(Nama, j + 1, Len(Nama) - j - 2)
				If IsNumeric(K) = True Then
					'UPGRADE_WARNING: Couldn't resolve default property of object arrnum. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					arrnum = GetArrNameNum(Nama) + 1
					'UPGRADE_WARNING: Couldn't resolve default property of object arrnum. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					CArrName = Left(Nama, j - 1) & "(" & arrnum & ")"
				Else
					CArrName = Nama & "(1)"
					Exit Function
				End If
			Else
				CArrName = Nama & "(1)"
				Exit Function
			End If
		End If
	End Function
	
	Public Function GetDirName(ByRef DirPath As String) As Object
		Dim cf As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object cf. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		cf = InStrRev(DirPath, "\", -1)
		'UPGRADE_WARNING: Couldn't resolve default property of object cf. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		GetDirName = Mid(DirPath, cf + 1)
	End Function
	
	Public Function GetArrNameNum(ByRef Nama As Object) As Short
		Dim j As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object Nama. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		j = InStrRev(Nama, "(", Len(Nama) - 1)
		'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Nama. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		GetArrNameNum = CShort(Mid(Nama, j + 1, Len(Nama) - j - 2))
	End Function
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Rounding Nombor
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Function RoundNum(ByRef numVal As Double, ByRef numDigits As Short) As Double
		RoundNum = Int(numVal * (10 ^ numDigits) + 0.5) / (10 ^ numDigits)
	End Function
End Module