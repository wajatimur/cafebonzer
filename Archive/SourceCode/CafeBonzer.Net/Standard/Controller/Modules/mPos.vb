Option Strict Off
Option Explicit On
Module mServices
	Public Sub LoadPosCatCB(ByRef iCboxName As AxMSComctlLib.AxImageCombo, ByRef ImgList As AxMSComctlLib.AxImageList)
		Dim t As Object
		Dim iconName As Object
		'rekodset kosong.. terus tambah "None"
		Rs = uSDB.OpenRecordset("pos-category", DAO.RecordsetTypeEnum.dbOpenSnapshot)
		If Rs.BOF = True Then GoTo AddNone
		With Rs
			.MoveFirst()
			Do Until .EOF = True
				'UPGRADE_WARNING: Couldn't resolve default property of object iconName. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				iconName = "folder"
				
				'check the availablelity of the icon
				For t = 1 To ImgList.ListImages.Count
					'UPGRADE_WARNING: Couldn't resolve default property of object iconName. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					If LCase(ImgList.ListImages(t).Key) = LCase(.Fields("Name").Value) Then iconName = LCase(.Fields("Name").Value)
				Next t
				'tambah ke dalam imagecombo
				iCboxName.ComboItems.Add( , "g" & .Fields("id").Value, .Fields("Name"), iconName, iconName, 1)
				.MoveNext()
			Loop 
		End With
		
AddNone: 
		iCboxName.ComboItems.Add( , VS(1), VS(1), "none", "none")
		iCboxName.ComboItems.Item(VS(1)).Selected = True
	End Sub
	
	Public Function LoadPosItmCB(ByRef iCboxName As AxMSComctlLib.AxImageCombo, ByRef GrpID As Object, ByRef ImgList As AxMSComctlLib.AxImageList) As Object
		Dim ret As Object
		Dim g As Object
		Dim tInt, tWd As Short
		Dim Rss As DAO.Recordset
		Dim CbItm As MSComctlLib.ComboItem
		Rss = uSDB.OpenRecordset("pos-items", DAO.RecordsetTypeEnum.dbOpenSnapshot)
		
		iCboxName.ComboItems.Clear()
		'UPGRADE_WARNING: Couldn't resolve default property of object GrpID. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Rss.Filter = "groupid = '" & GrpID & "'"
		Rs = Rss.OpenRecordset
		
		If Rs.BOF = True Then Exit Function
		With Rs
			.MoveFirst()
			Do Until .EOF = True
				CbItm = iCboxName.ComboItems.Add( , .Fields("id"), .Fields("Nama"),  ,  , 1)
				CbItm.let_Tag(.Fields("Harga"))
				.MoveNext()
			Loop 
		End With
		
		'resize the list width
		For g = 1 To iCboxName.ComboItems.Count
			'UPGRADE_ISSUE: Form method FrmMain.TextWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			tWd = FrmMain.TextWidth(iCboxName.ComboItems(g).Text)
			If tWd > tInt Then tInt = tWd
		Next g
		tInt = (tInt / 15) + 40
		'UPGRADE_WARNING: Couldn't resolve default property of object ret. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ret = SendMessage(iCboxName.hWnd, CB_SETDROPPEDWIDTH, tInt, 0)
		
		'default select
		iCboxName.ComboItems(1).Selected = True
		iCboxName.CtlRefresh()
	End Function
End Module