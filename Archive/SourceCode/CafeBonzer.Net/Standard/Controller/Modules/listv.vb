Option Strict Off
Option Explicit On
Module mControls
	'==================================================================
	' Aplication codename : CafeBonzer
	' Programmer          : Azri Jamil a.k.a wajatimur
	' Module Name         : Controls
	' Description         :
	'==================================================================
	
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Cek jika nama telah digunakan dalam lv1
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Function CekDuplicate(ByRef Nama As Object) As Boolean
		Dim j As Object
		Dim Ltm As MSComctlLib.ListItems
		Ltm = FrmMain.DefInstance.Lv1.ListItems
		CekDuplicate = False
		
		If UniAgents.Count = 0 Then CekDuplicate = False : Exit Function
		For j = 1 To UniAgents.Count
			'UPGRADE_WARNING: Couldn't resolve default property of object Nama. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			If Ltm.Item(j).Text = Nama Then CekDuplicate = True : Exit Function
		Next j
	End Function
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Dapatkan subitem semasa
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Function SelSubItm(ByRef Index As Short) As String
		If UniAgents.Count = 0 Then Exit Function
		SelSubItm = FrmMain.DefInstance.Lv1.SelectedItem.SubItems(Index)
	End Function
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Dapatkan key semasa
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Function SelKey() As String
		If UniAgents.Count = 0 Then Exit Function
		SelKey = FrmMain.DefInstance.Lv1.SelectedItem.Key
	End Function
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Dapatkan index semasa
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Function SelIndex() As Short
		If UniAgents.Count = 0 Then Exit Function
		SelIndex = FrmMain.DefInstance.Lv1.SelectedItem.Index
	End Function
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Dapatkan Text semasa
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Function SelText() As String
		If UniAgents.Count = 0 Then Exit Function
		SelText = FrmMain.DefInstance.Lv1.SelectedItem.Text
	End Function
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Dapatkan Tag semasa
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Function SelTag() As String
		If UniAgents.Count = 0 Then Exit Function
		'UPGRADE_WARNING: Couldn't resolve default property of object FrmMain.Lv1.SelectedItem.Tag. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		SelTag = FrmMain.DefInstance.Lv1.SelectedItem.Tag
	End Function
	
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Copy dari ListItems ke ListItems
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Sub CopyItems(ByRef FromItems As MSComctlLib.ListItems, ByRef ToItems As MSComctlLib.ListItems, Optional ByRef CopyDataOnly As Boolean = True)
		Dim H As Object
		Dim g As Object
		If FromItems.Count = 0 Then Exit Sub
		
		For g = 1 To FromItems.Count
			ToItems(g).Text = FromItems(g).Text
			ToItems(g).Key = FromItems(g).Key
			ToItems(g).let_Tag(FromItems(g).Tag)
			If CopyDataOnly = False Then
				ToItems(g).Bold = FromItems(g).Bold
				ToItems(g).Checked = FromItems(g).Checked
				ToItems(g).ForeColor = System.Convert.ToUInt32(System.Drawing.ColorTranslator.FromOle(System.Convert.ToInt32(FromItems(g).ForeColor)))
				ToItems(g).Ghosted = FromItems(g).Ghosted
				'UPGRADE_WARNING: Couldn't resolve default property of object FromItems().Icon. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				ToItems(g).Icon = FromItems(g).Icon
				'UPGRADE_WARNING: Couldn't resolve default property of object FromItems().SmallIcon. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				ToItems(g).SmallIcon = FromItems(g).SmallIcon
				ToItems(g).ToolTipText = FromItems(g).ToolTipText
			End If
			
			For H = 1 To FromItems(1).ListSubItems.Count
				ToItems(g).ListSubItems(H).Text = FromItems(g).Text
				ToItems(g).ListSubItems(H).Key = FromItems(g).Key
				ToItems(g).ListSubItems(H).let_Tag(FromItems(g).Tag)
				If CopyDataOnly = False Then
					ToItems(g).ListSubItems(H).Bold = FromItems(g).Bold
					ToItems(g).ListSubItems(H).ForeColor = System.Convert.ToUInt32(System.Drawing.ColorTranslator.FromOle(System.Convert.ToInt32(FromItems(g).ForeColor)))
					ToItems(g).ListSubItems(H).ToolTipText = FromItems(g).ToolTipText
				End If
			Next H
		Next g
	End Sub
	
	
	
	'dapatkan key semasa
	Function TvSelKey(ByRef TrView As AxMSComctlLib.AxTreeView) As String
		Dim g As Object
		For g = 1 To TrView.Nodes.Count
			If TrView.Nodes.Item(g).Selected = True Then
				TvSelKey = TrView.Nodes.Item(g).Key
			End If
		Next g
	End Function
	'dapatkan index semasa
	Function TvSelIndex(ByRef TrView As AxMSComctlLib.AxTreeView) As Short
		Dim g As Object
		For g = 1 To TrView.Nodes.Count
			If TrView.Nodes.Item(g).Selected = True Then
				TvSelIndex = TrView.Nodes.Item(g).Index
			End If
		Next g
	End Function
	'dapatkan text semasa
	Function TvSelText(ByRef TrView As AxMSComctlLib.AxTreeView) As String
		Dim g As Object
		For g = 1 To TrView.Nodes.Count
			If TrView.Nodes.Item(g).Selected = True Then
				TvSelText = TrView.Nodes.Item(g).Text
			End If
		Next g
	End Function
End Module