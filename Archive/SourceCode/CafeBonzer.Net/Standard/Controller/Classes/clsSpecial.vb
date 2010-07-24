Option Strict Off
Option Explicit On
Friend Class clsSpecial
	Private Lv As AxMSComctlLib.AxListView
	Private c_Cb As System.Windows.Forms.ComboBox
	Private StackMatrix As New Collection
	
	
	Public Sub Init(ByRef Lview As AxMSComctlLib.AxListView)
		Lv = Lview
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object Lv may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
		Lv = Nothing
		'UPGRADE_NOTE: Object c_Cb may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
		c_Cb = Nothing
		'UPGRADE_NOTE: Object StackMatrix may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
		StackMatrix = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Public Function ItemAdd(Optional ByRef Key As Object = Nothing, Optional ByRef Text As Object = Nothing, Optional ByRef SmallIcon As Object = Nothing) As MSComctlLib.ListItem
		'On Error GoTo ErrInt
		Dim NewItm As MSComctlLib.ListItem
		Dim NewDts As New clsDataStore
		
		NewItm = Lv.ListItems.Add( , Key, Text,  , SmallIcon)
		'UPGRADE_WARNING: Couldn't resolve default property of object Key. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		NewDts.Name = Key
		StackMatrix.Add(NewDts, Key)
		ItemAdd = NewItm
		'UPGRADE_NOTE: Object NewItm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
		NewItm = Nothing
		'UPGRADE_NOTE: Object NewDts may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
		NewDts = Nothing
		Exit Function
		
ErrInt: 
		ErrLog(Err, "clsSpecial | AddItem")
	End Function
	
	Public Sub ItemRemove(ByRef Key As Object)
		StackMatrix.Remove(Key)
		Lv.ListItems.Remove(Key)
	End Sub
	
	Public Sub ItemClear()
		Lv.ListItems.Clear()
		'UPGRADE_NOTE: Object StackMatrix may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
		StackMatrix = Nothing
	End Sub
	
	
	Public Sub MatrixAdd(ByRef ParentKey As Object, ByRef Data As Object)
		Dim DTS As clsDataStore
		
		DTS = StackMatrix.Item(ParentKey)
		DTS.Add(Data, DTS.Count + 1)
		'UPGRADE_NOTE: Object DTS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
		DTS = Nothing
	End Sub
	
	Public Sub MatrixExpand(ByRef Parent As MSComctlLib.ListItem)
		Dim a As Object
		On Error GoTo ErrInt
		Dim t_Rect As RECT
		Dim l_TwipsX, l_TwipsY As Integer
		Dim v_Col As Object
		Dim l_mPos As Integer
		Dim DTS As clsDataStore
		
		If Parent.Key = "" Then Exit Sub
		DTS = StackMatrix.Item(Parent.Key)
		If DTS.Count = 0 Then Exit Sub
		
		l_mPos = 1
		l_TwipsX = VB6.TwipsPerPixelX
		l_TwipsY = VB6.TwipsPerPixelY
		c_Cb = FrmAgnInfo.DefInstance.DynaCombo
		
		SetParent(c_Cb.Handle.ToInt32, Lv.hWnd)
		GetSubItemRect(Lv.hWnd, Parent.Index - 1, l_mPos, LVIR_LABEL, t_Rect)
		
		'set position
		With t_Rect
			c_Cb.Left = VB6.TwipsToPixelsX((.Left_Renamed * l_TwipsX))
			c_Cb.Top = VB6.TwipsToPixelsY((.Top * l_TwipsY))
			c_Cb.Width = VB6.TwipsToPixelsX((.Right_Renamed - .Left_Renamed) * l_TwipsX)
		End With
		
		'load all data to combobox
		c_Cb.Items.Clear()
		If DTS.Count = 0 Then
			c_Cb.Items.Add("No Printer")
		Else
			c_Cb.Items.Add("All Printer")
			For a = 1 To DTS.Count
				'UPGRADE_WARNING: Couldn't resolve default property of object DTS(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				c_Cb.Items.Add(DTS.Data(a))
			Next 
		End If
		c_Cb.SelectedIndex = 0
		c_Cb.Visible = True
		'UPGRADE_NOTE: Object DTS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
		DTS = Nothing
		Exit Sub
		
ErrInt: 
		ErrLog(Err, "clsSpecial | MatrixExpand")
	End Sub
	
	Public Sub MatrixClear(ByRef Key As Object)
		On Error GoTo ErrInt
		Dim DTS As clsDataStore
		Lv.ListItems.Clear()
		DTS = StackMatrix.Item(Key)
		DTS.Clear()
		'UPGRADE_NOTE: Object DTS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
		DTS = Nothing
		Exit Sub
ErrInt: 
		ErrLog(Err, "clsSpecial | MatrixClear")
	End Sub
	
	
	Private Function GetSubItemRect(ByVal hWndLV As Integer, ByVal iItem As Integer, ByVal iSubItem As Integer, ByVal code As Integer, ByRef lpRect As RECT) As Boolean
		lpRect.Top = iSubItem
		lpRect.Left_Renamed = code
		'UPGRADE_WARNING: Couldn't resolve default property of object lpRect. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		GetSubItemRect = SendMessage(hWndLV, LVM_GETSUBITEMRECT, iItem, lpRect)
	End Function
End Class