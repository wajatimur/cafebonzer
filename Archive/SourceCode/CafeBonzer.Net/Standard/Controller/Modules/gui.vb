Option Strict Off
Option Explicit On
Module mInterface
	Public Im As New VisualSuite1.VsIconMenu
	
	Public MgoInpt As New VisualSuite1.VsInput
	Public MgoSlv1 As New clsSpecial
	Public MglPageLast As Integer
	
	Private CbLv1 As AxMSComctlLib.AxListView
	Private CbLv2 As AxMSComctlLib.AxListView
	
	Public Enum eFrmMetric
		Pos_X = 1
		Pos_Y = 2
		Size_X = 3
		Size_Y = 4
		State = 5
	End Enum
	
	Public Sub LoadIconic()
		Im.Attach(FrmMain.DefInstance.Handle.ToInt32)
		With Im
			.HighlightStyle = VisualSuite1.ECPHighlightStyleConstants.ECPHighlightStyleGradient
			.ImageList = FrmMain.DefInstance.ImgList16
			.IconIndex(FrmMain.DefInstance.menu1penalaan.Text) = FrmMain.DefInstance.ImgList16.ListImages("penalaan").Index - 1
			.IconIndex(FrmMain.DefInstance.menu1logout.Text) = FrmMain.DefInstance.ImgList16.ListImages("logoff").Index - 1
			.IconIndex(FrmMain.DefInstance.menu3mesej.Text) = FrmMain.DefInstance.ImgList16.ListImages("mesej").Index - 1
			.IconIndex(FrmMain.DefInstance.menu3ctllock(0).Text) = FrmMain.DefInstance.ImgList16.ListImages("TerminalLock").Index - 1
			.IconIndex(FrmMain.DefInstance.menu3ctllock(1).Text) = FrmMain.DefInstance.ImgList16.ListImages("TerminalLock").Index - 1
			.IconIndex(FrmMain.DefInstance.menu3ctllock(2).Text) = FrmMain.DefInstance.ImgList16.ListImages("kuncibuka").Index - 1
			.IconIndex(FrmMain.DefInstance.menu3ctlwinexit(0).Text) = FrmMain.DefInstance.ImgList16.ListImages("off").Index - 1
			.IconIndex(FrmMain.DefInstance.menu3ctlwinexit(2).Text) = FrmMain.DefInstance.ImgList16.ListImages("boot").Index - 1
			.IconIndex(FrmMain.DefInstance.menu2bantuan.Text) = FrmMain.DefInstance.ImgList16.ListImages("help").Index - 1
			.IconIndex(FrmMain.DefInstance.menu2aplikasi.Text) = FrmMain.DefInstance.ImgList16.ListImages("info").Index - 1
			.IconIndex(FrmMain.DefInstance.menu3tiker.Text) = FrmMain.DefInstance.ImgList16.ListImages("hoi").Index - 1
			.IconIndex(FrmMain.DefInstance.menu2keluar.Text) = FrmMain.DefInstance.ImgList16.ListImages("power").Index - 1
			.IconIndex(FrmMain.DefInstance.menu3announce.Text) = FrmMain.DefInstance.ImgList16.ListImages("broad").Index - 1
			.IconIndex(FrmMain.DefInstance.menu4PosMgr.Text) = FrmMain.DefInstance.ImgList16.ListImages("gift").Index - 1
			.IconIndex(FrmMain.DefInstance.menu4Stat.Text) = FrmMain.DefInstance.ImgList16.ListImages("graft").Index - 1
			
			.IconIndex(FrmMain.DefInstance.pmenu1flog.Text) = FrmMain.DefInstance.ImgList16.ListImages("jalan1").Index - 1
			.IconIndex(FrmMain.DefInstance.pmenu1cancel.Text) = FrmMain.DefInstance.ImgList16.ListImages("no").Index - 1
			.IconIndex(FrmMain.DefInstance.pmenu1trans.Text) = FrmMain.DefInstance.ImgList16.ListImages("transfer").Index - 1
			.IconIndex(FrmMain.DefInstance.pmenu1terminal.Text) = FrmMain.DefInstance.ImgList16.ListImages("term").Index - 1
			.IconIndex(FrmMain.DefInstance.pmenu1cln.Text) = FrmMain.DefInstance.ImgList16.ListImages("TerminalClean").Index - 1
			.IconIndex(FrmMain.DefInstance.pmenu1ctl.Text) = FrmMain.DefInstance.ImgList16.ListImages("cpu").Index - 1
			.IconIndex(FrmMain.DefInstance.pmenu1ctlsub(0).Text) = FrmMain.DefInstance.ImgList16.ListImages("TerminalLock").Index - 1
			.IconIndex(FrmMain.DefInstance.pmenu1ctlsub(1).Text) = FrmMain.DefInstance.ImgList16.ListImages("kuncibuka").Index - 1
			.IconIndex(FrmMain.DefInstance.pmenu1ctlsub(2).Text) = FrmMain.DefInstance.ImgList16.ListImages("boot").Index - 1
			.IconIndex(FrmMain.DefInstance.pmenu1ctlsub(3).Text) = FrmMain.DefInstance.ImgList16.ListImages("off").Index - 1
		End With
	End Sub
	
	Public Sub UnloadIconic()
		Im.Detach()
	End Sub
	Sub StatText(Optional ByRef Panel As Short = 0, Optional ByRef Text As String = "")
		FrmMain.DefInstance.MainSbar.Panels(Panel).Text = Text
	End Sub
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' ComboBox Add With Trim
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Sub CbAddEx(ByRef Item As Object, ByRef Cbox As System.Windows.Forms.ComboBox)
		Dim a As Object
		For a = 0 To Cbox.Items.Count - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object Item. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Couldn't resolve default property of object a. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			If Trim(VB6.GetItemString(Cbox, a)) = Trim(Item) Then Exit Sub
		Next a
		'UPGRADE_WARNING: Couldn't resolve default property of object Item. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Cbox.Items.Add(Item)
	End Sub
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Tutup form dan bebaskan resource
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Sub CloseFrm(ByRef AnyFrm As System.Windows.Forms.Form)
		AnyFrm.Hide()
		AnyFrm.Close()
		'UPGRADE_NOTE: Object AnyFrm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
		AnyFrm = Nothing
	End Sub
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Benarkan nombor sahaja
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	'UPGRADE_NOTE: Reset was upgraded to Reset_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
	Public Sub NumOnly(ByRef TextControl As System.Windows.Forms.TextBox, Optional ByRef Reset_Renamed As Boolean = False)
		Dim Style As Integer
		
		Style = GetWindowLong(TextControl.Handle.ToInt32, GWL_STYLE)
		SetWindowLong(TextControl.Handle.ToInt32, GWL_STYLE, Style Or ES_NUMBER)
	End Sub
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Pengiraan beza objek antara objek kecil dan besar
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Function KiraBezaSaizX(ByRef ObjekBesar As Object, ByRef ObjekKecil As Object) As Integer
		'UPGRADE_WARNING: Couldn't resolve default property of object ObjekKecil.Width. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object ObjekBesar.Width. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		KiraBezaSaizX = ObjekBesar.Width - ObjekKecil.Width
	End Function
	Function KiraBezaSaizY(ByRef ObjekBesar As Object, ByRef ObjekKecil As Object) As Integer
		'UPGRADE_WARNING: Couldn't resolve default property of object ObjekKecil.Height. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object ObjekBesar.Height. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		KiraBezaSaizY = ObjekBesar.Height - ObjekKecil.Height
	End Function
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Simpan info metric bagi form
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Sub CbFrmMetricSave(ByRef FormName As System.Windows.Forms.Form)
		SetSimpan(FormName.Name & "saizx", CStr(VB6.PixelsToTwipsX(FormName.Width)))
		SetSimpan(FormName.Name & "saizy", CStr(VB6.PixelsToTwipsY(FormName.Height)))
		SetSimpan(FormName.Name & "posx", CStr(VB6.PixelsToTwipsX(FormName.Left)))
		SetSimpan(FormName.Name & "posY", CStr(VB6.PixelsToTwipsY(FormName.Top)))
		SetSimpan(FormName.Name & "state", CStr(FormName.WindowState))
	End Sub
	
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Load metric info
	'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Sub CbFrmMetricLoad(ByRef FormName As System.Windows.Forms.Form)
		FormName.Top = VB6.TwipsToPixelsY(CbFrmMetricGet(FormName, eFrmMetric.Pos_X))
		FormName.Left = VB6.TwipsToPixelsX(CbFrmMetricGet(FormName, eFrmMetric.Pos_X))
		FormName.Width = VB6.TwipsToPixelsX(CbFrmMetricGet(FormName, eFrmMetric.Size_X))
		FormName.Height = VB6.TwipsToPixelsY(CbFrmMetricGet(FormName, eFrmMetric.Size_Y))
		FormName.WindowState = CbFrmMetricGet(FormName, eFrmMetric.State)
	End Sub
	
	Public Function CbFrmMetricGet(ByRef FormName As System.Windows.Forms.Form, ByRef MetricType As eFrmMetric) As Integer
		Dim mtrc As Integer
		
		Select Case MetricType
			Case 1
				'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				mtrc = SetAmbil(FormName.Name & "posx", 0)
			Case 2
				'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				mtrc = SetAmbil(FormName.Name & "posy", 0)
			Case 3
				'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				mtrc = SetAmbil(FormName.Name & "saizx", 11460)
			Case 4
				'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				mtrc = SetAmbil(FormName.Name & "saizy", 8640)
			Case 5
				'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				mtrc = SetAmbil(FormName.Name & "state", 0)
		End Select
		CbFrmMetricGet = mtrc
	End Function
	
	Public Sub MoveFrm(ByRef hwnd As Integer)
		Dim lret As Object
		ReleaseCapture()
		'UPGRADE_WARNING: Couldn't resolve default property of object lret. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		lret = SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
	End Sub
	
	Public Sub PutOnTop(ByRef hwnd As Integer)
		Dim i As Integer
		i = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_WNDFLAGS)
	End Sub
	
	Public Sub ConstructPage(ByRef PageType As Integer)
		Dim a As Object
		On Error GoTo ErrInt
		
		CbLv1 = FrmAgnInfo.DefInstance.DynaLv(0)
		CbLv2 = FrmAgnInfo.DefInstance.DynaLv(1)
		
		If PageType = 0 And CbViewMode = 0 Then
			FrmMain.DefInstance.Lv1.View = MSComctlLib.ListViewConstants.lvwIcon
			FrmMain.DefInstance.Lv1.Arrange = MSComctlLib.ListArrangeConstants.lvwNone
			CbViewMode = 1
		ElseIf PageType = 0 And CbViewMode = 1 Then 
			FrmMain.DefInstance.Lv1.View = MSComctlLib.ListViewConstants.lvwReport
			FrmMain.DefInstance.Lv1.Arrange = MSComctlLib.ListArrangeConstants.lvwNone
			CbViewMode = 0
		ElseIf MglPageLast <> 0 And CbViewMode = 1 Then 
			FrmMain.DefInstance.Lv1.View = MSComctlLib.ListViewConstants.lvwReport
			FrmMain.DefInstance.Lv1.Arrange = MSComctlLib.ListArrangeConstants.lvwNone
			CbViewMode = 0
		End If
		
		MglPageLast = PageType
		StatText(3)
		
		Select Case PageType
			Case 0
				'FrmMain.Pages(0).ZOrder 0
				
			Case 1
				CbLv1.Width = VB6.TwipsToPixelsX(3690)
				CbLv2.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(CbLv1.Width))
				CbLv2.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(FrmAgnInfo.DefInstance.Pages(1).Width) - VB6.PixelsToTwipsX(CbLv2.Left) - 15)
				CbLv1.Visible = True
				CbLv2.Visible = True
				
				MgoSlv1.Init(CbLv1)
				MgoSlv1.ItemClear()
				CbLv2.ListItems.Clear()
				CbLv1.ColumnHeaders.Clear()
				CbLv2.ColumnHeaders.Clear()
				CbLv1.ColumnHeaders.Add( ,  , "Station", 1800)
				CbLv1.ColumnHeaders.Add( ,  , "Printers", 1500)
				CbLv2.ColumnHeaders.Add( ,  , "Printer", 2000)
				CbLv2.ColumnHeaders.Add( ,  , "Job ID", 800)
				CbLv2.ColumnHeaders.Add( ,  , "Job Title", 2800)
				CbLv2.ColumnHeaders.Add( ,  , "Status", 1500)
				CbLv2.ColumnHeaders.Add( ,  , "Total", 800)
				CbLv2.ColumnHeaders.Add( ,  , "Printed", 800)
				
				For a = 1 To UniAgents.Count
					UniAgents.Agents(a).AgnAddPage(1)
				Next a
				'FrmMain.Pages(1).ZOrder 0
				
			Case 2
				CbLv2.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(CbLv1.Left))
				CbLv2.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(FrmAgnInfo.DefInstance.Pages(1).Width) - 15)
				CbLv1.Visible = False
				CbLv2.Visible = True
				
				CbLv1.ListItems.Clear()
				CbLv2.ListItems.Clear()
				CbLv1.ColumnHeaders.Clear()
				CbLv2.ColumnHeaders.Clear()
				CbLv2.ColumnHeaders.Add( ,  , "Station", 1500)
				CbLv2.ColumnHeaders.Add( ,  , "Memory Load", 1500)
				CbLv2.ColumnHeaders.Add( ,  , "Physical Total", 1500)
				CbLv2.ColumnHeaders.Add( ,  , "Available", 1300)
				CbLv2.ColumnHeaders.Add( ,  , "Virtual Total", 1500)
				CbLv2.ColumnHeaders.Add( ,  , "Available", 1300)
				CbLv2.ColumnHeaders.Add( ,  , "Pagefile Total", 1500)
				CbLv2.ColumnHeaders.Add( ,  , "Available", 1300)
				
				For a = 1 To UniAgents.Count
					UniAgents.Agents(a).AgnAddPage(2)
				Next a
				'FrmMain.Pages(1).ZOrder 0
		End Select
		Exit Sub
		
ErrInt: 
		ErrLog(Err, "mGui | ConstructPage")
	End Sub
End Module