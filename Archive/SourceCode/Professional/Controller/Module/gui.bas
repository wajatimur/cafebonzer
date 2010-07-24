Attribute VB_Name = "mGui"
Public Im As New VsIconMenu
Public MgoInpt As New VsInput
Public MgoSlv1 As New clsSpecial
Public MglPageLast As Long

Private CbLv1 As ListView
Private CbLv2 As ListView

Public Enum eFrmMetric
    [Pos X] = 1
    [Pos Y] = 2
    [Size X] = 3
    [Size Y] = 4
    [State] = 5
End Enum

Sub StatText(Optional Panel As Integer = 0, Optional Text As String = "")
    FrmMain.MainSbar.Panels(Panel).Text = Text
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ComboBox Add With Trim
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub CbAddEx(Item, Cbox As ComboBox)
    For a = 0 To Cbox.ListCount - 1
        If Trim(Cbox.List(a)) = Trim(Item) Then Exit Sub
    Next a
    Cbox.AddItem Item
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Tutup form dan bebaskan resource
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub CloseFrm(AnyFrm As Form)
    AnyFrm.Hide
    Unload AnyFrm
    Set AnyFrm = Nothing
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Benarkan nombor sahaja
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub NumOnly(TextControl As TextBox)
    Dim Style As Long
    
    Style = GetWindowLong(TextControl.hwnd, GWL_STYLE)
    SetWindowLong TextControl.hwnd, GWL_STYLE, Style Or ES_NUMBER
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Simpan info metric bagi form
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub CbFrmMetricSave(FormName As Form)
    SetSimpan FormName.Name & "saizx", FormName.Width
    SetSimpan FormName.Name & "saizy", FormName.Height
    SetSimpan FormName.Name & "posx", FormName.Left
    SetSimpan FormName.Name & "posY", FormName.Top
    SetSimpan FormName.Name & "state", FormName.WindowState
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Load metric info
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub CbFrmMetricLoad(FormName As Form)
    FormName.Top = CbFrmMetricGet(FormName, [Pos X])
    FormName.Left = CbFrmMetricGet(FormName, [Pos X])
    FormName.Width = CbFrmMetricGet(FormName, [Size X])
    FormName.Height = CbFrmMetricGet(FormName, [Size Y])
    FormName.WindowState = CbFrmMetricGet(FormName, [State])
End Sub

Public Function CbFrmMetricGet(FormName As Form, MetricType As eFrmMetric) As Long
    Dim mtrc As Long
    
    Select Case MetricType
        Case 1: mtrc = SetAmbil(FormName.Name & "posx", 0)
        Case 2: mtrc = SetAmbil(FormName.Name & "posy", 0)
        Case 3: mtrc = SetAmbil(FormName.Name & "saizx", 11460)
        Case 4: mtrc = SetAmbil(FormName.Name & "saizy", 8640)
        Case 5: mtrc = SetAmbil(FormName.Name & "state", 0)
    End Select
    CbFrmMetricGet = mtrc
End Function

Public Sub MoveFrm(hwnd As Long)
    ReleaseCapture
    lret = SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Public Sub PutOnTop(hwnd As Long)
    Dim i As Long
    i = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_WNDFLAGS)
End Sub

Public Sub ConstructPage(PageType As Long)
'On Error GoTo ErrInt
    
    Set CbLv1 = FrmMain.DynaLv(0)
    Set CbLv2 = FrmMain.DynaLv(1)
    
    If PageType = 0 And CbViewMode = 0 Then
        FrmMain.Lv1.View = lvwIcon
        FrmMain.Lv1.Arrange = lvwNone
        CbViewMode = 1
    ElseIf PageType = 0 And CbViewMode = 1 Then
        FrmMain.Lv1.View = lvwReport
        FrmMain.Lv1.Arrange = lvwNone
        CbViewMode = 0
    ElseIf MglPageLast <> 0 And CbViewMode = 1 Then
        FrmMain.Lv1.View = lvwReport
        FrmMain.Lv1.Arrange = lvwNone
        CbViewMode = 0
    End If
    
    MglPageLast = PageType
    StatText 3
    
    Select Case PageType
        Case 0
            FrmMain.Pages(0).ZOrder 0
            
        Case 1
            CbLv1.Width = 3690
            CbLv2.Left = CbLv1.Width
            CbLv2.Width = FrmMain.Pages(1).Width - CbLv2.Left - 15
            CbLv1.Visible = True
            CbLv2.Visible = True
            
            MgoSlv1.Init CbLv1
            MgoSlv1.ItemClear
            CbLv2.ListItems.Clear
            CbLv1.ColumnHeaders.Clear
            CbLv2.ColumnHeaders.Clear
            CbLv1.ColumnHeaders.Add , , "Station", 1800
            CbLv1.ColumnHeaders.Add , , "Printers", 1500
            CbLv2.ColumnHeaders.Add , , "Printer", 2000
            CbLv2.ColumnHeaders.Add , , "Job ID", 800
            CbLv2.ColumnHeaders.Add , , "Job Title", 2800
            CbLv2.ColumnHeaders.Add , , "Status", 1500
            CbLv2.ColumnHeaders.Add , , "Total", 800
            CbLv2.ColumnHeaders.Add , , "Printed", 800
            
            For a = 1 To AgentCount
                UniAgents(a).AgnAddPage 1
            Next a
            FrmMain.Pages(1).ZOrder 0
            
        Case 2
            CbLv2.Left = CbLv1.Left
            CbLv2.Width = FrmMain.Pages(1).Width - 15
            CbLv1.Visible = False
            CbLv2.Visible = True
            
            CbLv1.ListItems.Clear
            CbLv2.ListItems.Clear
            CbLv1.ColumnHeaders.Clear
            CbLv2.ColumnHeaders.Clear
            CbLv2.ColumnHeaders.Add , , "Station", 1500
            CbLv2.ColumnHeaders.Add , , "Memory Load", 1500
            CbLv2.ColumnHeaders.Add , , "Physical Total", 1500
            CbLv2.ColumnHeaders.Add , , "Available", 1300
            CbLv2.ColumnHeaders.Add , , "Virtual Total", 1500
            CbLv2.ColumnHeaders.Add , , "Available", 1300
            CbLv2.ColumnHeaders.Add , , "Pagefile Total", 1500
            CbLv2.ColumnHeaders.Add , , "Available", 1300
            
            For a = 1 To AgentCount
                UniAgents(a).AgnAddPage 2
            Next a
            FrmMain.Pages(1).ZOrder 0
    End Select
Exit Sub

ErrInt:
    ErrLog Err, "mGui | ConstructPage"
End Sub


Public Sub InfoJobsEnum(AgentName, PrinterName)
On Error GoTo ErrInt
    Dim uA As clsAgent, oeJobs As clsAgInfoPrinterJob
    Dim tItm As ListItem
    
    CbLv2.ListItems.Clear
    If PrinterName = "No Printer" Then Exit Sub
    If AgentName = "" Then Exit Sub
    
    If PrinterName = "All Printer" Then
        Set uA = UniAgents.Agents(AgentName)
        For a = 1 To uA.AgentInfo.PrintersCount
            For c = 1 To uA.AgentInfo.Printers(a).JobsCount
                Set oeJobs = uA.AgentInfo.Printers(a).Jobs(c)
                Set tItm = CbLv2.ListItems.Add(, , oeJobs.PrinterName, , "paper")
                tItm.SubItems(1) = oeJobs.JobId
                tItm.SubItems(2) = oeJobs.Document
                tItm.SubItems(3) = oeJobs.Status
                tItm.SubItems(4) = oeJobs.TotalPages
                tItm.SubItems(5) = oeJobs.PagePrinted
            Next c
        Next a
    Else
        Set uA = UniAgents.Agents(AgentName)
        For c = 1 To uA.AgentInfo.Printers(PrinterName).JobsCount
            Set oeJobs = uA.AgentInfo.Printers(PrinterName).Jobs(c)
            Set tItm = CbLv2.ListItems.Add(, , oeJobs.PrinterName, , "paper")
            tItm.SubItems(1) = oeJobs.JobId
            tItm.SubItems(2) = oeJobs.Document
            tItm.SubItems(3) = oeJobs.Status
            tItm.SubItems(4) = oeJobs.TotalPages
            tItm.SubItems(5) = oeJobs.PagePrinted
        Next c
    End If
Exit Sub

ErrInt:
    ErrLog Err, "mGui | InfoJobsEnum"
End Sub


Public Sub LoadIconic()
    Im.Attach FrmMain.hwnd
    With Im
        .HighlightStyle = ECPHighlightStyleGradient
        .ImageList = FrmMain.imglist
        .IconIndex(FrmMain.Menu1Sub1(0).Caption) = FrmMain.imglist.ListImages("penalaan").Index - 1
        .IconIndex(FrmMain.Menu1Sub1(2).Caption) = FrmMain.imglist.ListImages("logoff").Index - 1
        .IconIndex(FrmMain.Menu1Sub1(3).Caption) = FrmMain.imglist.ListImages("power").Index - 1
        .IconIndex(FrmMain.Menu3Bcst.Caption) = FrmMain.imglist.ListImages("broad").Index - 1
        .IconIndex(FrmMain.Menu3BcstSub1(0).Caption) = FrmMain.imglist.ListImages("mesej").Index - 1
        .IconIndex(FrmMain.Menu3BcstSub1(1).Caption) = FrmMain.imglist.ListImages("hoi").Index - 1
        .IconIndex(FrmMain.Menu3CtlLock(0).Caption) = FrmMain.imglist.ListImages("lock").Index - 1
        .IconIndex(FrmMain.Menu3CtlLock(1).Caption) = FrmMain.imglist.ListImages("lock").Index - 1
        .IconIndex(FrmMain.Menu3CtlLock(2).Caption) = FrmMain.imglist.ListImages("kuncibuka").Index - 1
        .IconIndex(FrmMain.Menu3CtlWinexit(0).Caption) = FrmMain.imglist.ListImages("off").Index - 1
        .IconIndex(FrmMain.Menu3CtlWinexit(2).Caption) = FrmMain.imglist.ListImages("boot").Index - 1
        .IconIndex(FrmMain.MenuToolsSub(0).Caption) = FrmMain.imglist.ListImages("gift").Index - 1
        .IconIndex(FrmMain.MenuToolsSub(1).Caption) = FrmMain.imglist.ListImages("graft").Index - 1
        .IconIndex(FrmMain.MenuInfoSub(0).Caption) = FrmMain.imglist.ListImages("help").Index - 1
        .IconIndex(FrmMain.MenuInfoSub(2).Caption) = FrmMain.imglist.ListImages("info").Index - 1
        
        .IconIndex(FrmMain.pmenu1flog.Caption) = FrmMain.imglist.ListImages("jalan1").Index - 1
        .IconIndex(FrmMain.pmenu1cancel.Caption) = FrmMain.imglist.ListImages("no").Index - 1
        .IconIndex(FrmMain.pmenu1trans.Caption) = FrmMain.imglist.ListImages("transfer").Index - 1
        .IconIndex(FrmMain.pmenu1terminal.Caption) = FrmMain.imglist.ListImages("term").Index - 1
        .IconIndex(FrmMain.pmenu1cln.Caption) = FrmMain.imglist.ListImages("cleaning").Index - 1
        .IconIndex(FrmMain.pmenu1ctl.Caption) = FrmMain.imglist.ListImages("cpu").Index - 1
        .IconIndex(FrmMain.pmenu1ctlsub(0).Caption) = FrmMain.imglist.ListImages("lock").Index - 1
        .IconIndex(FrmMain.pmenu1ctlsub(1).Caption) = FrmMain.imglist.ListImages("kuncibuka").Index - 1
        .IconIndex(FrmMain.pmenu1ctlsub(2).Caption) = FrmMain.imglist.ListImages("boot").Index - 1
        .IconIndex(FrmMain.pmenu1ctlsub(3).Caption) = FrmMain.imglist.ListImages("off").Index - 1
    End With
End Sub

Public Sub UnloadIconic()
    Im.Detach
End Sub

