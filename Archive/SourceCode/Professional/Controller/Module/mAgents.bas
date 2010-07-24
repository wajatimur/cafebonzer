Attribute VB_Name = "mAgents"
Public UniAgents As New clsAgents
Public lSock As Long

Public Type udResource
    MemLoad As Long
    PhyTotal As Long
    PhyAvail As Long
    VirTotal As Long
    VirAvail As Long
    PageTotal As Long
    PageAvail As Long
End Type

Public Type udPrinterJob
    PrinterName As String * 32
    JobId As String * 16
    Status As String * 32
    Document As String * 64
    PagePrinted As Long
    TotalPages As Long
End Type

Public Type udRecoverInfo
    TerminateOK As Boolean
    AgentName As String * 32
    AgentStatus As String * 16
    CustomerName As String * 32
    CustomerType As String * 32
    CustomerTimeIn As String * 32
    CustomerTimeOut As String * 32
    CustomerFlag As String * 16
    JobsCount As Long
    Jobs() As udPrinterJob
End Type


''=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'' Selected Agent
''=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Function SelAgn() As clsAgent
    Set SelAgn = UniAgents(SelText)
End Function

''=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'' Parse Network Command
''=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub ParseCmd(DataRcv As String, SockIndex As Long)
    Dim s_CmdName As String, s_CmdVal As String
    Dim DataString As String, uA As clsAgent
    Dim xPos As Integer
    Dim yPos As Integer
    
    xPos = 0
    Do Until xPos = Len(DataRcv)
        xPos = InStr(xPos + 1, DataRcv, "/")
        yPos = InStr(xPos + 1, DataRcv, "/")
        
        If xPos = 0 Then Exit Do
        If yPos = 0 Then
            DataString = Mid(DataRcv, xPos)
            xPos = Len(DataRcv)
        Else
            DataString = Mid(DataRcv, xPos, yPos - xPos)
        End If
    Loop
    
    s_CmdName = CmdName(DataString)
    s_CmdVal = CmdValue(DataString)
    
    If s_CmdName = "cert" Then
        UniAgents.AgentCert SockIndex, s_CmdVal
    Else
        Set uA = UniAgents.AgentsByIndex(SockIndex)
        Select Case s_CmdName
            Case Is = "hoi"
                uA.NetPingReset
            Case Is = "mesej"
                Call FnMesej(DataString)
        End Select
        If Left(s_CmdName, 4) = "info" Then uA.AgnInfoParse DataString
    End If
End Sub


Public Sub UpdatePanel(KeyName)
    Dim uA As clsAgent
    
    FrmMain.SpgInfoLblB(0) = UniAgents.Count
    FrmMain.SpgInfoLblB(1) = UniAgents.CountStatus(UnUsed)
    
    If UniAgents.Count > 0 And KeyName <> "" Then
        Set uA = UniAgents(KeyName)
        FrmMain.SpgInfoLblD(0) = uA.AgentConnected
        FrmMain.SpgInfoLblD(1) = uA.AgentIPAdd
        FrmMain.SpgInfoLblD(2) = uA.AgentMAC
        If uA.AgentStatus = VS(3) Then
            FrmMain.SpgInfoLblD(3) = Crnc & Format(uA.CusGetPrice, "#0.00")
        Else
            FrmMain.SpgInfoLblD(3) = ""
        End If
        Set uA = Nothing
    Else
        FrmMain.SpgInfoLblD(0) = ""
        FrmMain.SpgInfoLblD(1) = ""
        FrmMain.SpgInfoLblD(2) = ""
        FrmMain.SpgInfoLblD(3) = ""
    End If
End Sub


Public Sub UpdateStat(sItem As ListItem)
    Dim sCustomerName As String, sCustomerTimeIn As String, sCustomerTimeOut As String
    Dim sAgentStatus As String, uA As clsAgent
    
    If sItem Is Nothing Then
        StatText 3
        Exit Sub
    End If
    
    Set uA = UniAgents(sItem.Text)
    sAgentStatus = uA.AgentStatus
    sCustomerName = uA.CustomerName
    sCustomerTimeIn = uA.CustomerTimeIn
    sCustomerTimeOut = uA.CustomerTimeOut
    
    If sAgentStatus = VS(4) Then
        StatText 3, ""
        Exit Sub
    End If
    
    'untuk prepaid
    If sCustomerTimeOut <> VS(1) And sAgentStatus = VS(3) Then
        StatText 3, "Time Left : " & uA.CusGetTimeLeft & " - " & sCustomerName
        Exit Sub
        Set uA = Nothing
    End If
    
    'jika prepaid telah tamat..
    If sCustomerTimeOut <> VS(1) And sAgentStatus = VS(5) Then
        StatText 3, "Time End - " & sCustomerName
        Exit Sub
        Set uA = Nothing
    End If
    
    'untuk pay as you go
    If sCustomerTimeOut = VS(1) And sAgentStatus = VS(3) Then
        StatText 3, Crnc & " " & uA.CusGetPrice & "  (" & uA.CusGetTimeUse & ")" & " - " & sCustomerName
        Exit Sub
        Set uA = Nothing
    End If
End Sub


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Jumlah Agent
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function AgentCount()
    AgentCount = FrmMain.Lv1.ListItems.Count
End Function

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Tukar ikon bagi agent dalam listview
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub AgentIcon(SockIndex, iconName, Optional SmallIconOnly As Boolean = False)
    For g = 1 To FrmMain.Lv1.ListItems.Count
        If FrmMain.Lv1.ListItems(g).Tag = CStr(SockIndex) Then
            FrmMain.Lv1.ListItems(g).SmallIcon = iconName
            If SmallIconOnly = False Then FrmMain.Lv1.ListItems(g).Icon = iconName
            Exit Sub
        End If
    Next g
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Tukar warna bagi agent dalam listview
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub AgentColour(AgentIndex, Warna)
    Dim lItm As ListItem
    Set lItm = FrmMain.Lv1.ListItems(AgentIndex)
    
    lItm.ForeColor = Warna
    For B = 1 To 7
        lItm.ListSubItems.Item(B).ForeColor = Warna
    Next B
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Dapatkan index dalam listview melalui SckID
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function GetAgentIndex(SockIndex)
    GetAgentIndex = 0
    For g = 1 To FrmMain.Lv1.ListItems.Count
        If FrmMain.Lv1.ListItems(g).Tag = CStr(SockIndex) Then GetAgentIndex = g: Exit Function
    Next g
End Function

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Dapatkan index dalam listview melalui Nama Terminal
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function GetAgentIndexB(Nama)
    GetAgentIndexB = 0
    For g = 1 To FrmMain.Lv1.ListItems.Count
        If FrmMain.Lv1.ListItems(g).Text = CStr(Nama) Then GetAgentIndexB = g: Exit Function
    Next g
End Function

''=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'' Prosedur untuk recovery
''=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub Recovery(Nama)
On Error GoTo ErrInt
    Dim uA As clsAgent
    Dim RecPath
    
    Set uA = UniAgents(Nama)
    RecPath = App.Path & "\Data\recstate.r"

    If Left(INIambil(RecPath, uA.AgentName, "terminateok"), 2) = "no" Then
        uA.CustomerName = INIambil(RecPath, uA.AgentName, 1)
        uA.CustomerTimeIn = INIambil(RecPath, uA.AgentName, 2)
        uA.CustomerTimeOut = INIambil(RecPath, uA.AgentName, 3)
        uA.CustomerType = INIambil(RecPath, uA.AgentName, 4)
        uA.AgentStatus = INIambil(RecPath, uA.AgentName, 5)
        'uA.ItemMain.SubItems(6) = INIambil(RecPath, uA.AgentName, 6)
        'uA.ItemMain.SubItems(7) = INIambil(RecPath, uA.AgentName, 7)
        
        uA.CustomerFlag = INIambil(RecPath, uA.AgentName, 90)
        
        If uA.AgentStatus = VS(5) Then
            uA.AgentSmallIcon = "tamat"
            uA.AgentIcon = "tamat"
            uA.AgentForeColor = vbBlue
        Else
            uA.AgentSmallIcon = "jalan"
            uA.AgentIcon = "jalan"
            uA.AgentForeColor = &H8000&
        End If
    End If
Exit Sub

ErrInt:
    ErrLog Err, "Recovery"
End Sub


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Prosedur untuk rekod recovery data
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub RecoveryGo(Item As ListItem)
On Error GoTo ErrInt
    Dim uA As clsAgent
    Dim RecPath: RecPath = App.Path & "\Data\recstate.r"
    Set uA = UniAgents(Item.Text)
    
    INIsimpan RecPath, uA.AgentName, 90, uA.CustomerFlag
    INIsimpan RecPath, uA.AgentName, 1, uA.CustomerName
    INIsimpan RecPath, uA.AgentName, 2, uA.CustomerTimeIn
    INIsimpan RecPath, uA.AgentName, 3, uA.CustomerTimeOut
    INIsimpan RecPath, uA.AgentName, 4, uA.CustomerType
    INIsimpan RecPath, uA.AgentName, 5, uA.AgentStatus
    'INIsimpan RecPath, uA.AgentName, 6, uA.ItemMain.SubItems(6)
    'INIsimpan RecPath, uA.AgentName, 7, uA.ItemMain.SubItems(7)
    
    If uA.AgentStatus = VS(3) Or uA.AgentStatus = VS(5) Then
        INIsimpan RecPath, uA.AgentName, "terminateok", "no"
    Else
        INIsimpan RecPath, uA.AgentName, "terminateok", "ok"
    End If
    Set uA = Nothing
Exit Sub

ErrInt:
    ErrLog Err, "RecoveryGo"
End Sub
