Attribute VB_Name = "MdlAgents"
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : MdlAgents
'    Project    : CafeBonzer
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Public lSock As Long

Public Enum EnAgentSymbol
    [User Online] = 0
    [User Offline] = 1
    [User Prepaid] = 2
    [User Prepaid Ended] = 3
    [Terminal Online] = 10
    [Terminal Offline] = 11
    [Terminal Lock] = 12
    [Terminal Cleaning] = 13
End Enum


''=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'' Update Main Panel
''=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub UpdatePanel(KeyName)
    Dim CuA As ClsAgent, CInfoListView As ListView, CInfoListView1 As ListView
    
    Set CInfoListView = FrmMain.InfoListView
    Set CInfoListView1 = FrmMain.InfoListView1
    
    CInfoListView1.ListItems("UNUSED").SubItems(1) = UniAgents.CountStatus(UnUsed)
    CInfoListView1.ListItems("TOTALCOUNT").SubItems(1) = UniAgents.Count
    CInfoListView1.ListItems("CONNECTEDCOUNT").SubItems(1) = UniAgents.CountStatus(Connected)
    
    If UniAgents.Count > 0 And KeyName <> "" Then
        Set CuA = UniAgents(KeyName)
        
        CInfoListView.ListItems("CONNECTION").SubItems(1) = CuA.AgentConnection
        CInfoListView.ListItems("LOCK").SubItems(1) = CuA.AgentLock
        CInfoListView.ListItems("STATUS").SubItems(1) = CuA.AgentStatus
        If CuA.AgentStatus = VS(1, 1) Then
            CInfoListView.ListItems("CURRENTUSAGE").SubItems(1) = Crnc & Format(CuA.CusGetUsage, "#0.00")
        Else
            CInfoListView.ListItems("CURRENTUSAGE").SubItems(1) = ""
        End If
        CInfoListView.ListItems("TERMCONNECTED").SubItems(1) = CuA.AgentConnected
        CInfoListView.ListItems("IPADDRESS").SubItems(1) = CuA.AgentIPAdd
        CInfoListView.ListItems("MACADDRESS").SubItems(1) = CuA.AgentMAC
        Set CuA = Nothing
    Else
        CInfoListView.ListItems("TERMCONNECTED").SubItems(1) = ""
        CInfoListView.ListItems("IPADDRESS").SubItems(1) = ""
        CInfoListView.ListItems("MACADDRESS").SubItems(1) = ""
        CInfoListView.ListItems("CURRENTUSAGE").SubItems(1) = ""
    End If
End Sub

''=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'' Update Stat
''=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub UpdateStat(sItem As ListItem)
    Dim sCustomerName As String, sCustomerTimeIn As String, sCustomerTimeOut As String
    Dim sAgentStatus As String, CuA As ClsAgent
    
    If sItem Is Nothing Then
        StatText 3
        Exit Sub
    End If
    
    Set CuA = UniAgents(sItem.Text)
    sAgentStatus = CuA.AgentStatus
    sCustomerName = CuA.CustomerName
    sCustomerTimeIn = CuA.CustomerTimeIn
    sCustomerTimeOut = CuA.CustomerTimeOut
    
    If sAgentStatus = VS(1, 2) Then
        StatText 3, ""
        Exit Sub
    End If
    
    If sCustomerTimeOut = VS(0, 1) Then
        StatText 3, Crnc & " " & CuA.CusGetUsage & "  (" & CuA.CusGetTimeUse & ")" & " - " & sCustomerName
    Else
        If sAgentStatus = VS(1, 1) Then
            StatText 3, "Time Left : " & CuA.CusGetTimeLeft & " - " & sCustomerName
        Else
            StatText 3, "Time End - " & sCustomerName
        End If
    End If

    Set CuA = Nothing
End Sub

''=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'' Selected Agent
''=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Function AgentSel() As ClsAgent
    Set AgentSel = UniAgents(SelText)
End Function

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Dapatkan index dalam listview melalui SckID
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function AgentGetIndex(SockIndex)
    AgentGetIndex = 0
    For g = 1 To FrmMain.ListView.ListItems.Count
        If FrmMain.ListView.ListItems(g).Tag = CStr(SockIndex) Then AgentGetIndex = g: Exit Function
    Next g
End Function

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Dapatkan index dalam listview melalui Nama Terminal
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function AgentGetIndexB(Nama)
    AgentGetIndexB = 0
    For g = 1 To FrmMain.ListView.ListItems.Count
        If FrmMain.ListView.ListItems(g).Text = CStr(Nama) Then AgentGetIndexB = g: Exit Function
    Next g
End Function
