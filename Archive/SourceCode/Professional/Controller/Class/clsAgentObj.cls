VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAgent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================
'  Aplication codename : CafeBonzer
'  Programmer          : Azri Jamil a.k.a wajatimur
'  Class Name          : Agent Object
'  Description         : Class for total control of agent
'==================================================================
Private ObjRoot As clsAgents
Private ObjInfo As New clsAgInfo
Private ObjSck As Socket

Private b_netMode As Boolean
Private i_sockIndex As Long
Private i_NetPingCount As Integer

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
Private l_CusPrinted As Long

Private LVm As ListView
Private LVd1 As ListView
Private LVd2 As ListView
Private LVmItem As ListItem
Private LVdItem1 As ListItem
Private LVdItem2 As ListItem


'=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*==*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=
'  MAIN PROCEDURE
'
'
'=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*==*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=
Private Sub Class_Initialize()
    Set LVm = FrmMain.Lv1
    Set LVd1 = FrmMain.DynaLv(0)
    Set LVd2 = FrmMain.DynaLv(1)
    
    ObjInfo.Init Me
    b_agCert = False
End Sub


'=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*==*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=
'  PUBLIC SUBS & FUNCTIONS
'
'
'=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*==*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Send Data
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Sub NetSend(Data As String)
    Send i_sockIndex, Data
End Sub
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Ping Client
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Function NetPing() As Integer
    Send i_sockIndex, "//hey"
    i_NetPingCount = i_NetPingCount + 1
    NetPing = i_NetPingCount
End Function
Public Sub NetPingReset()
    i_NetPingCount = 0
End Sub


'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
'
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Function Root() As clsAgents
    Set Root = ObjRoot
End Function
Public Function ItemMain() As ListItem
    Set ItemMain = LVmItem
End Function
Public Function ItemDyna1() As ListView
    Set ItemDyna1 = LVdItem1
End Function
Public Function ItemDyna2() As ListItem
    Set ItemDyna2 = LVdItem2
End Function


'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Initialize Agent
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Sub AgnInit(NewSocket As Socket)
    Set ObjSck = NewSocket
    
    b_netMode = True
    i_sockIndex = ObjSck.Index
    s_agIP = ObjSck.PeerAddress
End Sub

'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Reject Agent
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Sub AgnInitReject()
    ObjSck.Cleanup
    Set clsAgent = Nothing
End Sub

'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Initialize Certification
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Function AgnInitCert(sCommand As String, Root As clsAgents) As Boolean
    s_agVersion = SubVal(sCommand, "version")
    s_agName = SubVal(sCommand, "name")
    s_agName = zGetFreeName(s_agName)
    
    If s_agVersion >= "1.7.75" Then
        Set ObjRoot = Root
        Send i_sockIndex, "//cert:" & SubBuild("granted", 1)
        b_agCert = True
        AgnInitCert = True
        Call zAgentAdd(s_agName)
    Else
        Send i_sockIndex, "//cert:" & SubBuild("granted", 0) & SubBuild("info", "Version too old, please update !")
        AgnInitCert = False
    End If
End Function

Public Sub AgnRecover()
    Dim lRecordMax As Long
    Dim tIrec As udRecoverInfo
    
    iRecover.AgentName = s_agName
    iRecover.AgentStatus = s_agStatus
    iRecover.CustomerFlag = s_CusFlag
    iRecover.CustomerName = s_CusName
    iRecover.CustomerTimeIn = s_CusTimeIn
    iRecover.CustomerTimeOut = s_CusTimeOut
    iRecover.CustomerType = s_CusType
    
    If s_agStatus = VS(3) Or s_agStatus = VS(5) Then
        iRecover.TerminateOK = False
    Else
        iRecover.TerminateOK = True
    End If
    
    Open CbPathDatRecv For Random Lock Read Write As #1 Len = Len(iRecover)
    lRecordMax = LOF(1) \ Len(iRecover)
    
    If lRecordMax > 0 Then
        For j = 1 To lRecordMax
            Get #1, j, tIrec
            If tIrec.AgentName = s_agName Then Put #1, j, iRecover
        Next j
    Else
        'Put #1, , iRecover
    End If
    Close #1
End Sub

Public Sub AgnTransfer(NewLvIndex As Integer)
    Dim nLv As ListItem
    Dim uA As clsAgent
    
    Set nLv = FrmMain.Lv1.ListItems(NewLvIndex)
    Set uA = UniAgents(nLv.Text)
    'Set LVmItem = nLv
    'Send i_sockIndex, "//kunci:1"
    Send i_sockIndex, "//logout"
    
   'copy data
    uA.CustomerName = s_CusName
    uA.CustomerTimeIn = s_CusTimeIn
    uA.CustomerTimeOut = s_CusTimeOut
    uA.AgentStatus = s_agStatus
    uA.CustomerType = s_CusType
    uA.CustomerFlag = s_CusFlag
    uA.ItemMain.SubItems(6) = LVmItem.SubItems(6)
    uA.ItemMain.SubItems(7) = LVmItem.SubItems(7)
    
   'copy appearance
    uA.AgentForeColor = LVmItem.ForeColor
    uA.AgentSmallIcon = LVmItem.SmallIcon
    uA.AgentIcon = LVmItem.Icon
    'uA.NetSend "//kunci:0"
    uA.NetSend "//login"
    
   'reset current agent
    Call CusStop
End Sub


'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Remove Agent
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Sub AgnRemove()
'On Error GoTo ErrInt
    Select Case MglPageLast
        Case 1
            MgoSlv1.ItemRemove LVmItem.Key
        Case 2
            LVd2.ListItems.Remove LVmItem.Key
    End Select
    
    LVm.ListItems.Remove LVmItem.Index
    ObjSck.Cleanup
    Set clsAgent = Nothing
Exit Sub
ErrInt:
    ErrLog Err, "Class Agent | AgnRemove"
    Resume Next
End Sub

'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Parse Information From Agent
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Sub AgnInfoParse(DataCommand As String)
    Dim sCmdName As String, sCmdVal As String
    sCmdName = CmdName(DataCommand)
    sCmdVal = CmdValue(DataCommand)
    
    If sCmdName = "info.me" Then
        Dim eXp As Boolean
        eXp = (sCmdVal = "unlock") Or (sCmdVal = "cleanok")
        
        If eXp And s_agStatus = VS(4) Then mAgents.AgentIcon i_sockIndex, "aktif1"
        If eXp And s_agStatus = VS(3) Then mAgents.AgentIcon i_sockIndex, "jalan"
        If eXp And s_agStatus = VS(5) Then mAgents.AgentIcon i_sockIndex, "tamat"
        
        If sCmdVal = "lock" Then mAgents.AgentIcon i_sockIndex, "lock"
        If sCmdVal = "block" Then mAgents.AgentIcon i_sockIndex, "mouse"
    ElseIf sCmdName = "info.net" Then
        s_agMAC = SubVal(sCmdVal, "mac")
    Else
        AgentInfo.Parse DataCommand, Me
    End If
End Sub

Public Sub AgnAddPage(InfoType)
    Dim tItm As ListItem
    Select Case InfoType
        Case 1
            Set tItm = MgoSlv1.ItemAdd(s_agName, s_agName, "dump")
            tItm.Tag = i_sockIndex
            Set LVdItem1 = tItm
            If ObjInfo.PrintersCount > 0 Then
                For a = 1 To ObjInfo.PrintersCount
                    MgoSlv1.MatrixAdd tItm, ObjInfo.Printers(a).DeviceName
                Next a
                tItm.SubItems(1) = "All Printer"
            Else
                tItm.SubItems(1) = "No Printer"
            End If
            tItm.ListSubItems(1).ReportIcon = "printer"
            
            NetSend "//mon.switch:1"
        Case 2
            Set tItm = LVd2.ListItems.Add(, s_agName, s_agName, , "dump")
            tItm.Tag = i_sockIndex
            Set LVdItem2 = tItm
            NetSend "//mon.switch:2"
            tItm.SubItems(1) = AgentInfo.MemLoad & " %"
            For a = 2 To 7
                tItm.SubItems(a) = 0 & " Mb"
            Next a
    End Select
End Sub


'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Check Expires Customer for Prepaid & Fixed Time
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Function CusCheckExpired() As Boolean
    CusCheckExpired = False
    If s_agStatus = VS(3) And s_CusTimeOut <= Now Then
        'set ikon ke tamat
        Me.AgentSmallIcon = "tamat"
        'hantar ke agent supaya kunci
        Call NetSend("//kunci:1")
        'masukkan ke dalam lv utama
        Me.AgentStatus = VS(5)
        'tukar warna
        Me.AgentForeColor = vbBlue
        CusCheckExpired = True
    End If
End Function

Public Sub CusStatusUpdate()
 '{ updatekan harga dan hantar ke setiap terminal }'
    Cond1 = s_CusTimeOut = VS(1) And s_agStatus = VS(3)  'kira jumlah harga
    Cond2 = s_CusTimeOut <> VS(1) And s_agStatus = VS(3) 'kira masa tinggal
    Cond3 = s_CusTimeOut <> VS(1) And s_agStatus = VS(5) 'tamat digunakan
    If Cond1 Then
        hrg = CusGetPrice
        NetSend "//harga:" & hrg
    ElseIf Cond2 Then
        mt = CusGetTimeLeft
        NetSend "//tiker: [ Time Left : " & mt & " ] "
    ElseIf Cond3 Then
        NetSend "//tiker: [ Thank You ] "
    End If
    
 '{ Update jumlah page printed }'
    'If MglPageLast = 1 Then
    '    If AgentInfo.Printers(1).JobsCount > 0 Then ' ?????
    '        Dim tJob As clsAgInfoPrinterJob, l_Printed As Long
    '        For Each tJob In eJobs
    '            l_Printed = l_Printed + tJob.PagePrinted
    '        Next
    '        If l_Printed > l_CusPrinted Then
    '            l_CusPrinted = l_Printed
    '            LVmItem.SubItems(7) = l_CusPrinted
    '        End If
    '    End If
    'End If
End Sub

'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Get Price for Customer
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Function CusGetPrice() As String
    Dim Ai As Integer, pRate As Double, gi01 As Double, gi02 As Double
    
    'ambil harga mengikut skema
    If LCase(s_CusType) = LCase(VS(2)) Then
        'default pricing
        pRate = SetAmbil("harga")
    Else
        'pricing for current scheme
        pRate = uSDBe.DataFind("skema", "skema", "harga", s_CusType)
    End If
    
    'ambil nilai biasa jika harga skema tidak dapat diterima
    If pRate < 0 Then pRate = SetAmbil("harga")
    
    'pengiraan harga..
    gi01 = DateDiff("n", s_CusTimeIn, Now) '+ Ai
    gi02 = (gi01 * pRate) + SetAmbil("hargaex")

    'return value.. easy heh..
    gi02 = GetRoundUpVal(gi02)
    CusGetPrice = Format(gi02, "#0.00")
End Function

'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Get Time Usage for Customer
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Function CusGetTimeUse(Optional InMinute As Boolean = False) As String
    Dim Jam As String, Minit As String
    
    'dapatkan jumlah minit
    Minit = DateDiff("n", s_CusTimeIn, Now)
    If InMinute = True Then CusGetTimeUse = Minit: Exit Function
    
    'dapatkan jumlah jam
    Jam = Minit \ 60
    
    'dapatkan minit selepas jam
    Minit = Minit - (Jam * 60)
    
    'gabungkan kesemua
    CusGetTimeUse = Jam & " " & VS(7) & ", " & Minit & " " & VS(8)
End Function

Public Function CusGetTimeUseEx() As Long
    Dim Minit As Long, Saat As Long
  ' dapatkan jumlah minit
    Minit = DateDiff("n", s_CusTimeIn, Now)
    Saat = DateDiff("s", s_CusTimeIn, Now)
    CusGetTimeUseEx = Saat
End Function

'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Get Time Left for Customer
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Function CusGetTimeLeft()
    Dim s_Jam As String, s_Minit As String
    
    s_Minit = DateDiff("n", Now, s_CusTimeOut)
    s_Jam = s_Minit \ 60
    s_Minit = s_Minit - (s_Jam * 60)
    CusGetTimeLeft = s_Jam & " " & VS(7) & ", " & s_Minit & " " & VS(8)
End Function

'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Start Pay As Your Go
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Sub CusStartPAYG(CusName As String, CusType As String)
    Me.AgentStatus = VS(3)
    Me.CustomerName = CusName
    Me.CustomerType = CusType
    Me.CustomerFlag = "g"
    Me.CustomerTimeIn = Now
    Me.CustomerTimeOut = VS(1)
    
    LVmItem.SubItems(6) = VS(1)
    LVmItem.SubItems(7) = ""
    
    Call zAgentSetRun
    Call NetSend("//login:" & SubBuild("name", CusName))
End Sub

'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Start Prepaid
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Sub CusStartPPAID(CusName As String, CusType As String, PaidPrice As Double, Optional CusSambung As Boolean = False)
    Dim cRate As Double, cRateEx As Double, nVal As Double
    Dim gMin As Integer, gJam As Integer
    Dim MinTmp As String
    
    cRateEx = CDbl(SetAmbil("hargaex")) 'ambil harga extra
    cRate = CDbl(SetAmbil("harga"))     'ambil harga

    Me.AgentStatus = VS(3)
    Me.CustomerName = CusName
    Me.CustomerType = CusType
    Me.CustomerTimeIn = Now

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

    gMin = PaidPrice / cRate                                    'kira jumlah minit..
    Me.CustomerTimeOut = DateAdd("n", gMin, Now)          'bila akan berhenti..
    
    If gMin >= 60 Then
        gJam = gMin \ 60: gMin = gMin - (gJam * 60)
        If gMin = 0 Then MinTmp = "" Else MinTmp = " " & gMin & " " & LCase(VS(8))
        LVmItem.SubItems(6) = gJam & " " & LCase(VS(7)) & MinTmp
    Else
        LVmItem.SubItems(6) = gMin & " " & LCase(VS(8))
    End If
    LVmItem.SubItems(7) = ""
    
    Call zAgentSetRun
    Call NetSend("//kunci:0")
End Sub

'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Start Fixed Time
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Sub CusStartTIME(CusName As String, CusType As String, UseHour As Integer, UseMinute As Integer, Optional CusSambung As Boolean = False)
    Dim gMin As Integer, TmpMin As Integer
    Dim cRate As Double, cRateEx As Double, cPaid As Double
    
    gMin = (UseHour * 60) + UseMinute                               'jumlah masa yang ingin di gunakan dalam minit
    
    Me.AgentStatus = VS(3)
    Me.CustomerName = CusName
    Me.CustomerType = CusType
    Me.CustomerTimeIn = Now
    
    If CusSambung = True Then
        Select Case Left(s_CusFlag, 1)
        Case Is = "p"
            cPaid = CDbl(Mid(s_CusFlag, 2))
            cRate = CDbl(SetAmbil("harga"))
            cRateEx = CDbl(SetAmbil("hargaex"))
            TmpMin = cPaid / cRate
            Me.CustomerFlag = "f" & (TmpMin + gMin)
        Case id = "f"
            Me.CustomerFlag = "f" & (CDbl(Mid(s_CusFlag, 2)) + gMin)
        End Select
    Else
        Me.CustomerFlag = "f" & gMin
    End If
    
    Me.CustomerTimeOut = DateAdd("n", gMin, Now)                        'masukkan masa berhenti ke dalam subitems(3)
    sJam = IIf(UseHour <> 0, UseHour & " " & LCase(VS(7)) & " ", "")
    sMin = IIf(UseMinute <> 0, UseMinute & " " & LCase(VS(8)), "")
    LVmItem.SubItems(6) = sJam & sMin
    LVmItem.SubItems(7) = ""
    
    Call zAgentSetRun
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
    Me.ItemMain.SubItems(6) = ""
    Me.ItemMain.SubItems(7) = ""
    
    Me.AgentIcon = "jalan"
    Me.AgentSmallIcon = "aktif1"
    Me.AgentForeColor = vbBlack
    Me.AgentFontBold = False
    
    FrmMain.MainSbar.Panels(2) = ""
    FrmMain.MainSbar.Panels(3) = ""
End Sub


'=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*==*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=
'  PROPERTY SECTION
'
'
'=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*==*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Agent | Extender Info
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Get AgentInfo() As clsAgInfo
    Set AgentInfo = ObjInfo
End Property
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Agent | Connected Time
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Get AgentConnected() As String
    AgentConnected = s_agConnected
End Property
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Agent | Name
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Let AgentName(nValue As String)
    If nValue <> "" Then
        s_agName = nValue
        LVmItem.Text = s_agName
    End If
End Property
Public Property Get AgentName() As String
    AgentName = s_agName
End Property
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Certification for Agent
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Get AgentCertified() As Boolean
    AgentCertified = b_agCert
End Property
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
'  Agent Socket Index
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Get AgentSockIndex() As Integer
    AgentSockIndex = i_sockIndex
End Property
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Agent IP Address
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Get AgentIPAdd() As String
    AgentIPAdd = s_agIP
End Property
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Agent MAC Address
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Get AgentMAC() As String
    AgentMAC = s_agMAC
End Property
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Agent Small Icon
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Let AgentSmallIcon(iconName As String)
    If iconName <> "" Then LVmItem.SmallIcon = iconName
End Property
Public Property Get AgentSmallIcon() As String
    AgentSmallIcon = LVmItem.SmallIcon
End Property
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Agent Icon
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Let AgentIcon(iconName As String)
    If iconName <> "" Then LVmItem.Icon = iconName
End Property
Public Property Get AgentIcon() As String
    AgentIcon = LVmItem.Icon
End Property
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Agent Fore Colour
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Let AgentForeColor(colour As ColorConstants)
    On Error Resume Next
    LVmItem.ForeColor = colour
    For B = 1 To 7
        LVmItem.ListSubItems(B).ForeColor = colour
    Next B
End Property
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Agent Font Bold
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Let AgentFontBold(nValue As Boolean)
    On Error Resume Next
    LVmItem.Bold = nValue
    For B = 1 To 7
        LVmItem.ListSubItems(B).Bold = nValue
    Next B
End Property
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Agent Status
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Let AgentStatus(nValue As String)
    If s_agStatus = nValue Then Exit Property
    s_agStatus = nValue
    LVmItem.SubItems(1) = s_agStatus
End Property
Public Property Get AgentStatus() As String
    AgentStatus = s_agStatus
End Property

'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Customer Name
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Let CustomerName(nValue As String)
    If s_CusName = nValue Then Exit Property
    s_CusName = nValue
    LVmItem.SubItems(2) = s_CusName
End Property
Public Property Get CustomerName() As String
    CustomerName = s_CusName
End Property
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Customer Time In
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Let CustomerTimeIn(nValue As String)
    If s_CusTimeIn = nValue Then Exit Property
    s_CusTimeIn = nValue
    If s_CusTimeIn = VS(1) Or s_CusTimeIn = "" Then
        LVmItem.SubItems(4) = s_CusTimeIn
    Else
        LVmItem.SubItems(4) = TimeValue(s_CusTimeIn)
    End If
End Property
Public Property Get CustomerTimeIn() As String
    CustomerTimeIn = s_CusTimeIn
End Property
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Customer Time Out
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Let CustomerTimeOut(nValue As String)
    If s_CusTimeOut = nValue Then Exit Property
    s_CusTimeOut = nValue
    If s_CusTimeOut = VS(1) Or s_CusTimeOut = "" Then
        LVmItem.SubItems(5) = s_CusTimeOut
    Else
        LVmItem.SubItems(5) = TimeValue(s_CusTimeOut)
    End If
End Property
Public Property Get CustomerTimeOut() As String
    CustomerTimeOut = s_CusTimeOut
End Property
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Customer Scheme
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Let CustomerType(nValue As String)
    If s_CusType = nValue Then Exit Property
    s_CusType = nValue
    LVmItem.SubItems(3) = s_CusType
End Property
Public Property Get CustomerType() As String
    CustomerType = s_CusType
End Property
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Customer Flag
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Let CustomerFlag(nValue As String)
    If s_CusFlag = nValue Then Exit Property
    s_CusFlag = nValue
    LVmItem.ListSubItems(1).Tag = nValue
End Property
Public Property Get CustomerFlag() As String
    CustomerFlag = s_CusFlag
End Property
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Customer Printed
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Get CustomerPrinted() As Long
    CustomerPrinted = l_CusPrinted
End Property
'=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*==*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=



'=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*==*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=
'  PRIVATE SUBS & FUNCTIONS
'
'
'=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*==*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Add Agent to Listview
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Private Sub zAgentAdd(s_Nama As String)
    Set LVmItem = LVm.ListItems.Add(, s_Nama, s_Nama, "aktif1", "aktif1")
    LVmItem.Tag = i_sockIndex
    
    s_agStatus = VS(4)
    s_agConnected = Time
    LVmItem.SubItems(1) = s_agStatus
    LVm.Sorted = True: LVm.Sorted = False
End Sub

'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Set Agent View Run
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Private Sub zAgentSetRun()
    LVmItem.SmallIcon = "jalan"
    LVmItem.Icon = "jalan"
    Me.AgentForeColor = &H8000&
End Sub

'!!Unused
Private Sub zAgentSetUnUsed(Optional LockAgent As Boolean = False)
    LVmItem.ForeColor = vbBlack
    LVmItem.ListSubItems(1).Tag = ""
    For s = 1 To 7
        If LVmItem.SubItems(s) <> "" Then LVmItem.ListSubItems(s).ForeColor = vbBlack
        LVmItem.SubItems(s) = ""
    Next s
    Me.AgentStatus = VS(4)
    
    LVmItem.SmallIcon = "aktif1"
    LVmItem.Icon = "aktif1"
    
    If LockAgent = True Then Send i_sockIndex, "//kunci:1"
End Sub

Private Function zGetFreeName(Name As String) As String
Recheck:
    zGetFreeName = Name
    If CekDuplicate(Name) = True Then
        zGetFreeName = CArrName(Name)
        GoTo Recheck
    End If
End Function
'=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*==*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=
