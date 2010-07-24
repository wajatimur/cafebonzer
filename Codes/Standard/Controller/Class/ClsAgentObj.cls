VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsAgent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : clsAgent
'    Project    : CafeBonzer
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
'==================================================================
'  Aplication codename : CafeBonzer
'  Programmer          : Azri Jamil a.k.a wajatimur
'  Class Name          : Agent Object
'  Description         : Class for total control of agent
'==================================================================
Private ObjCommand As ClsCommand
Private ObjSck As Socket

Private BlnNetMode As Boolean
Private LngSocketIndex As Long
Private IntNetPingCount As Integer

Private BlnAgCertified As Boolean
Private StrAgVersion As String
Private StrAgName As String
Private StrAgMAC As String
Private StrAgIP As String
Private StrAgStatus As String
Private StrAgConnected As String
Private StrAgConnection As String
Private StrAgLock As String

Private StrCusId As String
Private StrCusName As String
Private StrCusType As String
Private StrCusFlag As String
Private StrCusTimeIn As String
Private StrCusTimeOut As String
Private StrCusUsage As String
Private LngCusPrinted As Long

Private ClistView As ListView
Private ClvItem As ListItem



'=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*==*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=
'  MAIN PROCEDURE
'
'
'=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*==*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=
Private Sub Class_Initialize()
    Set ClistView = FrmMain.ListView
    Set ObjCommand = New ClsCommand
    ObjCommand.ObjectInit Me
    BlnAgCertified = False
    StrAgConnection = "Offline"
End Sub


'=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*==*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=
'  PUBLIC SUBS & FUNCTIONS
'
'
'=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*==*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Commands
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Function Commands() As ClsCommand
    Set Commands = ObjCommand
End Function

'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Send Data
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Sub NetSend(Data As String)
    Send LngSocketIndex, Data
End Sub
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Ping Client
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Function NetPing() As Integer
    ObjCommand.NetPing
    IntNetPingCount = IntNetPingCount + 1
    NetPing = IntNetPingCount
End Function
Public Sub NetPingReset()
    IntNetPingCount = 0
End Sub
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
'
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Sub NetSetVariable(Name As String, Value As String)
    Select Case Name
        Case Is = "NETMAC"
            StrAgMAC = Value
    End Select
End Sub

'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
'
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Function ItemMain() As ListItem
    Set ItemMain = ClvItem
End Function

'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Initialize Agent
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Sub AgnInit(Online As Boolean, Optional NewSocket As Socket = Nothing, Optional AgentName As String = "")
    If Online = True Then
        Set ObjSck = NewSocket
        BlnNetMode = True
        LngSocketIndex = ObjSck.Index
        StrAgIP = ObjSck.PeerAddress
    Else
        StrAgName = AgentName
        Call zAgentAdd(StrAgName)
    End If
End Sub

'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Reject Agent
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Sub AgnInitReject()
    ObjSck.Cleanup
    Set ClsAgent = Nothing
End Sub

'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Initialize Certification
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Function AgnInitCert(sCommand As String, Optional OfflineSignal As Boolean = False) As Boolean ', Root As clsAgents) As Boolean
    StrAgVersion = CmdSubGet(sCommand, "AGENTVERSION")

    If StrAgVersion >= "1.7.75" Then
        ObjCommand.NetCert Granted
        BlnAgCertified = True
        AgnInitCert = True
        If OfflineSignal = False Then
            '{ Add Terminal to main ListView }'
            StrAgName = CmdSubGet(sCommand, "NAME")
            Call zAgentAdd(StrAgName)
        Else
            '{ Terminal alreay exist. Change symbol to notify }'
            If Me.AgentStatus = VS(1, 2) Then
                Call AgSetSymbol([Terminal Online])
            Else
                Call AgSetSymbol([User Online])
            End If
        End If
        StrAgConnection = "Online"
    Else
        ObjCommand.NetCert Denied, "Version too old, Please update!"
        AgnInitCert = False
    End If
End Function

Public Function AgnOffline()
    BlnAgCertified = False
    BlnNetMode = False
    LngSocketIndex = 0
    
    If Me.AgentStatus = VS(1, 2) Then
        Call AgSetSymbol([Terminal Offline])
    Else
        Call AgSetSymbol([User Offline])
    End If
    StrAgConnection = "Offline"
    ObjSck.Cleanup
    Set ObjSck = Nothing
End Function

Public Sub AgnRecover()
    Dim DbR As Database, RsR As Recordset
    
    Set DbR = OpenDatabase(CurIDBPath, False, False, ";pwd=nsb2003")
    Set RsR = DbR.OpenRecordset("TerminalRecover", dbOpenDynaset)
    
    With RsR
        .FindFirst "AgentName = '" & StrAgName & "'"
        If .NoMatch = False Then
            .Edit
        Else
            .AddNew
        End If
        !AgentName = StrAgName
        !AgentStatus = StrAgStatus
        !CustomerName = StrCusName
        !CustomerFlag = StrCusFlag
        !CustomerType = StrCusType
        !CustomerTimeIn = StrCusTimeIn
        !CustomerTimeOut = StrCusTimeOut
        !CustomerUsage = StrCusUsage
        !CustomerPrinted = LngCusPrinted
        .Update
    End With
    RsR.Close
    DbR.Close
End Sub

Public Sub AgnRecoverRemove()
    Dim DbR As Database, RsR As Recordset
    
    Set DbR = OpenDatabase(CurIDBPath, False, False, ";pwd=nsb2003")
    Set RsR = DbR.OpenRecordset("TerminalRecover", dbOpenDynaset)
    
    With RsR
        .FindFirst "AgentName = '" & StrAgName & "'"
        If .NoMatch = False Then .Delete
    End With
    RsR.Close
    DbR.Close
End Sub

Public Sub AgnTransfer(NewLvIndex As Integer)
    Dim nLv As ListItem
    Dim CuA As ClsAgent
    
    Set nLv = ClistView.ListItems(NewLvIndex)
    Set CuA = UniAgents(nLv.Text)
    ObjCommand.ConLogin LogOut
    
   'copy data
    CuA.CustomerName = StrCusName
    CuA.CustomerTimeIn = StrCusTimeIn
    CuA.CustomerTimeOut = StrCusTimeOut
    CuA.AgentStatus = StrAgStatus
    CuA.CustomerType = StrCusType
    CuA.CustomerFlag = StrCusFlag
    CuA.CustomerUsage = StrCusUsage
    CuA.CustomerPrinted = LngCusPrinted
    
   'copy appearance
    CuA.AgentForeColor = ClvItem.ForeColor
    CuA.AgentSmallIcon = ClvItem.SmallIcon
    CuA.AgentIcon = ClvItem.Icon
    CuA.Commands.ConLogin (LogIn)
    
   'reset current agent
    Call CusStop
    Call AgnRecoverRemove
End Sub


'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Remove Agent
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Sub AgnRemove()
On Error GoTo ErrInt
    ClistView.ListItems.Remove ClvItem.Index
    ObjSck.Cleanup
    Set ClsAgent = Nothing
Exit Sub
ErrInt:
    AppErrorLog Err, "Class Agent | AgnRemove"
    Resume Next
End Sub

'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Check Expires Customer for Prepaid & Fixed Time
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Function CusCheckExpired() As Boolean
    CusCheckExpired = False
    If StrAgStatus = VS(1, 1) And StrCusTimeOut <= Now Then
        Call AgSetSymbol([User Prepaid Ended])
        ObjCommand.ConLogin LogOut
        Me.AgentStatus = VS(1, 3)
        CusCheckExpired = True
    End If
End Function

Public Sub CusStatusUpdate()
    Dim StrPrice As String, StrTimeLeft As String

    Cond1 = StrCusTimeOut = VS(0, 1) And StrAgStatus = VS(1, 1) And BlnNetMode = True
    Cond2 = StrCusTimeOut <> VS(0, 1) And StrAgStatus = VS(1, 1) And BlnNetMode = True

    If Cond1 Then
        StrPrice = CusGetUsage
        ObjCommand.ConUsage PriceUse, Crnc & " " & StrPrice
    ElseIf Cond2 Then
        StrTimeLeft = CusGetTimeLeft
        ObjCommand.ConUsage TimeLeft, StrTimeLeft
    End If
    
    'NetSend "//tiker: [ Thank You ] " !! Implement at agent
End Sub

'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Get Price for Customer
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Function CusGetPrice() As String
    Dim StrPriceRate As String, StrUsedTime As String, StrUsedPrice As String
    
    StrPriceRate = CDataSe.DataFind("PriceScheme", "Scheme", "Price", StrCusType)
    StrPriceRate = StrPriceRate / 60
    If StrPriceRate < 0 Then StrPriceRate = CDataSe.DataFind("PriceScheme", "Scheme", "Price", StrCusType)
    StrUsedTime = DateDiff("n", StrCusTimeIn, Now)
    StrUsedPrice = StrUsedTime * StrPriceRate

    StrUsedPrice = GetRoundUpVal(CSng(StrUsedPrice))
    CusGetPrice = Format$(StrUsedPrice, "#0.00")
End Function

'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Get Time Usage for Customer
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Function CusGetTimeUse(Optional InMinute As Boolean = False) As String
    Dim StrHour As String, StrMinute As String
    
    StrMinute = DateDiff("n", StrCusTimeIn, Now)
    If InMinute = True Then CusGetTimeUse = StrMinute: Exit Function
    StrHour = StrMinute \ 60
    StrMinute = StrMinute - (StrHour * 60)
    CusGetTimeUse = StrHour & " " & VS(0, 2) & ", " & StrMinute & " " & VS(0, 3)
End Function

Public Function CusGetTimeUseEx() As Long
    Dim StrMinute As Long, Saat As Long

    StrMinute = DateDiff("n", StrCusTimeIn, Now)
    Saat = DateDiff("s", StrCusTimeIn, Now)
    CusGetTimeUseEx = Saat
End Function

'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Get Time Left for Customer
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Function CusGetTimeLeft()
    Dim StrHour As String, StrMinute As String
    
    StrMinute = DateDiff("n", Now, StrCusTimeOut)
    StrHour = StrMinute \ 60
    StrMinute = StrMinute - (StrHour * 60)
    CusGetTimeLeft = StrHour & " " & VS(0, 2) & ", " & StrMinute & " " & VS(0, 3)
End Function

Public Function CusGetUsage() As String
    Dim StrPriceRate As String, StrPriceExtra As String
    Dim StrUsedTime As String, StrUsedPrice As String
    
    StrPriceRate = CDataSe.DataFind("PriceScheme", "Scheme", "Price", StrCusType)
    StrPriceExtra = CDataSe.DataFind("PriceScheme", "Scheme", "Extra", StrCusType)
    StrPriceRate = StrPriceRate / 60
    a = StrTimeIn
    If StrCusTimeOut = VS(0, 1) Then
        StrUsedTime = DateDiff("n", StrCusTimeIn, Now)
        StrUsedPrice = StrUsedTime * StrPriceRate
        StrUsedPrice = GetRoundUpVal(CDbl(StrUsedPrice))
        If CDbl(StrUsedPrice) < CDbl(StrPriceExtra) Then StrUsedPrice = StrPriceExtra
    Else
        StrUsedTime = DateDiff("n", StrCusTimeIn, StrCusTimeOut)
        StrUsedPrice = StrUsedTime * StrPriceRate
        StrUsedPrice = GetRoundUpVal(CDbl(StrUsedPrice))
        'If CDbl(StrUsedPrice) < CDbl(StrPriceExtra) Then StrUsedPrice = StrPriceExtra
    End If
    CusGetUsage = Format$(StrUsedPrice, "#0.00")
End Function

Public Function CusGetUsageTime(Optional InMinute As Boolean = False) As String
    Dim StrHour As String, StrMinute As String
    
    If StrCusTimeOut = VS(0, 1) Then
        StrMinute = DateDiff("n", StrCusTimeIn, Now)
        If InMinute = True Then CusGetUsageTime = StrMinute: Exit Function
        StrHour = StrMinute \ 60
        StrMinute = StrMinute - (StrHour * 60)
        CusGetUsageTime = StrHour & " " & VS(0, 2) & ", " & StrMinute & " " & VS(0, 3)
    Else
        StrMinute = DateDiff("n", StrCusTimeIn, StrCusTimeOut)
        If InMinute = True Then CusGetUsageTime = StrMinute: Exit Function
        StrHour = StrMinute \ 60
        StrMinute = StrMinute - (StrHour * 60)
        CusGetUsageTime = StrHour & " " & VS(0, 2) & ", " & StrMinute & " " & VS(0, 3)
    End If
End Function

'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Start Pay As Your Go
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Sub CusStartPAYG(CusId As String, CusName As String, CusType As String)
    Me.AgentStatus = VS(1, 1)
    
    Me.CustomerName = CusName
    Me.CustomerType = CusType
    Me.CustomerTimeIn = Now
    Me.CustomerTimeOut = VS(0, 1)
    Me.CustomerUsage = VS(0, 1)
    Me.CustomerPrinted = 0
    
    If BlnAgCertified = True Then Call AgSetSymbol([User Online])
    If BlnAgCertified = False Then Call AgSetSymbol([User Offline])
    ObjCommand.ConLogin LogIn
End Sub

'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Start Prepaid
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Sub CusStartPPAID(CusId As String, CusName As String, CusType As String, PaidPrice As Double, Optional CusSambung As Boolean = False)
    Dim IntMinute As Integer, IntHour As Integer, DblPriceRate As Double
    Dim MinTmp As String
    
    DblPriceRate = CDataSe.DataFind("PriceScheme", "Scheme", "Price", CusType)
    DblPriceRate = DblPriceRate / 60

    Me.AgentStatus = VS(1, 1)
    Me.CustomerName = CusName
    Me.CustomerType = CusType
    Me.CustomerTimeIn = Now
    Me.CustomerPrinted = 0
    
    IntMinute = PaidPrice / DblPriceRate                       'kira jumlah minit..
    Me.CustomerTimeOut = DateAdd("n", IntMinute, Now)          'bila akan berhenti..
    
    If IntMinute >= 60 Then
        IntHour = IntMinute \ 60: IntMinute = IntMinute - (IntHour * 60)
        If IntMinute = 0 Then MinTmp = "" Else MinTmp = " " & IntMinute & " " & LCase$(VS(0, 3))
        Me.CustomerUsage = IntHour & " " & LCase$(VS(0, 2)) & MinTmp
    Else
        Me.CustomerUsage = IntMinute & " " & LCase$(VS(0, 3))
    End If
    
    If BlnAgCertified = True Then Call AgSetSymbol([User Online])
    If BlnAgCertified = False Then Call AgSetSymbol([User Offline])
    ObjCommand.ConLogin LogIn
End Sub

'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Stop \ Terminate Customer
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Sub CusStop()
    Me.AgentStatus = VS(1, 2)
    
    Me.CustomerId = ""
    Me.CustomerName = ""
    Me.CustomerTimeIn = ""
    Me.CustomerTimeOut = ""
    Me.CustomerType = ""
    Me.CustomerUsage = ""
    Me.CustomerPrinted = 0
    
    If BlnAgCertified = True Then Call AgSetSymbol([Terminal Online])
    If BlnAgCertified = False Then Call AgSetSymbol([Terminal Offline])
    
    Me.AgentForeColor = vbBlack
    Me.AgentFontBold = False
    
    If BlnNetMode = True Then ObjCommand.ConLogin (LogOut)
    
    FrmMain.MainSbar.Panels(2) = ""
    FrmMain.MainSbar.Panels(3) = ""
End Sub


'=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*==*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=
'  PROPERTY SECTION
'
'
'=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*==*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Agent | Connected Time
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Get AgentConnected() As String
    AgentConnected = StrAgConnected
End Property
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Agent | Connection Status
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Get AgentConnection() As String
    AgentConnection = StrAgConnection
End Property
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Agent | Lock Status
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Get AgentLock() As String
    AgentLock = StrAgLock
End Property
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Agent | Name
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Let AgentName(nValue As String)
    If nValue <> "" Then
        StrAgName = nValue
        ClvItem.Text = StrAgName
    End If
End Property
Public Property Get AgentName() As String
    AgentName = StrAgName
End Property
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Certification for Agent
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Get AgentCertified() As Boolean
    AgentCertified = BlnAgCertified
End Property
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
'  Agent Socket Index
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Get AgentSockIndex() As Integer
    AgentSockIndex = LngSocketIndex
End Property
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Agent IP Address
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Get AgentIPAdd() As String
    AgentIPAdd = StrAgIP
End Property
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Agent MAC Address
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Get AgentMAC() As String
    AgentMAC = StrAgMAC
End Property
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Agent Small Icon
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Let AgentSmallIcon(iconName As String)
    If iconName <> "" Then ClvItem.SmallIcon = iconName
End Property
Public Property Get AgentSmallIcon() As String
    AgentSmallIcon = ClvItem.SmallIcon
End Property
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Agent Icon
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Let AgentIcon(iconName As String)
    If iconName <> "" Then ClvItem.Icon = iconName
End Property
Public Property Get AgentIcon() As String
    AgentIcon = ClvItem.Icon
End Property
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Agent Fore Colour
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Let AgentForeColor(colour As ColorConstants)
    On Error Resume Next
    ClvItem.ForeColor = colour
    For B = 1 To 7
        ClvItem.ListSubItems(B).ForeColor = colour
    Next
End Property
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Agent Font Bold
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Let AgentFontBold(nValue As Boolean)
    On Error Resume Next
    ClvItem.Bold = nValue
    For B = 1 To 7
        ClvItem.ListSubItems(B).Bold = nValue
    Next
End Property
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Agent Status
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Let AgentStatus(nValue As String)
    If StrAgStatus = nValue Then Exit Property
    StrAgStatus = nValue
    ClvItem.SubItems(1) = StrAgStatus
End Property
Public Property Get AgentStatus() As String
    AgentStatus = StrAgStatus
End Property

'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Customer Id
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Let CustomerId(nValue As String)
    If StrCusId = nValue Then Exit Property
    StrCusId = nValue
End Property
Public Property Get CustomerId() As String
    CustomerId = StrCusId
End Property
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Customer Name
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Let CustomerName(nValue As String)
    If StrCusName = nValue Then Exit Property
    StrCusName = nValue
    ClvItem.SubItems(2) = StrCusName
End Property
Public Property Get CustomerName() As String
    CustomerName = StrCusName
End Property
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Customer Time In
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Let CustomerTimeIn(nValue As String)
    If StrCusTimeIn = nValue Then Exit Property
    StrCusTimeIn = nValue
    If StrCusTimeIn = VS(0, 1) Or StrCusTimeIn = "" Then
        ClvItem.SubItems(4) = StrCusTimeIn
    Else
        ClvItem.SubItems(4) = TimeValue(StrCusTimeIn)
    End If
End Property
Public Property Get CustomerTimeIn() As String
    CustomerTimeIn = StrCusTimeIn
End Property
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Customer Time Out
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Let CustomerTimeOut(nValue As String)
    If StrCusTimeOut = nValue Then Exit Property
    StrCusTimeOut = nValue
    If StrCusTimeOut = VS(0, 1) Or StrCusTimeOut = "" Then
        ClvItem.SubItems(5) = StrCusTimeOut
    Else
        ClvItem.SubItems(5) = TimeValue(StrCusTimeOut)
    End If
End Property
Public Property Get CustomerTimeOut() As String
    CustomerTimeOut = StrCusTimeOut
End Property
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Customer Scheme
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Let CustomerType(nValue As String)
    If StrCusType = nValue Then Exit Property
    StrCusType = nValue
    ClvItem.SubItems(3) = StrCusType
End Property
Public Property Get CustomerType() As String
    CustomerType = StrCusType
End Property
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Customer Flag
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Let CustomerFlag(nValue As String)
    If StrCusFlag = nValue Then Exit Property
    StrCusFlag = nValue
    ClvItem.ListSubItems(1).Tag = nValue
End Property
Public Property Get CustomerFlag() As String
    CustomerFlag = StrCusFlag
End Property
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Customer Usage
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Let CustomerUsage(nValue As String)
    If StrCusUsage = nValue Then Exit Property
    StrCusUsage = nValue
    ClvItem.SubItems(6) = nValue
End Property
Public Property Get CustomerUsage() As String
    CustomerUsage = StrCusUsage
End Property
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Customer Printed
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Property Let CustomerPrinted(nValue As Long)
    If LngCusPrinted = nValue Then Exit Property
    LngCusPrinted = nValue
    ClvItem.SubItems(7) = nValue
End Property
Public Property Get CustomerPrinted() As Long
    CustomerPrinted = LngCusPrinted
End Property
'=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*==*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=



'=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*==*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=
'  PRIVATE SUBS & FUNCTIONS
'
'
'=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*==*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Add Agent to Listview, plus load recover data
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Private Sub zAgentAdd(s_Nama As String)
    Dim DbR As Database, RsR As Recordset
    
    Set DbR = OpenDatabase(CurIDBPath, False, False, ";pwd=nsb2003")
    Set RsR = DbR.OpenRecordset("TerminalRecover", dbOpenDynaset)
    Set ClvItem = ClistView.ListItems.Add(, s_Nama, s_Nama, , "TerminalOnline")
    
    ClvItem.Tag = LngSocketIndex
    StrAgStatus = VS(1, 2)
    StrAgConnected = Time
    ClvItem.SubItems(1) = StrAgStatus
    ClistView.Sorted = True: ClistView.Sorted = False
    
    With RsR
        .FindFirst "AgentName = '" & StrAgName & "'"
        If .NoMatch = False Then
            Me.AgentStatus = !AgentStatus
            Me.CustomerName = !CustomerName
            Me.CustomerFlag = !CustomerFlag
            Me.CustomerType = !CustomerType
            Me.CustomerTimeIn = !CustomerTimeIn
            Me.CustomerTimeOut = !CustomerTimeOut
            Me.CustomerUsage = !CustomerUsage
            Me.CustomerPrinted = !CustomerPrinted
            If Me.AgentStatus = VS(1, 3) Then
                Call AgSetSymbol([User Prepaid Ended])
            Else
                If BlnAgCertified = True Then Call AgSetSymbol([User Online])
                If BlnAgCertified = False Then Call AgSetSymbol([User Offline])
            End If
        End If
    End With

    RsR.Close
    DbR.Close
End Sub

Public Sub AgStatusRequested(SubCommand As String)
    Dim StrStatusLock As String, StrStatusClean As String
    
    StrStatusLock = CmdSubGet(SubCommand, "LOCK")
    StrStatusClean = CmdSubGet(SubCommand, "CLEAN")
    
    If StrStatusLock = "1" Then
        StrAgLock = "Locked"
    Else
        StrAgLock = "Unlock"
    End If
    
    If StrAgStatus = VS(1, 1) Then
        If StrStatusLock = "1" Then
            AgSetSymbol [Terminal Lock]
        ElseIf StrStatusLock = "0" Then
            If BlnAgCertified = True Then
                AgSetSymbol [User Online]
            Else
                AgSetSymbol [User Offline]
            End If
        End If
        If StrStatusClean = "1" Then
            If BlnAgCertified = True Then
                AgSetSymbol [User Online]
            Else
                AgSetSymbol [User Offline]
            End If
        End If
            
    End If
    If StrAgStatus = VS(1, 2) Then
        If StrStatusLock = "1" Then
            AgSetSymbol [Terminal Lock]
        ElseIf StrStatusLock = "0" Then
            If BlnAgCertified = True Then
                AgSetSymbol [Terminal Online]
            Else
                AgSetSymbol [Terminal Offline]
            End If
        End If
        If StrStatusClean = "1" Then
            If BlnAgCertified = True Then
                AgSetSymbol [Terminal Online]
            Else
                AgSetSymbol [Terminal Offline]
            End If
        End If
    End If
    If StrAgStatus = VS(1, 3) Then
        If StrStatusClean = "1" Then AgSetSymbol [User Prepaid Ended]
    End If
End Sub

'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
' Set Agent Symbol
'=-=-===-===-=-==-=-=-=-=-==-==-=-=-=-=-=-=-=-=-==
Public Sub AgSetSymbol(Symbol As EnAgentSymbol)
    Select Case Symbol
    Case 0
        ClvItem.SmallIcon = "UserOnline"
        Me.AgentForeColor = &H8000&
    Case 1
        ClvItem.SmallIcon = "UserOffline"
        Me.AgentForeColor = &H8000&
    Case 2
        ClvItem.SmallIcon = "UserPrepaid"
        Me.AgentForeColor = &H8000&
    Case 3
        ClvItem.SmallIcon = "UserEnded"
        Me.AgentForeColor = vbBlue
    Case 10
        ClvItem.SmallIcon = "TerminalOnline"
    Case 11
        ClvItem.SmallIcon = "TerminalOffline"
    Case 12
        ClvItem.SmallIcon = "TerminalLock"
    Case 13
        ClvItem.SmallIcon = "TerminalClean"
    End Select
End Sub
'=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*==*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=
