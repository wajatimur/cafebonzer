VERSION 5.00
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Begin VB.Form FrmHost 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   585
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   1920
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmHost.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   585
   ScaleWidth      =   1920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SocketWrenchCtrl.Socket Socket 
      Index           =   0
      Left            =   105
      Top             =   75
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
      Binary          =   0   'False
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   2048
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1380
      Top             =   75
   End
   Begin VB.Timer Pinger 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   525
      Top             =   75
   End
   Begin VB.Timer NetTimer 
      Interval        =   10
      Left            =   960
      Top             =   75
   End
End
Attribute VB_Name = "FrmHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents EveAgents As clsAgents
Attribute EveAgents.VB_VarHelpID = -1

Private Sub Form_Load()
    Set EveAgents = UniAgents
End Sub

Private Sub EveAgents_AgentAdded(Agent As clsAgent)
    MainLog Agent.AgentName & " " & VS(9) & " " & Agent.AgentConnected
    Call UpdatePanel(SelText)
    Agent.AgnAddPage MglPageLast
End Sub

Private Sub EveAgents_AgentRemove(Agent As clsAgent)
    Call UpdatePanel(SelText)
    Call UpdateStat(Nothing)
End Sub

Private Sub EveAgents_InfoUpdated(Agent As clsAgent, InfoType As Long)
    Select Case InfoType
        Case 1
            
        Case 2
            If MglPageLast = 2 Then
                Dim AgI As clsAgInfo
                Set AgI = Agent.AgentInfo
                Agent.ItemDyna2.SubItems(1) = AgI.MemLoad & " %"
                Agent.ItemDyna2.SubItems(2) = (AgI.MemPhyTotal \ 1024) \ 1024 & " Mb"
                Agent.ItemDyna2.SubItems(3) = (AgI.MemPhyAvail \ 1024) \ 1024 & " Mb"
                Agent.ItemDyna2.SubItems(4) = (AgI.MemVirTotal \ 1024) \ 1024 & " Mb"
                Agent.ItemDyna2.SubItems(5) = (AgI.MemVirAvail \ 1024) \ 1024 & " Mb"
                Agent.ItemDyna2.SubItems(6) = (AgI.MemPageTotal \ 1024) \ 1024 & " Mb"
                Agent.ItemDyna2.SubItems(7) = (AgI.MemPageAvail \ 1024) \ 1024 & " Mb"
            End If
    End Select
End Sub



'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Tutup socket
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub Socket_Disconnect(Index As Integer)
    UniAgents.RemoveAgent Index
End Sub


'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Connection request
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub Socket_Accept(Index As Integer, SocketID As Integer)
On Error GoTo ErrTrap
    lSock = lSock + 1
    Load Socket(lSock)
    UniAgents.AddAgent Socket(lSock), CLng(SocketID)
Exit Sub
ErrTrap:
    ErrLog Err, "Socket | Connect"
End Sub


'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Socket - data terima
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub Socket_Read(Index As Integer, DataLength As Integer, IsUrgent As Integer)
On Error GoTo ErrTrap
    Dim DataRcv As String
    
    Socket(Index).Read DataRcv, DataLength
    Call ParseCmd(DataRcv, CLng(Index))
Exit Sub

ErrTrap:
    ErrLog Err, "Socket | Read", True
End Sub


'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Cek pengguna dalam LV
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub Timer1_Timer()
On Error GoTo ErrTrap
    Dim sItm As ListItem, uA As clsAgent
    
    If AgentCount = 0 Then
        pinger.Enabled = False
    Else
        For i = 1 To AgentCount
            Set sItm = FrmMain.Lv1.ListItems(i)
            Set uA = UniAgents(sItm.Text)
          '{ update infobar }'
            If sItm.Selected = True And MglPageLast = 0 Then Call UpdateStat(sItm)
          '{ save data untuk recovery }'
            Call RecoveryGo(sItm)
          '{ cek untuk masa tamat bagi yang disetkan masa }'
            Call uA.CusCheckExpired
          '{ updatekan status harga dan printer }'
            Call uA.CusStatusUpdate
            Set sItm = Nothing
            Set uA = Nothing
        Next i
    End If
Exit Sub
    
ErrTrap:
    ErrLog Err, "FrmHost | Timer1_Timer"
End Sub


Private Sub NetTimer_Timer()
On Error GoTo ErrInt
    Dim SID As Long, DTS As String, l_FndCnt As Long
    Dim TmpDst As clsDataStore

    If StackNetData.Count = 0 Then Exit Sub
    Set TmpDst = StackNetData(1)
    SID = TmpDst("sockindex")
    DTS = TmpDst("data")
    StackNetData.Remove (1)
    
    If FrmHost.Socket(SID).IsWritable = True Then
        FrmHost.Socket(SID).SendLen = Len(DTS)
        FrmHost.Socket(SID).SendData = DTS
    End If
Exit Sub
    
ErrInt:
    For a% = 0 To StackNetData.Count - 1
        If TmpDst.Name = SID Then
            StackNetData.Remove (a%)
        End If
    Next a%
    
    StatText 2
    StatText 3
End Sub


'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Pinger Timer - tukang ping
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub pinger_Timer()
On Error GoTo ErrTrap
    Dim AgName As String, AgToRemove() As clsAgent, PingCounter As Integer
    Dim nUpBound As Integer, j As Long
    
    ReDim AgToRemove(0)
    
    For j = 1 To AgentCount
        AgName = FrmMain.Lv1.ListItems(j).Text
        PingCounter = UniAgents(AgName).NetPing
        
        If PingCounter >= 6 Then
            nUpBound = nUpBound + 1
            ReDim Preserve AgToRemove(nUpBound)
            Set AgToRemove(nUpBound) = UniAgents.Agents(AgName)
            'AgToRemove(nUpBound) = AgName
        End If
    Next j
    
    Timer1.Enabled = False
    If UBound(AgToRemove) > 0 Then
        For j = 1 To UBound(AgToRemove)
            UniAgents.RemoveAgent AgToRemove(j).AgentSockIndex
        Next j
    End If
    Timer1.Enabled = True
Exit Sub
    
ErrTrap:
    ErrLog Err, "pinger_Timer", True
End Sub
