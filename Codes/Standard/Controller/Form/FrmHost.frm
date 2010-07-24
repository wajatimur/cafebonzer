VERSION 5.00
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "Cswsk32.ocx"
Begin VB.Form FrmSysHost 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   600
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   1905
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmHost.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleWidth      =   1905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SocketWrenchCtrl.Socket Socket 
      Index           =   0
      Left            =   90
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
   Begin VB.Timer TmrAgent 
      Interval        =   2000
      Left            =   1380
      Top             =   75
   End
   Begin VB.Timer TmrPing 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   525
      Top             =   75
   End
   Begin VB.Timer TmrNet 
      Interval        =   10
      Left            =   960
      Top             =   75
   End
End
Attribute VB_Name = "FrmSysHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : FrmSysHost
'    Project    : CafeBonzer
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Private WithEvents EveAgents As ClsAgents
Attribute EveAgents.VB_VarHelpID = -1

Private Sub Form_Load()
    Set EveAgents = UniAgents
End Sub


'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Tutup socket
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub Socket_Disconnect(Index As Integer)
    UniAgents.AgentDisconnect Index
End Sub
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Connection request
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub Socket_Accept(Index As Integer, SocketID As Integer)
On Error GoTo ErrTrap
    lSock = lSock + 1
    Load Socket(lSock)
    UniAgents.AgentAdd Socket(lSock), CLng(SocketID)
Exit Sub
ErrTrap:
    AppErrorLog Err, "Socket | Connect"
End Sub
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Socket - data terima
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub Socket_Read(Index As Integer, DataLength As Integer, IsUrgent As Integer)
On Error GoTo ErrTrap
    Dim DataRcv As String
    
    Socket(Index).Read DataRcv, DataLength
    Call CmdParse(DataRcv, CLng(Index))
Exit Sub

ErrTrap:
    AppErrorLog Err, "Socket | Read"
End Sub


'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Timer Agent
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub TmrAgent_Timer()
On Error GoTo ErrTrap
    If UniAgents.Count = 0 Then
        TmrPing.Enabled = False
    Else
        'UniAgents.AgentRecoverUsed
        UniAgents.AgentCheckUsed
    End If
Exit Sub

ErrTrap:
    AppErrorLog Err, "FrmSysHost | TmrAgent_Timer"
End Sub
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Timer Network
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub TmrNet_Timer()
On Error GoTo ErrInt
    Dim SID As Long, DTS As String, LngIdx As Long
    Dim TmpDst As ClsDataStore

    If StackNetData.Count = 0 Then Exit Sub
    Set TmpDst = StackNetData(1)
    SID = TmpDst("sockindex")
    DTS = StrCmdSep & TmpDst("data")
    StackNetData.Remove (1)
    
    If FrmSysHost.Socket(SID).IsWritable = True Then
        FrmSysHost.Socket(SID).SendLen = Len(DTS)
        FrmSysHost.Socket(SID).SendData = DTS
    End If
Exit Sub

ErrInt:
    For LngIdx = 0 To StackNetData.Count - 1
        If TmpDst.Name = SID Then
            StackNetData.Remove (LngIdx)
        End If
    Next
    
    StatText 2
    StatText 3
End Sub
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' TmrPing Timer - tukang ping
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub TmrPing_Timer()
On Error GoTo ErrTrap
    Dim LngIdxPing As Long, LngIdx As Long, CuA As ClsAgent
    
    For LngIdx = 1 To UniAgents.Count
        Set CuA = UniAgents(LngIdx)
        If CuA.AgentCertified = True Then
            LngIdxPing = CuA.NetPing
            If LngIdxPing >= 6 Then
                CuA.NetPingReset
                UniAgents.AgentDisconnect CuA.AgentSockIndex
            End If
        End If
    Next LngIdx
Exit Sub

ErrTrap:
    AppErrorLog Err, "TmrPing_Timer"
End Sub


'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Events | Agent Added
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub EveAgents_AgentAdded(Agent As ClsAgent)
    SecAppMainLog Agent.AgentName & " " & ST(2, 4) & " " & Agent.AgentConnected
End Sub
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Events | Agent Remove
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub EveAgents_AgentRemove(Agent As ClsAgent)
    Call UpdatePanel(SelText)
    Call UpdateStat(Nothing)
End Sub
