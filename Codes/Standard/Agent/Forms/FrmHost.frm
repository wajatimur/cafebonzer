VERSION 5.00
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "Cswsk32.ocx"
Begin VB.Form FrmHost 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CbAgentHost"
   ClientHeight    =   1455
   ClientLeft      =   255
   ClientTop       =   1695
   ClientWidth     =   3615
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmHost.frx":0000
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Source
   LinkTopic       =   "CbAgentHost"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   3615
   StartUpPosition =   1  'CenterOwner
   Begin SocketWrenchCtrl.Socket Socket 
      Left            =   1530
      Top             =   960
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
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
      RemotePort      =   8180
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   0   'False
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1020
      Top             =   960
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   990
      LinkItem        =   "DdeClientItem"
      TabIndex        =   0
      Top             =   60
      Width           =   2535
   End
   Begin VB.Timer Pinger 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   75
      Top             =   960
   End
   Begin VB.Timer Connecter 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   555
      Top             =   960
   End
   Begin VB.Image IconStatOn 
      Height          =   240
      Left            =   300
      Picture         =   "FrmHost.frx":6852
      ToolTipText     =   "Not Connected"
      Top             =   60
      Width           =   240
   End
   Begin VB.Image IconStatClean 
      Height          =   240
      Left            =   540
      Picture         =   "FrmHost.frx":6DDC
      ToolTipText     =   "Not Connected"
      Top             =   60
      Width           =   240
   End
   Begin VB.Image IconStatOff 
      Height          =   240
      Left            =   60
      Picture         =   "FrmHost.frx":7366
      ToolTipText     =   "Not Connected"
      Top             =   60
      Width           =   240
   End
End
Attribute VB_Name = "FrmHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : FrmHost
'    Project    : CafeBonzerAG
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>


'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Timer] - Pinger & Condition Checker
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Private Sub Pinger_Timer()
On Error GoTo ErrInt
    Call NetPing
    'Call AgentInfoStatus
Exit Sub

ErrInt:
    AppErrorLog Err, "Pinger | Timer"
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Timer] - Connecter
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Private Sub Connecter_Timer()
On Error GoTo ErrInt
    Socket.HostAddress = StrNetHost
    Socket.RemotePort = StrNetPort
    Socket.Connect
Exit Sub

ErrInt:
    AppErrorLog Err, "Connecter | Timer"
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Socket] - Socket Close Event
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Private Sub Socket_Disconnect()
    Call NetClose
    Call NetConnect
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Socket] - On Connect Event
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Private Sub Socket_Connect()
On Error GoTo ErrInt
    Connecter.Enabled = False
    NetSend "010030" & CmdSubPut("NAME", SysInfoGetName) & CmdSubPut("AGENTVERSION", (App.Major & "." & StrAppBuild))
Exit Sub

ErrInt:
    AppErrorLog Err, "Socket | Connect"
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Socket] - Read Event
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Private Sub Socket_Read(DataLength As Integer, IsUrgent As Integer)
On Error GoTo ErrInt
    Dim DataRcv As String, DataRcvLen As Long
    
    Socket.Read DataRcv, DataLength
    DataRcvLen = Len(Trim$(DataRcv))
    If DataRcvLen = 0 Then Exit Sub
    Call CmdParse(DataRcv)
Exit Sub

ErrInt:
    If Err.Number = 24504 Then
        Call NetClose
        Call NetConnect
    Else
        AppErrorLog Err, "Socket | Read"
    End If
End Sub

Private Sub Timer1_Timer()
    Static StcCnt As Integer
    
    If StcCnt = 0 Then
        'IPCSendData "TFrmMain", "STATUSONLINE"
        StcCnt = 1
    Else
        'IPCSendData "TFrmMain", "STATUSOFFLINE"
        StcCnt = 0
    End If
End Sub
