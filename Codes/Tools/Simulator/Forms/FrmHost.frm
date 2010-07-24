VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Begin VB.Form FrmHost 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agent Simulator"
   ClientHeight    =   6105
   ClientLeft      =   255
   ClientTop       =   1695
   ClientWidth     =   8190
   ClipControls    =   0   'False
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   8190
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      Caption         =   "No Ping"
      Height          =   435
      Left            =   6945
      TabIndex        =   9
      Top             =   4095
      Width           =   1110
   End
   Begin SocketWrenchCtrl.Socket Socket 
      Index           =   0
      Left            =   1785
      Top             =   90
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   5
      Binary          =   0   'False
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   1024
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   1000
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   8180
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   300
      Left            =   6900
      TabIndex        =   8
      Top             =   4710
      Width           =   1230
   End
   Begin VB.Frame Frame2 
      Caption         =   "Statistic"
      Height          =   1995
      Left            =   3570
      TabIndex        =   7
      Top             =   4035
      Width           =   3210
   End
   Begin VB.CommandButton CmdBut 
      Caption         =   "DisConnect"
      Height          =   435
      Index           =   1
      Left            =   6900
      TabIndex        =   4
      Top             =   5100
      Width           =   1230
   End
   Begin VB.CommandButton CmdBut 
      Caption         =   "Connect"
      Height          =   435
      Index           =   0
      Left            =   6900
      TabIndex        =   3
      Top             =   5610
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      Caption         =   "Settings"
      Height          =   1995
      Left            =   75
      TabIndex        =   2
      Top             =   4035
      Width           =   3420
      Begin VB.TextBox Txt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1530
         TabIndex        =   5
         Text            =   "5"
         Top             =   285
         Width           =   675
      End
      Begin VB.Label Lbl 
         Caption         =   "Total Agent :"
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   6
         Top             =   315
         Width           =   1215
      End
   End
   Begin VB.ListBox List 
      Appearance      =   0  'Flat
      Height          =   3930
      Left            =   2985
      TabIndex        =   1
      Top             =   60
      Width           =   5160
   End
   Begin VB.Timer Monitor 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   1350
      Top             =   75
   End
   Begin VB.Timer Pinger 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   510
      Top             =   75
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   90
      Top             =   75
   End
   Begin VB.Timer Connecter 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   930
      Top             =   75
   End
   Begin MSComctlLib.ListView mLv 
      Height          =   3930
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Width           =   2850
      _ExtentX        =   5027
      _ExtentY        =   6932
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Socket Id"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "FrmHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private l_idx As Long
Private l_wIdx As Long
Private b_wait As Boolean


Private Sub Check1_Click()
    NoReply = Check1.Value
End Sub

Private Sub CmdBut_Click(Index As Integer)
    Select Case Index
        Case 0
            Call NetConnect
        Case 1
            List.Clear
            Call NetClose
    End Select
End Sub

Private Sub Command1_Click()
    List.Clear
End Sub

Private Sub Form_Terminate()
    Call Tutup
End Sub


'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Timer] - Connecter
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Private Sub Connecter_Timer()
On Error GoTo ErrInt
    Dim tItm As ListItem
    
    If lTotalAgent = 0 Then
        For a = 1 To Txt(0)
            Load Socket(a)
            Set tItm = mLv.ListItems.Add(, , Socket(a).Handle)
            tItm.SubItems(1) = "Loaded"
        Next a
        lTotalAgent = Txt(0)
    End If
    

    For a = 1 To lTotalAgent
        Socket(a).HostName = SetGet("nomborip", "127.0.0.1")
        Socket(a).RemotePort = SetGet("porttempatan", "8180")
        Socket(a).Connect
        mLv.ListItems(a).Text = Socket(a).Handle
        mLv.ListItems(a).SubItems(1) = "Connecting"
    Next a

Exit Sub

ErrInt:
    ErrHand Err, "Connecter_Timer"
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Timer] - Monitoring
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Private Sub Monitor_Timer()
  ' Switching mode
  ' 1 = printer only
  ' 2 = resource + printer
  ' 3 = app + printer
  ' 4 = traffic + printer
  
    If bConnected = False Then Exit Sub
    b_MonPrinter = SetGet("mon.printer", True)
    b_MonResource = SetGet("mon.resource", True)
    b_MonApp = SetGet("mon.app", True)
    b_MonTraffic = SetGet("mon.traffic", True)
    
    If b_MonPrinter = True Then MonPrinter
    
    Select Case l_MonSwitch
    Case 2
        If b_MonResource = True Then MonResource
    Case 3
        If b_MonApp = True Then MonApp
    End Select
End Sub


'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Timer] - Pinger & Condition Checker
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Private Sub Pinger_Timer()
   'ping the server
    For a = 1 To Txt(0)
        Call NetPing(a)
        If bConLock = 1 Then NetSend a, "/info.me:lock": idx = 0
    Next a
End Sub


'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Socket] - Socket Close Event
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Private Sub Socket_Disconnect(Index As Integer)
    NetClose
    NetConnect
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Socket] - On Connect Event
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Private Sub Socket_Connect(Index As Integer)
On Error GoTo ErrInt
  ' hentikan penyambung automatik
    Connecter.Enabled = False
  ' certified station
    NetSend Index, "/cert:" & SubBuild("name", (MyName & Index)) & SubBuild("version", "2.0.0")
    mLv.ListItems(Index).SubItems(1) = "Certifying"
Exit Sub

ErrInt:
    Call ErrHand(Err, "Socket_Connect")
End Sub


'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Socket] - Read Event
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Private Sub Socket_Read(Index As Integer, DataLength As Integer, IsUrgent As Integer)
On Error GoTo ErrInt
    Dim DataRcv As String
    Dim DataTmp As String
    Dim DataRcvLen As Long
    Dim xPos As Long
    Dim yPos As Long
    
    Socket(Index).Read DataRcv, DataLength
    DataRcvLen = Len(DataRcv)
    xPos = 0

    Do Until xPos = DataRcvLen
        xPos = InStr(xPos + 1, DataRcv, "//")
        yPos = InStr(xPos + 1, DataRcv, "//")
        
        If xPos = 0 Then Exit Sub
        If yPos = 0 Then
            DataTmp = Mid(DataRcv, xPos)
            xPos = Len(DataRcv)
        Else
            DataTmp = Mid(DataRcv, xPos, yPos - xPos)
        End If
        
        Call CmdParse(Index, DataTmp)
    Loop
Exit Sub
    
ErrInt:
    If Err.Number = 24504 Then
        NetClose
        NetConnect
    Else
        ErrHand Err, "Sock_DataArrival"
    End If
End Sub



'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Timer] - Timer2 | For Ticker Message Timer
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Private Sub Timer2_Timer()
    Timer2.Enabled = False
    bConTick = False
End Sub
