Option Strict Off
Option Explicit On
Friend Class FrmSysHost
	Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "
	Public Sub New()
		MyBase.New()
		If m_vb6FormDefInstance Is Nothing Then
			If m_InitializingDefInstance Then
				m_vb6FormDefInstance = Me
			Else
				Try 
					'For the start-up form, the first instance created is the default instance.
					If System.Reflection.Assembly.GetExecutingAssembly.EntryPoint.DeclaringType Is Me.GetType Then
						m_vb6FormDefInstance = Me
					End If
				Catch
				End Try
			End If
		End If
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents _Socket_0 As AxSocketWrenchCtrl.AxSocket
	Public WithEvents Timer1 As System.Windows.Forms.Timer
	Public WithEvents Pinger As System.Windows.Forms.Timer
	Public WithEvents NetTimer As System.Windows.Forms.Timer
	Public WithEvents Socket As AxSocketArray.AxSocketArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmSysHost))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me._Socket_0 = New AxSocketWrenchCtrl.AxSocket
		Me.Timer1 = New System.Windows.Forms.Timer(components)
		Me.Pinger = New System.Windows.Forms.Timer(components)
		Me.NetTimer = New System.Windows.Forms.Timer(components)
		Me.Socket = New AxSocketArray.AxSocketArray(components)
		CType(Me._Socket_0, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Socket, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.ControlBox = False
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.ClientSize = New System.Drawing.Size(128, 39)
		Me.Location = New System.Drawing.Point(17, 94)
		Me.KeyPreview = True
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.Enabled = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "FrmSysHost"
		_Socket_0.OcxState = CType(resources.GetObject("_Socket_0.OcxState"), System.Windows.Forms.AxHost.State)
		Me._Socket_0.Location = New System.Drawing.Point(7, 5)
		Me._Socket_0.Name = "_Socket_0"
		Me.Timer1.Interval = 2000
		Me.Timer1.Enabled = True
		Me.Pinger.Enabled = False
		Me.Pinger.Interval = 4000
		Me.NetTimer.Interval = 10
		Me.NetTimer.Enabled = True
		Me.Controls.Add(_Socket_0)
		Me.Socket.SetIndex(_Socket_0, CType(0, Short))
		CType(Me.Socket, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me._Socket_0, System.ComponentModel.ISupportInitialize).EndInit()
	End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As FrmSysHost
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As FrmSysHost
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New FrmSysHost()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	Private WithEvents EveAgents As clsAgents
	
	Private Sub FrmSysHost_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		EveAgents = UniAgents
	End Sub
	
	Private Sub EveAgents_AgentAdded(ByRef Agent As clsAgent) Handles EveAgents.AgentAdded
		MainLog(Agent.AgentName & " " & VS(9) & " " & Agent.AgentConnected)
		Call UpdatePanel(SelText)
		Agent.AgnAddPage(MglPageLast)
	End Sub
	
	Private Sub EveAgents_AgentRemove_Renamed(ByRef Agent As clsAgent) Handles EveAgents.AgentRemove_Renamed
		Call UpdatePanel(SelText)
		Call UpdateStat(Nothing)
	End Sub
	
	Private Sub EveAgents_InfoUpdated(ByRef Agent As clsAgent, ByRef InfoType As Integer) Handles EveAgents.InfoUpdated
		Select Case InfoType
			Case 1
				
			Case 2
				
		End Select
	End Sub
	
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Tutup socket
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Private Sub Socket_DisconnectEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Socket.DisconnectEvent
		Dim Index As Short = Socket.GetIndex(eventSender)
		UniAgents.AgentDisconnect(Index)
	End Sub
	
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Connection request
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Private Sub Socket_AcceptEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxSocketWrenchCtrl._DSocketWrenchEvents_AcceptEvent) Handles Socket.AcceptEvent
		Dim Index As Short = Socket.GetIndex(eventSender)
		On Error GoTo ErrTrap
		lSock = lSock + 1
		Socket.Load(lSock)
		UniAgents.AgentAdd(Socket(lSock), CInt(eventArgs.SocketID))
		Exit Sub
ErrTrap: 
		ErrLog(Err, "Socket | Connect")
	End Sub
	
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Socket - data terima
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Private Sub Socket_ReadEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxSocketWrenchCtrl._DSocketWrenchEvents_ReadEvent) Handles Socket.ReadEvent
		Dim Index As Short = Socket.GetIndex(eventSender)
		On Error GoTo ErrTrap
		Dim DataRcv As String
		
		Socket(Index).Read(DataRcv, eventArgs.DataLength)
		Call ParseCmd(DataRcv, CInt(Index))
		Exit Sub
		
ErrTrap: 
		ErrLog(Err, "Socket | Read", True)
	End Sub
	
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Cek pengguna dalam LV
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
		On Error GoTo ErrTrap
		Dim sItm As MSComctlLib.ListItem
		Dim uA As clsAgent
		
		If UniAgents.Count = 0 Then
			Pinger.Enabled = False
		Else
			UniAgents.AgentRecoverUsed()
			UniAgents.AgentCheckUsed()
		End If
		Exit Sub
		
ErrTrap: 
		ErrLog(Err, "FrmSysHost | Timer1_Timer")
	End Sub
	
	
	Private Sub NetTimer_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles NetTimer.Tick
		Dim a As Short
		On Error GoTo ErrInt
		Dim SID, l_FndCnt As Integer
		Dim DTS As String
		Dim TmpDst As clsDataStore
		
		If StackNetData.Count() = 0 Then Exit Sub
		TmpDst = StackNetData.Item(1)
		'UPGRADE_WARNING: Couldn't resolve default property of object TmpDst(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		SID = TmpDst.Data("sockindex")
		'UPGRADE_WARNING: Couldn't resolve default property of object TmpDst(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		DTS = TmpDst.Data("data")
		StackNetData.Remove((1))
		
		If FrmSysHost.DefInstance.Socket(SID).IsWritable = True Then
			FrmSysHost.DefInstance.Socket(SID).SendLen = Len(DTS)
			FrmSysHost.DefInstance.Socket(SID).SendData = DTS
		End If
		Exit Sub
		
ErrInt: 
		For a = 0 To StackNetData.Count() - 1
			If CDbl(TmpDst.Name) = SID Then
				StackNetData.Remove((a))
			End If
		Next a
		
		StatText(2)
		StatText(3)
	End Sub
	
	
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Pinger Timer - tukang ping
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Private Sub pinger_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles pinger.Tick
		On Error GoTo ErrTrap
		Dim PingCounter As Short
		Dim j As Integer
		Dim uA As clsAgent
		
		For j = 1 To UniAgents.Count
			uA = UniAgents.Agents(j)
			If uA.AgentCertified = True Then
				PingCounter = uA.NetPing
				If PingCounter >= 6 Then
					uA.NetPingReset()
					UniAgents.AgentDisconnect((uA.AgentSockIndex))
				End If
			End If
		Next j
		Exit Sub
		
ErrTrap: 
		ErrLog(Err, "pinger_Timer", True)
	End Sub
End Class