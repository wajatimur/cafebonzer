Option Strict Off
Option Explicit On
Friend Class FrmSysDbg
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
	Public WithEvents List3 As System.Windows.Forms.ListBox
	Public WithEvents Frame3 As System.Windows.Forms.GroupBox
	Public WithEvents Command1 As System.Windows.Forms.Button
	Public WithEvents List2 As System.Windows.Forms.ListBox
	Public WithEvents Frame2 As System.Windows.Forms.GroupBox
	Public WithEvents List1 As System.Windows.Forms.ListBox
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents Timer1 As System.Windows.Forms.Timer
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmSysDbg))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.Frame3 = New System.Windows.Forms.GroupBox
		Me.List3 = New System.Windows.Forms.ListBox
		Me.Command1 = New System.Windows.Forms.Button
		Me.Frame2 = New System.Windows.Forms.GroupBox
		Me.List2 = New System.Windows.Forms.ListBox
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.List1 = New System.Windows.Forms.ListBox
		Me.Timer1 = New System.Windows.Forms.Timer(components)
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Internal Process Viewer"
		Me.ClientSize = New System.Drawing.Size(614, 311)
		Me.Location = New System.Drawing.Point(17, 113)
		Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.KeyPreview = True
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
		Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "FrmSysDbg"
		Me.Frame3.Text = "Forms"
		Me.Frame3.Size = New System.Drawing.Size(195, 265)
		Me.Frame3.Location = New System.Drawing.Point(413, 2)
		Me.Frame3.TabIndex = 5
		Me.Frame3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame3.BackColor = System.Drawing.SystemColors.Control
		Me.Frame3.Enabled = True
		Me.Frame3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame3.Visible = True
		Me.Frame3.Name = "Frame3"
		Me.List3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.List3.Size = New System.Drawing.Size(176, 231)
		Me.List3.Location = New System.Drawing.Point(8, 16)
		Me.List3.TabIndex = 6
		Me.List3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.List3.BackColor = System.Drawing.SystemColors.Window
		Me.List3.CausesValidation = True
		Me.List3.Enabled = True
		Me.List3.ForeColor = System.Drawing.SystemColors.WindowText
		Me.List3.IntegralHeight = True
		Me.List3.Cursor = System.Windows.Forms.Cursors.Default
		Me.List3.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.List3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.List3.Sorted = False
		Me.List3.TabStop = True
		Me.List3.Visible = True
		Me.List3.MultiColumn = False
		Me.List3.Name = "List3"
		Me.Command1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command1.Text = "Bye"
		Me.Command1.Size = New System.Drawing.Size(72, 32)
		Me.Command1.Location = New System.Drawing.Point(536, 274)
		Me.Command1.TabIndex = 4
		Me.Command1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command1.BackColor = System.Drawing.SystemColors.Control
		Me.Command1.CausesValidation = True
		Me.Command1.Enabled = True
		Me.Command1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command1.TabStop = True
		Me.Command1.Name = "Command1"
		Me.Frame2.Text = "Public Var"
		Me.Frame2.Size = New System.Drawing.Size(195, 265)
		Me.Frame2.Location = New System.Drawing.Point(209, 2)
		Me.Frame2.TabIndex = 2
		Me.Frame2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame2.BackColor = System.Drawing.SystemColors.Control
		Me.Frame2.Enabled = True
		Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame2.Visible = True
		Me.Frame2.Name = "Frame2"
		Me.List2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.List2.Size = New System.Drawing.Size(176, 231)
		Me.List2.Location = New System.Drawing.Point(8, 16)
		Me.List2.TabIndex = 3
		Me.List2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.List2.BackColor = System.Drawing.SystemColors.Window
		Me.List2.CausesValidation = True
		Me.List2.Enabled = True
		Me.List2.ForeColor = System.Drawing.SystemColors.WindowText
		Me.List2.IntegralHeight = True
		Me.List2.Cursor = System.Windows.Forms.Cursors.Default
		Me.List2.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.List2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.List2.Sorted = False
		Me.List2.TabStop = True
		Me.List2.Visible = True
		Me.List2.MultiColumn = False
		Me.List2.Name = "List2"
		Me.Frame1.Text = "Socket"
		Me.Frame1.Size = New System.Drawing.Size(195, 265)
		Me.Frame1.Location = New System.Drawing.Point(6, 2)
		Me.Frame1.TabIndex = 0
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Name = "Frame1"
		Me.List1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.List1.Size = New System.Drawing.Size(176, 231)
		Me.List1.Location = New System.Drawing.Point(8, 16)
		Me.List1.TabIndex = 1
		Me.List1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.List1.BackColor = System.Drawing.SystemColors.Window
		Me.List1.CausesValidation = True
		Me.List1.Enabled = True
		Me.List1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.List1.IntegralHeight = True
		Me.List1.Cursor = System.Windows.Forms.Cursors.Default
		Me.List1.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.List1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.List1.Sorted = False
		Me.List1.TabStop = True
		Me.List1.Visible = True
		Me.List1.MultiColumn = False
		Me.List1.Name = "List1"
		Me.Timer1.Enabled = False
		Me.Timer1.Interval = 100
		Me.Controls.Add(Frame3)
		Me.Controls.Add(Command1)
		Me.Controls.Add(Frame2)
		Me.Controls.Add(Frame1)
		Me.Frame3.Controls.Add(List3)
		Me.Frame2.Controls.Add(List2)
		Me.Frame1.Controls.Add(List1)
	End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As FrmSysDbg
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As FrmSysDbg
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New FrmSysDbg()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Timer1.Enabled = False
		Me.Close()
	End Sub
	
	Private Sub FrmSysDbg_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Timer1.Enabled = True
	End Sub
	
	'UPGRADE_WARNING: Form event FrmSysDbg.Unload has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2065"'
	Private Sub FrmSysDbg_Closed(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Closed
		Timer1.Enabled = False
	End Sub
	
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
		Dim Form_Renamed As Object
		Dim u As Object
		On Error GoTo ErrInt
		List1.Items.Clear()
		For u = 1 To UniAgents.SockCount
			'UPGRADE_ISSUE: VBControlExtender property UniAgents.Socks.Index was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			List1.Items.Add("Socket > " & UniAgents.Socks(u).CtlHandle & " (" & UniAgents.Socks(u).Index & ")")
		Next u
		
		List2.Items.Clear()
		List2.Items.Add("cbUser > " & CbUserName)
		List2.Items.Add("cbAkses > " & CbUserAccess)
		List2.Items.Add("cbDemo > " & CbDemoMode)
		List2.Items.Add("cbDrvStr > " & CbDrvStr)
		List2.Items.Add("cbMsgRcv > " & CbMsgRcv)
		List2.Items.Add("cbConsole > " & CbConsole)
		List2.Items.Add("cbLogUser > " & CbLogUser)
		List2.Items.Add("lSock > " & lSock)
		
		List3.Items.Clear()
		'UPGRADE_ISSUE: Forms collection was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2068"'
		For	Each Form_Renamed In Forms
			'UPGRADE_WARNING: Couldn't resolve default property of object Form.Name. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			List3.Items.Add(Form_Renamed.Name)
		Next Form_Renamed
		Exit Sub
ErrInt: 
		Timer1.Enabled = False
		FrmMain.DefInstance.Text = "Internal Process Viewer - Error Detected !"
		ErrLog(Err, "Debug Windows - Timer1_Timer")
	End Sub
End Class