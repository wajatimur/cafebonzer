Option Strict Off
Option Explicit On
Friend Class FrmSysConsole
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
	Public WithEvents Stat1 As AxMSComctlLib.AxStatusBar
	Public WithEvents List1 As System.Windows.Forms.ListBox
	Public WithEvents Text2 As System.Windows.Forms.TextBox
	Public WithEvents Text1 As System.Windows.Forms.TextBox
	Public WithEvents Picture1 As System.Windows.Forms.Panel
	Public WithEvents Timer1 As System.Windows.Forms.Timer
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmSysConsole))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.Stat1 = New AxMSComctlLib.AxStatusBar
		Me.Picture1 = New System.Windows.Forms.Panel
		Me.List1 = New System.Windows.Forms.ListBox
		Me.Text2 = New System.Windows.Forms.TextBox
		Me.Text1 = New System.Windows.Forms.TextBox
		Me.Timer1 = New System.Windows.Forms.Timer(components)
		CType(Me.Stat1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.Text = "Cafebonzer - Terminal"
		Me.ClientSize = New System.Drawing.Size(378, 262)
		Me.Location = New System.Drawing.Point(4, 23)
		Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Icon = CType(resources.GetObject("FrmSysConsole.Icon"), System.Drawing.Icon)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "FrmSysConsole"
		Stat1.OcxState = CType(resources.GetObject("Stat1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.Stat1.Dock = System.Windows.Forms.DockStyle.Bottom
		Me.Stat1.Size = New System.Drawing.Size(378, 24)
		Me.Stat1.Location = New System.Drawing.Point(0, 238)
		Me.Stat1.TabIndex = 4
		Me.Stat1.Name = "Stat1"
		Me.Picture1.BackColor = System.Drawing.Color.Black
		Me.Picture1.Size = New System.Drawing.Size(374, 233)
		Me.Picture1.Location = New System.Drawing.Point(2, 3)
		Me.Picture1.TabIndex = 2
		Me.Picture1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Picture1.Dock = System.Windows.Forms.DockStyle.None
		Me.Picture1.CausesValidation = True
		Me.Picture1.Enabled = True
		Me.Picture1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Picture1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Picture1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Picture1.TabStop = True
		Me.Picture1.Visible = True
		Me.Picture1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Picture1.Name = "Picture1"
		Me.List1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.List1.BackColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me.List1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.List1.ForeColor = System.Drawing.Color.FromARGB(0, 192, 0)
		Me.List1.Size = New System.Drawing.Size(371, 187)
		Me.List1.Location = New System.Drawing.Point(-1, 21)
		Me.List1.TabIndex = 3
		Me.List1.TabStop = False
		Me.List1.CausesValidation = True
		Me.List1.Enabled = True
		Me.List1.IntegralHeight = True
		Me.List1.Cursor = System.Windows.Forms.Cursors.Default
		Me.List1.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.List1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.List1.Sorted = False
		Me.List1.Visible = True
		Me.List1.MultiColumn = False
		Me.List1.Name = "List1"
		Me.Text2.AutoSize = False
		Me.Text2.BackColor = System.Drawing.Color.FromARGB(255, 255, 128)
		Me.Text2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text2.ForeColor = System.Drawing.Color.FromARGB(0, 192, 0)
		Me.Text2.Size = New System.Drawing.Size(367, 19)
		Me.Text2.Location = New System.Drawing.Point(2, 209)
		Me.Text2.TabIndex = 0
		Me.Text2.AcceptsReturn = True
		Me.Text2.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text2.CausesValidation = True
		Me.Text2.Enabled = True
		Me.Text2.HideSelection = True
		Me.Text2.ReadOnly = False
		Me.Text2.Maxlength = 0
		Me.Text2.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text2.MultiLine = False
		Me.Text2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text2.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text2.TabStop = True
		Me.Text2.Visible = True
		Me.Text2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Text2.Name = "Text2"
		Me.Text1.AutoSize = False
		Me.Text1.BackColor = System.Drawing.Color.FromARGB(224, 224, 224)
		Me.Text1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text1.ForeColor = System.Drawing.Color.FromARGB(0, 192, 0)
		Me.Text1.Size = New System.Drawing.Size(371, 21)
		Me.Text1.Location = New System.Drawing.Point(-1, 0)
		Me.Text1.ReadOnly = True
		Me.Text1.MultiLine = True
		Me.Text1.TabIndex = 1
		Me.Text1.TabStop = False
		Me.Text1.AcceptsReturn = True
		Me.Text1.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text1.CausesValidation = True
		Me.Text1.Enabled = True
		Me.Text1.HideSelection = True
		Me.Text1.Maxlength = 0
		Me.Text1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text1.Visible = True
		Me.Text1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.Text1.Name = "Text1"
		Me.Timer1.Enabled = False
		Me.Timer1.Interval = 10
		Me.Controls.Add(Stat1)
		Me.Controls.Add(Picture1)
		Me.Picture1.Controls.Add(List1)
		Me.Picture1.Controls.Add(Text2)
		Me.Picture1.Controls.Add(Text1)
		CType(Me.Stat1, System.ComponentModel.ISupportInitialize).EndInit()
	End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As FrmSysConsole
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As FrmSysConsole
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New FrmSysConsole()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	Dim Mm As String
	Dim Asiap As Boolean
	Public CurSocket As Integer
	
	Sub wr(ByRef Ayat As String, Optional ByRef CuciDulu As Boolean = True)
		Mm = Ayat
		Timer1.Enabled = True
		Asiap = False
		If CuciDulu = True Then
			Text1.Text = ""
		End If
	End Sub
	
	Private Sub FrmSysConsole_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Text1.BackColor = System.Drawing.Color.Black
		Text2.BackColor = System.Drawing.Color.Black
		List1.BackColor = System.Drawing.Color.Black
		
		Echo = True
		Asiap = True
		CbConsole = True
		If SelText <> "" Then CurSocket = CInt(SelTag)
		wr("Welcome To Console")
	End Sub
	
	'UPGRADE_WARNING: Form event FrmSysConsole.QueryUnload has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2065"'
	Private Sub FrmSysConsole_Closing(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
		Dim Cancel As Short = eventArgs.Cancel
		CbConsole = False
		eventArgs.Cancel = Cancel
	End Sub
	
	'UPGRADE_WARNING: Event FrmSysConsole.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2075"'
	Private Sub FrmSysConsole_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		On Error Resume Next
		Picture1.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Width) - 170)
		Picture1.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) - (480 + VB6.PixelsToTwipsY(Stat1.Height)))
		Text1.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Picture1.Width) - 50)
		Text2.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Picture1.Width) - 50)
		Text2.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Picture1.Height) - (VB6.PixelsToTwipsY(Text2.Height) + 20))
		List1.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Picture1.Width) - 50)
		List1.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Picture1.Height) - (VB6.PixelsToTwipsY(Text1.Height) + VB6.PixelsToTwipsY(Text2.Height)))
	End Sub
	
	Private Sub Text2_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles Text2.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		If Asiap = False Then GoTo EventExitSub
		
		If KeyAscii = 13 And Text2.Text <> "" Then
			ProcessIn()
		End If
		If KeyAscii = 27 Then
			FrmSysConsole.DefInstance.Hide() : CbConsole = False
			'UPGRADE_NOTE: Object FrmSysConsole may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
			FrmSysConsole.DefInstance = Nothing
		End If
EventExitSub: 
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
		Dim rd As Object
		Dim pj As Object
		Static idx As Short
		'UPGRADE_WARNING: Couldn't resolve default property of object pj. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		pj = Len(Mm)
		'UPGRADE_WARNING: Couldn't resolve default property of object pj. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If idx = pj Then
			Timer1.Enabled = False
			Asiap = True
			idx = 0
			Exit Sub
		End If
		idx = idx + 1
		'UPGRADE_WARNING: Couldn't resolve default property of object rd. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		rd = Mid(Mm, idx, 1)
		'UPGRADE_WARNING: Couldn't resolve default property of object rd. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Text1.Text = Text1.Text & " " & rd
	End Sub
	
	
	Private Sub ProcessIn()
		If Text1.Text <> "" Then List1.Items.Add(Text1.Text)
		If Echo = True Then wr((Text2.Text)) Else Text1.Text = ""
		Fetch((Text2.Text))
		Text2.Text = ""
	End Sub
	
	
	Public Sub Fetch(ByRef arahan As String)
		If arahan = "/debug" Then FrmSysDbg.DefInstance.Show() : Exit Sub
		If arahan = "/flush" Then FrmSysConsole.DefInstance.List1.Items.Clear() : Exit Sub
		If Mid(arahan, 1, 8) = "/dkeydrv" Then DkeyVar(Mid(arahan, 10, 1)) : Exit Sub
		If Mid(arahan, 1, 5) = "/echo" Then DisEcho(Mid(arahan, 7)) : Exit Sub
		If Mid(arahan, 1, 6) = "/mesej" Then SendMesej(Mid(arahan, 8), CurSocket) : Exit Sub
		If Mid(arahan, 1, 4) = "/cur" Then CurrentHook() : Exit Sub
		If Mid(arahan, 1, 5) = "/hook" Then Hook(Mid(arahan, 7)) : Exit Sub
		If CurSocket <> 0 Then Send(CurSocket, arahan) : Exit Sub
		wr("Command Not Found")
	End Sub
End Class