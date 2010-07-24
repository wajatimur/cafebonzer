Option Strict Off
Option Explicit On
Friend Class FrmAgnMsg
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
	Public WithEvents tosend As System.Windows.Forms.TextBox
	Public WithEvents BtnOk As XpButton
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents server As System.Windows.Forms.Label
	Public WithEvents Image1 As System.Windows.Forms.PictureBox
	Public WithEvents rcv As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmAgnMsg))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.tosend = New System.Windows.Forms.TextBox
		Me.BtnOk = New XpButton
		Me.Label1 = New System.Windows.Forms.Label
		Me.server = New System.Windows.Forms.Label
		Me.Image1 = New System.Windows.Forms.PictureBox
		Me.rcv = New System.Windows.Forms.Label
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Server - Message"
		Me.ClientSize = New System.Drawing.Size(301, 236)
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
		Me.Name = "FrmAgnMsg"
		Me.tosend.AutoSize = False
		Me.tosend.Font = New System.Drawing.Font("Verdana", 9!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.tosend.Size = New System.Drawing.Size(293, 64)
		Me.tosend.Location = New System.Drawing.Point(4, 128)
		Me.tosend.TabIndex = 1
		Me.ToolTip1.SetToolTip(Me.tosend, "Sila tekan enter selepas menulis mesej")
		Me.tosend.AcceptsReturn = True
		Me.tosend.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.tosend.BackColor = System.Drawing.SystemColors.Window
		Me.tosend.CausesValidation = True
		Me.tosend.Enabled = True
		Me.tosend.ForeColor = System.Drawing.SystemColors.WindowText
		Me.tosend.HideSelection = True
		Me.tosend.ReadOnly = False
		Me.tosend.Maxlength = 0
		Me.tosend.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.tosend.MultiLine = False
		Me.tosend.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.tosend.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.tosend.TabStop = True
		Me.tosend.Visible = True
		Me.tosend.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.tosend.Name = "tosend"
		Me.BtnOk.Size = New System.Drawing.Size(32, 32)
		Me.BtnOk.Location = New System.Drawing.Point(264, 200)
		Me.BtnOk.TabIndex = 4
		Me.ToolTip1.SetToolTip(Me.BtnOk, "Shutdown")
		Me.BtnOk.TX = ""
		Me.BtnOk.ENAB = -1
		Me.BtnOk.COLTYPE = 2
		Me.BtnOk.FOCUSR = -1
		Me.BtnOk.BCOL = 16777215
		Me.BtnOk.BCOLO = 16777215
		Me.BtnOk.FCOL = 0
		Me.BtnOk.FCOLO = 0
		Me.BtnOk.MCOL = 16777215
		Me.BtnOk.MPTR = 1
		Me.BtnOk.MICON = 0
		Me.BtnOk.PICN = 0
		Me.BtnOk.UMCOL = -1
		Me.BtnOk.SOFT = 0
		Me.BtnOk.PICPOS = 0
		Me.BtnOk.NGREY = 0
		Me.BtnOk.FX = 0
		Me.BtnOk.HAND = 0
		Me.BtnOk.CHECK = 0
		Me.BtnOk.Name = "BtnOk"
		Me.Label1.Text = "Please enter your message below.."
		Me.Label1.Size = New System.Drawing.Size(197, 13)
		Me.Label1.Location = New System.Drawing.Point(6, 106)
		Me.Label1.TabIndex = 3
		Me.Label1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.BackColor = System.Drawing.Color.Transparent
		Me.Label1.Enabled = True
		Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.server.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.server.ForeColor = System.Drawing.Color.FromARGB(0, 0, 192)
		Me.server.Size = New System.Drawing.Size(241, 22)
		Me.server.Location = New System.Drawing.Point(39, 7)
		Me.server.TabIndex = 2
		Me.server.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.server.BackColor = System.Drawing.Color.Transparent
		Me.server.Enabled = True
		Me.server.Cursor = System.Windows.Forms.Cursors.Default
		Me.server.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.server.UseMnemonic = True
		Me.server.Visible = True
		Me.server.AutoSize = False
		Me.server.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.server.Name = "server"
		Me.Image1.Size = New System.Drawing.Size(32, 32)
		Me.Image1.Location = New System.Drawing.Point(3, 1)
		Me.Image1.Image = CType(resources.GetObject("Image1.Image"), System.Drawing.Image)
		Me.Image1.Enabled = True
		Me.Image1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Image1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.Image1.Visible = True
		Me.Image1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Image1.Name = "Image1"
		Me.rcv.Font = New System.Drawing.Font("Verdana", 9!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.rcv.Size = New System.Drawing.Size(293, 64)
		Me.rcv.Location = New System.Drawing.Point(4, 35)
		Me.rcv.TabIndex = 0
		Me.rcv.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.rcv.BackColor = System.Drawing.Color.Transparent
		Me.rcv.Enabled = True
		Me.rcv.ForeColor = System.Drawing.SystemColors.ControlText
		Me.rcv.Cursor = System.Windows.Forms.Cursors.Default
		Me.rcv.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.rcv.UseMnemonic = True
		Me.rcv.Visible = True
		Me.rcv.AutoSize = False
		Me.rcv.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.rcv.Name = "rcv"
		Me.Controls.Add(tosend)
		Me.Controls.Add(BtnOk)
		Me.Controls.Add(Label1)
		Me.Controls.Add(server)
		Me.Controls.Add(Image1)
		Me.Controls.Add(rcv)
	End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As FrmAgnMsg
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As FrmAgnMsg
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New FrmAgnMsg()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	Private Sub Asx_ButtonClick(ByVal ButtonIndex As Short, ByVal ButtonKey As String)
		FrmAgnMsg.DefInstance.Hide()
		CbMsgRcv = False
	End Sub
	
	Private Sub tosend_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles tosend.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		If KeyCode = System.Windows.Forms.Keys.Return Then
			AgentSel.NetSend("//mesej:Server:" & tosend.Text)
			tosend.Text = ""
		ElseIf KeyCode = System.Windows.Forms.Keys.Escape Then 
			FrmAgnMsg.DefInstance.Hide()
			CbMsgRcv = False
		End If
	End Sub
End Class