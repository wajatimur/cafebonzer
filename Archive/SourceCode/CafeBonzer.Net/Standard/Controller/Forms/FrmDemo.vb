Option Strict Off
Option Explicit On
Friend Class FrmSysDemo
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
	Public WithEvents pbDay As AxMSComctlLib.AxProgressBar
	Public WithEvents Image1 As System.Windows.Forms.PictureBox
	Public WithEvents lbDayleft As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents _MainBtn_0 As XpButton
	Public WithEvents _MainBtn_1 As XpButton
	Public WithEvents _MainBtn_2 As XpButton
	Public WithEvents MainBtn As XpButtonArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmSysDemo))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.pbDay = New AxMSComctlLib.AxProgressBar
		Me.Image1 = New System.Windows.Forms.PictureBox
		Me.lbDayleft = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me._MainBtn_0 = New XpButton
		Me._MainBtn_1 = New XpButton
		Me._MainBtn_2 = New XpButton
		Me.MainBtn = New XpButtonArray(components)
		CType(Me.pbDay, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.MainBtn, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.ControlBox = False
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.ClientSize = New System.Drawing.Size(268, 183)
		Me.Location = New System.Drawing.Point(17, 94)
		Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.KeyPreview = True
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
		Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.Enabled = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "FrmSysDemo"
		Me.Frame1.Size = New System.Drawing.Size(260, 149)
		Me.Frame1.Location = New System.Drawing.Point(5, -2)
		Me.Frame1.TabIndex = 0
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Name = "Frame1"
		pbDay.OcxState = CType(resources.GetObject("pbDay.OcxState"), System.Windows.Forms.AxHost.State)
		Me.pbDay.Size = New System.Drawing.Size(245, 13)
		Me.pbDay.Location = New System.Drawing.Point(7, 129)
		Me.pbDay.TabIndex = 3
		Me.pbDay.Name = "pbDay"
		Me.Image1.Size = New System.Drawing.Size(200, 28)
		Me.Image1.Location = New System.Drawing.Point(3, 13)
		Me.Image1.Image = CType(resources.GetObject("Image1.Image"), System.Drawing.Image)
		Me.Image1.Enabled = True
		Me.Image1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Image1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.Image1.Visible = True
		Me.Image1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Image1.Name = "Image1"
		Me.lbDayleft.Text = "0 Days Left"
		Me.lbDayleft.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lbDayleft.Size = New System.Drawing.Size(87, 16)
		Me.lbDayleft.Location = New System.Drawing.Point(8, 108)
		Me.lbDayleft.TabIndex = 2
		Me.lbDayleft.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lbDayleft.BackColor = System.Drawing.Color.Transparent
		Me.lbDayleft.Enabled = True
		Me.lbDayleft.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lbDayleft.Cursor = System.Windows.Forms.Cursors.Default
		Me.lbDayleft.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lbDayleft.UseMnemonic = True
		Me.lbDayleft.Visible = True
		Me.lbDayleft.AutoSize = False
		Me.lbDayleft.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lbDayleft.Name = "lbDayleft"
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label1.Text = "Welcome to Cafebonzer. This is an unregistered (trial) version of CafeBonzer. You only can use it for 9 days."
		Me.Label1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Size = New System.Drawing.Size(240, 53)
		Me.Label1.Location = New System.Drawing.Point(8, 49)
		Me.Label1.TabIndex = 1
		Me.Label1.BackColor = System.Drawing.SystemColors.Control
		Me.Label1.Enabled = True
		Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me._MainBtn_0.Size = New System.Drawing.Size(59, 29)
		Me._MainBtn_0.Location = New System.Drawing.Point(63, 151)
		Me._MainBtn_0.TabIndex = 4
		Me.ToolTip1.SetToolTip(Me._MainBtn_0, "Buy CafeBonzer.")
		Me._MainBtn_0.TX = "Buy"
		Me._MainBtn_0.ENAB = -1
		Me._MainBtn_0.COLTYPE = 1
		Me._MainBtn_0.FOCUSR = -1
		Me._MainBtn_0.BCOL = 12632256
		Me._MainBtn_0.BCOLO = 12632256
		Me._MainBtn_0.FCOL = 0
		Me._MainBtn_0.FCOLO = 0
		Me._MainBtn_0.MCOL = 16777215
		Me._MainBtn_0.MPTR = 1
		Me._MainBtn_0.MICON = 0
		Me._MainBtn_0.PICN = 0
		Me._MainBtn_0.UMCOL = -1
		Me._MainBtn_0.SOFT = 0
		Me._MainBtn_0.PICPOS = 0
		Me._MainBtn_0.NGREY = 0
		Me._MainBtn_0.FX = 0
		Me._MainBtn_0.HAND = 0
		Me._MainBtn_0.CHECK = 0
		Me._MainBtn_0.Name = "_MainBtn_0"
		Me._MainBtn_1.Size = New System.Drawing.Size(58, 29)
		Me._MainBtn_1.Location = New System.Drawing.Point(123, 151)
		Me._MainBtn_1.TabIndex = 5
		Me.ToolTip1.SetToolTip(Me._MainBtn_1, "Evaluate CafeBonzer.")
		Me._MainBtn_1.TX = "Try"
		Me._MainBtn_1.ENAB = -1
		Me._MainBtn_1.COLTYPE = 1
		Me._MainBtn_1.FOCUSR = -1
		Me._MainBtn_1.BCOL = 12632256
		Me._MainBtn_1.BCOLO = 12632256
		Me._MainBtn_1.FCOL = 0
		Me._MainBtn_1.FCOLO = 0
		Me._MainBtn_1.MCOL = 16777215
		Me._MainBtn_1.MPTR = 1
		Me._MainBtn_1.MICON = 0
		Me._MainBtn_1.PICN = 0
		Me._MainBtn_1.UMCOL = -1
		Me._MainBtn_1.SOFT = 0
		Me._MainBtn_1.PICPOS = 0
		Me._MainBtn_1.NGREY = 0
		Me._MainBtn_1.FX = 0
		Me._MainBtn_1.HAND = 0
		Me._MainBtn_1.CHECK = 0
		Me._MainBtn_1.Name = "_MainBtn_1"
		Me._MainBtn_2.Size = New System.Drawing.Size(82, 29)
		Me._MainBtn_2.Location = New System.Drawing.Point(182, 151)
		Me._MainBtn_2.TabIndex = 6
		Me.ToolTip1.SetToolTip(Me._MainBtn_2, "Register or Obtain a Full Version.")
		Me._MainBtn_2.TX = "Register"
		Me._MainBtn_2.ENAB = -1
		Me._MainBtn_2.COLTYPE = 1
		Me._MainBtn_2.FOCUSR = -1
		Me._MainBtn_2.BCOL = 12632256
		Me._MainBtn_2.BCOLO = 12632256
		Me._MainBtn_2.FCOL = 0
		Me._MainBtn_2.FCOLO = 0
		Me._MainBtn_2.MCOL = 16777215
		Me._MainBtn_2.MPTR = 1
		Me._MainBtn_2.MICON = 0
		Me._MainBtn_2.PICN = 0
		Me._MainBtn_2.UMCOL = -1
		Me._MainBtn_2.SOFT = 0
		Me._MainBtn_2.PICPOS = 0
		Me._MainBtn_2.NGREY = 0
		Me._MainBtn_2.FX = 0
		Me._MainBtn_2.HAND = 0
		Me._MainBtn_2.CHECK = 0
		Me._MainBtn_2.Name = "_MainBtn_2"
		Me.Controls.Add(Frame1)
		Me.Controls.Add(_MainBtn_0)
		Me.Controls.Add(_MainBtn_1)
		Me.Controls.Add(_MainBtn_2)
		Me.Frame1.Controls.Add(pbDay)
		Me.Frame1.Controls.Add(Image1)
		Me.Frame1.Controls.Add(lbDayleft)
		Me.Frame1.Controls.Add(Label1)
		Me.MainBtn.SetIndex(_MainBtn_0, CType(0, Short))
		Me.MainBtn.SetIndex(_MainBtn_1, CType(1, Short))
		Me.MainBtn.SetIndex(_MainBtn_2, CType(2, Short))
		CType(Me.MainBtn, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.pbDay, System.ComponentModel.ISupportInitialize).EndInit()
	End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As FrmSysDemo
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As FrmSysDemo
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New FrmSysDemo()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	
	Private Sub Asx1_ButtonClick(ByVal ButtonIndex As Short, ByVal ButtonKey As String)
		
	End Sub
	
	Private Sub FrmSysDemo_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim LngDayLeft, LngDayUse As Integer
		
		PutOnTop(FrmSysDemo.DefInstance.Handle.ToInt32)
		'UPGRADE_WARNING: Couldn't resolve default property of object SetGetDb(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		LngDayUse = SetGetDb("demoday", 10)
		If LngDayUse > 9 Then pbDay.Value = 9 : Exit Sub
		
		LngDayLeft = 9 - LngDayUse
		lbDayleft.Text = LngDayLeft & " Days Left"
		pbDay.Value = LngDayUse
	End Sub
	
	Private Sub RegisterIt()
		Dim sNamaDaftar, sNomborDaftar As String
		If CbDrvStr = "" Then CbDrvStr = "a:"
		
		If ValidateDisk(CbDrvStr) = True Then
			sNamaDaftar = GetName(CbDrvStr)
			sNomborDaftar = GetKey(CbDrvStr)
			SetSimpan("namadaftar", sNamaDaftar)
			SetSimpan("nombordaftar", sNomborDaftar)
			'DemoMode = False
			CbDemoMode = False
			SetSaveDb("demo", False)
			MsgBox(MB(6), MsgBoxStyle.OKOnly, CbMsgWarn)
			Me.Close()
		End If
	End Sub
	
	Private Sub MainBtn_Click(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles MainBtn.Click
		Dim Index As Short = MainBtn.GetIndex(Sender)
		Select Case Index
			Case 0
				'open website (registration page)
				'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
				If Len(Dir(VB6.GetPath & "\buy.htm", FileAttribute.Normal)) = 0 Then
					Call ShellExecute(Me.Handle.ToInt32, "open", "http://www.nematix.net", vbNullString, vbNullString, SW_NORMAL)
				Else
					Call ShellExecute(Me.Handle.ToInt32, "open", VB6.GetPath & "\buy.htm", vbNullString, vbNullString, SW_NORMAL)
				End If
				FrmSysDemo.DefInstance.Close()
				'UPGRADE_WARNING: Couldn't resolve default property of object SetGetDb(demoday, 10). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				If SetGetDb("demoday", 10) > 9 Then
					Keluar(False)
					End
				End If
			Case 1
				'UPGRADE_WARNING: Couldn't resolve default property of object SetGetDb(demoday, 10). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				If SetGetDb("demoday", 10) > 9 Then
					MsgBox(MB(3), MsgBoxStyle.OKOnly, "CafeBonzer")
					End
				End If
				Me.Close()
			Case 2
				Call RegisterIt()
		End Select
	End Sub
End Class