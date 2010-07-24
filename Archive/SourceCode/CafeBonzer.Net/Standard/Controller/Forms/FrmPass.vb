Option Strict Off
Option Explicit On
Friend Class FrmPass
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
	Public WithEvents LblInfo As Label3D
	Public WithEvents Line3D As Line3D
	Public WithEvents TxtPass As System.Windows.Forms.TextBox
	Public WithEvents TxtUser As System.Windows.Forms.TextBox
	Public WithEvents LblPass As System.Windows.Forms.Label
	Public WithEvents LblUser As System.Windows.Forms.Label
	Public WithEvents LblBuild As System.Windows.Forms.Label
	Public WithEvents imgLogo As System.Windows.Forms.PictureBox
	Public WithEvents LblCopy As System.Windows.Forms.Label
	Public WithEvents imgPass As System.Windows.Forms.PictureBox
	Public WithEvents imgBg As System.Windows.Forms.PictureBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmPass))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.LblInfo = New Label3D
		Me.Line3D = New Line3D
		Me.TxtPass = New System.Windows.Forms.TextBox
		Me.TxtUser = New System.Windows.Forms.TextBox
		Me.LblPass = New System.Windows.Forms.Label
		Me.LblUser = New System.Windows.Forms.Label
		Me.LblBuild = New System.Windows.Forms.Label
		Me.imgLogo = New System.Windows.Forms.PictureBox
		Me.LblCopy = New System.Windows.Forms.Label
		Me.imgPass = New System.Windows.Forms.PictureBox
		Me.imgBg = New System.Windows.Forms.PictureBox
		Me.ControlBox = False
		Me.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.ClientSize = New System.Drawing.Size(270, 157)
		Me.Location = New System.Drawing.Point(17, 94)
		Me.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.KeyPreview = True
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
		Me.Enabled = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "FrmPass"
		Me.LblInfo.Size = New System.Drawing.Size(187, 14)
		Me.LblInfo.Location = New System.Drawing.Point(5, 138)
		Me.LblInfo.TabIndex = 4
		Me.LblInfo.Name = "LblInfo"
		Me.Line3D.Size = New System.Drawing.Size(267, 3)
		Me.Line3D.Location = New System.Drawing.Point(2, 130)
		Me.Line3D.TabIndex = 3
		Me.Line3D.horizon = -1
		Me.Line3D.Name = "Line3D"
		Me.TxtPass.AutoSize = False
		Me.TxtPass.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.TxtPass.BackColor = System.Drawing.Color.White
		Me.TxtPass.Font = New System.Drawing.Font("Wingdings", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
		Me.TxtPass.Size = New System.Drawing.Size(160, 25)
		Me.TxtPass.IMEMode = System.Windows.Forms.ImeMode.Disable
		Me.TxtPass.Location = New System.Drawing.Point(98, 98)
		Me.TxtPass.PasswordChar = ChrW(108)
		Me.TxtPass.TabIndex = 1
		Me.ToolTip1.SetToolTip(Me.TxtPass, "Press ESC to exit..")
		Me.TxtPass.AcceptsReturn = True
		Me.TxtPass.CausesValidation = True
		Me.TxtPass.Enabled = True
		Me.TxtPass.ForeColor = System.Drawing.SystemColors.WindowText
		Me.TxtPass.HideSelection = True
		Me.TxtPass.ReadOnly = False
		Me.TxtPass.Maxlength = 0
		Me.TxtPass.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.TxtPass.MultiLine = False
		Me.TxtPass.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.TxtPass.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.TxtPass.TabStop = True
		Me.TxtPass.Visible = True
		Me.TxtPass.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.TxtPass.Name = "TxtPass"
		Me.TxtUser.AutoSize = False
		Me.TxtUser.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.TxtUser.BackColor = System.Drawing.Color.White
		Me.TxtUser.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.TxtUser.Size = New System.Drawing.Size(160, 24)
		Me.TxtUser.IMEMode = System.Windows.Forms.ImeMode.Disable
		Me.TxtUser.Location = New System.Drawing.Point(98, 66)
		Me.TxtUser.TabIndex = 0
		Me.TxtUser.AcceptsReturn = True
		Me.TxtUser.CausesValidation = True
		Me.TxtUser.Enabled = True
		Me.TxtUser.ForeColor = System.Drawing.SystemColors.WindowText
		Me.TxtUser.HideSelection = True
		Me.TxtUser.ReadOnly = False
		Me.TxtUser.Maxlength = 0
		Me.TxtUser.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.TxtUser.MultiLine = False
		Me.TxtUser.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.TxtUser.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.TxtUser.TabStop = True
		Me.TxtUser.Visible = True
		Me.TxtUser.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.TxtUser.Name = "TxtUser"
		Me.LblPass.Text = "Password :"
		Me.LblPass.Size = New System.Drawing.Size(70, 19)
		Me.LblPass.Location = New System.Drawing.Point(27, 101)
		Me.LblPass.TabIndex = 7
		Me.LblPass.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.LblPass.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.LblPass.BackColor = System.Drawing.Color.Transparent
		Me.LblPass.Enabled = True
		Me.LblPass.ForeColor = System.Drawing.SystemColors.ControlText
		Me.LblPass.Cursor = System.Windows.Forms.Cursors.Default
		Me.LblPass.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.LblPass.UseMnemonic = True
		Me.LblPass.Visible = True
		Me.LblPass.AutoSize = False
		Me.LblPass.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.LblPass.Name = "LblPass"
		Me.LblUser.Text = "Username :"
		Me.LblUser.Size = New System.Drawing.Size(73, 19)
		Me.LblUser.Location = New System.Drawing.Point(23, 69)
		Me.LblUser.TabIndex = 6
		Me.LblUser.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.LblUser.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.LblUser.BackColor = System.Drawing.Color.Transparent
		Me.LblUser.Enabled = True
		Me.LblUser.ForeColor = System.Drawing.SystemColors.ControlText
		Me.LblUser.Cursor = System.Windows.Forms.Cursors.Default
		Me.LblUser.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.LblUser.UseMnemonic = True
		Me.LblUser.Visible = True
		Me.LblUser.AutoSize = False
		Me.LblUser.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.LblUser.Name = "LblUser"
		Me.LblBuild.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.LblBuild.Text = "Build 1.7.42"
		Me.LblBuild.Font = New System.Drawing.Font("Verdana", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.LblBuild.ForeColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me.LblBuild.Size = New System.Drawing.Size(71, 13)
		Me.LblBuild.Location = New System.Drawing.Point(194, 140)
		Me.LblBuild.TabIndex = 5
		Me.LblBuild.BackColor = System.Drawing.Color.Transparent
		Me.LblBuild.Enabled = True
		Me.LblBuild.Cursor = System.Windows.Forms.Cursors.Default
		Me.LblBuild.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.LblBuild.UseMnemonic = True
		Me.LblBuild.Visible = True
		Me.LblBuild.AutoSize = False
		Me.LblBuild.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.LblBuild.Name = "LblBuild"
		Me.imgLogo.Size = New System.Drawing.Size(200, 28)
		Me.imgLogo.Location = New System.Drawing.Point(39, 6)
		Me.imgLogo.Image = CType(resources.GetObject("imgLogo.Image"), System.Drawing.Image)
		Me.imgLogo.Enabled = True
		Me.imgLogo.Cursor = System.Windows.Forms.Cursors.Default
		Me.imgLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.imgLogo.Visible = True
		Me.imgLogo.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.imgLogo.Name = "imgLogo"
		Me.LblCopy.Text = "Nematix Technology© 1996-2002"
		Me.LblCopy.Font = New System.Drawing.Font("Verdana", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.LblCopy.ForeColor = System.Drawing.Color.White
		Me.LblCopy.Size = New System.Drawing.Size(182, 13)
		Me.LblCopy.Location = New System.Drawing.Point(9, 37)
		Me.LblCopy.TabIndex = 2
		Me.LblCopy.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.LblCopy.BackColor = System.Drawing.Color.Transparent
		Me.LblCopy.Enabled = True
		Me.LblCopy.Cursor = System.Windows.Forms.Cursors.Default
		Me.LblCopy.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.LblCopy.UseMnemonic = True
		Me.LblCopy.Visible = True
		Me.LblCopy.AutoSize = False
		Me.LblCopy.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.LblCopy.Name = "LblCopy"
		Me.imgPass.Size = New System.Drawing.Size(32, 32)
		Me.imgPass.Location = New System.Drawing.Point(5, 4)
		Me.imgPass.Image = CType(resources.GetObject("imgPass.Image"), System.Drawing.Image)
		Me.imgPass.Enabled = True
		Me.imgPass.Cursor = System.Windows.Forms.Cursors.Default
		Me.imgPass.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.imgPass.Visible = True
		Me.imgPass.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.imgPass.Name = "imgPass"
		Me.imgBg.Size = New System.Drawing.Size(310, 60)
		Me.imgBg.Location = New System.Drawing.Point(-1, -3)
		Me.imgBg.Image = CType(resources.GetObject("imgBg.Image"), System.Drawing.Image)
		Me.imgBg.Enabled = True
		Me.imgBg.Cursor = System.Windows.Forms.Cursors.Default
		Me.imgBg.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.imgBg.Visible = True
		Me.imgBg.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.imgBg.Name = "imgBg"
		Me.Controls.Add(LblInfo)
		Me.Controls.Add(Line3D)
		Me.Controls.Add(TxtPass)
		Me.Controls.Add(TxtUser)
		Me.Controls.Add(LblPass)
		Me.Controls.Add(LblUser)
		Me.Controls.Add(LblBuild)
		Me.Controls.Add(imgLogo)
		Me.Controls.Add(LblCopy)
		Me.Controls.Add(imgPass)
		Me.Controls.Add(imgBg)
	End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As FrmPass
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As FrmPass
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New FrmPass()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	
	Private Sub FrmPass_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyUp
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		If KeyCode = System.Windows.Forms.Keys.Return Then
			If CekPass(TxtUser, TxtPass) = True Then
				FrmPass.DefInstance.Close()
				FrmMain.DefInstance.Show()
				Exit Sub
			Else
				LblInfo.Caption = "Access denied !"
				TxtPass.Text = ""
				Exit Sub
			End If
		End If
		If KeyCode = System.Windows.Forms.Keys.Escape Then
			Keluar(False)
		End If
	End Sub
	
	Private Sub FrmPass_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		LblBuild.Text = "Build " & System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).FileMajorPart & "." & CbAppBuild
	End Sub
End Class