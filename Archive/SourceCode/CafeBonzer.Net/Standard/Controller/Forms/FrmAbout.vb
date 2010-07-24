Option Strict Off
Option Explicit On
Friend Class FrmAbout
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
	Public WithEvents Timer1 As System.Windows.Forms.Timer
	Public WithEvents lblnama As System.Windows.Forms.Label
	Public WithEvents lblkedai As System.Windows.Forms.Label
	Public WithEvents lblemail As System.Windows.Forms.Label
	Public WithEvents Frame2 As System.Windows.Forms.GroupBox
	Public WithEvents Label3D2 As Label3D
	Public WithEvents lblCopy As System.Windows.Forms.Label
	Public WithEvents ImgClose As System.Windows.Forms.PictureBox
	Public WithEvents sc As System.Windows.Forms.Panel
	Public WithEvents Image1 As System.Windows.Forms.PictureBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmAbout))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.Timer1 = New System.Windows.Forms.Timer(components)
		Me.Frame2 = New System.Windows.Forms.GroupBox
		Me.lblnama = New System.Windows.Forms.Label
		Me.lblkedai = New System.Windows.Forms.Label
		Me.lblemail = New System.Windows.Forms.Label
		Me.Label3D2 = New Label3D
		Me.sc = New System.Windows.Forms.Panel
		Me.lblCopy = New System.Windows.Forms.Label
		Me.ImgClose = New System.Windows.Forms.PictureBox
		Me.Image1 = New System.Windows.Forms.PictureBox
		Me.ControlBox = False
		Me.BackColor = System.Drawing.Color.Black
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.ClientSize = New System.Drawing.Size(352, 131)
		Me.Location = New System.Drawing.Point(3, 3)
		Me.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "FrmAbout"
		Me.Timer1.Enabled = False
		Me.Timer1.Interval = 7
		Me.Frame2.Text = "This Product Is Licensed To"
		Me.Frame2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Italic Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame2.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me.Frame2.Size = New System.Drawing.Size(249, 76)
		Me.Frame2.Location = New System.Drawing.Point(229, 250)
		Me.Frame2.TabIndex = 1
		Me.Frame2.BackColor = System.Drawing.SystemColors.Control
		Me.Frame2.Enabled = True
		Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame2.Visible = True
		Me.Frame2.Name = "Frame2"
		Me.lblnama.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblnama.Text = "Nama :"
		Me.lblnama.Size = New System.Drawing.Size(228, 15)
		Me.lblnama.Location = New System.Drawing.Point(9, 18)
		Me.lblnama.TabIndex = 4
		Me.lblnama.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblnama.BackColor = System.Drawing.Color.Transparent
		Me.lblnama.Enabled = True
		Me.lblnama.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblnama.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblnama.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblnama.UseMnemonic = True
		Me.lblnama.Visible = True
		Me.lblnama.AutoSize = False
		Me.lblnama.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblnama.Name = "lblnama"
		Me.lblkedai.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblkedai.Text = "Cc :"
		Me.lblkedai.Size = New System.Drawing.Size(227, 15)
		Me.lblkedai.Location = New System.Drawing.Point(10, 35)
		Me.lblkedai.TabIndex = 3
		Me.lblkedai.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblkedai.BackColor = System.Drawing.Color.Transparent
		Me.lblkedai.Enabled = True
		Me.lblkedai.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblkedai.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblkedai.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblkedai.UseMnemonic = True
		Me.lblkedai.Visible = True
		Me.lblkedai.AutoSize = False
		Me.lblkedai.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblkedai.Name = "lblkedai"
		Me.lblemail.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblemail.Text = "Email :"
		Me.lblemail.Size = New System.Drawing.Size(228, 15)
		Me.lblemail.Location = New System.Drawing.Point(10, 52)
		Me.lblemail.TabIndex = 2
		Me.lblemail.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblemail.BackColor = System.Drawing.Color.Transparent
		Me.lblemail.Enabled = True
		Me.lblemail.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblemail.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblemail.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblemail.UseMnemonic = True
		Me.lblemail.Visible = True
		Me.lblemail.AutoSize = False
		Me.lblemail.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblemail.Name = "lblemail"
		Me.Label3D2.Size = New System.Drawing.Size(137, 16)
		Me.Label3D2.Location = New System.Drawing.Point(50, 246)
		Me.Label3D2.TabIndex = 5
		Me.Label3D2.Name = "Label3D2"
		Me.sc.BackColor = System.Drawing.Color.Black
		Me.sc.Font = New System.Drawing.Font("Verdana", 9!, System.Drawing.FontStyle.Italic Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.sc.ForeColor = System.Drawing.SystemColors.WindowText
		Me.sc.Size = New System.Drawing.Size(347, 127)
		Me.sc.Location = New System.Drawing.Point(3, 2)
		Me.sc.TabIndex = 0
		Me.sc.Dock = System.Windows.Forms.DockStyle.None
		Me.sc.CausesValidation = True
		Me.sc.Enabled = True
		Me.sc.Cursor = System.Windows.Forms.Cursors.Default
		Me.sc.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.sc.TabStop = True
		Me.sc.Visible = True
		Me.sc.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.sc.Name = "sc"
		Me.lblCopy.Text = "Copyright Nematix Technology© 1996-2002"
		Me.lblCopy.Font = New System.Drawing.Font("Verdana", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblCopy.ForeColor = System.Drawing.Color.White
		Me.lblCopy.Size = New System.Drawing.Size(237, 13)
		Me.lblCopy.Location = New System.Drawing.Point(1, 112)
		Me.lblCopy.TabIndex = 6
		Me.lblCopy.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblCopy.BackColor = System.Drawing.Color.Transparent
		Me.lblCopy.Enabled = True
		Me.lblCopy.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblCopy.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblCopy.UseMnemonic = True
		Me.lblCopy.Visible = True
		Me.lblCopy.AutoSize = False
		Me.lblCopy.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblCopy.Name = "lblCopy"
		Me.ImgClose.Size = New System.Drawing.Size(32, 32)
		Me.ImgClose.Location = New System.Drawing.Point(316, 96)
		Me.ImgClose.Image = CType(resources.GetObject("ImgClose.Image"), System.Drawing.Image)
		Me.ToolTip1.SetToolTip(Me.ImgClose, "In Business Time Is Money")
		Me.ImgClose.Enabled = True
		Me.ImgClose.Cursor = System.Windows.Forms.Cursors.Default
		Me.ImgClose.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.ImgClose.Visible = True
		Me.ImgClose.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.ImgClose.Name = "ImgClose"
		Me.Image1.Size = New System.Drawing.Size(153, 16)
		Me.Image1.Location = New System.Drawing.Point(47, 263)
		Me.Image1.Image = CType(resources.GetObject("Image1.Image"), System.Drawing.Image)
		Me.Image1.Enabled = True
		Me.Image1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Image1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.Image1.Visible = True
		Me.Image1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Image1.Name = "Image1"
		Me.Controls.Add(Frame2)
		Me.Controls.Add(Label3D2)
		Me.Controls.Add(sc)
		Me.Controls.Add(Image1)
		Me.Frame2.Controls.Add(lblnama)
		Me.Frame2.Controls.Add(lblkedai)
		Me.Frame2.Controls.Add(lblemail)
		Me.sc.Controls.Add(lblCopy)
		Me.sc.Controls.Add(ImgClose)
	End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As FrmAbout
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As FrmAbout
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New FrmAbout()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	Dim xHalf, yHalf As Integer
	Dim Quat1, Quat2 As Integer
	
	Private idx As Integer
	Private idxx As Integer
	Private idxName As Integer
	
	Private Sub FrmAbout_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		xHalf = VB6.PixelsToTwipsX(sc.Width) / 2
		yHalf = VB6.PixelsToTwipsY(sc.Height) / 2
		Quat1 = xHalf / 2
		Quat2 = Quat1 + xHalf
		
		Timer1.Enabled = True
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		lblnama.Text = SetAmbil("namadaftar")
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		lblkedai.Text = SetAmbil("namacc")
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		lblemail.Text = SetAmbil("emailpengguna")
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(demo). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If SetAmbil("demo") = "True" Then FrmMain.DefInstance.Text = FrmMain.DefInstance.Text & " UNREGISTERED"
	End Sub
	
	Private Sub ImgClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ImgClose.Click
		Timer1.Enabled = False
		'UPGRADE_NOTE: Object FrmAbout may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
		FrmAbout.DefInstance = Nothing
		Me.Close()
	End Sub
	
	Private Sub sc_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles sc.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		yHalf = Y
	End Sub
	
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
		Dim curX, curY As Integer
		Dim curXX, curYY As Integer
		Dim nameStr(12) As String
		Dim nameStr2(12) As String
		
		nameStr(0) = "[ Azri Jamil ]"
		nameStr(1) = "wajatimur@bootbox.net"
		nameStr(2) = "Nematix Technology"
		nameStr(3) = "registered to"
		nameStr(4) = "> user"
		nameStr(5) = "> organisation"
		nameStr(6) = "> email"
		nameStr(7) = "thanks to"
		nameStr(8) = "> maui"
		nameStr(9) = "> bent/toilet"
		nameStr(10) = "> lemang/lembing"
		nameStr(11) = "> adie/comot"
		nameStr(12) = "[ end ]"
		
		nameStr2(0) = "programmer/author"
		nameStr2(1) = "email"
		nameStr2(2) = "copyright"
		nameStr2(3) = ""
		nameStr2(4) = lblnama.Text
		nameStr2(5) = lblkedai.Text
		nameStr2(6) = lblemail.Text
		nameStr2(7) = ""
		nameStr2(8) = "bsd guru"
		nameStr2(9) = "script"
		nameStr2(10) = "tester"
		nameStr2(11) = "tester"
		nameStr2(12) = ""
		
		'UPGRADE_ISSUE: PictureBox method sc.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		sc.Cls()
		
		'>> Line 1 Counter
		If idxx >= xHalf And idxx <= xHalf + Quat1 Then
			'walking speed
			idxx = idxx + 30
		ElseIf idx >= xHalf + Quat1 Then 
			'walkout speed
			idxx = idxx + 650
		Else
			'walkin speed
			idxx = idxx + 300
		End If
		
		curXX = VB6.PixelsToTwipsX(sc.Width) - idxx
		curYY = yHalf - 150
		
		'Line 2 Counter
		If idx >= xHalf And idx <= xHalf + Quat1 Then
			'walking speed
			idx = idx + 35
		ElseIf idx >= xHalf + Quat1 Then 
			'walkout speed
			idx = idx + 650
		Else
			'walking speed
			idx = idx + 450
		End If
		
		curX = VB6.PixelsToTwipsX(sc.Width) - idx
		curY = yHalf - 150
		
		'>> Line 1
		'UPGRADE_ISSUE: PictureBox property sc.CurrentX was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		sc.CurrentX = curXX + 100
		'UPGRADE_ISSUE: PictureBox property sc.CurrentY was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		sc.CurrentY = curYY - 250
		sc.ForeColor = System.Drawing.Color.White
		sc.Font = VB6.FontChangeBold(sc.Font, False)
		'UPGRADE_ISSUE: PictureBox method sc.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		sc.Print(nameStr2(idxName))
		
		'>> Line 2
		'UPGRADE_ISSUE: PictureBox property sc.CurrentX was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		sc.CurrentX = curX - (Rnd() * 90)
		'UPGRADE_ISSUE: PictureBox property sc.CurrentY was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		sc.CurrentY = curY - (Rnd() * 90)
		sc.ForeColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
		sc.Font = VB6.FontChangeBold(sc.Font, False)
		'UPGRADE_ISSUE: PictureBox method sc.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		sc.Print(nameStr(idxName))
		
		'UPGRADE_ISSUE: PictureBox property sc.CurrentX was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		sc.CurrentX = curX + (Rnd() * 90)
		'UPGRADE_ISSUE: PictureBox property sc.CurrentY was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		sc.CurrentY = curY + (Rnd() * 90)
		sc.ForeColor = System.Drawing.ColorTranslator.FromOle(&H808080)
		sc.Font = VB6.FontChangeBold(sc.Font, False)
		'UPGRADE_ISSUE: PictureBox method sc.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		sc.Print(nameStr(idxName))
		
		'UPGRADE_ISSUE: PictureBox property sc.CurrentX was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		sc.CurrentX = curX
		'UPGRADE_ISSUE: PictureBox property sc.CurrentY was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		sc.CurrentY = curY - (Rnd() * 10)
		sc.ForeColor = System.Drawing.Color.White
		sc.Font = VB6.FontChangeBold(sc.Font, True)
		'UPGRADE_ISSUE: PictureBox method sc.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		sc.Print(nameStr(idxName))
		
		'AntiAlias 1, 600, sc.Width, 1200, 4
		
		'UPGRADE_ISSUE: PictureBox method sc.TextWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		If idxx >= VB6.PixelsToTwipsX(sc.Width) + (sc.TextWidth(nameStr(idxName))) Then
			idx = 0
			idxx = 0
			idxName = idxName + 1
			If idxName = UBound(nameStr) + 1 Then idxName = 0
		End If
	End Sub
	
	'UPGRADE_NOTE: step was upgraded to step_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
	Sub AntiAlias(ByRef X1 As Integer, ByRef Y1 As Integer, ByRef X2 As Integer, ByRef Y2 As Integer, ByRef step_Renamed As Integer)
		Dim cp As Object
		Dim xx As Object
		Dim yy As Object
		Dim i As Object
		Dim Avg As Object
		Dim X As Object
		Dim Y As Object
		For Y = Y1 To Y2 Step 100
			For X = X1 To X2 Step 100
				'UPGRADE_WARNING: Couldn't resolve default property of object Avg. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				Avg = 0
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				i = 0
				'UPGRADE_WARNING: Couldn't resolve default property of object Y. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				For yy = Y - step_Renamed To Y + step_Renamed
					'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					For xx = X - step_Renamed To X + step_Renamed
						'UPGRADE_ISSUE: PictureBox method sc.Point was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
						'UPGRADE_WARNING: Couldn't resolve default property of object cp. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
						cp = sc.Point(xx, yy)
						'UPGRADE_WARNING: Couldn't resolve default property of object cp. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
						'UPGRADE_WARNING: Couldn't resolve default property of object Avg. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
						Avg = Avg + (cp * cp)
						'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
						i = i + 1
					Next 
				Next 
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Avg. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				Avg = System.Math.Sqrt(Avg / i)
				'UPGRADE_ISSUE: PictureBox method sc.PSet was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
				sc.PSet (X, Y), Avg
			Next 
		Next 
		
	End Sub
End Class