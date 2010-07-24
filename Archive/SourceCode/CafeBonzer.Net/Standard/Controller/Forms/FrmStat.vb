Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class FrmStat
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
	Public WithEvents _StatLine_1 As Line3D
	Public WithEvents cbHari As System.Windows.Forms.ComboBox
	Public WithEvents cbBulan As System.Windows.Forms.ComboBox
	Public WithEvents cbTahun As System.Windows.Forms.ComboBox
	Public WithEvents StatBtn As XpButton
	Public WithEvents StatSesBtn As XpButton
	Public WithEvents _StatLbl_2 As System.Windows.Forms.Label
	Public WithEvents _StatLbl_3 As System.Windows.Forms.Label
	Public WithEvents CurSession As System.Windows.Forms.Label
	Public WithEvents _StatLbl_1 As System.Windows.Forms.Label
	Public WithEvents _StatLbl_0 As System.Windows.Forms.Label
	Public WithEvents StatFme As System.Windows.Forms.GroupBox
	Public WithEvents StatBar As AxMSComctlLib.AxStatusBar
	Public WithEvents _GenHdr_0 As System.Windows.Forms.Label
	Public WithEvents _GenLbl_4 As System.Windows.Forms.Label
	Public WithEvents _GenLbl_3 As System.Windows.Forms.Label
	Public WithEvents lblJualanPos As System.Windows.Forms.Label
	Public WithEvents lblJualanPc As System.Windows.Forms.Label
	Public WithEvents _GenLbl_2 As System.Windows.Forms.Label
	Public WithEvents _GenLbl_1 As System.Windows.Forms.Label
	Public WithEvents lblUntung As System.Windows.Forms.Label
	Public WithEvents lblJualan As System.Windows.Forms.Label
	Public WithEvents lblModal As System.Windows.Forms.Label
	Public WithEvents _GenLbl_0 As System.Windows.Forms.Label
	Public WithEvents _GenHdr_1 As System.Windows.Forms.Label
	Public WithEvents _GenLbl_6 As System.Windows.Forms.Label
	Public WithEvents lblServis As System.Windows.Forms.Label
	Public WithEvents lblPungut As System.Windows.Forms.Label
	Public WithEvents _GenLbl_5 As System.Windows.Forms.Label
	Public WithEvents _GenHdr_2 As System.Windows.Forms.Label
	Public WithEvents _GrafDay_0 As System.Windows.Forms.Label
	Public WithEvents _GrafDay_1 As System.Windows.Forms.Label
	Public WithEvents _GrafDay_2 As System.Windows.Forms.Label
	Public WithEvents _GrafDay_3 As System.Windows.Forms.Label
	Public WithEvents _GrafDay_4 As System.Windows.Forms.Label
	Public WithEvents _GrafDay_5 As System.Windows.Forms.Label
	Public WithEvents _GrafDay_6 As System.Windows.Forms.Label
	Public WithEvents GrafDock As System.Windows.Forms.Panel
	Public WithEvents _Bar1_0 As AxMSComctlLib.AxProgressBar
	Public WithEvents _Bar1_1 As AxMSComctlLib.AxProgressBar
	Public WithEvents _Bar1_2 As AxMSComctlLib.AxProgressBar
	Public WithEvents _Bar1_3 As AxMSComctlLib.AxProgressBar
	Public WithEvents _Bar1_4 As AxMSComctlLib.AxProgressBar
	Public WithEvents _Bar1_5 As AxMSComctlLib.AxProgressBar
	Public WithEvents _Bar1_6 As AxMSComctlLib.AxProgressBar
	Public WithEvents GrafHigh As System.Windows.Forms.Label
	Public WithEvents Graf1 As System.Windows.Forms.Panel
	Public WithEvents _StatLine_0 As Line3D
	Public WithEvents _StatTab_TabPage0 As System.Windows.Forms.TabPage
	Public WithEvents Lv1 As AxMSComctlLib.AxListView
	Public WithEvents _StatTab_TabPage1 As System.Windows.Forms.TabPage
	Public WithEvents Lv2 As AxMSComctlLib.AxListView
	Public WithEvents _StatTab_TabPage2 As System.Windows.Forms.TabPage
	Public WithEvents cbHari2 As System.Windows.Forms.ComboBox
	Public WithEvents Lv3 As AxMSComctlLib.AxListView
	Public WithEvents SlsBtn As XpButton
	Public WithEvents _SlsLbl_0 As System.Windows.Forms.Label
	Public WithEvents _StatTab_TabPage3 As System.Windows.Forms.TabPage
	Public WithEvents PosCmbDate As System.Windows.Forms.ComboBox
	Public WithEvents PosCmbItems As System.Windows.Forms.ComboBox
	Public WithEvents PosLV1 As AxMSComctlLib.AxListView
	Public WithEvents SrvBtn As XpButton
	Public WithEvents _SrvLbl_0 As System.Windows.Forms.Label
	Public WithEvents _SrvLbl_1 As System.Windows.Forms.Label
	Public WithEvents _StatTab_TabPage4 As System.Windows.Forms.TabPage
	Public WithEvents StatTab As System.Windows.Forms.TabControl
	Public WithEvents Bar1 As AxProgressBarArray.AxProgressBarArray
	Public WithEvents GenHdr As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents GenLbl As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents GrafDay As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents SlsLbl As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents SrvLbl As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents StatLbl As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents StatLine As Line3DArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmStat))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.StatFme = New System.Windows.Forms.GroupBox
		Me._StatLine_1 = New Line3D
		Me.cbHari = New System.Windows.Forms.ComboBox
		Me.cbBulan = New System.Windows.Forms.ComboBox
		Me.cbTahun = New System.Windows.Forms.ComboBox
		Me.StatBtn = New XpButton
		Me.StatSesBtn = New XpButton
		Me._StatLbl_2 = New System.Windows.Forms.Label
		Me._StatLbl_3 = New System.Windows.Forms.Label
		Me.CurSession = New System.Windows.Forms.Label
		Me._StatLbl_1 = New System.Windows.Forms.Label
		Me._StatLbl_0 = New System.Windows.Forms.Label
		Me.StatBar = New AxMSComctlLib.AxStatusBar
		Me.StatTab = New System.Windows.Forms.TabControl
		Me._StatTab_TabPage0 = New System.Windows.Forms.TabPage
		Me._GenHdr_0 = New System.Windows.Forms.Label
		Me._GenLbl_4 = New System.Windows.Forms.Label
		Me._GenLbl_3 = New System.Windows.Forms.Label
		Me.lblJualanPos = New System.Windows.Forms.Label
		Me.lblJualanPc = New System.Windows.Forms.Label
		Me._GenLbl_2 = New System.Windows.Forms.Label
		Me._GenLbl_1 = New System.Windows.Forms.Label
		Me.lblUntung = New System.Windows.Forms.Label
		Me.lblJualan = New System.Windows.Forms.Label
		Me.lblModal = New System.Windows.Forms.Label
		Me._GenLbl_0 = New System.Windows.Forms.Label
		Me._GenHdr_1 = New System.Windows.Forms.Label
		Me._GenLbl_6 = New System.Windows.Forms.Label
		Me.lblServis = New System.Windows.Forms.Label
		Me.lblPungut = New System.Windows.Forms.Label
		Me._GenLbl_5 = New System.Windows.Forms.Label
		Me._GenHdr_2 = New System.Windows.Forms.Label
		Me.Graf1 = New System.Windows.Forms.Panel
		Me.GrafDock = New System.Windows.Forms.Panel
		Me._GrafDay_0 = New System.Windows.Forms.Label
		Me._GrafDay_1 = New System.Windows.Forms.Label
		Me._GrafDay_2 = New System.Windows.Forms.Label
		Me._GrafDay_3 = New System.Windows.Forms.Label
		Me._GrafDay_4 = New System.Windows.Forms.Label
		Me._GrafDay_5 = New System.Windows.Forms.Label
		Me._GrafDay_6 = New System.Windows.Forms.Label
		Me._Bar1_0 = New AxMSComctlLib.AxProgressBar
		Me._Bar1_1 = New AxMSComctlLib.AxProgressBar
		Me._Bar1_2 = New AxMSComctlLib.AxProgressBar
		Me._Bar1_3 = New AxMSComctlLib.AxProgressBar
		Me._Bar1_4 = New AxMSComctlLib.AxProgressBar
		Me._Bar1_5 = New AxMSComctlLib.AxProgressBar
		Me._Bar1_6 = New AxMSComctlLib.AxProgressBar
		Me.GrafHigh = New System.Windows.Forms.Label
		Me._StatLine_0 = New Line3D
		Me._StatTab_TabPage1 = New System.Windows.Forms.TabPage
		Me.Lv1 = New AxMSComctlLib.AxListView
		Me._StatTab_TabPage2 = New System.Windows.Forms.TabPage
		Me.Lv2 = New AxMSComctlLib.AxListView
		Me._StatTab_TabPage3 = New System.Windows.Forms.TabPage
		Me.cbHari2 = New System.Windows.Forms.ComboBox
		Me.Lv3 = New AxMSComctlLib.AxListView
		Me.SlsBtn = New XpButton
		Me._SlsLbl_0 = New System.Windows.Forms.Label
		Me._StatTab_TabPage4 = New System.Windows.Forms.TabPage
		Me.PosCmbDate = New System.Windows.Forms.ComboBox
		Me.PosCmbItems = New System.Windows.Forms.ComboBox
		Me.PosLV1 = New AxMSComctlLib.AxListView
		Me.SrvBtn = New XpButton
		Me._SrvLbl_0 = New System.Windows.Forms.Label
		Me._SrvLbl_1 = New System.Windows.Forms.Label
		Me.Bar1 = New AxProgressBarArray.AxProgressBarArray(components)
		Me.GenHdr = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.GenLbl = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.GrafDay = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.SlsLbl = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.SrvLbl = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.StatLbl = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.StatLine = New Line3DArray(components)
		CType(Me.StatBar, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me._Bar1_0, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me._Bar1_1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me._Bar1_2, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me._Bar1_3, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me._Bar1_4, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me._Bar1_5, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me._Bar1_6, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Lv1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Lv2, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Lv3, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.PosLV1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Bar1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.GenHdr, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.GenLbl, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.GrafDay, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.SlsLbl, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.SrvLbl, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.StatLbl, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.StatLine, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Text = "CafeBonzer - Statistic"
		Me.ClientSize = New System.Drawing.Size(652, 512)
		Me.Location = New System.Drawing.Point(17, 113)
		Me.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Icon = CType(resources.GetObject("FrmStat.Icon"), System.Drawing.Icon)
		Me.KeyPreview = True
		Me.MaximizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
		Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
		Me.ControlBox = True
		Me.Enabled = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "FrmStat"
		Me.StatFme.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me.StatFme.Text = "Option"
		Me.StatFme.ForeColor = System.Drawing.Color.Blue
		Me.StatFme.Size = New System.Drawing.Size(645, 121)
		Me.StatFme.Location = New System.Drawing.Point(4, 365)
		Me.StatFme.TabIndex = 2
		Me.StatFme.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.StatFme.Enabled = True
		Me.StatFme.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.StatFme.Visible = True
		Me.StatFme.Name = "StatFme"
		Me._StatLine_1.Size = New System.Drawing.Size(3, 108)
		Me._StatLine_1.Location = New System.Drawing.Point(165, 10)
		Me._StatLine_1.TabIndex = 56
		Me._StatLine_1.horizon = 0
		Me._StatLine_1.Name = "_StatLine_1"
		Me.cbHari.BackColor = System.Drawing.Color.FromARGB(192, 192, 255)
		Me.cbHari.Size = New System.Drawing.Size(92, 21)
		Me.cbHari.Location = New System.Drawing.Point(64, 84)
		Me.cbHari.TabIndex = 49
		Me.cbHari.Text = "cbHari"
		Me.cbHari.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cbHari.CausesValidation = True
		Me.cbHari.Enabled = True
		Me.cbHari.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cbHari.IntegralHeight = True
		Me.cbHari.Cursor = System.Windows.Forms.Cursors.Default
		Me.cbHari.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cbHari.Sorted = False
		Me.cbHari.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cbHari.TabStop = True
		Me.cbHari.Visible = True
		Me.cbHari.Name = "cbHari"
		Me.cbBulan.BackColor = System.Drawing.Color.FromARGB(192, 192, 255)
		Me.cbBulan.Size = New System.Drawing.Size(92, 21)
		Me.cbBulan.Location = New System.Drawing.Point(64, 54)
		Me.cbBulan.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cbBulan.TabIndex = 5
		Me.cbBulan.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cbBulan.CausesValidation = True
		Me.cbBulan.Enabled = True
		Me.cbBulan.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cbBulan.IntegralHeight = True
		Me.cbBulan.Cursor = System.Windows.Forms.Cursors.Default
		Me.cbBulan.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cbBulan.Sorted = False
		Me.cbBulan.TabStop = True
		Me.cbBulan.Visible = True
		Me.cbBulan.Name = "cbBulan"
		Me.cbTahun.BackColor = System.Drawing.Color.FromARGB(192, 192, 255)
		Me.cbTahun.Size = New System.Drawing.Size(92, 21)
		Me.cbTahun.Location = New System.Drawing.Point(64, 23)
		Me.cbTahun.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cbTahun.TabIndex = 3
		Me.cbTahun.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cbTahun.CausesValidation = True
		Me.cbTahun.Enabled = True
		Me.cbTahun.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cbTahun.IntegralHeight = True
		Me.cbTahun.Cursor = System.Windows.Forms.Cursors.Default
		Me.cbTahun.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cbTahun.Sorted = False
		Me.cbTahun.TabStop = True
		Me.cbTahun.Visible = True
		Me.cbTahun.Name = "cbTahun"
		Me.StatBtn.Size = New System.Drawing.Size(35, 29)
		Me.StatBtn.Location = New System.Drawing.Point(604, 87)
		Me.StatBtn.TabIndex = 57
		Me.ToolTip1.SetToolTip(Me.StatBtn, "Save settings and exit.")
		Me.StatBtn.TX = ""
		Me.StatBtn.ENAB = -1
		Me.StatBtn.COLTYPE = 1
		Me.StatBtn.FOCUSR = -1
		Me.StatBtn.BCOL = 12632256
		Me.StatBtn.BCOLO = 12632256
		Me.StatBtn.FCOL = 0
		Me.StatBtn.FCOLO = 0
		Me.StatBtn.MCOL = 16777215
		Me.StatBtn.MPTR = 1
		Me.StatBtn.MICON = 0
		Me.StatBtn.PICN = 0
		Me.StatBtn.UMCOL = -1
		Me.StatBtn.SOFT = 0
		Me.StatBtn.PICPOS = 0
		Me.StatBtn.NGREY = 0
		Me.StatBtn.FX = 0
		Me.StatBtn.HAND = 0
		Me.StatBtn.CHECK = 0
		Me.StatBtn.Name = "StatBtn"
		Me.StatSesBtn.Size = New System.Drawing.Size(28, 24)
		Me.StatSesBtn.Location = New System.Drawing.Point(304, 38)
		Me.StatSesBtn.TabIndex = 58
		Me.ToolTip1.SetToolTip(Me.StatSesBtn, "Delete selected employee.")
		Me.StatSesBtn.TX = ""
		Me.StatSesBtn.ENAB = -1
		Me.StatSesBtn.COLTYPE = 1
		Me.StatSesBtn.FOCUSR = -1
		Me.StatSesBtn.BCOL = 12632256
		Me.StatSesBtn.BCOLO = 12632256
		Me.StatSesBtn.FCOL = 0
		Me.StatSesBtn.FCOLO = 0
		Me.StatSesBtn.MCOL = 16777215
		Me.StatSesBtn.MPTR = 1
		Me.StatSesBtn.MICON = 0
		Me.StatSesBtn.PICN = 0
		Me.StatSesBtn.UMCOL = -1
		Me.StatSesBtn.SOFT = 0
		Me.StatSesBtn.PICPOS = 0
		Me.StatSesBtn.NGREY = 0
		Me.StatSesBtn.FX = 0
		Me.StatSesBtn.HAND = 0
		Me.StatSesBtn.CHECK = 0
		Me.StatSesBtn.Name = "StatSesBtn"
		Me._StatLbl_2.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._StatLbl_2.Text = "Date :"
		Me._StatLbl_2.Size = New System.Drawing.Size(38, 13)
		Me._StatLbl_2.Location = New System.Drawing.Point(21, 85)
		Me._StatLbl_2.TabIndex = 50
		Me._StatLbl_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._StatLbl_2.BackColor = System.Drawing.Color.Transparent
		Me._StatLbl_2.Enabled = True
		Me._StatLbl_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._StatLbl_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._StatLbl_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._StatLbl_2.UseMnemonic = True
		Me._StatLbl_2.Visible = True
		Me._StatLbl_2.AutoSize = False
		Me._StatLbl_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._StatLbl_2.Name = "_StatLbl_2"
		Me._StatLbl_3.Text = "Current Session Date :"
		Me._StatLbl_3.Size = New System.Drawing.Size(137, 16)
		Me._StatLbl_3.Location = New System.Drawing.Point(179, 18)
		Me._StatLbl_3.TabIndex = 47
		Me._StatLbl_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._StatLbl_3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._StatLbl_3.BackColor = System.Drawing.Color.Transparent
		Me._StatLbl_3.Enabled = True
		Me._StatLbl_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._StatLbl_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._StatLbl_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._StatLbl_3.UseMnemonic = True
		Me._StatLbl_3.Visible = True
		Me._StatLbl_3.AutoSize = False
		Me._StatLbl_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._StatLbl_3.Name = "_StatLbl_3"
		Me.CurSession.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.CurSession.BackColor = System.Drawing.Color.FromARGB(192, 255, 192)
		Me.CurSession.Text = "12/12/2000"
		Me.CurSession.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.CurSession.ForeColor = System.Drawing.Color.Black
		Me.CurSession.Size = New System.Drawing.Size(109, 20)
		Me.CurSession.Location = New System.Drawing.Point(192, 40)
		Me.CurSession.TabIndex = 46
		Me.CurSession.Enabled = True
		Me.CurSession.Cursor = System.Windows.Forms.Cursors.Default
		Me.CurSession.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.CurSession.UseMnemonic = True
		Me.CurSession.Visible = True
		Me.CurSession.AutoSize = False
		Me.CurSession.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.CurSession.Name = "CurSession"
		Me._StatLbl_1.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._StatLbl_1.Text = "Month :"
		Me._StatLbl_1.Size = New System.Drawing.Size(44, 16)
		Me._StatLbl_1.Location = New System.Drawing.Point(15, 57)
		Me._StatLbl_1.TabIndex = 6
		Me._StatLbl_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._StatLbl_1.BackColor = System.Drawing.Color.Transparent
		Me._StatLbl_1.Enabled = True
		Me._StatLbl_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._StatLbl_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._StatLbl_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._StatLbl_1.UseMnemonic = True
		Me._StatLbl_1.Visible = True
		Me._StatLbl_1.AutoSize = False
		Me._StatLbl_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._StatLbl_1.Name = "_StatLbl_1"
		Me._StatLbl_0.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._StatLbl_0.Text = "Year :"
		Me._StatLbl_0.Size = New System.Drawing.Size(44, 16)
		Me._StatLbl_0.Location = New System.Drawing.Point(15, 26)
		Me._StatLbl_0.TabIndex = 4
		Me._StatLbl_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._StatLbl_0.BackColor = System.Drawing.Color.Transparent
		Me._StatLbl_0.Enabled = True
		Me._StatLbl_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._StatLbl_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._StatLbl_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._StatLbl_0.UseMnemonic = True
		Me._StatLbl_0.Visible = True
		Me._StatLbl_0.AutoSize = False
		Me._StatLbl_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._StatLbl_0.Name = "_StatLbl_0"
		StatBar.OcxState = CType(resources.GetObject("StatBar.OcxState"), System.Windows.Forms.AxHost.State)
		Me.StatBar.Dock = System.Windows.Forms.DockStyle.Bottom
		Me.StatBar.Size = New System.Drawing.Size(652, 23)
		Me.StatBar.Location = New System.Drawing.Point(0, 489)
		Me.StatBar.TabIndex = 1
		Me.StatBar.Name = "StatBar"
		Me.StatTab.Size = New System.Drawing.Size(647, 363)
		Me.StatTab.Location = New System.Drawing.Point(3, 1)
		Me.StatTab.TabIndex = 0
		Me.StatTab.Appearance = System.Windows.Forms.TabAppearance.FlatButtons
		Me.StatTab.ItemSize = New System.Drawing.Size(42, 18)
		Me.StatTab.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me.StatTab.ForeColor = System.Drawing.Color.FromARGB(0, 0, 255)
		Me.StatTab.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Underline Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.StatTab.Name = "StatTab"
		Me._StatTab_TabPage0.Text = "Monthly\Daily Overview"
		Me._GenHdr_0.BackColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me._GenHdr_0.Text = " Monthly Statistic"
		Me._GenHdr_0.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GenHdr_0.ForeColor = System.Drawing.Color.White
		Me._GenHdr_0.Size = New System.Drawing.Size(276, 17)
		Me._GenHdr_0.Location = New System.Drawing.Point(14, 40)
		Me._GenHdr_0.TabIndex = 17
		Me._GenHdr_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._GenHdr_0.Enabled = True
		Me._GenHdr_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._GenHdr_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GenHdr_0.UseMnemonic = True
		Me._GenHdr_0.Visible = True
		Me._GenHdr_0.AutoSize = False
		Me._GenHdr_0.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._GenHdr_0.Name = "_GenHdr_0"
		Me._GenLbl_4.Text = "Service && Maintenance :"
		Me._GenLbl_4.Size = New System.Drawing.Size(140, 13)
		Me._GenLbl_4.Location = New System.Drawing.Point(22, 198)
		Me._GenLbl_4.TabIndex = 18
		Me._GenLbl_4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GenLbl_4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._GenLbl_4.BackColor = System.Drawing.Color.Transparent
		Me._GenLbl_4.Enabled = True
		Me._GenLbl_4.ForeColor = System.Drawing.SystemColors.ControlText
		Me._GenLbl_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._GenLbl_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GenLbl_4.UseMnemonic = True
		Me._GenLbl_4.Visible = True
		Me._GenLbl_4.AutoSize = False
		Me._GenLbl_4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._GenLbl_4.Name = "_GenLbl_4"
		Me._GenLbl_3.Text = "PC Rent Sales :"
		Me._GenLbl_3.Size = New System.Drawing.Size(91, 13)
		Me._GenLbl_3.Location = New System.Drawing.Point(70, 169)
		Me._GenLbl_3.TabIndex = 19
		Me._GenLbl_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GenLbl_3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._GenLbl_3.BackColor = System.Drawing.Color.Transparent
		Me._GenLbl_3.Enabled = True
		Me._GenLbl_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._GenLbl_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._GenLbl_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GenLbl_3.UseMnemonic = True
		Me._GenLbl_3.Visible = True
		Me._GenLbl_3.AutoSize = False
		Me._GenLbl_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._GenLbl_3.Name = "_GenLbl_3"
		Me.lblJualanPos.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblJualanPos.BackColor = System.Drawing.Color.FromARGB(192, 255, 192)
		Me.lblJualanPos.Text = "RM 0.00"
		Me.lblJualanPos.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblJualanPos.ForeColor = System.Drawing.Color.Black
		Me.lblJualanPos.Size = New System.Drawing.Size(98, 20)
		Me.lblJualanPos.Location = New System.Drawing.Point(169, 195)
		Me.lblJualanPos.TabIndex = 20
		Me.lblJualanPos.Enabled = True
		Me.lblJualanPos.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblJualanPos.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblJualanPos.UseMnemonic = True
		Me.lblJualanPos.Visible = True
		Me.lblJualanPos.AutoSize = False
		Me.lblJualanPos.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.lblJualanPos.Name = "lblJualanPos"
		Me.lblJualanPc.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblJualanPc.BackColor = System.Drawing.Color.FromARGB(192, 255, 192)
		Me.lblJualanPc.Text = "RM 0.00"
		Me.lblJualanPc.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblJualanPc.ForeColor = System.Drawing.Color.Black
		Me.lblJualanPc.Size = New System.Drawing.Size(98, 20)
		Me.lblJualanPc.Location = New System.Drawing.Point(169, 166)
		Me.lblJualanPc.TabIndex = 21
		Me.lblJualanPc.Enabled = True
		Me.lblJualanPc.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblJualanPc.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblJualanPc.UseMnemonic = True
		Me.lblJualanPc.Visible = True
		Me.lblJualanPc.AutoSize = False
		Me.lblJualanPc.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.lblJualanPc.Name = "lblJualanPc"
		Me._GenLbl_2.Text = "Monthly Profit :"
		Me._GenLbl_2.Size = New System.Drawing.Size(121, 13)
		Me._GenLbl_2.Location = New System.Drawing.Point(23, 133)
		Me._GenLbl_2.TabIndex = 22
		Me._GenLbl_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GenLbl_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._GenLbl_2.BackColor = System.Drawing.Color.Transparent
		Me._GenLbl_2.Enabled = True
		Me._GenLbl_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._GenLbl_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._GenLbl_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GenLbl_2.UseMnemonic = True
		Me._GenLbl_2.Visible = True
		Me._GenLbl_2.AutoSize = False
		Me._GenLbl_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._GenLbl_2.Name = "_GenLbl_2"
		Me._GenLbl_1.Text = "Monthly Sales :"
		Me._GenLbl_1.Size = New System.Drawing.Size(121, 13)
		Me._GenLbl_1.Location = New System.Drawing.Point(23, 100)
		Me._GenLbl_1.TabIndex = 23
		Me._GenLbl_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GenLbl_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._GenLbl_1.BackColor = System.Drawing.Color.Transparent
		Me._GenLbl_1.Enabled = True
		Me._GenLbl_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._GenLbl_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._GenLbl_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GenLbl_1.UseMnemonic = True
		Me._GenLbl_1.Visible = True
		Me._GenLbl_1.AutoSize = False
		Me._GenLbl_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._GenLbl_1.Name = "_GenLbl_1"
		Me.lblUntung.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblUntung.BackColor = System.Drawing.Color.FromARGB(192, 255, 192)
		Me.lblUntung.Text = "RM 0.00"
		Me.lblUntung.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblUntung.ForeColor = System.Drawing.Color.Black
		Me.lblUntung.Size = New System.Drawing.Size(115, 21)
		Me.lblUntung.Location = New System.Drawing.Point(153, 130)
		Me.lblUntung.TabIndex = 24
		Me.lblUntung.Enabled = True
		Me.lblUntung.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblUntung.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblUntung.UseMnemonic = True
		Me.lblUntung.Visible = True
		Me.lblUntung.AutoSize = False
		Me.lblUntung.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.lblUntung.Name = "lblUntung"
		Me.lblJualan.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblJualan.BackColor = System.Drawing.Color.FromARGB(192, 255, 192)
		Me.lblJualan.Text = "RM 0.00"
		Me.lblJualan.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblJualan.ForeColor = System.Drawing.Color.Black
		Me.lblJualan.Size = New System.Drawing.Size(115, 21)
		Me.lblJualan.Location = New System.Drawing.Point(153, 98)
		Me.lblJualan.TabIndex = 25
		Me.lblJualan.Enabled = True
		Me.lblJualan.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblJualan.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblJualan.UseMnemonic = True
		Me.lblJualan.Visible = True
		Me.lblJualan.AutoSize = False
		Me.lblJualan.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.lblJualan.Name = "lblJualan"
		Me.lblModal.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblModal.BackColor = System.Drawing.Color.FromARGB(192, 255, 192)
		Me.lblModal.Text = "RM 0.00"
		Me.lblModal.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblModal.ForeColor = System.Drawing.Color.Black
		Me.lblModal.Size = New System.Drawing.Size(115, 21)
		Me.lblModal.Location = New System.Drawing.Point(153, 68)
		Me.lblModal.TabIndex = 26
		Me.lblModal.Enabled = True
		Me.lblModal.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblModal.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblModal.UseMnemonic = True
		Me.lblModal.Visible = True
		Me.lblModal.AutoSize = False
		Me.lblModal.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.lblModal.Name = "lblModal"
		Me._GenLbl_0.Text = "Monthly Overhead :"
		Me._GenLbl_0.Size = New System.Drawing.Size(121, 13)
		Me._GenLbl_0.Location = New System.Drawing.Point(24, 70)
		Me._GenLbl_0.TabIndex = 27
		Me._GenLbl_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GenLbl_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._GenLbl_0.BackColor = System.Drawing.Color.Transparent
		Me._GenLbl_0.Enabled = True
		Me._GenLbl_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._GenLbl_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._GenLbl_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GenLbl_0.UseMnemonic = True
		Me._GenLbl_0.Visible = True
		Me._GenLbl_0.AutoSize = False
		Me._GenLbl_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._GenLbl_0.Name = "_GenLbl_0"
		Me._GenHdr_1.BackColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me._GenHdr_1.Text = " Daily Average Graft"
		Me._GenHdr_1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GenHdr_1.ForeColor = System.Drawing.Color.White
		Me._GenHdr_1.Size = New System.Drawing.Size(309, 17)
		Me._GenHdr_1.Location = New System.Drawing.Point(322, 40)
		Me._GenHdr_1.TabIndex = 45
		Me._GenHdr_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._GenHdr_1.Enabled = True
		Me._GenHdr_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._GenHdr_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GenHdr_1.UseMnemonic = True
		Me._GenHdr_1.Visible = True
		Me._GenHdr_1.AutoSize = False
		Me._GenHdr_1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._GenHdr_1.Name = "_GenHdr_1"
		Me._GenLbl_6.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._GenLbl_6.Text = "Services Collection :"
		Me._GenLbl_6.Size = New System.Drawing.Size(124, 11)
		Me._GenLbl_6.Location = New System.Drawing.Point(38, 308)
		Me._GenLbl_6.TabIndex = 51
		Me._GenLbl_6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GenLbl_6.BackColor = System.Drawing.Color.Transparent
		Me._GenLbl_6.Enabled = True
		Me._GenLbl_6.ForeColor = System.Drawing.SystemColors.ControlText
		Me._GenLbl_6.Cursor = System.Windows.Forms.Cursors.Default
		Me._GenLbl_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GenLbl_6.UseMnemonic = True
		Me._GenLbl_6.Visible = True
		Me._GenLbl_6.AutoSize = False
		Me._GenLbl_6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._GenLbl_6.Name = "_GenLbl_6"
		Me.lblServis.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblServis.BackColor = System.Drawing.Color.FromARGB(192, 255, 192)
		Me.lblServis.Text = "RM 0.00"
		Me.lblServis.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblServis.ForeColor = System.Drawing.Color.Black
		Me.lblServis.Size = New System.Drawing.Size(98, 20)
		Me.lblServis.Location = New System.Drawing.Point(170, 306)
		Me.lblServis.TabIndex = 52
		Me.lblServis.Enabled = True
		Me.lblServis.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblServis.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblServis.UseMnemonic = True
		Me.lblServis.Visible = True
		Me.lblServis.AutoSize = False
		Me.lblServis.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.lblServis.Name = "lblServis"
		Me.lblPungut.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblPungut.BackColor = System.Drawing.Color.FromARGB(192, 255, 192)
		Me.lblPungut.Text = "RM 0.00"
		Me.lblPungut.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblPungut.ForeColor = System.Drawing.Color.Black
		Me.lblPungut.Size = New System.Drawing.Size(98, 20)
		Me.lblPungut.Location = New System.Drawing.Point(170, 273)
		Me.lblPungut.TabIndex = 53
		Me.lblPungut.Enabled = True
		Me.lblPungut.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblPungut.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblPungut.UseMnemonic = True
		Me.lblPungut.Visible = True
		Me.lblPungut.AutoSize = False
		Me.lblPungut.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.lblPungut.Name = "lblPungut"
		Me._GenLbl_5.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._GenLbl_5.Text = "PC Collection :"
		Me._GenLbl_5.Size = New System.Drawing.Size(123, 13)
		Me._GenLbl_5.Location = New System.Drawing.Point(39, 275)
		Me._GenLbl_5.TabIndex = 54
		Me._GenLbl_5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GenLbl_5.BackColor = System.Drawing.Color.Transparent
		Me._GenLbl_5.Enabled = True
		Me._GenLbl_5.ForeColor = System.Drawing.SystemColors.ControlText
		Me._GenLbl_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._GenLbl_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GenLbl_5.UseMnemonic = True
		Me._GenLbl_5.Visible = True
		Me._GenLbl_5.AutoSize = False
		Me._GenLbl_5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._GenLbl_5.Name = "_GenLbl_5"
		Me._GenHdr_2.BackColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me._GenHdr_2.Text = " Today's Collection"
		Me._GenHdr_2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GenHdr_2.ForeColor = System.Drawing.Color.White
		Me._GenHdr_2.Size = New System.Drawing.Size(276, 17)
		Me._GenHdr_2.Location = New System.Drawing.Point(14, 244)
		Me._GenHdr_2.TabIndex = 55
		Me._GenHdr_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._GenHdr_2.Enabled = True
		Me._GenHdr_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._GenHdr_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GenHdr_2.UseMnemonic = True
		Me._GenHdr_2.Visible = True
		Me._GenHdr_2.AutoSize = False
		Me._GenHdr_2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._GenHdr_2.Name = "_GenHdr_2"
		Me.Graf1.BackColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me.Graf1.Size = New System.Drawing.Size(282, 196)
		Me.Graf1.Location = New System.Drawing.Point(337, 67)
		Me.Graf1.TabIndex = 28
		Me.Graf1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Graf1.Dock = System.Windows.Forms.DockStyle.None
		Me.Graf1.CausesValidation = True
		Me.Graf1.Enabled = True
		Me.Graf1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Graf1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Graf1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Graf1.TabStop = True
		Me.Graf1.Visible = True
		Me.Graf1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Graf1.Name = "Graf1"
		Me.GrafDock.BackColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me.GrafDock.ForeColor = System.Drawing.SystemColors.WindowText
		Me.GrafDock.Size = New System.Drawing.Size(283, 28)
		Me.GrafDock.Location = New System.Drawing.Point(-3, 171)
		Me.GrafDock.TabIndex = 29
		Me.GrafDock.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.GrafDock.Dock = System.Windows.Forms.DockStyle.None
		Me.GrafDock.CausesValidation = True
		Me.GrafDock.Enabled = True
		Me.GrafDock.Cursor = System.Windows.Forms.Cursors.Default
		Me.GrafDock.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.GrafDock.TabStop = True
		Me.GrafDock.Visible = True
		Me.GrafDock.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.GrafDock.Name = "GrafDock"
		Me._GrafDay_0.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._GrafDay_0.Text = "A"
		Me._GrafDay_0.Font = New System.Drawing.Font("Verdana", 14.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GrafDay_0.ForeColor = System.Drawing.Color.FromARGB(192, 192, 255)
		Me._GrafDay_0.Size = New System.Drawing.Size(22, 22)
		Me._GrafDay_0.Location = New System.Drawing.Point(29, -2)
		Me._GrafDay_0.TabIndex = 36
		Me._GrafDay_0.BackColor = System.Drawing.Color.Transparent
		Me._GrafDay_0.Enabled = True
		Me._GrafDay_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._GrafDay_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GrafDay_0.UseMnemonic = True
		Me._GrafDay_0.Visible = True
		Me._GrafDay_0.AutoSize = False
		Me._GrafDay_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._GrafDay_0.Name = "_GrafDay_0"
		Me._GrafDay_1.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._GrafDay_1.Text = "I"
		Me._GrafDay_1.Font = New System.Drawing.Font("Verdana", 14.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GrafDay_1.ForeColor = System.Drawing.Color.FromARGB(192, 192, 255)
		Me._GrafDay_1.Size = New System.Drawing.Size(22, 22)
		Me._GrafDay_1.Location = New System.Drawing.Point(64, -2)
		Me._GrafDay_1.TabIndex = 35
		Me._GrafDay_1.BackColor = System.Drawing.Color.Transparent
		Me._GrafDay_1.Enabled = True
		Me._GrafDay_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._GrafDay_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GrafDay_1.UseMnemonic = True
		Me._GrafDay_1.Visible = True
		Me._GrafDay_1.AutoSize = False
		Me._GrafDay_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._GrafDay_1.Name = "_GrafDay_1"
		Me._GrafDay_2.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._GrafDay_2.Text = "S"
		Me._GrafDay_2.Font = New System.Drawing.Font("Verdana", 14.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GrafDay_2.ForeColor = System.Drawing.Color.FromARGB(192, 192, 255)
		Me._GrafDay_2.Size = New System.Drawing.Size(22, 22)
		Me._GrafDay_2.Location = New System.Drawing.Point(97, -2)
		Me._GrafDay_2.TabIndex = 34
		Me._GrafDay_2.BackColor = System.Drawing.Color.Transparent
		Me._GrafDay_2.Enabled = True
		Me._GrafDay_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._GrafDay_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GrafDay_2.UseMnemonic = True
		Me._GrafDay_2.Visible = True
		Me._GrafDay_2.AutoSize = False
		Me._GrafDay_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._GrafDay_2.Name = "_GrafDay_2"
		Me._GrafDay_3.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._GrafDay_3.Text = "R"
		Me._GrafDay_3.Font = New System.Drawing.Font("Verdana", 14.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GrafDay_3.ForeColor = System.Drawing.Color.FromARGB(192, 192, 255)
		Me._GrafDay_3.Size = New System.Drawing.Size(22, 22)
		Me._GrafDay_3.Location = New System.Drawing.Point(132, -2)
		Me._GrafDay_3.TabIndex = 33
		Me._GrafDay_3.BackColor = System.Drawing.Color.Transparent
		Me._GrafDay_3.Enabled = True
		Me._GrafDay_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._GrafDay_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GrafDay_3.UseMnemonic = True
		Me._GrafDay_3.Visible = True
		Me._GrafDay_3.AutoSize = False
		Me._GrafDay_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._GrafDay_3.Name = "_GrafDay_3"
		Me._GrafDay_4.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._GrafDay_4.Text = "K"
		Me._GrafDay_4.Font = New System.Drawing.Font("Verdana", 14.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GrafDay_4.ForeColor = System.Drawing.Color.FromARGB(192, 192, 255)
		Me._GrafDay_4.Size = New System.Drawing.Size(22, 22)
		Me._GrafDay_4.Location = New System.Drawing.Point(168, -1)
		Me._GrafDay_4.TabIndex = 32
		Me._GrafDay_4.BackColor = System.Drawing.Color.Transparent
		Me._GrafDay_4.Enabled = True
		Me._GrafDay_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._GrafDay_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GrafDay_4.UseMnemonic = True
		Me._GrafDay_4.Visible = True
		Me._GrafDay_4.AutoSize = False
		Me._GrafDay_4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._GrafDay_4.Name = "_GrafDay_4"
		Me._GrafDay_5.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._GrafDay_5.Text = "J"
		Me._GrafDay_5.Font = New System.Drawing.Font("Verdana", 14.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GrafDay_5.ForeColor = System.Drawing.Color.FromARGB(192, 192, 255)
		Me._GrafDay_5.Size = New System.Drawing.Size(22, 22)
		Me._GrafDay_5.Location = New System.Drawing.Point(199, -1)
		Me._GrafDay_5.TabIndex = 31
		Me._GrafDay_5.BackColor = System.Drawing.Color.Transparent
		Me._GrafDay_5.Enabled = True
		Me._GrafDay_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._GrafDay_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GrafDay_5.UseMnemonic = True
		Me._GrafDay_5.Visible = True
		Me._GrafDay_5.AutoSize = False
		Me._GrafDay_5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._GrafDay_5.Name = "_GrafDay_5"
		Me._GrafDay_6.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._GrafDay_6.Text = "S"
		Me._GrafDay_6.Font = New System.Drawing.Font("Verdana", 14.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GrafDay_6.ForeColor = System.Drawing.Color.FromARGB(192, 192, 255)
		Me._GrafDay_6.Size = New System.Drawing.Size(22, 22)
		Me._GrafDay_6.Location = New System.Drawing.Point(233, -1)
		Me._GrafDay_6.TabIndex = 30
		Me._GrafDay_6.BackColor = System.Drawing.Color.Transparent
		Me._GrafDay_6.Enabled = True
		Me._GrafDay_6.Cursor = System.Windows.Forms.Cursors.Default
		Me._GrafDay_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GrafDay_6.UseMnemonic = True
		Me._GrafDay_6.Visible = True
		Me._GrafDay_6.AutoSize = False
		Me._GrafDay_6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._GrafDay_6.Name = "_GrafDay_6"
		_Bar1_0.OcxState = CType(resources.GetObject("_Bar1_0.OcxState"), System.Windows.Forms.AxHost.State)
		Me._Bar1_0.Size = New System.Drawing.Size(25, 164)
		Me._Bar1_0.Location = New System.Drawing.Point(24, 17)
		Me._Bar1_0.TabIndex = 37
		Me._Bar1_0.Name = "_Bar1_0"
		_Bar1_1.OcxState = CType(resources.GetObject("_Bar1_1.OcxState"), System.Windows.Forms.AxHost.State)
		Me._Bar1_1.Size = New System.Drawing.Size(25, 164)
		Me._Bar1_1.Location = New System.Drawing.Point(58, 17)
		Me._Bar1_1.TabIndex = 38
		Me._Bar1_1.Name = "_Bar1_1"
		_Bar1_2.OcxState = CType(resources.GetObject("_Bar1_2.OcxState"), System.Windows.Forms.AxHost.State)
		Me._Bar1_2.Size = New System.Drawing.Size(25, 164)
		Me._Bar1_2.Location = New System.Drawing.Point(93, 17)
		Me._Bar1_2.TabIndex = 39
		Me._Bar1_2.Name = "_Bar1_2"
		_Bar1_3.OcxState = CType(resources.GetObject("_Bar1_3.OcxState"), System.Windows.Forms.AxHost.State)
		Me._Bar1_3.Size = New System.Drawing.Size(25, 164)
		Me._Bar1_3.Location = New System.Drawing.Point(128, 17)
		Me._Bar1_3.TabIndex = 40
		Me._Bar1_3.Name = "_Bar1_3"
		_Bar1_4.OcxState = CType(resources.GetObject("_Bar1_4.OcxState"), System.Windows.Forms.AxHost.State)
		Me._Bar1_4.Size = New System.Drawing.Size(25, 164)
		Me._Bar1_4.Location = New System.Drawing.Point(161, 17)
		Me._Bar1_4.TabIndex = 41
		Me._Bar1_4.Name = "_Bar1_4"
		_Bar1_5.OcxState = CType(resources.GetObject("_Bar1_5.OcxState"), System.Windows.Forms.AxHost.State)
		Me._Bar1_5.Size = New System.Drawing.Size(25, 164)
		Me._Bar1_5.Location = New System.Drawing.Point(195, 17)
		Me._Bar1_5.TabIndex = 42
		Me._Bar1_5.Name = "_Bar1_5"
		_Bar1_6.OcxState = CType(resources.GetObject("_Bar1_6.OcxState"), System.Windows.Forms.AxHost.State)
		Me._Bar1_6.Size = New System.Drawing.Size(25, 164)
		Me._Bar1_6.Location = New System.Drawing.Point(228, 17)
		Me._Bar1_6.TabIndex = 43
		Me._Bar1_6.Name = "_Bar1_6"
		Me.GrafHigh.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.GrafHigh.Size = New System.Drawing.Size(51, 11)
		Me.GrafHigh.Location = New System.Drawing.Point(4, 3)
		Me.GrafHigh.TabIndex = 44
		Me.GrafHigh.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.GrafHigh.BackColor = System.Drawing.Color.Transparent
		Me.GrafHigh.Enabled = True
		Me.GrafHigh.ForeColor = System.Drawing.SystemColors.ControlText
		Me.GrafHigh.Cursor = System.Windows.Forms.Cursors.Default
		Me.GrafHigh.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.GrafHigh.UseMnemonic = True
		Me.GrafHigh.Visible = True
		Me.GrafHigh.AutoSize = False
		Me.GrafHigh.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.GrafHigh.Name = "GrafHigh"
		Me._StatLine_0.Size = New System.Drawing.Size(3, 335)
		Me._StatLine_0.Location = New System.Drawing.Point(303, 24)
		Me._StatLine_0.TabIndex = 48
		Me._StatLine_0.horizon = 0
		Me._StatLine_0.Name = "_StatLine_0"
		Me._StatTab_TabPage1.Text = "Stations"
		Lv1.OcxState = CType(resources.GetObject("Lv1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.Lv1.Size = New System.Drawing.Size(623, 315)
		Me.Lv1.Location = New System.Drawing.Point(11, 36)
		Me.Lv1.TabIndex = 16
		Me.Lv1.Name = "Lv1"
		Me._StatTab_TabPage2.Text = "Customers"
		Lv2.OcxState = CType(resources.GetObject("Lv2.OcxState"), System.Windows.Forms.AxHost.State)
		Me.Lv2.Size = New System.Drawing.Size(623, 315)
		Me.Lv2.Location = New System.Drawing.Point(11, 36)
		Me.Lv2.TabIndex = 15
		Me.Lv2.Name = "Lv2"
		Me._StatTab_TabPage3.Text = "Sales Record"
		Me.cbHari2.BackColor = System.Drawing.Color.FromARGB(192, 192, 255)
		Me.cbHari2.Size = New System.Drawing.Size(108, 21)
		Me.cbHari2.Location = New System.Drawing.Point(53, 35)
		Me.cbHari2.TabIndex = 12
		Me.cbHari2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cbHari2.CausesValidation = True
		Me.cbHari2.Enabled = True
		Me.cbHari2.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cbHari2.IntegralHeight = True
		Me.cbHari2.Cursor = System.Windows.Forms.Cursors.Default
		Me.cbHari2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cbHari2.Sorted = False
		Me.cbHari2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cbHari2.TabStop = True
		Me.cbHari2.Visible = True
		Me.cbHari2.Name = "cbHari2"
		Lv3.OcxState = CType(resources.GetObject("Lv3.OcxState"), System.Windows.Forms.AxHost.State)
		Me.Lv3.Size = New System.Drawing.Size(625, 287)
		Me.Lv3.Location = New System.Drawing.Point(11, 65)
		Me.Lv3.TabIndex = 13
		Me.Lv3.Name = "Lv3"
		Me.SlsBtn.Size = New System.Drawing.Size(28, 24)
		Me.SlsBtn.Location = New System.Drawing.Point(605, 32)
		Me.SlsBtn.TabIndex = 59
		Me.ToolTip1.SetToolTip(Me.SlsBtn, "Delete selected employee.")
		Me.SlsBtn.TX = ""
		Me.SlsBtn.ENAB = -1
		Me.SlsBtn.COLTYPE = 1
		Me.SlsBtn.FOCUSR = -1
		Me.SlsBtn.BCOL = 12632256
		Me.SlsBtn.BCOLO = 12632256
		Me.SlsBtn.FCOL = 0
		Me.SlsBtn.FCOLO = 0
		Me.SlsBtn.MCOL = 16777215
		Me.SlsBtn.MPTR = 1
		Me.SlsBtn.MICON = 0
		Me.SlsBtn.PICN = 0
		Me.SlsBtn.UMCOL = -1
		Me.SlsBtn.SOFT = 0
		Me.SlsBtn.PICPOS = 0
		Me.SlsBtn.NGREY = 0
		Me.SlsBtn.FX = 0
		Me.SlsBtn.HAND = 0
		Me.SlsBtn.CHECK = 0
		Me.SlsBtn.Name = "SlsBtn"
		Me._SlsLbl_0.Text = "Date :"
		Me._SlsLbl_0.Size = New System.Drawing.Size(38, 13)
		Me._SlsLbl_0.Location = New System.Drawing.Point(11, 37)
		Me._SlsLbl_0.TabIndex = 14
		Me._SlsLbl_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._SlsLbl_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._SlsLbl_0.BackColor = System.Drawing.Color.Transparent
		Me._SlsLbl_0.Enabled = True
		Me._SlsLbl_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._SlsLbl_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._SlsLbl_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._SlsLbl_0.UseMnemonic = True
		Me._SlsLbl_0.Visible = True
		Me._SlsLbl_0.AutoSize = False
		Me._SlsLbl_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._SlsLbl_0.Name = "_SlsLbl_0"
		Me._StatTab_TabPage4.Text = "Services Record"
		Me.PosCmbDate.BackColor = System.Drawing.Color.FromARGB(192, 192, 255)
		Me.PosCmbDate.Size = New System.Drawing.Size(108, 21)
		Me.PosCmbDate.Location = New System.Drawing.Point(53, 35)
		Me.PosCmbDate.TabIndex = 9
		Me.PosCmbDate.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.PosCmbDate.CausesValidation = True
		Me.PosCmbDate.Enabled = True
		Me.PosCmbDate.ForeColor = System.Drawing.SystemColors.WindowText
		Me.PosCmbDate.IntegralHeight = True
		Me.PosCmbDate.Cursor = System.Windows.Forms.Cursors.Default
		Me.PosCmbDate.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.PosCmbDate.Sorted = False
		Me.PosCmbDate.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.PosCmbDate.TabStop = True
		Me.PosCmbDate.Visible = True
		Me.PosCmbDate.Name = "PosCmbDate"
		Me.PosCmbItems.BackColor = System.Drawing.Color.FromARGB(192, 192, 255)
		Me.PosCmbItems.Size = New System.Drawing.Size(108, 21)
		Me.PosCmbItems.Location = New System.Drawing.Point(283, 34)
		Me.PosCmbItems.TabIndex = 8
		Me.PosCmbItems.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.PosCmbItems.CausesValidation = True
		Me.PosCmbItems.Enabled = True
		Me.PosCmbItems.ForeColor = System.Drawing.SystemColors.WindowText
		Me.PosCmbItems.IntegralHeight = True
		Me.PosCmbItems.Cursor = System.Windows.Forms.Cursors.Default
		Me.PosCmbItems.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.PosCmbItems.Sorted = False
		Me.PosCmbItems.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.PosCmbItems.TabStop = True
		Me.PosCmbItems.Visible = True
		Me.PosCmbItems.Name = "PosCmbItems"
		PosLV1.OcxState = CType(resources.GetObject("PosLV1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.PosLV1.Size = New System.Drawing.Size(625, 287)
		Me.PosLV1.Location = New System.Drawing.Point(11, 65)
		Me.PosLV1.TabIndex = 7
		Me.PosLV1.Name = "PosLV1"
		Me.SrvBtn.Size = New System.Drawing.Size(28, 24)
		Me.SrvBtn.Location = New System.Drawing.Point(395, 33)
		Me.SrvBtn.TabIndex = 60
		Me.ToolTip1.SetToolTip(Me.SrvBtn, "Delete selected employee.")
		Me.SrvBtn.TX = ""
		Me.SrvBtn.ENAB = -1
		Me.SrvBtn.COLTYPE = 1
		Me.SrvBtn.FOCUSR = -1
		Me.SrvBtn.BCOL = 12632256
		Me.SrvBtn.BCOLO = 12632256
		Me.SrvBtn.FCOL = 0
		Me.SrvBtn.FCOLO = 0
		Me.SrvBtn.MCOL = 16777215
		Me.SrvBtn.MPTR = 1
		Me.SrvBtn.MICON = 0
		Me.SrvBtn.PICN = 0
		Me.SrvBtn.UMCOL = -1
		Me.SrvBtn.SOFT = 0
		Me.SrvBtn.PICPOS = 0
		Me.SrvBtn.NGREY = 0
		Me.SrvBtn.FX = 0
		Me.SrvBtn.HAND = 0
		Me.SrvBtn.CHECK = 0
		Me.SrvBtn.Name = "SrvBtn"
		Me._SrvLbl_0.Text = "Date :"
		Me._SrvLbl_0.Size = New System.Drawing.Size(38, 13)
		Me._SrvLbl_0.Location = New System.Drawing.Point(11, 37)
		Me._SrvLbl_0.TabIndex = 11
		Me._SrvLbl_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._SrvLbl_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._SrvLbl_0.BackColor = System.Drawing.Color.Transparent
		Me._SrvLbl_0.Enabled = True
		Me._SrvLbl_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._SrvLbl_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._SrvLbl_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._SrvLbl_0.UseMnemonic = True
		Me._SrvLbl_0.Visible = True
		Me._SrvLbl_0.AutoSize = False
		Me._SrvLbl_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._SrvLbl_0.Name = "_SrvLbl_0"
		Me._SrvLbl_1.Text = "Filter By Items :"
		Me._SrvLbl_1.Size = New System.Drawing.Size(94, 13)
		Me._SrvLbl_1.Location = New System.Drawing.Point(185, 37)
		Me._SrvLbl_1.TabIndex = 10
		Me._SrvLbl_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._SrvLbl_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._SrvLbl_1.BackColor = System.Drawing.Color.Transparent
		Me._SrvLbl_1.Enabled = True
		Me._SrvLbl_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._SrvLbl_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._SrvLbl_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._SrvLbl_1.UseMnemonic = True
		Me._SrvLbl_1.Visible = True
		Me._SrvLbl_1.AutoSize = False
		Me._SrvLbl_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._SrvLbl_1.Name = "_SrvLbl_1"
		Me.Controls.Add(StatFme)
		Me.Controls.Add(StatBar)
		Me.Controls.Add(StatTab)
		Me.StatFme.Controls.Add(_StatLine_1)
		Me.StatFme.Controls.Add(cbHari)
		Me.StatFme.Controls.Add(cbBulan)
		Me.StatFme.Controls.Add(cbTahun)
		Me.StatFme.Controls.Add(StatBtn)
		Me.StatFme.Controls.Add(StatSesBtn)
		Me.StatFme.Controls.Add(_StatLbl_2)
		Me.StatFme.Controls.Add(_StatLbl_3)
		Me.StatFme.Controls.Add(CurSession)
		Me.StatFme.Controls.Add(_StatLbl_1)
		Me.StatFme.Controls.Add(_StatLbl_0)
		Me.StatTab.Controls.Add(_StatTab_TabPage0)
		Me.StatTab.Controls.Add(_StatTab_TabPage1)
		Me.StatTab.Controls.Add(_StatTab_TabPage2)
		Me.StatTab.Controls.Add(_StatTab_TabPage3)
		Me.StatTab.Controls.Add(_StatTab_TabPage4)
		Me._StatTab_TabPage0.Controls.Add(_GenHdr_0)
		Me._StatTab_TabPage0.Controls.Add(_GenLbl_4)
		Me._StatTab_TabPage0.Controls.Add(_GenLbl_3)
		Me._StatTab_TabPage0.Controls.Add(lblJualanPos)
		Me._StatTab_TabPage0.Controls.Add(lblJualanPc)
		Me._StatTab_TabPage0.Controls.Add(_GenLbl_2)
		Me._StatTab_TabPage0.Controls.Add(_GenLbl_1)
		Me._StatTab_TabPage0.Controls.Add(lblUntung)
		Me._StatTab_TabPage0.Controls.Add(lblJualan)
		Me._StatTab_TabPage0.Controls.Add(lblModal)
		Me._StatTab_TabPage0.Controls.Add(_GenLbl_0)
		Me._StatTab_TabPage0.Controls.Add(_GenHdr_1)
		Me._StatTab_TabPage0.Controls.Add(_GenLbl_6)
		Me._StatTab_TabPage0.Controls.Add(lblServis)
		Me._StatTab_TabPage0.Controls.Add(lblPungut)
		Me._StatTab_TabPage0.Controls.Add(_GenLbl_5)
		Me._StatTab_TabPage0.Controls.Add(_GenHdr_2)
		Me._StatTab_TabPage0.Controls.Add(Graf1)
		Me._StatTab_TabPage0.Controls.Add(_StatLine_0)
		Me.Graf1.Controls.Add(GrafDock)
		Me.Graf1.Controls.Add(_Bar1_0)
		Me.Graf1.Controls.Add(_Bar1_1)
		Me.Graf1.Controls.Add(_Bar1_2)
		Me.Graf1.Controls.Add(_Bar1_3)
		Me.Graf1.Controls.Add(_Bar1_4)
		Me.Graf1.Controls.Add(_Bar1_5)
		Me.Graf1.Controls.Add(_Bar1_6)
		Me.Graf1.Controls.Add(GrafHigh)
		Me.GrafDock.Controls.Add(_GrafDay_0)
		Me.GrafDock.Controls.Add(_GrafDay_1)
		Me.GrafDock.Controls.Add(_GrafDay_2)
		Me.GrafDock.Controls.Add(_GrafDay_3)
		Me.GrafDock.Controls.Add(_GrafDay_4)
		Me.GrafDock.Controls.Add(_GrafDay_5)
		Me.GrafDock.Controls.Add(_GrafDay_6)
		Me._StatTab_TabPage1.Controls.Add(Lv1)
		Me._StatTab_TabPage2.Controls.Add(Lv2)
		Me._StatTab_TabPage3.Controls.Add(cbHari2)
		Me._StatTab_TabPage3.Controls.Add(Lv3)
		Me._StatTab_TabPage3.Controls.Add(SlsBtn)
		Me._StatTab_TabPage3.Controls.Add(_SlsLbl_0)
		Me._StatTab_TabPage4.Controls.Add(PosCmbDate)
		Me._StatTab_TabPage4.Controls.Add(PosCmbItems)
		Me._StatTab_TabPage4.Controls.Add(PosLV1)
		Me._StatTab_TabPage4.Controls.Add(SrvBtn)
		Me._StatTab_TabPage4.Controls.Add(_SrvLbl_0)
		Me._StatTab_TabPage4.Controls.Add(_SrvLbl_1)
		Me.Bar1.SetIndex(_Bar1_0, CType(0, Short))
		Me.Bar1.SetIndex(_Bar1_1, CType(1, Short))
		Me.Bar1.SetIndex(_Bar1_2, CType(2, Short))
		Me.Bar1.SetIndex(_Bar1_3, CType(3, Short))
		Me.Bar1.SetIndex(_Bar1_4, CType(4, Short))
		Me.Bar1.SetIndex(_Bar1_5, CType(5, Short))
		Me.Bar1.SetIndex(_Bar1_6, CType(6, Short))
		Me.GenHdr.SetIndex(_GenHdr_2, CType(2, Short))
		Me.GenHdr.SetIndex(_GenHdr_1, CType(1, Short))
		Me.GenHdr.SetIndex(_GenHdr_0, CType(0, Short))
		Me.GenLbl.SetIndex(_GenLbl_5, CType(5, Short))
		Me.GenLbl.SetIndex(_GenLbl_6, CType(6, Short))
		Me.GenLbl.SetIndex(_GenLbl_0, CType(0, Short))
		Me.GenLbl.SetIndex(_GenLbl_1, CType(1, Short))
		Me.GenLbl.SetIndex(_GenLbl_2, CType(2, Short))
		Me.GenLbl.SetIndex(_GenLbl_3, CType(3, Short))
		Me.GenLbl.SetIndex(_GenLbl_4, CType(4, Short))
		Me.GrafDay.SetIndex(_GrafDay_0, CType(0, Short))
		Me.GrafDay.SetIndex(_GrafDay_1, CType(1, Short))
		Me.GrafDay.SetIndex(_GrafDay_2, CType(2, Short))
		Me.GrafDay.SetIndex(_GrafDay_3, CType(3, Short))
		Me.GrafDay.SetIndex(_GrafDay_4, CType(4, Short))
		Me.GrafDay.SetIndex(_GrafDay_5, CType(5, Short))
		Me.GrafDay.SetIndex(_GrafDay_6, CType(6, Short))
		Me.SlsLbl.SetIndex(_SlsLbl_0, CType(0, Short))
		Me.SrvLbl.SetIndex(_SrvLbl_0, CType(0, Short))
		Me.SrvLbl.SetIndex(_SrvLbl_1, CType(1, Short))
		Me.StatLbl.SetIndex(_StatLbl_2, CType(2, Short))
		Me.StatLbl.SetIndex(_StatLbl_3, CType(3, Short))
		Me.StatLbl.SetIndex(_StatLbl_1, CType(1, Short))
		Me.StatLbl.SetIndex(_StatLbl_0, CType(0, Short))
		Me.StatLine.SetIndex(_StatLine_1, CType(1, Short))
		Me.StatLine.SetIndex(_StatLine_0, CType(0, Short))
		CType(Me.StatLine, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.StatLbl, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.SrvLbl, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.SlsLbl, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.GrafDay, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.GenLbl, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.GenHdr, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Bar1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.PosLV1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Lv3, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Lv2, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Lv1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me._Bar1_6, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me._Bar1_5, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me._Bar1_4, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me._Bar1_3, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me._Bar1_2, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me._Bar1_1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me._Bar1_0, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.StatBar, System.ComponentModel.ISupportInitialize).EndInit()
	End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As FrmStat
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As FrmStat
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New FrmStat()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	Private Rs As DAO.Recordset
	Private DBloaded As Boolean
	
	'UPGRADE_NOTE: Modal was upgraded to Modal_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
	Private Modal_Renamed As Double
	Private Jualan As Double
	Private Untung As Double
	Private sTahun As String
	Private sBulan As String
	
	
	Private Sub FrmStat_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim a As Object
		'{ Reset all date container }'
		cbTahun.Items.Clear()
		cbBulan.Items.Clear()
		cbHari.Items.Clear()
		cbHari2.Items.Clear()
		PosCmbDate.Items.Clear()
		
		'{ load all year }'
		Call LoadYear(cbTahun)
		
		'{ select current year }'
		For a = 0 To cbTahun.Items.Count - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object a. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			If CDbl(VB6.GetItemString(cbTahun, a)) = Year(Today) Then cbTahun.SelectedIndex = a
		Next a
		
		'{ display currents session }'
		CurSession.Text = OpenSessionCur
	End Sub
	
	'UPGRADE_WARNING: Event cbTahun.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2075"'
	Private Sub cbTahun_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbTahun.SelectedIndexChanged
		Dim a As Object
		'{ Reset date container }'
		cbBulan.Items.Clear()
		cbHari.Items.Clear()
		cbHari2.Items.Clear()
		
		'{ load all month }'
		sTahun = cbTahun.Text
		Call LoadMonth(cbBulan)
		
		'{ select current month }'
		For a = 0 To cbBulan.Items.Count - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object a. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			If CDbl(VB6.GetItemString(cbBulan, a)) = Month(Today) Then cbBulan.SelectedIndex = a
		Next a
	End Sub
	
	'UPGRADE_WARNING: Event cbBulan.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2075"'
	Private Sub cbBulan_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbBulan.SelectedIndexChanged
		'{ Reset date container }'
		Lv1.ListItems.Clear()
		Lv2.ListItems.Clear()
		Lv3.ListItems.Clear()
		PosLV1.ListItems.Clear()
		
		sBulan = cbBulan.Text
		Call LoadDate()
		
		Call StatKewangan()
		Call StatTerminal()
		Call StatPelanggan()
		Call StatHarian()
		Call StatPOS()
		'UPGRADE_WARNING: Couldn't resolve default property of object GetBulan(sBulan). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		GenHdr(0).Text = " Monthly Statistic - " & GetBulan(sBulan) & " \ " & sTahun
	End Sub
	
	'UPGRADE_WARNING: Event cbHari.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2075"'
	Private Sub cbHari_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbHari.SelectedIndexChanged
		If cbHari.Text = "" Then Exit Sub
		If IsDate(cbHari.Text) = False Then Exit Sub
		Lv3.ListItems.Clear()
		PosLV1.ListItems.Clear()
		
		cbHari2.Text = cbHari.Text
		Call StatHarian(cbHari.Text, False, False)
		Call StatHarian(cbHari2.Text, True)
	End Sub
	
	'UPGRADE_WARNING: Event cbHari2.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2075"'
	Private Sub cbHari2_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbHari2.SelectedIndexChanged
		If cbHari2.Text = "" Then Exit Sub
		If IsDate(cbHari.Text) = False Then Exit Sub
		Lv3.ListItems.Clear()
		
		cbHari.Text = cbHari2.Text
		Call StatHarian(cbHari2.Text, True)
		Call StatHarian(cbHari.Text, False, False)
	End Sub
	
	Private Sub SlsBtn_Click(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles SlsBtn.Click
		'ret = ShellExecute(Me.hwnd, "open", App.Path & "\CafeReport.exe", "pc-usage", vbNullString, SW_NORMAL)
		'If ret <= 32 Then MsgBox MB(20), vbCritical, CbMsgWarn
		Call LoadModule(mApplication.EnuModule.CafeReport)
	End Sub
	
	Private Sub SrvBtn_Click(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles SrvBtn.Click
		Rs = uSDB.OpenRecordset("pos-items", DAO.RecordsetTypeEnum.dbOpenSnapshot)
		PosCmbItems.Items.Clear()
		
		If Rs.BOF = True Then Exit Sub
		With Rs
			Do Until .EOF = True
				PosCmbItems.Items.Add(.Fields("Nama").Value)
				.MoveNext()
			Loop 
		End With
	End Sub
	
	Private Sub StatBtn_Click(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles StatBtn.Click
		DBloaded = False
		FrmStat.DefInstance.Close()
	End Sub
	
	'UPGRADE_WARNING: Event PosCmbDate.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2075"'
	Private Sub PosCmbDate_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles PosCmbDate.SelectedIndexChanged
		If PosCmbDate.Text = "" Then Exit Sub
		If Not IsDate(PosCmbDate.Text) Then Exit Sub
		
		PosLV1.ListItems.Clear()
		Call StatPOS(PosCmbDate.Text, PosCmbItems.Text)
	End Sub
	
	'UPGRADE_WARNING: Event PosCmbItems.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2075"'
	Private Sub PosCmbItems_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles PosCmbItems.SelectedIndexChanged
		If PosCmbItems.Text = "" Then Exit Sub
		
		PosLV1.ListItems.Clear()
		Call StatPOS(PosCmbDate.Text, PosCmbItems.Text)
	End Sub
	
	
	
	Public Sub LoadYear(ByRef Cbox As System.Windows.Forms.ComboBox)
		Dim Rss, Rs As DAO.Recordset
		Dim SqlQ As String
		
		'{ add list month to pc sales combo }'
		Rss = uIDB.OpenRecordset("pc-harian", DAO.RecordsetTypeEnum.dbOpenSnapshot)
		With Rss
			Do Until .EOF = True
				CbAddEx(.Fields("Tahun"), Cbox)
				.MoveNext()
			Loop 
		End With
		
		'{ add list of previous date to pos combo }'
		Rs = uIDB.OpenRecordset("pos-usage", DAO.RecordsetTypeEnum.dbOpenSnapshot)
		Do Until Rs.BOF = True
			With Rs
				.MoveFirst()
				CbAddEx(.Fields("Tahun"), Cbox)
				SqlQ = "tahun = " & .Fields("Tahun").Value & " AND bulan = " & .Fields("Bulan").Value & " AND hari <> " & .Fields("Hari").Value '& "'"
			End With
			Rs = RsFilter(Rs, SqlQ)
		Loop 
	End Sub
	
	Public Sub LoadMonth(ByRef Cbox As System.Windows.Forms.ComboBox)
		Dim Rss, Rs As DAO.Recordset
		Dim SqlQ As String
		
		'{ add list month to pc sales combo }'
		Rss = uIDB.OpenRecordset("pc-harian", DAO.RecordsetTypeEnum.dbOpenSnapshot)
		With Rss
			Do Until .EOF = True
				If .Fields("Tahun").Value = sTahun Then CbAddEx(.Fields("Bulan"), Cbox)
				.MoveNext()
			Loop 
		End With
		
		'{ add list of previous date to pos combo }'
		Rs = uIDB.OpenRecordset("pos-usage", DAO.RecordsetTypeEnum.dbOpenSnapshot)
		Do Until Rs.BOF = True
			With Rs
				.MoveFirst()
				If .Fields("Tahun").Value = sTahun Then CbAddEx(.Fields("Bulan"), Cbox)
				SqlQ = "tahun = " & .Fields("Tahun").Value & " AND bulan = " & .Fields("Bulan").Value & " AND hari <> " & .Fields("Hari").Value '& "'"
			End With
			Rs = RsFilter(Rs, SqlQ)
		Loop 
	End Sub
	
	Public Sub LoadDate()
		Dim Rss, Rs As DAO.Recordset
		Dim tDate As String
		Dim SqlQ As String
		
		cbHari.Items.Clear()
		cbHari2.Items.Clear()
		
		'{ add list of previous date to pc sales combo }'
		Rss = uIDB.OpenRecordset("pc-harian", DAO.RecordsetTypeEnum.dbOpenSnapshot)
		With Rss
			Do Until .EOF = True
				If .Fields("Tahun").Value = sTahun And .Fields("Bulan").Value = sBulan Then
					tDate = GetSystemDate(.Fields("Hari").Value, .Fields("Bulan").Value, .Fields("Tahun").Value)
					cbHari.Items.Add(tDate)
					cbHari2.Items.Add(tDate)
				End If
				.MoveNext()
			Loop 
		End With
		
		'{ add list of previous date to pos combo }'
		Rs = uIDB.OpenRecordset("pos-usage", DAO.RecordsetTypeEnum.dbOpenSnapshot)
		Do Until Rs.BOF = True
			With Rs
				.MoveFirst()
				If .Fields("Tahun").Value = sTahun And .Fields("Bulan").Value = sBulan Then
					tDate = GetSystemDate(.Fields("Hari").Value, .Fields("Bulan").Value, .Fields("Tahun").Value)
					PosCmbDate.Items.Add(tDate)
				End If
				SqlQ = "tahun = " & .Fields("Tahun").Value & " AND bulan = " & .Fields("Bulan").Value & " AND hari <> " & .Fields("Hari").Value '& "'"
			End With
			Rs = RsFilter(Rs, SqlQ)
		Loop 
		
		'{ display to combo todays date }'
		If CDbl(sTahun) = Year(Today) And CDbl(sBulan) = Month(Today) Then
			cbHari.Text = CStr(Today)
		Else
			cbHari.Text = VB6.GetItemString(cbHari, 0)
		End If
		cbHari2.Text = cbHari.Text
		PosCmbDate.Text = VB6.GetItemString(PosCmbDate, 0)
	End Sub
	
	
	
	
	Public Sub StatKewangan()
		Dim p As Object
		Dim dh As Object
		Dim H As Object
		Dim K As Object
		Dim Rss As DAO.Recordset
		Dim SqlQ As String
		Dim JualanPc, JualanPos As Double
		Dim BilLain, Sewa, Gaji, Bil, Biggest As Object
		Dim tmpVal As Double
		Dim cTahun, cBulan As Object
		
		Modal_Renamed = 0 : Jualan = 0 : Untung = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object Biggest. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Biggest = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object cTahun. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		cTahun = Year(CDate(sTahun))
		'UPGRADE_WARNING: Couldn't resolve default property of object cBulan. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		cBulan = Month(CDate(sBulan))
		'cHari = Day(OpenSessionCur)
		
		'ambil jumlah kesemua modal
		'UPGRADE_WARNING: Couldn't resolve default property of object uSDBe.DataCount(pekerja-list). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		For K = 0 To uSDBe.DataCount("pekerja-list") - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object uSDBe.DataGet(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Gaji. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			Gaji = Gaji + CDbl(uSDBe.DataGet("pekerja-list", "gaji", K))
		Next K
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Sewa. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Sewa = CDbl(SetAmbil("sewa"))
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Bil. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Bil = CDbl(SetAmbil("bil"))
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object BilLain. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		BilLain = CDbl(SetAmbil("billain"))
		'UPGRADE_WARNING: Couldn't resolve default property of object BilLain. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Bil. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Sewa. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Gaji. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Modal_Renamed = Gaji + Sewa + Bil + BilLain
		lblModal.Text = Crnc & " " & VB6.Format(Modal_Renamed, "#0.00")
		
		
		'Query-Code untuk filter bulan semasa
		SqlQ = "tahun = " & sTahun & " AND bulan = " & sBulan '& "'"
		
		'ambil jumlah jualan - PC
		Rss = uIDB.OpenRecordset("pc-harian", DAO.RecordsetTypeEnum.dbOpenSnapshot)
		Rss.Filter = SqlQ
		Rs = Rss.OpenRecordset
		With Rs
			Do Until .EOF = True
				JualanPc = JualanPc + .Fields("pungutan").Value
				.MoveNext()
			Loop 
		End With
		lblJualanPc.Text = Crnc & " " & VB6.Format(JualanPc, "#0.00")
		'ambil jumlah jualan - POS
		Rss = uIDB.OpenRecordset("pos-usage", DAO.RecordsetTypeEnum.dbOpenSnapshot)
		Rss.Filter = SqlQ
		Rs = Rss.OpenRecordset
		With Rs
			Do Until .EOF = True
				JualanPos = JualanPos + .Fields("Harga").Value
				.MoveNext()
			Loop 
		End With
		lblJualanPos.Text = Crnc & " " & VB6.Format(JualanPos, "#0.00")
		
		'pemer jumlah harga
		Jualan = JualanPc + JualanPos
		lblJualan.Text = Crnc & " " & VB6.Format(Jualan, "#0.00")
		
		'ambil data untuk bar1
		Rs = uIDB.OpenRecordset("pc-grafminggu", DAO.RecordsetTypeEnum.dbOpenSnapshot)
		If Rs.BOF = True Then Exit Sub
		Rs.FindFirst(SqlQ)
		If Rs.NoMatch = False Then
			For H = 1 To 7
				'UPGRADE_WARNING: Couldn't resolve default property of object H. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Choose(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				'UPGRADE_WARNING: Couldn't resolve default property of object dh. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				dh = Choose(H, "Ahad", "Isnin", "Selasa", "Rabu", "Khamis", "Jumaat", "Sabtu")
				tmpVal = Rs.Fields(dh).Value
				'UPGRADE_WARNING: Couldn't resolve default property of object Biggest. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				If tmpVal > Biggest Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Biggest. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					Biggest = tmpVal
					For p = 0 To 6
						'UPGRADE_WARNING: Couldn't resolve default property of object Biggest. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
						Bar1(p).Max = Biggest
					Next p
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object H. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				Bar1(H - 1).Value = tmpVal '(tmpVal * Biggest) / 100
				'UPGRADE_WARNING: Couldn't resolve default property of object Biggest. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				GrafHigh.Text = Crnc & " " & VB6.Format(Biggest, "#0.00")
			Next H
		End If
		
		'kira keuntungan dan juga filter... pasti enakkk
		Untung = Jualan - Modal_Renamed
		lblUntung.Text = Crnc & " " & VB6.Format(Untung, "#0.00")
		If Jualan < Modal_Renamed Then
			lblJualan.ForeColor = System.Drawing.ColorTranslator.FromOle(&H8080FF)
			lblUntung.ForeColor = System.Drawing.ColorTranslator.FromOle(&H8080FF)
		End If
		
		'UPGRADE_NOTE: Object Rss may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
		Rss = Nothing
	End Sub
	
	Public Sub StatTerminal()
		Dim g As Object
		Dim Rss As DAO.Recordset
		Dim tItm As MSComctlLib.ListItem
		Dim SqlQ As String
		
		'Query-Code untuk filter bulan semasa
		SqlQ = "tahun = " & sTahun & " AND bulan = " & sBulan '& "'"
		
		Rss = uIDB.OpenRecordset("pc-bulanan", DAO.RecordsetTypeEnum.dbOpenSnapshot)
		Rss.Filter = SqlQ
		Rs = Rss.OpenRecordset
		If Rs.BOF = True Then Exit Sub
		With Rs
			.MoveLast()
			.MoveFirst()
			For g = 1 To .RecordCount
				tItm = Lv1.ListItems.Add( ,  , .Fields("NamaPc"))
				tItm.SubItems(1) = .Fields("JumlahMasa").Value & " Minit"
				tItm.SubItems(2) = Crnc & " " & VB6.Format(.Fields("JumlahBayar").Value, "#0.00")
				.MoveNext()
			Next g
		End With
		
		'UPGRADE_NOTE: Object Rss may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
		Rss = Nothing
		'UPGRADE_NOTE: Object tItm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
		tItm = Nothing
	End Sub
	
	Public Sub StatPelanggan()
		Dim tItm As MSComctlLib.ListItem
		
		Rs = uSDB.OpenRecordset("pelanggan-list", DAO.RecordsetTypeEnum.dbOpenSnapshot)
		
		If Rs.BOF = True Then Exit Sub
		With Rs
			.MoveFirst()
			Do Until .EOF = True
				tItm = Lv2.ListItems.Add( ,  , .Fields("Nama"))
				tItm.SubItems(1) = .Fields("lawat").Value & " kali"
				tItm.SubItems(2) = .Fields("tarikhakhir").Value
				tItm.SubItems(3) = .Fields("JumlahMasa").Value
				tItm.SubItems(4) = Crnc & " " & VB6.Format(.Fields("JumlahBayar").Value, "#0.00")
				.MoveNext()
			Loop 
		End With
	End Sub
	
	
	Public Sub StatHarian(Optional ByRef TarikhStr As String = "", Optional ByRef LoadListOnly As Boolean = False, Optional ByRef LoadDetail As Boolean = True)
		Dim Rss As DAO.Recordset
		Dim SqlQ, TrkhTmp, SqlQx As String
		Dim nItm As MSComctlLib.ListItem
		
		If TarikhStr = "" Then
			TrkhTmp = cbHari.Text
		Else
			TrkhTmp = TarikhStr
		End If
		
		SqlQ = "tahun =  '" & Year(CDate(TrkhTmp)) & "' AND bulan = '" & Month(CDate(TrkhTmp)) & "' AND hari = '" & VB.Day(CDate(TrkhTmp)) & "'"
		SqlQx = "tahun =  " & Year(CDate(TrkhTmp)) & " AND bulan = " & Month(CDate(TrkhTmp)) & " AND hari = " & VB.Day(CDate(TrkhTmp))
		
		If LoadListOnly = True Then GoTo ListOnly
		Rs = uIDB.OpenRecordset("pc-harian", DAO.RecordsetTypeEnum.dbOpenSnapshot)
		If Rs.BOF = True Then Exit Sub
		With Rs
			.FindFirst(SqlQx)
			If .NoMatch = False Then
				lblPungut.Text = Crnc & " " & VB6.Format(.Fields("pungutan").Value, "#0.00")
				lblServis.Text = Crnc & " " & VB6.Format(.Fields("pungutanservis").Value, "#0.00")
			End If
		End With
		If LoadDetail = False Then Exit Sub
		
ListOnly: 
		Rss = uIDB.OpenRecordset("pc-usage", DAO.RecordsetTypeEnum.dbOpenSnapshot)
		Rss.Filter = SqlQx
		Rss.Sort = "masuk"
		Rs = Rss.OpenRecordset
		If Rs.BOF = True Then Exit Sub
		Do While Rs.EOF <> True
			With Rs
				nItm = Lv3.ListItems.Add( ,  , TrkhTmp)
				nItm.SubItems(1) = .Fields("PcName").Value
				nItm.SubItems(2) = .Fields("Nama").Value
				nItm.SubItems(3) = .Fields("masuk").Value
				nItm.SubItems(4) = .Fields("Keluar").Value
				nItm.SubItems(5) = Crnc & " " & .Fields("Harga").Value
			End With
			Rs.MoveNext()
		Loop 
		
		'UPGRADE_NOTE: Object Rss may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
		Rss = Nothing
		'UPGRADE_NOTE: Object nItm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
		nItm = Nothing
	End Sub
	
	Public Sub StatPOS(Optional ByRef FilterDate As String = "", Optional ByRef FilterItem As String = "")
		Dim tItm As MSComctlLib.ListItem
		Dim SqlQ As String
		
		If FilterDate = "" Then
			SqlQ = "tahun = " & sTahun & " AND bulan = " & sBulan '& "'"
		Else
			SqlQ = "tahun = " & Year(CDate(FilterDate)) & " AND bulan = " & Month(CDate(FilterDate)) & " AND hari = " & VB.Day(CDate(FilterDate)) '& "'"
		End If
		If FilterItem <> "" Then SqlQ = SqlQ & " AND item = '" & FilterItem & "'"
		
		Rs = uIDB.OpenRecordset("pos-usage", DAO.RecordsetTypeEnum.dbOpenSnapshot)
		Rs.Filter = SqlQ
		Rs = Rs.OpenRecordset
		
		If Rs.BOF = True Then Exit Sub
		With Rs
			.MoveFirst()
			Do Until .EOF = True
				tItm = PosLV1.ListItems.Add( ,  , GetSystemDate(.Fields("Hari").Value, .Fields("Bulan").Value, .Fields("Tahun").Value))
				tItm.SubItems(1) = .Fields("GroupId").Value
				tItm.SubItems(2) = .Fields("transid").Value
				tItm.SubItems(3) = .Fields("Item").Value
				tItm.SubItems(4) = .Fields("qty").Value
				tItm.SubItems(5) = .Fields("Harga").Value
				.MoveNext()
			Loop 
		End With
		
		'UPGRADE_NOTE: Object tItm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
		tItm = Nothing
	End Sub
	
	Private Sub StatSesBtn_Click(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles StatSesBtn.Click
		Dim ret As Object
		Dim mSj As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object mSj. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		mSj = "Close current session ?"
		'UPGRADE_WARNING: Couldn't resolve default property of object ret. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ret = MsgBox(mSj, MsgBoxStyle.OKCancel, CbMsgApp)
		If ret = MsgBoxResult.OK Then
			uSDBe.DbSaveSetting("lastsession", OpenSessionCur)
			OpenSessionCur = CStr(Today)
			uSDBe.DbSaveSetting("opensession", OpenSessionCur)
		End If
	End Sub
End Class