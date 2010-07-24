Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class FrmAgnMgr
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
	Public WithEvents _AgsMnu_3 As XpButton
	Public WithEvents _AgsMnu_0 As XpButton
	Public WithEvents _AgsMnu_1 As XpButton
	Public WithEvents _AgsMnu_2 As XpButton
	Public WithEvents _AgsCmd_0 As XpButton
	Public WithEvents MainPdck As PageDock
	Public WithEvents _CptMnu_0 As XpButton
	Public WithEvents _CptMnu_1 As XpButton
	Public WithEvents _CptMnu_2 As XpButton
	Public WithEvents _CptMnu_3 As XpButton
	Public WithEvents _CptMnu_4 As XpButton
	Public WithEvents MainPhld As PageHolder
	Public WithEvents MainBnrCap As System.Windows.Forms.Label
	Public WithEvents MainBnrLbl As System.Windows.Forms.Label
	Public WithEvents MainBnr As System.Windows.Forms.Panel
	Public WithEvents MainLne As Line3D
	Public WithEvents LstVw1 As AxMSComctlLib.AxListView
	Public WithEvents GcnCmdCB As System.Windows.Forms.ComboBox
	Public WithEvents GcnList As System.Windows.Forms.ListBox
	Public WithEvents GcnBtnClr As System.Windows.Forms.PictureBox
	Public WithEvents GcnCmdBtn As System.Windows.Forms.PictureBox
	Public WithEvents _GcnHdr_1 As System.Windows.Forms.Label
	Public WithEvents _GcnHdr_0 As System.Windows.Forms.Label
	Public WithEvents _Pages_0 As System.Windows.Forms.Panel
	Public WithEvents _GenOpt2_0 As System.Windows.Forms.CheckBox
	Public WithEvents _GenOpt2_1 As System.Windows.Forms.CheckBox
	Public WithEvents _GenOpt2_2 As System.Windows.Forms.CheckBox
	Public WithEvents _GenOpt2_3 As System.Windows.Forms.CheckBox
	Public WithEvents _GenOpt1_0 As System.Windows.Forms.CheckBox
	Public WithEvents GenWelcome As System.Windows.Forms.TextBox
	Public WithEvents GenPass2 As System.Windows.Forms.TextBox
	Public WithEvents GenPass1 As System.Windows.Forms.TextBox
	Public WithEvents _GenOpt1_1 As System.Windows.Forms.CheckBox
	Public WithEvents GenNetName As System.Windows.Forms.TextBox
	Public WithEvents GenNetIP As System.Windows.Forms.TextBox
	Public WithEvents GenNetPort As System.Windows.Forms.TextBox
	Public WithEvents _GenHdr_3 As System.Windows.Forms.Label
	Public WithEvents _GenHdr_2 As System.Windows.Forms.Label
	Public WithEvents _GenMiscLbl_0 As System.Windows.Forms.Label
	Public WithEvents _GenHdr_1 As System.Windows.Forms.Label
	Public WithEvents _GenPassLbl_1 As System.Windows.Forms.Label
	Public WithEvents _GenPassLbl_0 As System.Windows.Forms.Label
	Public WithEvents _GenHdr_0 As System.Windows.Forms.Label
	Public WithEvents _GenNetLbl_0 As System.Windows.Forms.Label
	Public WithEvents _GenNetLbl_2 As System.Windows.Forms.Label
	Public WithEvents _GenNetLbl_1 As System.Windows.Forms.Label
	Public WithEvents _Pages_1 As System.Windows.Forms.Panel
	Public WithEvents _Pages_2 As System.Windows.Forms.Panel
	Public WithEvents _Pages_4 As System.Windows.Forms.Panel
	Public WithEvents _Pages_3 As System.Windows.Forms.Panel
	Public WithEvents AgsCmd As XpButtonArray
	Public WithEvents AgsMnu As XpButtonArray
	Public WithEvents CptMnu As XpButtonArray
	Public WithEvents GcnHdr As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents GenHdr As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents GenMiscLbl As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents GenNetLbl As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents GenOpt1 As Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray
	Public WithEvents GenOpt2 As Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray
	Public WithEvents GenPassLbl As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents Mnu2Sel As Microsoft.VisualBasic.Compatibility.VB6.MenuItemArray
	Public WithEvents Pages As Microsoft.VisualBasic.Compatibility.VB6.PanelArray
	Public WithEvents Mnu1Rfsh As System.Windows.Forms.MenuItem
	Public WithEvents Mnu1Close As System.Windows.Forms.MenuItem
	Public WithEvents Mnu1 As System.Windows.Forms.MenuItem
	Public WithEvents _Mnu2Sel_0 As System.Windows.Forms.MenuItem
	Public WithEvents _Mnu2Sel_1 As System.Windows.Forms.MenuItem
	Public WithEvents _Mnu2Sel_2 As System.Windows.Forms.MenuItem
	Public WithEvents _Mnu2Sel_3 As System.Windows.Forms.MenuItem
	Public WithEvents _Mnu2Sel_4 As System.Windows.Forms.MenuItem
	Public WithEvents Mnu2 As System.Windows.Forms.MenuItem
	Public MainMenu1 As System.Windows.Forms.MainMenu
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmAgnMgr))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.MainPhld = New PageHolder
		Me.MainPdck = New PageDock
		Me._AgsMnu_3 = New XpButton
		Me._AgsMnu_0 = New XpButton
		Me._AgsMnu_1 = New XpButton
		Me._AgsMnu_2 = New XpButton
		Me._AgsCmd_0 = New XpButton
		Me._CptMnu_0 = New XpButton
		Me._CptMnu_1 = New XpButton
		Me._CptMnu_2 = New XpButton
		Me._CptMnu_3 = New XpButton
		Me._CptMnu_4 = New XpButton
		Me.MainBnr = New System.Windows.Forms.Panel
		Me.MainBnrCap = New System.Windows.Forms.Label
		Me.MainBnrLbl = New System.Windows.Forms.Label
		Me.MainLne = New Line3D
		Me.LstVw1 = New AxMSComctlLib.AxListView
		Me._Pages_0 = New System.Windows.Forms.Panel
		Me.GcnCmdCB = New System.Windows.Forms.ComboBox
		Me.GcnList = New System.Windows.Forms.ListBox
		Me.GcnBtnClr = New System.Windows.Forms.PictureBox
		Me.GcnCmdBtn = New System.Windows.Forms.PictureBox
		Me._GcnHdr_1 = New System.Windows.Forms.Label
		Me._GcnHdr_0 = New System.Windows.Forms.Label
		Me._Pages_1 = New System.Windows.Forms.Panel
		Me._GenOpt2_0 = New System.Windows.Forms.CheckBox
		Me._GenOpt2_1 = New System.Windows.Forms.CheckBox
		Me._GenOpt2_2 = New System.Windows.Forms.CheckBox
		Me._GenOpt2_3 = New System.Windows.Forms.CheckBox
		Me._GenOpt1_0 = New System.Windows.Forms.CheckBox
		Me.GenWelcome = New System.Windows.Forms.TextBox
		Me.GenPass2 = New System.Windows.Forms.TextBox
		Me.GenPass1 = New System.Windows.Forms.TextBox
		Me._GenOpt1_1 = New System.Windows.Forms.CheckBox
		Me.GenNetName = New System.Windows.Forms.TextBox
		Me.GenNetIP = New System.Windows.Forms.TextBox
		Me.GenNetPort = New System.Windows.Forms.TextBox
		Me._GenHdr_3 = New System.Windows.Forms.Label
		Me._GenHdr_2 = New System.Windows.Forms.Label
		Me._GenMiscLbl_0 = New System.Windows.Forms.Label
		Me._GenHdr_1 = New System.Windows.Forms.Label
		Me._GenPassLbl_1 = New System.Windows.Forms.Label
		Me._GenPassLbl_0 = New System.Windows.Forms.Label
		Me._GenHdr_0 = New System.Windows.Forms.Label
		Me._GenNetLbl_0 = New System.Windows.Forms.Label
		Me._GenNetLbl_2 = New System.Windows.Forms.Label
		Me._GenNetLbl_1 = New System.Windows.Forms.Label
		Me._Pages_2 = New System.Windows.Forms.Panel
		Me._Pages_4 = New System.Windows.Forms.Panel
		Me._Pages_3 = New System.Windows.Forms.Panel
		Me.AgsCmd = New XpButtonArray(components)
		Me.AgsMnu = New XpButtonArray(components)
		Me.CptMnu = New XpButtonArray(components)
		Me.GcnHdr = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.GenHdr = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.GenMiscLbl = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.GenNetLbl = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.GenOpt1 = New Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray(components)
		Me.GenOpt2 = New Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray(components)
		Me.GenPassLbl = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.Mnu2Sel = New Microsoft.VisualBasic.Compatibility.VB6.MenuItemArray(components)
		Me.Pages = New Microsoft.VisualBasic.Compatibility.VB6.PanelArray(components)
		Me.MainMenu1 = New System.Windows.Forms.MainMenu
		Me.Mnu1 = New System.Windows.Forms.MenuItem
		Me.Mnu1Rfsh = New System.Windows.Forms.MenuItem
		Me.Mnu1Close = New System.Windows.Forms.MenuItem
		Me.Mnu2 = New System.Windows.Forms.MenuItem
		Me._Mnu2Sel_0 = New System.Windows.Forms.MenuItem
		Me._Mnu2Sel_1 = New System.Windows.Forms.MenuItem
		Me._Mnu2Sel_2 = New System.Windows.Forms.MenuItem
		Me._Mnu2Sel_3 = New System.Windows.Forms.MenuItem
		Me._Mnu2Sel_4 = New System.Windows.Forms.MenuItem
		CType(Me.LstVw1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.AgsCmd, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.AgsMnu, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.CptMnu, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.GcnHdr, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.GenHdr, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.GenMiscLbl, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.GenNetLbl, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.GenOpt1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.GenOpt2, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.GenPassLbl, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Mnu2Sel, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Pages, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Text = "CafeBonzer - Agent Manager"
		Me.ClientSize = New System.Drawing.Size(687, 465)
		Me.Location = New System.Drawing.Point(10, 48)
		Me.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Icon = CType(resources.GetObject("FrmAgnMgr.Icon"), System.Drawing.Icon)
		Me.MaximizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "FrmAgnMgr"
		Me.MainPhld.Dock = System.Windows.Forms.DockStyle.Bottom
		Me.MainPhld.Size = New System.Drawing.Size(687, 64)
		Me.MainPhld.Location = New System.Drawing.Point(0, 401)
		Me.MainPhld.TabIndex = 36
		Me.MainPhld.HldrTxt = "Control Option"
		Me.MainPhld.HldrTxtClr = 16777215
		Me.MainPhld.HldrLne = -1
		Me.MainPhld.PageHeight = 960
		Me.MainPhld.Name = "MainPhld"
		Me.MainPdck.Size = New System.Drawing.Size(19, 41)
		Me.MainPdck.Location = New System.Drawing.Point(668, 23)
		Me.MainPdck.TabIndex = 42
		Me.MainPdck.HldrBtnPos = 0
		Me.MainPdck.HldrLne = -1
		Me.MainPdck.PageState = 1
		Me.MainPdck.PageWidth = 10305
		Me.MainPdck.Name = "MainPdck"
		Me._AgsMnu_3.Size = New System.Drawing.Size(28, 28)
		Me._AgsMnu_3.Location = New System.Drawing.Point(121, 6)
		Me._AgsMnu_3.TabIndex = 47
		Me._AgsMnu_3.TX = ""
		Me._AgsMnu_3.ENAB = -1
		Me._AgsMnu_3.COLTYPE = 2
		Me._AgsMnu_3.FOCUSR = -1
		Me._AgsMnu_3.BCOL = 16777215
		Me._AgsMnu_3.BCOLO = 16777215
		Me._AgsMnu_3.FCOL = 0
		Me._AgsMnu_3.FCOLO = 0
		Me._AgsMnu_3.MCOL = 16777215
		Me._AgsMnu_3.MPTR = 1
		Me._AgsMnu_3.MICON = 0
		Me._AgsMnu_3.PICN = 0
		Me._AgsMnu_3.UMCOL = -1
		Me._AgsMnu_3.SOFT = 0
		Me._AgsMnu_3.PICPOS = 0
		Me._AgsMnu_3.NGREY = 0
		Me._AgsMnu_3.FX = 0
		Me._AgsMnu_3.HAND = 0
		Me._AgsMnu_3.CHECK = 0
		Me._AgsMnu_3.Name = "_AgsMnu_3"
		Me._AgsMnu_0.Size = New System.Drawing.Size(28, 28)
		Me._AgsMnu_0.Location = New System.Drawing.Point(31, 6)
		Me._AgsMnu_0.TabIndex = 46
		Me._AgsMnu_0.TX = ""
		Me._AgsMnu_0.ENAB = -1
		Me._AgsMnu_0.COLTYPE = 2
		Me._AgsMnu_0.FOCUSR = -1
		Me._AgsMnu_0.BCOL = 16777215
		Me._AgsMnu_0.BCOLO = 16777215
		Me._AgsMnu_0.FCOL = 0
		Me._AgsMnu_0.FCOLO = 0
		Me._AgsMnu_0.MCOL = 16777215
		Me._AgsMnu_0.MPTR = 1
		Me._AgsMnu_0.MICON = 0
		Me._AgsMnu_0.PICN = 0
		Me._AgsMnu_0.UMCOL = -1
		Me._AgsMnu_0.SOFT = 0
		Me._AgsMnu_0.PICPOS = 0
		Me._AgsMnu_0.NGREY = 0
		Me._AgsMnu_0.FX = 0
		Me._AgsMnu_0.HAND = 0
		Me._AgsMnu_0.CHECK = 0
		Me._AgsMnu_0.Name = "_AgsMnu_0"
		Me._AgsMnu_1.Size = New System.Drawing.Size(28, 28)
		Me._AgsMnu_1.Location = New System.Drawing.Point(61, 6)
		Me._AgsMnu_1.TabIndex = 45
		Me._AgsMnu_1.TX = ""
		Me._AgsMnu_1.ENAB = -1
		Me._AgsMnu_1.COLTYPE = 2
		Me._AgsMnu_1.FOCUSR = -1
		Me._AgsMnu_1.BCOL = 16777215
		Me._AgsMnu_1.BCOLO = 16777215
		Me._AgsMnu_1.FCOL = 0
		Me._AgsMnu_1.FCOLO = 0
		Me._AgsMnu_1.MCOL = 16777215
		Me._AgsMnu_1.MPTR = 1
		Me._AgsMnu_1.MICON = 0
		Me._AgsMnu_1.PICN = 0
		Me._AgsMnu_1.UMCOL = -1
		Me._AgsMnu_1.SOFT = 0
		Me._AgsMnu_1.PICPOS = 0
		Me._AgsMnu_1.NGREY = 0
		Me._AgsMnu_1.FX = 0
		Me._AgsMnu_1.HAND = 0
		Me._AgsMnu_1.CHECK = 0
		Me._AgsMnu_1.Name = "_AgsMnu_1"
		Me._AgsMnu_2.Size = New System.Drawing.Size(28, 28)
		Me._AgsMnu_2.Location = New System.Drawing.Point(91, 6)
		Me._AgsMnu_2.TabIndex = 44
		Me._AgsMnu_2.TX = ""
		Me._AgsMnu_2.ENAB = -1
		Me._AgsMnu_2.COLTYPE = 2
		Me._AgsMnu_2.FOCUSR = -1
		Me._AgsMnu_2.BCOL = 16777215
		Me._AgsMnu_2.BCOLO = 16777215
		Me._AgsMnu_2.FCOL = 0
		Me._AgsMnu_2.FCOLO = 0
		Me._AgsMnu_2.MCOL = 16777215
		Me._AgsMnu_2.MPTR = 1
		Me._AgsMnu_2.MICON = 0
		Me._AgsMnu_2.PICN = 0
		Me._AgsMnu_2.UMCOL = -1
		Me._AgsMnu_2.SOFT = 0
		Me._AgsMnu_2.PICPOS = 0
		Me._AgsMnu_2.NGREY = 0
		Me._AgsMnu_2.FX = 0
		Me._AgsMnu_2.HAND = 0
		Me._AgsMnu_2.CHECK = 0
		Me._AgsMnu_2.Name = "_AgsMnu_2"
		Me._AgsCmd_0.Size = New System.Drawing.Size(64, 28)
		Me._AgsCmd_0.Location = New System.Drawing.Point(618, 6)
		Me._AgsCmd_0.TabIndex = 43
		Me._AgsCmd_0.TX = "Send"
		Me._AgsCmd_0.ENAB = -1
		Me._AgsCmd_0.COLTYPE = 2
		Me._AgsCmd_0.FOCUSR = -1
		Me._AgsCmd_0.BCOL = 16777215
		Me._AgsCmd_0.BCOLO = 16777215
		Me._AgsCmd_0.FCOL = 0
		Me._AgsCmd_0.FCOLO = 0
		Me._AgsCmd_0.MCOL = 16777215
		Me._AgsCmd_0.MPTR = 1
		Me._AgsCmd_0.MICON = 0
		Me._AgsCmd_0.PICN = 0
		Me._AgsCmd_0.UMCOL = -1
		Me._AgsCmd_0.SOFT = 0
		Me._AgsCmd_0.PICPOS = 0
		Me._AgsCmd_0.NGREY = 0
		Me._AgsCmd_0.FX = 0
		Me._AgsCmd_0.HAND = 0
		Me._AgsCmd_0.CHECK = 0
		Me._AgsCmd_0.Name = "_AgsCmd_0"
		Me._CptMnu_0.Size = New System.Drawing.Size(32, 32)
		Me._CptMnu_0.Location = New System.Drawing.Point(6, 27)
		Me._CptMnu_0.TabIndex = 41
		Me.ToolTip1.SetToolTip(Me._CptMnu_0, "UnLock")
		Me._CptMnu_0.TX = ""
		Me._CptMnu_0.ENAB = -1
		Me._CptMnu_0.COLTYPE = 2
		Me._CptMnu_0.FOCUSR = -1
		Me._CptMnu_0.BCOL = 16777215
		Me._CptMnu_0.BCOLO = 16777215
		Me._CptMnu_0.FCOL = 0
		Me._CptMnu_0.FCOLO = 0
		Me._CptMnu_0.MCOL = 16777215
		Me._CptMnu_0.MPTR = 1
		Me._CptMnu_0.MICON = 0
		Me._CptMnu_0.PICN = 0
		Me._CptMnu_0.UMCOL = -1
		Me._CptMnu_0.SOFT = 0
		Me._CptMnu_0.PICPOS = 0
		Me._CptMnu_0.NGREY = 0
		Me._CptMnu_0.FX = 0
		Me._CptMnu_0.HAND = 0
		Me._CptMnu_0.CHECK = 0
		Me._CptMnu_0.Name = "_CptMnu_0"
		Me._CptMnu_1.Size = New System.Drawing.Size(32, 32)
		Me._CptMnu_1.Location = New System.Drawing.Point(41, 27)
		Me._CptMnu_1.TabIndex = 40
		Me.ToolTip1.SetToolTip(Me._CptMnu_1, "Lock")
		Me._CptMnu_1.TX = ""
		Me._CptMnu_1.ENAB = -1
		Me._CptMnu_1.COLTYPE = 2
		Me._CptMnu_1.FOCUSR = -1
		Me._CptMnu_1.BCOL = 16777215
		Me._CptMnu_1.BCOLO = 16777215
		Me._CptMnu_1.FCOL = 0
		Me._CptMnu_1.FCOLO = 0
		Me._CptMnu_1.MCOL = 16777215
		Me._CptMnu_1.MPTR = 1
		Me._CptMnu_1.MICON = 0
		Me._CptMnu_1.PICN = 0
		Me._CptMnu_1.UMCOL = -1
		Me._CptMnu_1.SOFT = 0
		Me._CptMnu_1.PICPOS = 0
		Me._CptMnu_1.NGREY = 0
		Me._CptMnu_1.FX = 0
		Me._CptMnu_1.HAND = 0
		Me._CptMnu_1.CHECK = 0
		Me._CptMnu_1.Name = "_CptMnu_1"
		Me._CptMnu_2.Size = New System.Drawing.Size(32, 32)
		Me._CptMnu_2.Location = New System.Drawing.Point(75, 27)
		Me._CptMnu_2.TabIndex = 39
		Me.ToolTip1.SetToolTip(Me._CptMnu_2, "Shutdown")
		Me._CptMnu_2.TX = ""
		Me._CptMnu_2.ENAB = -1
		Me._CptMnu_2.COLTYPE = 2
		Me._CptMnu_2.FOCUSR = -1
		Me._CptMnu_2.BCOL = 16777215
		Me._CptMnu_2.BCOLO = 16777215
		Me._CptMnu_2.FCOL = 0
		Me._CptMnu_2.FCOLO = 0
		Me._CptMnu_2.MCOL = 16777215
		Me._CptMnu_2.MPTR = 1
		Me._CptMnu_2.MICON = 0
		Me._CptMnu_2.PICN = 0
		Me._CptMnu_2.UMCOL = -1
		Me._CptMnu_2.SOFT = 0
		Me._CptMnu_2.PICPOS = 0
		Me._CptMnu_2.NGREY = 0
		Me._CptMnu_2.FX = 0
		Me._CptMnu_2.HAND = 0
		Me._CptMnu_2.CHECK = 0
		Me._CptMnu_2.Name = "_CptMnu_2"
		Me._CptMnu_3.Size = New System.Drawing.Size(32, 32)
		Me._CptMnu_3.Location = New System.Drawing.Point(110, 27)
		Me._CptMnu_3.TabIndex = 38
		Me.ToolTip1.SetToolTip(Me._CptMnu_3, "Restart")
		Me._CptMnu_3.TX = ""
		Me._CptMnu_3.ENAB = -1
		Me._CptMnu_3.COLTYPE = 2
		Me._CptMnu_3.FOCUSR = -1
		Me._CptMnu_3.BCOL = 16777215
		Me._CptMnu_3.BCOLO = 16777215
		Me._CptMnu_3.FCOL = 0
		Me._CptMnu_3.FCOLO = 0
		Me._CptMnu_3.MCOL = 16777215
		Me._CptMnu_3.MPTR = 1
		Me._CptMnu_3.MICON = 0
		Me._CptMnu_3.PICN = 0
		Me._CptMnu_3.UMCOL = -1
		Me._CptMnu_3.SOFT = 0
		Me._CptMnu_3.PICPOS = 0
		Me._CptMnu_3.NGREY = 0
		Me._CptMnu_3.FX = 0
		Me._CptMnu_3.HAND = 0
		Me._CptMnu_3.CHECK = 0
		Me._CptMnu_3.Name = "_CptMnu_3"
		Me._CptMnu_4.Size = New System.Drawing.Size(32, 32)
		Me._CptMnu_4.Location = New System.Drawing.Point(145, 27)
		Me._CptMnu_4.TabIndex = 37
		Me.ToolTip1.SetToolTip(Me._CptMnu_4, "Close")
		Me._CptMnu_4.TX = ""
		Me._CptMnu_4.ENAB = -1
		Me._CptMnu_4.COLTYPE = 2
		Me._CptMnu_4.FOCUSR = -1
		Me._CptMnu_4.BCOL = 16777215
		Me._CptMnu_4.BCOLO = 16777215
		Me._CptMnu_4.FCOL = 0
		Me._CptMnu_4.FCOLO = 0
		Me._CptMnu_4.MCOL = 16777215
		Me._CptMnu_4.MPTR = 1
		Me._CptMnu_4.MICON = 0
		Me._CptMnu_4.PICN = 0
		Me._CptMnu_4.UMCOL = -1
		Me._CptMnu_4.SOFT = 0
		Me._CptMnu_4.PICPOS = 0
		Me._CptMnu_4.NGREY = 0
		Me._CptMnu_4.FX = 0
		Me._CptMnu_4.HAND = 0
		Me._CptMnu_4.CHECK = 0
		Me._CptMnu_4.Name = "_CptMnu_4"
		Me.MainBnr.BackColor = System.Drawing.Color.White
		Me.MainBnr.Size = New System.Drawing.Size(482, 54)
		Me.MainBnr.Location = New System.Drawing.Point(203, 5)
		Me.MainBnr.BackgroundImage = CType(resources.GetObject("MainBnr.BackgroundImage"), System.Drawing.Image)
		Me.MainBnr.TabIndex = 26
		Me.MainBnr.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.MainBnr.Dock = System.Windows.Forms.DockStyle.None
		Me.MainBnr.CausesValidation = True
		Me.MainBnr.Enabled = True
		Me.MainBnr.ForeColor = System.Drawing.SystemColors.ControlText
		Me.MainBnr.Cursor = System.Windows.Forms.Cursors.Default
		Me.MainBnr.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.MainBnr.TabStop = True
		Me.MainBnr.Visible = True
		Me.MainBnr.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.MainBnr.Name = "MainBnr"
		Me.MainBnrCap.BackColor = System.Drawing.Color.Transparent
		Me.MainBnrCap.Text = "Send general command to agent."
		Me.MainBnrCap.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me.MainBnrCap.Size = New System.Drawing.Size(191, 13)
		Me.MainBnrCap.Location = New System.Drawing.Point(56, 24)
		Me.MainBnrCap.TabIndex = 28
		Me.MainBnrCap.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.MainBnrCap.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.MainBnrCap.Enabled = True
		Me.MainBnrCap.Cursor = System.Windows.Forms.Cursors.Default
		Me.MainBnrCap.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.MainBnrCap.UseMnemonic = True
		Me.MainBnrCap.Visible = True
		Me.MainBnrCap.AutoSize = True
		Me.MainBnrCap.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.MainBnrCap.Name = "MainBnrCap"
		Me.MainBnrLbl.BackColor = System.Drawing.Color.Transparent
		Me.MainBnrLbl.Text = "General Control"
		Me.MainBnrLbl.Font = New System.Drawing.Font("Verdana", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.MainBnrLbl.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me.MainBnrLbl.Size = New System.Drawing.Size(138, 18)
		Me.MainBnrLbl.Location = New System.Drawing.Point(56, 4)
		Me.MainBnrLbl.TabIndex = 27
		Me.MainBnrLbl.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.MainBnrLbl.Enabled = True
		Me.MainBnrLbl.Cursor = System.Windows.Forms.Cursors.Default
		Me.MainBnrLbl.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.MainBnrLbl.UseMnemonic = True
		Me.MainBnrLbl.Visible = True
		Me.MainBnrLbl.AutoSize = True
		Me.MainBnrLbl.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.MainBnrLbl.Name = "MainBnrLbl"
		Me.MainLne.Size = New System.Drawing.Size(687, 3)
		Me.MainLne.Location = New System.Drawing.Point(0, -1)
		Me.MainLne.TabIndex = 1
		Me.MainLne.horizon = -1
		Me.MainLne.Name = "MainLne"
		LstVw1.OcxState = CType(resources.GetObject("LstVw1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.LstVw1.Size = New System.Drawing.Size(198, 395)
		Me.LstVw1.Location = New System.Drawing.Point(2, 5)
		Me.LstVw1.TabIndex = 2
		Me.LstVw1.Name = "LstVw1"
		Me._Pages_0.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me._Pages_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Pages_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._Pages_0.Size = New System.Drawing.Size(482, 336)
		Me._Pages_0.Location = New System.Drawing.Point(204, 62)
		Me._Pages_0.TabIndex = 0
		Me._Pages_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Pages_0.Enabled = True
		Me._Pages_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._Pages_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Pages_0.Visible = True
		Me._Pages_0.Name = "_Pages_0"
		Me.GcnCmdCB.Size = New System.Drawing.Size(433, 21)
		Me.GcnCmdCB.Location = New System.Drawing.Point(19, 233)
		Me.GcnCmdCB.TabIndex = 35
		Me.GcnCmdCB.Text = "block:1"
		Me.GcnCmdCB.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.GcnCmdCB.BackColor = System.Drawing.SystemColors.Window
		Me.GcnCmdCB.CausesValidation = True
		Me.GcnCmdCB.Enabled = True
		Me.GcnCmdCB.ForeColor = System.Drawing.SystemColors.WindowText
		Me.GcnCmdCB.IntegralHeight = True
		Me.GcnCmdCB.Cursor = System.Windows.Forms.Cursors.Default
		Me.GcnCmdCB.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.GcnCmdCB.Sorted = False
		Me.GcnCmdCB.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.GcnCmdCB.TabStop = True
		Me.GcnCmdCB.Visible = True
		Me.GcnCmdCB.Name = "GcnCmdCB"
		Me.GcnList.BackColor = System.Drawing.Color.FromARGB(224, 224, 224)
		Me.GcnList.Enabled = False
		Me.GcnList.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me.GcnList.Size = New System.Drawing.Size(435, 163)
		Me.GcnList.Location = New System.Drawing.Point(20, 33)
		Me.GcnList.TabIndex = 32
		Me.GcnList.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.GcnList.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.GcnList.CausesValidation = True
		Me.GcnList.IntegralHeight = True
		Me.GcnList.Cursor = System.Windows.Forms.Cursors.Default
		Me.GcnList.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.GcnList.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.GcnList.Sorted = False
		Me.GcnList.TabStop = True
		Me.GcnList.Visible = True
		Me.GcnList.MultiColumn = False
		Me.GcnList.Name = "GcnList"
		Me.GcnBtnClr.Size = New System.Drawing.Size(16, 16)
		Me.GcnBtnClr.Location = New System.Drawing.Point(459, 33)
		Me.GcnBtnClr.Image = CType(resources.GetObject("GcnBtnClr.Image"), System.Drawing.Image)
		Me.ToolTip1.SetToolTip(Me.GcnBtnClr, "Send Command")
		Me.GcnBtnClr.Enabled = True
		Me.GcnBtnClr.Cursor = System.Windows.Forms.Cursors.Default
		Me.GcnBtnClr.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.GcnBtnClr.Visible = True
		Me.GcnBtnClr.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.GcnBtnClr.Name = "GcnBtnClr"
		Me.GcnCmdBtn.Size = New System.Drawing.Size(16, 16)
		Me.GcnCmdBtn.Location = New System.Drawing.Point(456, 235)
		Me.GcnCmdBtn.Image = CType(resources.GetObject("GcnCmdBtn.Image"), System.Drawing.Image)
		Me.ToolTip1.SetToolTip(Me.GcnCmdBtn, "Send Command")
		Me.GcnCmdBtn.Enabled = True
		Me.GcnCmdBtn.Cursor = System.Windows.Forms.Cursors.Default
		Me.GcnCmdBtn.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.GcnCmdBtn.Visible = True
		Me.GcnCmdBtn.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.GcnCmdBtn.Name = "GcnCmdBtn"
		Me._GcnHdr_1.BackColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me._GcnHdr_1.Text = " Custom Command"
		Me._GcnHdr_1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GcnHdr_1.ForeColor = System.Drawing.Color.White
		Me._GcnHdr_1.Size = New System.Drawing.Size(472, 18)
		Me._GcnHdr_1.Location = New System.Drawing.Point(5, 206)
		Me._GcnHdr_1.TabIndex = 34
		Me._GcnHdr_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._GcnHdr_1.Enabled = True
		Me._GcnHdr_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._GcnHdr_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GcnHdr_1.UseMnemonic = True
		Me._GcnHdr_1.Visible = True
		Me._GcnHdr_1.AutoSize = False
		Me._GcnHdr_1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._GcnHdr_1.Name = "_GcnHdr_1"
		Me._GcnHdr_0.BackColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me._GcnHdr_0.Text = " Summary"
		Me._GcnHdr_0.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GcnHdr_0.ForeColor = System.Drawing.Color.White
		Me._GcnHdr_0.Size = New System.Drawing.Size(472, 18)
		Me._GcnHdr_0.Location = New System.Drawing.Point(4, 7)
		Me._GcnHdr_0.TabIndex = 33
		Me._GcnHdr_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._GcnHdr_0.Enabled = True
		Me._GcnHdr_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._GcnHdr_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GcnHdr_0.UseMnemonic = True
		Me._GcnHdr_0.Visible = True
		Me._GcnHdr_0.AutoSize = False
		Me._GcnHdr_0.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._GcnHdr_0.Name = "_GcnHdr_0"
		Me._Pages_1.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me._Pages_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Pages_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._Pages_1.Size = New System.Drawing.Size(482, 336)
		Me._Pages_1.Location = New System.Drawing.Point(204, 62)
		Me._Pages_1.TabIndex = 3
		Me._Pages_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Pages_1.Enabled = True
		Me._Pages_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Pages_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Pages_1.Visible = True
		Me._Pages_1.Name = "_Pages_1"
		Me._GenOpt2_0.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me._GenOpt2_0.Text = "Monitor print activity."
		Me._GenOpt2_0.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GenOpt2_0.ForeColor = System.Drawing.Color.Black
		Me._GenOpt2_0.Size = New System.Drawing.Size(190, 21)
		Me._GenOpt2_0.Location = New System.Drawing.Point(263, 29)
		Me._GenOpt2_0.TabIndex = 23
		Me.ToolTip1.SetToolTip(Me._GenOpt2_0, "Monitor print activity.")
		Me._GenOpt2_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._GenOpt2_0.CausesValidation = True
		Me._GenOpt2_0.Enabled = True
		Me._GenOpt2_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._GenOpt2_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GenOpt2_0.Appearance = System.Windows.Forms.Appearance.Normal
		Me._GenOpt2_0.TabStop = True
		Me._GenOpt2_0.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._GenOpt2_0.Visible = True
		Me._GenOpt2_0.Name = "_GenOpt2_0"
		Me._GenOpt2_1.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me._GenOpt2_1.Text = "Monitor system resource."
		Me._GenOpt2_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GenOpt2_1.ForeColor = System.Drawing.Color.Black
		Me._GenOpt2_1.Size = New System.Drawing.Size(190, 21)
		Me._GenOpt2_1.Location = New System.Drawing.Point(263, 52)
		Me._GenOpt2_1.TabIndex = 22
		Me.ToolTip1.SetToolTip(Me._GenOpt2_1, "Monitor print activity.")
		Me._GenOpt2_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._GenOpt2_1.CausesValidation = True
		Me._GenOpt2_1.Enabled = True
		Me._GenOpt2_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._GenOpt2_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GenOpt2_1.Appearance = System.Windows.Forms.Appearance.Normal
		Me._GenOpt2_1.TabStop = True
		Me._GenOpt2_1.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._GenOpt2_1.Visible = True
		Me._GenOpt2_1.Name = "_GenOpt2_1"
		Me._GenOpt2_2.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me._GenOpt2_2.Text = "Monitor applications."
		Me._GenOpt2_2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GenOpt2_2.ForeColor = System.Drawing.Color.Black
		Me._GenOpt2_2.Size = New System.Drawing.Size(190, 21)
		Me._GenOpt2_2.Location = New System.Drawing.Point(263, 75)
		Me._GenOpt2_2.TabIndex = 21
		Me.ToolTip1.SetToolTip(Me._GenOpt2_2, "Monitor process & applications.")
		Me._GenOpt2_2.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._GenOpt2_2.CausesValidation = True
		Me._GenOpt2_2.Enabled = True
		Me._GenOpt2_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._GenOpt2_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GenOpt2_2.Appearance = System.Windows.Forms.Appearance.Normal
		Me._GenOpt2_2.TabStop = True
		Me._GenOpt2_2.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._GenOpt2_2.Visible = True
		Me._GenOpt2_2.Name = "_GenOpt2_2"
		Me._GenOpt2_3.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me._GenOpt2_3.Text = "Monitor network traffic."
		Me._GenOpt2_3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GenOpt2_3.ForeColor = System.Drawing.Color.Black
		Me._GenOpt2_3.Size = New System.Drawing.Size(190, 21)
		Me._GenOpt2_3.Location = New System.Drawing.Point(263, 98)
		Me._GenOpt2_3.TabIndex = 20
		Me.ToolTip1.SetToolTip(Me._GenOpt2_3, "Monitor network traffic.")
		Me._GenOpt2_3.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._GenOpt2_3.CausesValidation = True
		Me._GenOpt2_3.Enabled = True
		Me._GenOpt2_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._GenOpt2_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GenOpt2_3.Appearance = System.Windows.Forms.Appearance.Normal
		Me._GenOpt2_3.TabStop = True
		Me._GenOpt2_3.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._GenOpt2_3.Visible = True
		Me._GenOpt2_3.Name = "_GenOpt2_3"
		Me._GenOpt1_0.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me._GenOpt1_0.Text = "Autostart on windows begin."
		Me._GenOpt1_0.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GenOpt1_0.ForeColor = System.Drawing.Color.Black
		Me._GenOpt1_0.Size = New System.Drawing.Size(190, 21)
		Me._GenOpt1_0.Location = New System.Drawing.Point(264, 163)
		Me._GenOpt1_0.TabIndex = 18
		Me.ToolTip1.SetToolTip(Me._GenOpt1_0, "Automatic start CafeBonzer")
		Me._GenOpt1_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._GenOpt1_0.CausesValidation = True
		Me._GenOpt1_0.Enabled = True
		Me._GenOpt1_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._GenOpt1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GenOpt1_0.Appearance = System.Windows.Forms.Appearance.Normal
		Me._GenOpt1_0.TabStop = True
		Me._GenOpt1_0.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._GenOpt1_0.Visible = True
		Me._GenOpt1_0.Name = "_GenOpt1_0"
		Me.GenWelcome.AutoSize = False
		Me.GenWelcome.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.GenWelcome.BackColor = System.Drawing.Color.White
		Me.GenWelcome.Size = New System.Drawing.Size(187, 21)
		Me.GenWelcome.IMEMode = System.Windows.Forms.ImeMode.Disable
		Me.GenWelcome.Location = New System.Drawing.Point(277, 210)
		Me.GenWelcome.TabIndex = 17
		Me.GenWelcome.Text = ":: CafeBonzer Agent R1 ::"
		Me.GenWelcome.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.GenWelcome.AcceptsReturn = True
		Me.GenWelcome.CausesValidation = True
		Me.GenWelcome.Enabled = True
		Me.GenWelcome.ForeColor = System.Drawing.SystemColors.WindowText
		Me.GenWelcome.HideSelection = True
		Me.GenWelcome.ReadOnly = False
		Me.GenWelcome.Maxlength = 0
		Me.GenWelcome.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.GenWelcome.MultiLine = False
		Me.GenWelcome.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.GenWelcome.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.GenWelcome.TabStop = True
		Me.GenWelcome.Visible = True
		Me.GenWelcome.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.GenWelcome.Name = "GenWelcome"
		Me.GenPass2.AutoSize = False
		Me.GenPass2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.GenPass2.BackColor = System.Drawing.Color.White
		Me.GenPass2.Font = New System.Drawing.Font("Wingdings", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
		Me.GenPass2.Size = New System.Drawing.Size(114, 21)
		Me.GenPass2.IMEMode = System.Windows.Forms.ImeMode.Disable
		Me.GenPass2.Location = New System.Drawing.Point(107, 209)
		Me.GenPass2.PasswordChar = ChrW(108)
		Me.GenPass2.TabIndex = 13
		Me.GenPass2.AcceptsReturn = True
		Me.GenPass2.CausesValidation = True
		Me.GenPass2.Enabled = True
		Me.GenPass2.ForeColor = System.Drawing.SystemColors.WindowText
		Me.GenPass2.HideSelection = True
		Me.GenPass2.ReadOnly = False
		Me.GenPass2.Maxlength = 0
		Me.GenPass2.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.GenPass2.MultiLine = False
		Me.GenPass2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.GenPass2.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.GenPass2.TabStop = True
		Me.GenPass2.Visible = True
		Me.GenPass2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.GenPass2.Name = "GenPass2"
		Me.GenPass1.AutoSize = False
		Me.GenPass1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.GenPass1.BackColor = System.Drawing.Color.White
		Me.GenPass1.Font = New System.Drawing.Font("Wingdings", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
		Me.GenPass1.Size = New System.Drawing.Size(114, 21)
		Me.GenPass1.IMEMode = System.Windows.Forms.ImeMode.Disable
		Me.GenPass1.Location = New System.Drawing.Point(107, 182)
		Me.GenPass1.PasswordChar = ChrW(108)
		Me.GenPass1.TabIndex = 12
		Me.GenPass1.AcceptsReturn = True
		Me.GenPass1.CausesValidation = True
		Me.GenPass1.Enabled = True
		Me.GenPass1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.GenPass1.HideSelection = True
		Me.GenPass1.ReadOnly = False
		Me.GenPass1.Maxlength = 0
		Me.GenPass1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.GenPass1.MultiLine = False
		Me.GenPass1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.GenPass1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.GenPass1.TabStop = True
		Me.GenPass1.Visible = True
		Me.GenPass1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.GenPass1.Name = "GenPass1"
		Me._GenOpt1_1.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me._GenOpt1_1.Text = "Retrive default password on start."
		Me._GenOpt1_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GenOpt1_1.ForeColor = System.Drawing.Color.Black
		Me._GenOpt1_1.Size = New System.Drawing.Size(190, 21)
		Me._GenOpt1_1.Location = New System.Drawing.Point(23, 155)
		Me._GenOpt1_1.TabIndex = 11
		Me.ToolTip1.SetToolTip(Me._GenOpt1_1, "Retrive default password from server when windows start.")
		Me._GenOpt1_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._GenOpt1_1.CausesValidation = True
		Me._GenOpt1_1.Enabled = True
		Me._GenOpt1_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._GenOpt1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GenOpt1_1.Appearance = System.Windows.Forms.Appearance.Normal
		Me._GenOpt1_1.TabStop = True
		Me._GenOpt1_1.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._GenOpt1_1.Visible = True
		Me._GenOpt1_1.Name = "_GenOpt1_1"
		Me.GenNetName.AutoSize = False
		Me.GenNetName.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.GenNetName.BackColor = System.Drawing.Color.White
		Me.GenNetName.Size = New System.Drawing.Size(110, 21)
		Me.GenNetName.Location = New System.Drawing.Point(109, 33)
		Me.GenNetName.TabIndex = 6
		Me.GenNetName.Text = "Cake"
		Me.GenNetName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.GenNetName.AcceptsReturn = True
		Me.GenNetName.CausesValidation = True
		Me.GenNetName.Enabled = True
		Me.GenNetName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.GenNetName.HideSelection = True
		Me.GenNetName.ReadOnly = False
		Me.GenNetName.Maxlength = 0
		Me.GenNetName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.GenNetName.MultiLine = False
		Me.GenNetName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.GenNetName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.GenNetName.TabStop = True
		Me.GenNetName.Visible = True
		Me.GenNetName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.GenNetName.Name = "GenNetName"
		Me.GenNetIP.AutoSize = False
		Me.GenNetIP.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.GenNetIP.BackColor = System.Drawing.Color.White
		Me.GenNetIP.Size = New System.Drawing.Size(110, 21)
		Me.GenNetIP.Location = New System.Drawing.Point(109, 86)
		Me.GenNetIP.TabIndex = 5
		Me.GenNetIP.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.GenNetIP.AcceptsReturn = True
		Me.GenNetIP.CausesValidation = True
		Me.GenNetIP.Enabled = True
		Me.GenNetIP.ForeColor = System.Drawing.SystemColors.WindowText
		Me.GenNetIP.HideSelection = True
		Me.GenNetIP.ReadOnly = False
		Me.GenNetIP.Maxlength = 0
		Me.GenNetIP.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.GenNetIP.MultiLine = False
		Me.GenNetIP.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.GenNetIP.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.GenNetIP.TabStop = True
		Me.GenNetIP.Visible = True
		Me.GenNetIP.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.GenNetIP.Name = "GenNetIP"
		Me.GenNetPort.AutoSize = False
		Me.GenNetPort.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.GenNetPort.BackColor = System.Drawing.Color.White
		Me.GenNetPort.Size = New System.Drawing.Size(110, 21)
		Me.GenNetPort.Location = New System.Drawing.Point(109, 59)
		Me.GenNetPort.TabIndex = 4
		Me.GenNetPort.Text = "56266"
		Me.GenNetPort.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.GenNetPort.AcceptsReturn = True
		Me.GenNetPort.CausesValidation = True
		Me.GenNetPort.Enabled = True
		Me.GenNetPort.ForeColor = System.Drawing.SystemColors.WindowText
		Me.GenNetPort.HideSelection = True
		Me.GenNetPort.ReadOnly = False
		Me.GenNetPort.Maxlength = 0
		Me.GenNetPort.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.GenNetPort.MultiLine = False
		Me.GenNetPort.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.GenNetPort.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.GenNetPort.TabStop = True
		Me.GenNetPort.Visible = True
		Me.GenNetPort.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.GenNetPort.Name = "GenNetPort"
		Me._GenHdr_3.BackColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me._GenHdr_3.Text = " Miscelaneous"
		Me._GenHdr_3.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GenHdr_3.ForeColor = System.Drawing.Color.White
		Me._GenHdr_3.Size = New System.Drawing.Size(213, 18)
		Me._GenHdr_3.Location = New System.Drawing.Point(253, 136)
		Me._GenHdr_3.TabIndex = 25
		Me._GenHdr_3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._GenHdr_3.Enabled = True
		Me._GenHdr_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._GenHdr_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GenHdr_3.UseMnemonic = True
		Me._GenHdr_3.Visible = True
		Me._GenHdr_3.AutoSize = False
		Me._GenHdr_3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._GenHdr_3.Name = "_GenHdr_3"
		Me._GenHdr_2.BackColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me._GenHdr_2.Text = " Pc monitoring"
		Me._GenHdr_2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GenHdr_2.ForeColor = System.Drawing.Color.White
		Me._GenHdr_2.Size = New System.Drawing.Size(213, 18)
		Me._GenHdr_2.Location = New System.Drawing.Point(253, 7)
		Me._GenHdr_2.TabIndex = 24
		Me._GenHdr_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._GenHdr_2.Enabled = True
		Me._GenHdr_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._GenHdr_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GenHdr_2.UseMnemonic = True
		Me._GenHdr_2.Visible = True
		Me._GenHdr_2.AutoSize = False
		Me._GenHdr_2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._GenHdr_2.Name = "_GenHdr_2"
		Me._GenMiscLbl_0.Text = "Welcome Message :"
		Me._GenMiscLbl_0.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GenMiscLbl_0.Size = New System.Drawing.Size(97, 14)
		Me._GenMiscLbl_0.Location = New System.Drawing.Point(265, 191)
		Me._GenMiscLbl_0.TabIndex = 19
		Me._GenMiscLbl_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._GenMiscLbl_0.BackColor = System.Drawing.Color.Transparent
		Me._GenMiscLbl_0.Enabled = True
		Me._GenMiscLbl_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._GenMiscLbl_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._GenMiscLbl_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GenMiscLbl_0.UseMnemonic = True
		Me._GenMiscLbl_0.Visible = True
		Me._GenMiscLbl_0.AutoSize = True
		Me._GenMiscLbl_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._GenMiscLbl_0.Name = "_GenMiscLbl_0"
		Me._GenHdr_1.BackColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me._GenHdr_1.Text = " Agent Password"
		Me._GenHdr_1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GenHdr_1.ForeColor = System.Drawing.Color.White
		Me._GenHdr_1.Size = New System.Drawing.Size(213, 18)
		Me._GenHdr_1.Location = New System.Drawing.Point(16, 127)
		Me._GenHdr_1.TabIndex = 16
		Me._GenHdr_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._GenHdr_1.Enabled = True
		Me._GenHdr_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._GenHdr_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GenHdr_1.UseMnemonic = True
		Me._GenHdr_1.Visible = True
		Me._GenHdr_1.AutoSize = False
		Me._GenHdr_1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._GenHdr_1.Name = "_GenHdr_1"
		Me._GenPassLbl_1.Text = "Retype :"
		Me._GenPassLbl_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GenPassLbl_1.ForeColor = System.Drawing.Color.FromARGB(192, 0, 0)
		Me._GenPassLbl_1.Size = New System.Drawing.Size(57, 14)
		Me._GenPassLbl_1.Location = New System.Drawing.Point(23, 212)
		Me._GenPassLbl_1.TabIndex = 15
		Me._GenPassLbl_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._GenPassLbl_1.BackColor = System.Drawing.Color.Transparent
		Me._GenPassLbl_1.Enabled = True
		Me._GenPassLbl_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._GenPassLbl_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GenPassLbl_1.UseMnemonic = True
		Me._GenPassLbl_1.Visible = True
		Me._GenPassLbl_1.AutoSize = False
		Me._GenPassLbl_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._GenPassLbl_1.Name = "_GenPassLbl_1"
		Me._GenPassLbl_0.Text = "Password :"
		Me._GenPassLbl_0.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GenPassLbl_0.ForeColor = System.Drawing.Color.FromARGB(192, 0, 0)
		Me._GenPassLbl_0.Size = New System.Drawing.Size(57, 14)
		Me._GenPassLbl_0.Location = New System.Drawing.Point(23, 185)
		Me._GenPassLbl_0.TabIndex = 14
		Me._GenPassLbl_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._GenPassLbl_0.BackColor = System.Drawing.Color.Transparent
		Me._GenPassLbl_0.Enabled = True
		Me._GenPassLbl_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._GenPassLbl_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GenPassLbl_0.UseMnemonic = True
		Me._GenPassLbl_0.Visible = True
		Me._GenPassLbl_0.AutoSize = False
		Me._GenPassLbl_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._GenPassLbl_0.Name = "_GenPassLbl_0"
		Me._GenHdr_0.BackColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me._GenHdr_0.Text = " Network Configuration"
		Me._GenHdr_0.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GenHdr_0.ForeColor = System.Drawing.Color.White
		Me._GenHdr_0.Size = New System.Drawing.Size(213, 18)
		Me._GenHdr_0.Location = New System.Drawing.Point(10, 7)
		Me._GenHdr_0.TabIndex = 10
		Me._GenHdr_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._GenHdr_0.Enabled = True
		Me._GenHdr_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._GenHdr_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GenHdr_0.UseMnemonic = True
		Me._GenHdr_0.Visible = True
		Me._GenHdr_0.AutoSize = False
		Me._GenHdr_0.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._GenHdr_0.Name = "_GenHdr_0"
		Me._GenNetLbl_0.Text = "Computer Name :"
		Me._GenNetLbl_0.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GenNetLbl_0.Size = New System.Drawing.Size(82, 14)
		Me._GenNetLbl_0.Location = New System.Drawing.Point(21, 35)
		Me._GenNetLbl_0.TabIndex = 9
		Me._GenNetLbl_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._GenNetLbl_0.BackColor = System.Drawing.Color.Transparent
		Me._GenNetLbl_0.Enabled = True
		Me._GenNetLbl_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._GenNetLbl_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._GenNetLbl_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GenNetLbl_0.UseMnemonic = True
		Me._GenNetLbl_0.Visible = True
		Me._GenNetLbl_0.AutoSize = True
		Me._GenNetLbl_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._GenNetLbl_0.Name = "_GenNetLbl_0"
		Me._GenNetLbl_2.Text = "Server IP :"
		Me._GenNetLbl_2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GenNetLbl_2.Size = New System.Drawing.Size(50, 14)
		Me._GenNetLbl_2.Location = New System.Drawing.Point(21, 88)
		Me._GenNetLbl_2.TabIndex = 8
		Me._GenNetLbl_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._GenNetLbl_2.BackColor = System.Drawing.Color.Transparent
		Me._GenNetLbl_2.Enabled = True
		Me._GenNetLbl_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._GenNetLbl_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._GenNetLbl_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GenNetLbl_2.UseMnemonic = True
		Me._GenNetLbl_2.Visible = True
		Me._GenNetLbl_2.AutoSize = True
		Me._GenNetLbl_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._GenNetLbl_2.Name = "_GenNetLbl_2"
		Me._GenNetLbl_1.Text = "Server Port :"
		Me._GenNetLbl_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GenNetLbl_1.Size = New System.Drawing.Size(61, 14)
		Me._GenNetLbl_1.Location = New System.Drawing.Point(22, 61)
		Me._GenNetLbl_1.TabIndex = 7
		Me._GenNetLbl_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._GenNetLbl_1.BackColor = System.Drawing.Color.Transparent
		Me._GenNetLbl_1.Enabled = True
		Me._GenNetLbl_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._GenNetLbl_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._GenNetLbl_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GenNetLbl_1.UseMnemonic = True
		Me._GenNetLbl_1.Visible = True
		Me._GenNetLbl_1.AutoSize = True
		Me._GenNetLbl_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._GenNetLbl_1.Name = "_GenNetLbl_1"
		Me._Pages_2.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me._Pages_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Pages_2.ForeColor = System.Drawing.SystemColors.WindowText
		Me._Pages_2.Size = New System.Drawing.Size(482, 336)
		Me._Pages_2.Location = New System.Drawing.Point(204, 62)
		Me._Pages_2.TabIndex = 29
		Me._Pages_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Pages_2.Enabled = True
		Me._Pages_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._Pages_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Pages_2.Visible = True
		Me._Pages_2.Name = "_Pages_2"
		Me._Pages_4.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me._Pages_4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Pages_4.ForeColor = System.Drawing.SystemColors.WindowText
		Me._Pages_4.Size = New System.Drawing.Size(482, 336)
		Me._Pages_4.Location = New System.Drawing.Point(204, 62)
		Me._Pages_4.TabIndex = 31
		Me._Pages_4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Pages_4.Enabled = True
		Me._Pages_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._Pages_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Pages_4.Visible = True
		Me._Pages_4.Name = "_Pages_4"
		Me._Pages_3.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me._Pages_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Pages_3.ForeColor = System.Drawing.SystemColors.WindowText
		Me._Pages_3.Size = New System.Drawing.Size(482, 336)
		Me._Pages_3.Location = New System.Drawing.Point(204, 62)
		Me._Pages_3.TabIndex = 30
		Me._Pages_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Pages_3.Enabled = True
		Me._Pages_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._Pages_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Pages_3.Visible = True
		Me._Pages_3.Name = "_Pages_3"
		Me.Mnu1.Text = "Menu"
		Me.Mnu1.Checked = False
		Me.Mnu1.Enabled = True
		Me.Mnu1.Visible = True
		Me.Mnu1.MDIList = False
		Me.Mnu1Rfsh.Text = "Refresh"
		Me.Mnu1Rfsh.Checked = False
		Me.Mnu1Rfsh.Enabled = True
		Me.Mnu1Rfsh.Visible = True
		Me.Mnu1Rfsh.MDIList = False
		Me.Mnu1Close.Text = "Close"
		Me.Mnu1Close.Checked = False
		Me.Mnu1Close.Enabled = True
		Me.Mnu1Close.Visible = True
		Me.Mnu1Close.MDIList = False
		Me.Mnu2.Text = "Select"
		Me.Mnu2.Checked = False
		Me.Mnu2.Enabled = True
		Me.Mnu2.Visible = True
		Me.Mnu2.MDIList = False
		Me._Mnu2Sel_0.Text = "Select All"
		Me._Mnu2Sel_0.Checked = False
		Me._Mnu2Sel_0.Enabled = True
		Me._Mnu2Sel_0.Visible = True
		Me._Mnu2Sel_0.MDIList = False
		Me._Mnu2Sel_1.Text = "DeSelect All"
		Me._Mnu2Sel_1.Checked = False
		Me._Mnu2Sel_1.Enabled = True
		Me._Mnu2Sel_1.Visible = True
		Me._Mnu2Sel_1.MDIList = False
		Me._Mnu2Sel_2.Text = "Select Unused"
		Me._Mnu2Sel_2.Checked = False
		Me._Mnu2Sel_2.Enabled = True
		Me._Mnu2Sel_2.Visible = True
		Me._Mnu2Sel_2.MDIList = False
		Me._Mnu2Sel_3.Text = "Select Used"
		Me._Mnu2Sel_3.Checked = False
		Me._Mnu2Sel_3.Enabled = True
		Me._Mnu2Sel_3.Visible = True
		Me._Mnu2Sel_3.MDIList = False
		Me._Mnu2Sel_4.Text = "Select Unlock"
		Me._Mnu2Sel_4.Checked = False
		Me._Mnu2Sel_4.Enabled = True
		Me._Mnu2Sel_4.Visible = True
		Me._Mnu2Sel_4.MDIList = False
		Me.Controls.Add(MainPhld)
		Me.Controls.Add(MainBnr)
		Me.Controls.Add(MainLne)
		Me.Controls.Add(LstVw1)
		Me.Controls.Add(_Pages_0)
		Me.Controls.Add(_Pages_1)
		Me.Controls.Add(_Pages_2)
		Me.Controls.Add(_Pages_4)
		Me.Controls.Add(_Pages_3)
		Me.MainPhld.Controls.Add(MainPdck)
		Me.MainPhld.Controls.Add(_CptMnu_0)
		Me.MainPhld.Controls.Add(_CptMnu_1)
		Me.MainPhld.Controls.Add(_CptMnu_2)
		Me.MainPhld.Controls.Add(_CptMnu_3)
		Me.MainPhld.Controls.Add(_CptMnu_4)
		Me.MainPdck.Controls.Add(_AgsMnu_3)
		Me.MainPdck.Controls.Add(_AgsMnu_0)
		Me.MainPdck.Controls.Add(_AgsMnu_1)
		Me.MainPdck.Controls.Add(_AgsMnu_2)
		Me.MainPdck.Controls.Add(_AgsCmd_0)
		Me.MainBnr.Controls.Add(MainBnrCap)
		Me.MainBnr.Controls.Add(MainBnrLbl)
		Me._Pages_0.Controls.Add(GcnCmdCB)
		Me._Pages_0.Controls.Add(GcnList)
		Me._Pages_0.Controls.Add(GcnBtnClr)
		Me._Pages_0.Controls.Add(GcnCmdBtn)
		Me._Pages_0.Controls.Add(_GcnHdr_1)
		Me._Pages_0.Controls.Add(_GcnHdr_0)
		Me._Pages_1.Controls.Add(_GenOpt2_0)
		Me._Pages_1.Controls.Add(_GenOpt2_1)
		Me._Pages_1.Controls.Add(_GenOpt2_2)
		Me._Pages_1.Controls.Add(_GenOpt2_3)
		Me._Pages_1.Controls.Add(_GenOpt1_0)
		Me._Pages_1.Controls.Add(GenWelcome)
		Me._Pages_1.Controls.Add(GenPass2)
		Me._Pages_1.Controls.Add(GenPass1)
		Me._Pages_1.Controls.Add(_GenOpt1_1)
		Me._Pages_1.Controls.Add(GenNetName)
		Me._Pages_1.Controls.Add(GenNetIP)
		Me._Pages_1.Controls.Add(GenNetPort)
		Me._Pages_1.Controls.Add(_GenHdr_3)
		Me._Pages_1.Controls.Add(_GenHdr_2)
		Me._Pages_1.Controls.Add(_GenMiscLbl_0)
		Me._Pages_1.Controls.Add(_GenHdr_1)
		Me._Pages_1.Controls.Add(_GenPassLbl_1)
		Me._Pages_1.Controls.Add(_GenPassLbl_0)
		Me._Pages_1.Controls.Add(_GenHdr_0)
		Me._Pages_1.Controls.Add(_GenNetLbl_0)
		Me._Pages_1.Controls.Add(_GenNetLbl_2)
		Me._Pages_1.Controls.Add(_GenNetLbl_1)
		Me.AgsCmd.SetIndex(_AgsCmd_0, CType(0, Short))
		Me.AgsMnu.SetIndex(_AgsMnu_3, CType(3, Short))
		Me.AgsMnu.SetIndex(_AgsMnu_0, CType(0, Short))
		Me.AgsMnu.SetIndex(_AgsMnu_1, CType(1, Short))
		Me.AgsMnu.SetIndex(_AgsMnu_2, CType(2, Short))
		Me.CptMnu.SetIndex(_CptMnu_0, CType(0, Short))
		Me.CptMnu.SetIndex(_CptMnu_1, CType(1, Short))
		Me.CptMnu.SetIndex(_CptMnu_2, CType(2, Short))
		Me.CptMnu.SetIndex(_CptMnu_3, CType(3, Short))
		Me.CptMnu.SetIndex(_CptMnu_4, CType(4, Short))
		Me.GcnHdr.SetIndex(_GcnHdr_1, CType(1, Short))
		Me.GcnHdr.SetIndex(_GcnHdr_0, CType(0, Short))
		Me.GenHdr.SetIndex(_GenHdr_3, CType(3, Short))
		Me.GenHdr.SetIndex(_GenHdr_2, CType(2, Short))
		Me.GenHdr.SetIndex(_GenHdr_1, CType(1, Short))
		Me.GenHdr.SetIndex(_GenHdr_0, CType(0, Short))
		Me.GenMiscLbl.SetIndex(_GenMiscLbl_0, CType(0, Short))
		Me.GenNetLbl.SetIndex(_GenNetLbl_0, CType(0, Short))
		Me.GenNetLbl.SetIndex(_GenNetLbl_2, CType(2, Short))
		Me.GenNetLbl.SetIndex(_GenNetLbl_1, CType(1, Short))
		Me.GenOpt1.SetIndex(_GenOpt1_0, CType(0, Short))
		Me.GenOpt1.SetIndex(_GenOpt1_1, CType(1, Short))
		Me.GenOpt2.SetIndex(_GenOpt2_0, CType(0, Short))
		Me.GenOpt2.SetIndex(_GenOpt2_1, CType(1, Short))
		Me.GenOpt2.SetIndex(_GenOpt2_2, CType(2, Short))
		Me.GenOpt2.SetIndex(_GenOpt2_3, CType(3, Short))
		Me.GenPassLbl.SetIndex(_GenPassLbl_1, CType(1, Short))
		Me.GenPassLbl.SetIndex(_GenPassLbl_0, CType(0, Short))
		Me.Mnu2Sel.SetIndex(_Mnu2Sel_0, CType(0, Short))
		Me.Mnu2Sel.SetIndex(_Mnu2Sel_1, CType(1, Short))
		Me.Mnu2Sel.SetIndex(_Mnu2Sel_2, CType(2, Short))
		Me.Mnu2Sel.SetIndex(_Mnu2Sel_3, CType(3, Short))
		Me.Mnu2Sel.SetIndex(_Mnu2Sel_4, CType(4, Short))
		Me.Pages.SetIndex(_Pages_0, CType(0, Short))
		Me.Pages.SetIndex(_Pages_1, CType(1, Short))
		Me.Pages.SetIndex(_Pages_2, CType(2, Short))
		Me.Pages.SetIndex(_Pages_4, CType(4, Short))
		Me.Pages.SetIndex(_Pages_3, CType(3, Short))
		CType(Me.Pages, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Mnu2Sel, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.GenPassLbl, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.GenOpt2, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.GenOpt1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.GenNetLbl, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.GenMiscLbl, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.GenHdr, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.GcnHdr, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.CptMnu, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.AgsMnu, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.AgsCmd, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.LstVw1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.Mnu1.Index = 0
		Me.Mnu2.Index = 1
		MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem(){Me.Mnu1, Me.Mnu2})
		Me.Mnu1Rfsh.Index = 0
		Me.Mnu1Close.Index = 1
		Mnu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem(){Me.Mnu1Rfsh, Me.Mnu1Close})
		Me._Mnu2Sel_0.Index = 0
		Me._Mnu2Sel_1.Index = 1
		Me._Mnu2Sel_2.Index = 2
		Me._Mnu2Sel_3.Index = 3
		Me._Mnu2Sel_4.Index = 4
		Mnu2.MenuItems.AddRange(New System.Windows.Forms.MenuItem(){Me._Mnu2Sel_0, Me._Mnu2Sel_1, Me._Mnu2Sel_2, Me._Mnu2Sel_3, Me._Mnu2Sel_4})
		Me.Menu = MainMenu1
	End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As FrmAgnMgr
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As FrmAgnMgr
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New FrmAgnMgr()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	'UPGRADE_WARNING: Lower bound of array sBnrLabel was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1033"'
	Private sBnrLabel(4) As String
	
	
	''*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
	'' FUNCTION
	''
	''
	''*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
	Private Sub FrmAgnMgr_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		sBnrLabel(1) = "General Settings | Main settings for agent."
		sBnrLabel(2) = "Appearance Settings | Set appearance for agent."
		sBnrLabel(3) = "Security Settings 1 | Protect your system."
		sBnrLabel(4) = "Security Settings 2 | More power to protect."
		
		Call LoadAgents()
	End Sub
	
	Private Sub MainPdck_PageFliped(ByVal Sender As System.Object, ByVal e As PageDock.PageFlipedEventArgs) Handles MainPdck.PageFliped
		Dim Flipped As Boolean = e.Flipped
		If Flipped = False Then
			MainBnrLbl.Text = "Agent Configuration"
			MainBnrCap.Text = "General Settings | Main settings for agent"
			Pages(1).BringToFront()
		Else
			MainBnrLbl.Text = "General Control"
			MainBnrCap.Text = "Send general command to agent."
			Pages(0).BringToFront()
		End If
	End Sub
	
	Private Sub MainPhld_PageFlip(ByVal Sender As System.Object, ByVal e As PageHolder.PageFlipEventArgs) Handles MainPhld.PageFlip
		Dim Collapse As Boolean = e.Collapse
		LstVw1.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(MainPhld.Top) - 100)
	End Sub
	
	Private Sub AgsMnu_Click(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles AgsMnu.Click
		Dim Index As Short = AgsMnu.GetIndex(Sender)
		MainBnrCap.Text = sBnrLabel(Index + 1)
		Pages(Index + 1).BringToFront()
	End Sub
	
	Private Sub GcnCmdBtn_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles GcnCmdBtn.Click
		Dim s_Cmd2Send As String
		GcnCmdCB.Text = Trim(GcnCmdCB.Text)
		If GcnCmdCB.Text <> "" Then
			If VB.Left(GcnCmdCB.Text, 2) <> "//" Then
				s_Cmd2Send = GcnCmdCB.Text
			Else
				s_Cmd2Send = "//" & GcnCmdCB.Text
			End If
			Call SendSel(s_Cmd2Send)
		End If
	End Sub
	
	Private Sub GcnCmdCB_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles GcnCmdCB.KeyUp
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		If KeyCode = System.Windows.Forms.Keys.Return Then
			Call GcnCmdBtn_Click(GcnCmdBtn, New System.EventArgs())
		End If
	End Sub
	
	
	
	''*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
	'' MENU
	''
	''
	''*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
	Public Sub Mnu1Close_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Mnu1Close.Popup
		Mnu1Close_Click(eventSender, eventArgs)
	End Sub
	Public Sub Mnu1Close_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Mnu1Close.Click
		Me.Close()
	End Sub
	
	
	
	''*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
	'' FUNCTION
	''
	''
	''*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
	Private Function LoadAgents() As Integer
		Dim a As Short
		Dim mLv As AxMSComctlLib.AxListView
		Dim mItm As MSComctlLib.ListItem
		Dim nItm As MSComctlLib.ListItem
		mLv = FrmMain.DefInstance.Lv1
		
		If mLv.ListItems.Count = 0 Then
			LoadAgents = 1
			Call GcnSmr("No agent loaded !")
			Exit Function
		End If
		
		For a = 1 To mLv.ListItems.Count
			mItm = mLv.ListItems(a)
			nItm = LstVw1.ListItems.Add( , mItm.Text, mItm.Text)
			nItm.SubItems(1) = mItm.SubItems(1)
			nItm.let_Tag(mItm.Tag)
		Next a
		Call GcnSmr(mLv.ListItems.Count & " agent loaded !")
	End Function
	
	Private Function SendSel(ByRef sCommand As Object) As Integer
		Dim a As Short
		Dim l_AgentSel As Integer
		
		If LstVw1.ListItems.Count = 0 Then
			SendSel = 1
			Exit Function
		End If
		
		For a = 1 To LstVw1.ListItems.Count
			If LstVw1.ListItems(a).Selected = True Then
				l_AgentSel = l_AgentSel + 1
				'//Send Command
			End If
		Next a
		
		If l_AgentSel = 0 Then SendSel = 2
	End Function
	
	'UPGRADE_NOTE: Text was upgraded to Text_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
	Private Sub GcnSmr(ByRef Text_Renamed As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object Text_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If Trim(Text_Renamed) = "" Then Exit Sub
		'UPGRADE_WARNING: Couldn't resolve default property of object Text_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		GcnList.Items.Add(">> " & Text_Renamed)
		'UPGRADE_ISSUE: ListBox property GcnList.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2059"'
		GcnList.SelectedIndex = GcnList.NewIndex
	End Sub
End Class