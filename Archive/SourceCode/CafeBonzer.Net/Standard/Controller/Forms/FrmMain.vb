Option Strict Off
Option Explicit On
Friend Class FrmMain
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
		Form_Initialize_renamed()
	End Sub
	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			Static fTerminateCalled As Boolean
			If Not fTerminateCalled Then
				Form_Terminate_renamed()
				fTerminateCalled = True
			End If
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents _MainDbtn_5 As System.Windows.Forms.PictureBox
	Public WithEvents _MainDbtn_4 As System.Windows.Forms.PictureBox
	Public WithEvents _MainDbtn_3 As System.Windows.Forms.PictureBox
	Public WithEvents _MainDbtn_2 As System.Windows.Forms.PictureBox
	Public WithEvents _MainDbtn_1 As System.Windows.Forms.PictureBox
	Public WithEvents _MainDbtn_0 As System.Windows.Forms.PictureBox
	Public WithEvents DockBar As System.Windows.Forms.Panel
	Public WithEvents MainPdock As PageDock
	Public WithEvents ImgList32 As AxMSComctlLib.AxImageList
	Public WithEvents ImgList16 As AxMSComctlLib.AxImageList
	Public WithEvents ImgListSnm As AxMSComctlLib.AxImageList
	Public WithEvents MainSbar As AxMSComctlLib.AxStatusBar
	Public WithEvents uLine3D1 As Line3D
	Public WithEvents _SubPagesMnu_3 As XpButton
	Public WithEvents _SubPagesMnu_2 As XpButton
	Public WithEvents _SubPagesMnu_1 As XpButton
	Public WithEvents _SubPagesMnu_0 As XpButton
	Public WithEvents SerTxtJumlah As System.Windows.Forms.TextBox
	Public WithEvents SerTxtQty As System.Windows.Forms.TextBox
	Public WithEvents SerTxtBaki As System.Windows.Forms.TextBox
	Public WithEvents SerTxtBayar As System.Windows.Forms.TextBox
	Public WithEvents SerScroll1 As System.Windows.Forms.VScrollBar
	Public WithEvents SerTxtTotalItm As System.Windows.Forms.TextBox
	Public WithEvents SerTxtPriItm As System.Windows.Forms.TextBox
	Public WithEvents SerImgCb2 As AxMSComctlLib.AxImageCombo
	Public WithEvents SerImgCb1 As AxMSComctlLib.AxImageCombo
	Public WithEvents SerLv1 As AxMSComctlLib.AxListView
	Public WithEvents SerAddBtn As XpButton
	Public WithEvents _SerBtn_0 As XpButton
	Public WithEvents _SerBtn_1 As XpButton
	Public WithEvents _SerLbl_3 As System.Windows.Forms.Label
	Public WithEvents _SerLbl_2 As System.Windows.Forms.Label
	Public WithEvents _SerLbl_1 As System.Windows.Forms.Label
	Public WithEvents _SerLbl_0 As System.Windows.Forms.Label
	Public WithEvents _SerLbl_4 As System.Windows.Forms.Label
	Public WithEvents _SerLbl_5 As System.Windows.Forms.Label
	Public WithEvents _SerLbl_6 As System.Windows.Forms.Label
	Public WithEvents _SerLbl_7 As System.Windows.Forms.Label
	Public WithEvents _SubPages_1 As System.Windows.Forms.Panel
	Public WithEvents MainLog As System.Windows.Forms.ListBox
	Public WithEvents _SubPages_3 As System.Windows.Forms.Panel
	Public WithEvents MainNote As System.Windows.Forms.TextBox
	Public WithEvents _MainNoteBtn_0 As XpButton
	Public WithEvents _MainNoteBtn_1 As XpButton
	Public WithEvents _SubPages_2 As System.Windows.Forms.Panel
	Public WithEvents uLine3D2 As Line3D
	Public WithEvents _SpgInfoLblD_3 As System.Windows.Forms.Label
	Public WithEvents _SpgInfoLblD_2 As System.Windows.Forms.Label
	Public WithEvents _SpgInfoLblD_1 As System.Windows.Forms.Label
	Public WithEvents _SpgInfoLblD_0 As System.Windows.Forms.Label
	Public WithEvents _SpgInfoLblC_3 As System.Windows.Forms.Label
	Public WithEvents _SpgInfoLblC_2 As System.Windows.Forms.Label
	Public WithEvents _SpgInfoLblC_1 As System.Windows.Forms.Label
	Public WithEvents _SpgInfoLblC_0 As System.Windows.Forms.Label
	Public WithEvents _SpgInfoHdr_1 As System.Windows.Forms.PictureBox
	Public WithEvents _SpgInfoHdr_0 As System.Windows.Forms.PictureBox
	Public WithEvents _SpgInfoLblB_1 As System.Windows.Forms.Label
	Public WithEvents _SpgInfoLblA_1 As System.Windows.Forms.Label
	Public WithEvents _SpgInfoLblB_0 As System.Windows.Forms.Label
	Public WithEvents _SpgInfoLblA_0 As System.Windows.Forms.Label
	Public WithEvents _SubPages_0 As System.Windows.Forms.Panel
	Public WithEvents MainPhold As PageHolder
	Public WithEvents Lv1 As AxMSComctlLib.AxListView
	Public WithEvents MainDbtn As Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray
	Public WithEvents MainNoteBtn As XpButtonArray
	Public WithEvents SerBtn As XpButtonArray
	Public WithEvents SerLbl As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents SpgInfoHdr As Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray
	Public WithEvents SpgInfoLblA As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents SpgInfoLblB As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents SpgInfoLblC As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents SpgInfoLblD As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents SubPages As Microsoft.VisualBasic.Compatibility.VB6.PanelArray
	Public WithEvents SubPagesMnu As XpButtonArray
	Public WithEvents menu3clnsub As Microsoft.VisualBasic.Compatibility.VB6.MenuItemArray
	Public WithEvents menu3ctllock As Microsoft.VisualBasic.Compatibility.VB6.MenuItemArray
	Public WithEvents menu3ctlwinexit As Microsoft.VisualBasic.Compatibility.VB6.MenuItemArray
	Public WithEvents menu4envsub As Microsoft.VisualBasic.Compatibility.VB6.MenuItemArray
	Public WithEvents menu4mon As Microsoft.VisualBasic.Compatibility.VB6.MenuItemArray
	Public WithEvents pmenu1clnsub As Microsoft.VisualBasic.Compatibility.VB6.MenuItemArray
	Public WithEvents pmenu1ctlsub As Microsoft.VisualBasic.Compatibility.VB6.MenuItemArray
	Public WithEvents menu1penalaan As System.Windows.Forms.MenuItem
	Public WithEvents menu1sep1 As System.Windows.Forms.MenuItem
	Public WithEvents menu1logout As System.Windows.Forms.MenuItem
	Public WithEvents menu2keluar As System.Windows.Forms.MenuItem
	Public WithEvents menu1 As System.Windows.Forms.MenuItem
	Public WithEvents menu3mesej As System.Windows.Forms.MenuItem
	Public WithEvents menu3tiker As System.Windows.Forms.MenuItem
	Public WithEvents menu3announce As System.Windows.Forms.MenuItem
	Public WithEvents _menu3ctllock_0 As System.Windows.Forms.MenuItem
	Public WithEvents _menu3ctllock_1 As System.Windows.Forms.MenuItem
	Public WithEvents _menu3ctllock_2 As System.Windows.Forms.MenuItem
	Public WithEvents menu3sep2 As System.Windows.Forms.MenuItem
	Public WithEvents _menu3ctlwinexit_0 As System.Windows.Forms.MenuItem
	Public WithEvents _menu3ctlwinexit_1 As System.Windows.Forms.MenuItem
	Public WithEvents _menu3ctlwinexit_2 As System.Windows.Forms.MenuItem
	Public WithEvents _menu3ctlwinexit_3 As System.Windows.Forms.MenuItem
	Public WithEvents menu3ctl As System.Windows.Forms.MenuItem
	Public WithEvents _menu3clnsub_0 As System.Windows.Forms.MenuItem
	Public WithEvents _menu3clnsub_1 As System.Windows.Forms.MenuItem
	Public WithEvents _menu3clnsub_2 As System.Windows.Forms.MenuItem
	Public WithEvents _menu3clnsub_3 As System.Windows.Forms.MenuItem
	Public WithEvents _menu3clnsub_4 As System.Windows.Forms.MenuItem
	Public WithEvents _menu3clnsub_5 As System.Windows.Forms.MenuItem
	Public WithEvents menu3cln As System.Windows.Forms.MenuItem
	Public WithEvents menu3sep1 As System.Windows.Forms.MenuItem
	Public WithEvents menu3AgMgr As System.Windows.Forms.MenuItem
	Public WithEvents menu3 As System.Windows.Forms.MenuItem
	Public WithEvents _menu4mon_0 As System.Windows.Forms.MenuItem
	Public WithEvents _menu4mon_1 As System.Windows.Forms.MenuItem
	Public WithEvents _menu4mon_2 As System.Windows.Forms.MenuItem
	Public WithEvents _menu4mon_3 As System.Windows.Forms.MenuItem
	Public WithEvents menu4sep1 As System.Windows.Forms.MenuItem
	Public WithEvents _menu4envsub_0 As System.Windows.Forms.MenuItem
	Public WithEvents _menu4envsub_1 As System.Windows.Forms.MenuItem
	Public WithEvents menu4env As System.Windows.Forms.MenuItem
	Public WithEvents menu4 As System.Windows.Forms.MenuItem
	Public WithEvents menu4PosMgr As System.Windows.Forms.MenuItem
	Public WithEvents menu4Stat As System.Windows.Forms.MenuItem
	Public WithEvents menu5Console As System.Windows.Forms.MenuItem
	Public WithEvents menu5 As System.Windows.Forms.MenuItem
	Public WithEvents menu2bantuan As System.Windows.Forms.MenuItem
	Public WithEvents menu2sep1 As System.Windows.Forms.MenuItem
	Public WithEvents menu2aplikasi As System.Windows.Forms.MenuItem
	Public WithEvents menu2 As System.Windows.Forms.MenuItem
	Public WithEvents pmenu1flog As System.Windows.Forms.MenuItem
	Public WithEvents pmenu1flout As System.Windows.Forms.MenuItem
	Public WithEvents psep2 As System.Windows.Forms.MenuItem
	Public WithEvents pmenu1cancel As System.Windows.Forms.MenuItem
	Public WithEvents pmenu1trans As System.Windows.Forms.MenuItem
	Public WithEvents pmenu1terminal As System.Windows.Forms.MenuItem
	Public WithEvents psep1 As System.Windows.Forms.MenuItem
	Public WithEvents _pmenu1clnsub_0 As System.Windows.Forms.MenuItem
	Public WithEvents _pmenu1clnsub_1 As System.Windows.Forms.MenuItem
	Public WithEvents _pmenu1clnsub_2 As System.Windows.Forms.MenuItem
	Public WithEvents _pmenu1clnsub_3 As System.Windows.Forms.MenuItem
	Public WithEvents _pmenu1clnsub_4 As System.Windows.Forms.MenuItem
	Public WithEvents _pmenu1clnsub_5 As System.Windows.Forms.MenuItem
	Public WithEvents pmenu1cln As System.Windows.Forms.MenuItem
	Public WithEvents _pmenu1ctlsub_0 As System.Windows.Forms.MenuItem
	Public WithEvents _pmenu1ctlsub_1 As System.Windows.Forms.MenuItem
	Public WithEvents _pmenu1ctlsub_2 As System.Windows.Forms.MenuItem
	Public WithEvents _pmenu1ctlsub_3 As System.Windows.Forms.MenuItem
	Public WithEvents pmenu1ctl As System.Windows.Forms.MenuItem
	Public WithEvents popmenu1 As System.Windows.Forms.MenuItem
	Public MainMenu1 As System.Windows.Forms.MainMenu
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmMain))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.MainPdock = New PageDock
		Me.DockBar = New System.Windows.Forms.Panel
		Me._MainDbtn_5 = New System.Windows.Forms.PictureBox
		Me._MainDbtn_4 = New System.Windows.Forms.PictureBox
		Me._MainDbtn_3 = New System.Windows.Forms.PictureBox
		Me._MainDbtn_2 = New System.Windows.Forms.PictureBox
		Me._MainDbtn_1 = New System.Windows.Forms.PictureBox
		Me._MainDbtn_0 = New System.Windows.Forms.PictureBox
		Me.ImgList32 = New AxMSComctlLib.AxImageList
		Me.ImgList16 = New AxMSComctlLib.AxImageList
		Me.ImgListSnm = New AxMSComctlLib.AxImageList
		Me.MainSbar = New AxMSComctlLib.AxStatusBar
		Me.MainPhold = New PageHolder
		Me.uLine3D1 = New Line3D
		Me._SubPagesMnu_3 = New XpButton
		Me._SubPagesMnu_2 = New XpButton
		Me._SubPagesMnu_1 = New XpButton
		Me._SubPagesMnu_0 = New XpButton
		Me._SubPages_1 = New System.Windows.Forms.Panel
		Me.SerTxtJumlah = New System.Windows.Forms.TextBox
		Me.SerTxtQty = New System.Windows.Forms.TextBox
		Me.SerTxtBaki = New System.Windows.Forms.TextBox
		Me.SerTxtBayar = New System.Windows.Forms.TextBox
		Me.SerScroll1 = New System.Windows.Forms.VScrollBar
		Me.SerTxtTotalItm = New System.Windows.Forms.TextBox
		Me.SerTxtPriItm = New System.Windows.Forms.TextBox
		Me.SerImgCb2 = New AxMSComctlLib.AxImageCombo
		Me.SerImgCb1 = New AxMSComctlLib.AxImageCombo
		Me.SerLv1 = New AxMSComctlLib.AxListView
		Me.SerAddBtn = New XpButton
		Me._SerBtn_0 = New XpButton
		Me._SerBtn_1 = New XpButton
		Me._SerLbl_3 = New System.Windows.Forms.Label
		Me._SerLbl_2 = New System.Windows.Forms.Label
		Me._SerLbl_1 = New System.Windows.Forms.Label
		Me._SerLbl_0 = New System.Windows.Forms.Label
		Me._SerLbl_4 = New System.Windows.Forms.Label
		Me._SerLbl_5 = New System.Windows.Forms.Label
		Me._SerLbl_6 = New System.Windows.Forms.Label
		Me._SerLbl_7 = New System.Windows.Forms.Label
		Me._SubPages_3 = New System.Windows.Forms.Panel
		Me.MainLog = New System.Windows.Forms.ListBox
		Me._SubPages_2 = New System.Windows.Forms.Panel
		Me.MainNote = New System.Windows.Forms.TextBox
		Me._MainNoteBtn_0 = New XpButton
		Me._MainNoteBtn_1 = New XpButton
		Me._SubPages_0 = New System.Windows.Forms.Panel
		Me.uLine3D2 = New Line3D
		Me._SpgInfoLblD_3 = New System.Windows.Forms.Label
		Me._SpgInfoLblD_2 = New System.Windows.Forms.Label
		Me._SpgInfoLblD_1 = New System.Windows.Forms.Label
		Me._SpgInfoLblD_0 = New System.Windows.Forms.Label
		Me._SpgInfoLblC_3 = New System.Windows.Forms.Label
		Me._SpgInfoLblC_2 = New System.Windows.Forms.Label
		Me._SpgInfoLblC_1 = New System.Windows.Forms.Label
		Me._SpgInfoLblC_0 = New System.Windows.Forms.Label
		Me._SpgInfoHdr_1 = New System.Windows.Forms.PictureBox
		Me._SpgInfoHdr_0 = New System.Windows.Forms.PictureBox
		Me._SpgInfoLblB_1 = New System.Windows.Forms.Label
		Me._SpgInfoLblA_1 = New System.Windows.Forms.Label
		Me._SpgInfoLblB_0 = New System.Windows.Forms.Label
		Me._SpgInfoLblA_0 = New System.Windows.Forms.Label
		Me.Lv1 = New AxMSComctlLib.AxListView
		Me.MainDbtn = New Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray(components)
		Me.MainNoteBtn = New XpButtonArray(components)
		Me.SerBtn = New XpButtonArray(components)
		Me.SerLbl = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.SpgInfoHdr = New Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray(components)
		Me.SpgInfoLblA = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.SpgInfoLblB = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.SpgInfoLblC = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.SpgInfoLblD = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.SubPages = New Microsoft.VisualBasic.Compatibility.VB6.PanelArray(components)
		Me.SubPagesMnu = New XpButtonArray(components)
		Me.menu3clnsub = New Microsoft.VisualBasic.Compatibility.VB6.MenuItemArray(components)
		Me.menu3ctllock = New Microsoft.VisualBasic.Compatibility.VB6.MenuItemArray(components)
		Me.menu3ctlwinexit = New Microsoft.VisualBasic.Compatibility.VB6.MenuItemArray(components)
		Me.menu4envsub = New Microsoft.VisualBasic.Compatibility.VB6.MenuItemArray(components)
		Me.menu4mon = New Microsoft.VisualBasic.Compatibility.VB6.MenuItemArray(components)
		Me.pmenu1clnsub = New Microsoft.VisualBasic.Compatibility.VB6.MenuItemArray(components)
		Me.pmenu1ctlsub = New Microsoft.VisualBasic.Compatibility.VB6.MenuItemArray(components)
		Me.MainMenu1 = New System.Windows.Forms.MainMenu
		Me.menu1 = New System.Windows.Forms.MenuItem
		Me.menu1penalaan = New System.Windows.Forms.MenuItem
		Me.menu1sep1 = New System.Windows.Forms.MenuItem
		Me.menu1logout = New System.Windows.Forms.MenuItem
		Me.menu2keluar = New System.Windows.Forms.MenuItem
		Me.menu3 = New System.Windows.Forms.MenuItem
		Me.menu3announce = New System.Windows.Forms.MenuItem
		Me.menu3mesej = New System.Windows.Forms.MenuItem
		Me.menu3tiker = New System.Windows.Forms.MenuItem
		Me.menu3ctl = New System.Windows.Forms.MenuItem
		Me._menu3ctllock_0 = New System.Windows.Forms.MenuItem
		Me._menu3ctllock_1 = New System.Windows.Forms.MenuItem
		Me._menu3ctllock_2 = New System.Windows.Forms.MenuItem
		Me.menu3sep2 = New System.Windows.Forms.MenuItem
		Me._menu3ctlwinexit_0 = New System.Windows.Forms.MenuItem
		Me._menu3ctlwinexit_1 = New System.Windows.Forms.MenuItem
		Me._menu3ctlwinexit_2 = New System.Windows.Forms.MenuItem
		Me._menu3ctlwinexit_3 = New System.Windows.Forms.MenuItem
		Me.menu3cln = New System.Windows.Forms.MenuItem
		Me._menu3clnsub_0 = New System.Windows.Forms.MenuItem
		Me._menu3clnsub_1 = New System.Windows.Forms.MenuItem
		Me._menu3clnsub_2 = New System.Windows.Forms.MenuItem
		Me._menu3clnsub_3 = New System.Windows.Forms.MenuItem
		Me._menu3clnsub_4 = New System.Windows.Forms.MenuItem
		Me._menu3clnsub_5 = New System.Windows.Forms.MenuItem
		Me.menu3sep1 = New System.Windows.Forms.MenuItem
		Me.menu3AgMgr = New System.Windows.Forms.MenuItem
		Me.menu4 = New System.Windows.Forms.MenuItem
		Me._menu4mon_0 = New System.Windows.Forms.MenuItem
		Me._menu4mon_1 = New System.Windows.Forms.MenuItem
		Me._menu4mon_2 = New System.Windows.Forms.MenuItem
		Me._menu4mon_3 = New System.Windows.Forms.MenuItem
		Me.menu4sep1 = New System.Windows.Forms.MenuItem
		Me.menu4env = New System.Windows.Forms.MenuItem
		Me._menu4envsub_0 = New System.Windows.Forms.MenuItem
		Me._menu4envsub_1 = New System.Windows.Forms.MenuItem
		Me.menu5 = New System.Windows.Forms.MenuItem
		Me.menu4PosMgr = New System.Windows.Forms.MenuItem
		Me.menu4Stat = New System.Windows.Forms.MenuItem
		Me.menu5Console = New System.Windows.Forms.MenuItem
		Me.menu2 = New System.Windows.Forms.MenuItem
		Me.menu2bantuan = New System.Windows.Forms.MenuItem
		Me.menu2sep1 = New System.Windows.Forms.MenuItem
		Me.menu2aplikasi = New System.Windows.Forms.MenuItem
		Me.popmenu1 = New System.Windows.Forms.MenuItem
		Me.pmenu1flog = New System.Windows.Forms.MenuItem
		Me.pmenu1flout = New System.Windows.Forms.MenuItem
		Me.psep2 = New System.Windows.Forms.MenuItem
		Me.pmenu1cancel = New System.Windows.Forms.MenuItem
		Me.pmenu1trans = New System.Windows.Forms.MenuItem
		Me.pmenu1terminal = New System.Windows.Forms.MenuItem
		Me.psep1 = New System.Windows.Forms.MenuItem
		Me.pmenu1cln = New System.Windows.Forms.MenuItem
		Me._pmenu1clnsub_0 = New System.Windows.Forms.MenuItem
		Me._pmenu1clnsub_1 = New System.Windows.Forms.MenuItem
		Me._pmenu1clnsub_2 = New System.Windows.Forms.MenuItem
		Me._pmenu1clnsub_3 = New System.Windows.Forms.MenuItem
		Me._pmenu1clnsub_4 = New System.Windows.Forms.MenuItem
		Me._pmenu1clnsub_5 = New System.Windows.Forms.MenuItem
		Me.pmenu1ctl = New System.Windows.Forms.MenuItem
		Me._pmenu1ctlsub_0 = New System.Windows.Forms.MenuItem
		Me._pmenu1ctlsub_1 = New System.Windows.Forms.MenuItem
		Me._pmenu1ctlsub_2 = New System.Windows.Forms.MenuItem
		Me._pmenu1ctlsub_3 = New System.Windows.Forms.MenuItem
		CType(Me.ImgList32, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.ImgList16, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.ImgListSnm, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.MainSbar, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.SerImgCb2, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.SerImgCb1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.SerLv1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Lv1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.MainDbtn, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.MainNoteBtn, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.SerBtn, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.SerLbl, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.SpgInfoHdr, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.SpgInfoLblA, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.SpgInfoLblB, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.SpgInfoLblC, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.SpgInfoLblD, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.SubPages, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.SubPagesMnu, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.menu3clnsub, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.menu3ctllock, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.menu3ctlwinexit, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.menu4envsub, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.menu4mon, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.pmenu1clnsub, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.pmenu1ctlsub, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me.Text = "CafeBonzer v2.0"
		Me.ClientSize = New System.Drawing.Size(756, 530)
		Me.Location = New System.Drawing.Point(11, 30)
		Me.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Icon = CType(resources.GetObject("FrmMain.Icon"), System.Drawing.Icon)
		Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
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
		Me.Name = "FrmMain"
		Me.MainPdock.Dock = System.Windows.Forms.DockStyle.Right
		Me.MainPdock.Size = New System.Drawing.Size(78, 505)
		Me.MainPdock.Location = New System.Drawing.Point(678, 0)
		Me.MainPdock.TabIndex = 0
		Me.MainPdock.HldrBtnPos = 1
		Me.MainPdock.HldrLne = -1
		Me.MainPdock.PageState = 0
		Me.MainPdock.PageWidth = 1170
		Me.MainPdock.Name = "MainPdock"
		Me.DockBar.BackColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me.DockBar.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.DockBar.Size = New System.Drawing.Size(53, 504)
		Me.DockBar.Location = New System.Drawing.Point(25, 0)
		Me.DockBar.TabIndex = 1
		Me.DockBar.Tag = "subcontainer"
		Me.DockBar.Dock = System.Windows.Forms.DockStyle.None
		Me.DockBar.CausesValidation = True
		Me.DockBar.Enabled = True
		Me.DockBar.ForeColor = System.Drawing.SystemColors.ControlText
		Me.DockBar.Cursor = System.Windows.Forms.Cursors.Default
		Me.DockBar.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.DockBar.TabStop = True
		Me.DockBar.Visible = True
		Me.DockBar.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.DockBar.Name = "DockBar"
		Me._MainDbtn_5.Size = New System.Drawing.Size(32, 32)
		Me._MainDbtn_5.Location = New System.Drawing.Point(6, 167)
		Me._MainDbtn_5.Image = CType(resources.GetObject("_MainDbtn_5.Image"), System.Drawing.Image)
		Me._MainDbtn_5.Enabled = True
		Me._MainDbtn_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._MainDbtn_5.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._MainDbtn_5.Visible = True
		Me._MainDbtn_5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._MainDbtn_5.Name = "_MainDbtn_5"
		Me._MainDbtn_4.Size = New System.Drawing.Size(32, 32)
		Me._MainDbtn_4.Location = New System.Drawing.Point(6, 135)
		Me._MainDbtn_4.Image = CType(resources.GetObject("_MainDbtn_4.Image"), System.Drawing.Image)
		Me._MainDbtn_4.Enabled = True
		Me._MainDbtn_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._MainDbtn_4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._MainDbtn_4.Visible = True
		Me._MainDbtn_4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._MainDbtn_4.Name = "_MainDbtn_4"
		Me._MainDbtn_3.Size = New System.Drawing.Size(32, 32)
		Me._MainDbtn_3.Location = New System.Drawing.Point(6, 102)
		Me._MainDbtn_3.Image = CType(resources.GetObject("_MainDbtn_3.Image"), System.Drawing.Image)
		Me._MainDbtn_3.Enabled = True
		Me._MainDbtn_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._MainDbtn_3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._MainDbtn_3.Visible = True
		Me._MainDbtn_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._MainDbtn_3.Name = "_MainDbtn_3"
		Me._MainDbtn_2.Size = New System.Drawing.Size(32, 32)
		Me._MainDbtn_2.Location = New System.Drawing.Point(6, 70)
		Me._MainDbtn_2.Image = CType(resources.GetObject("_MainDbtn_2.Image"), System.Drawing.Image)
		Me._MainDbtn_2.Enabled = True
		Me._MainDbtn_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._MainDbtn_2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._MainDbtn_2.Visible = True
		Me._MainDbtn_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._MainDbtn_2.Name = "_MainDbtn_2"
		Me._MainDbtn_1.Size = New System.Drawing.Size(32, 32)
		Me._MainDbtn_1.Location = New System.Drawing.Point(6, 37)
		Me._MainDbtn_1.Image = CType(resources.GetObject("_MainDbtn_1.Image"), System.Drawing.Image)
		Me._MainDbtn_1.Enabled = True
		Me._MainDbtn_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._MainDbtn_1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._MainDbtn_1.Visible = True
		Me._MainDbtn_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._MainDbtn_1.Name = "_MainDbtn_1"
		Me._MainDbtn_0.Size = New System.Drawing.Size(32, 32)
		Me._MainDbtn_0.Location = New System.Drawing.Point(6, 4)
		Me._MainDbtn_0.Image = CType(resources.GetObject("_MainDbtn_0.Image"), System.Drawing.Image)
		Me._MainDbtn_0.Enabled = True
		Me._MainDbtn_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._MainDbtn_0.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._MainDbtn_0.Visible = True
		Me._MainDbtn_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._MainDbtn_0.Name = "_MainDbtn_0"
		ImgList32.OcxState = CType(resources.GetObject("ImgList32.OcxState"), System.Windows.Forms.AxHost.State)
		Me.ImgList32.Location = New System.Drawing.Point(637, 191)
		Me.ImgList32.Name = "ImgList32"
		ImgList16.OcxState = CType(resources.GetObject("ImgList16.OcxState"), System.Windows.Forms.AxHost.State)
		Me.ImgList16.Location = New System.Drawing.Point(637, 230)
		Me.ImgList16.Name = "ImgList16"
		ImgListSnm.OcxState = CType(resources.GetObject("ImgListSnm.OcxState"), System.Windows.Forms.AxHost.State)
		Me.ImgListSnm.Location = New System.Drawing.Point(637, 267)
		Me.ImgListSnm.Name = "ImgListSnm"
		MainSbar.OcxState = CType(resources.GetObject("MainSbar.OcxState"), System.Windows.Forms.AxHost.State)
		Me.MainSbar.Dock = System.Windows.Forms.DockStyle.Bottom
		Me.MainSbar.Size = New System.Drawing.Size(756, 25)
		Me.MainSbar.Location = New System.Drawing.Point(0, 505)
		Me.MainSbar.TabIndex = 50
		Me.MainSbar.Name = "MainSbar"
		Me.MainPhold.Size = New System.Drawing.Size(677, 198)
		Me.MainPhold.Location = New System.Drawing.Point(0, 308)
		Me.MainPhold.TabIndex = 2
		Me.MainPhold.HldrTxt = "Toolbox"
		Me.MainPhold.HldrTxtClr = 16777215
		Me.MainPhold.HldrLne = -1
		Me.MainPhold.PageHeight = 2970
		Me.MainPhold.Name = "MainPhold"
		Me.uLine3D1.Size = New System.Drawing.Size(3, 173)
		Me.uLine3D1.Location = New System.Drawing.Point(31, 23)
		Me.uLine3D1.TabIndex = 7
		Me.uLine3D1.horizon = 0
		Me.uLine3D1.Name = "uLine3D1"
		Me._SubPagesMnu_3.Size = New System.Drawing.Size(27, 27)
		Me._SubPagesMnu_3.Location = New System.Drawing.Point(2, 109)
		Me._SubPagesMnu_3.TabIndex = 6
		Me.ToolTip1.SetToolTip(Me._SubPagesMnu_3, "Log")
		Me._SubPagesMnu_3.TX = ""
		Me._SubPagesMnu_3.ENAB = -1
		Me._SubPagesMnu_3.COLTYPE = 1
		Me._SubPagesMnu_3.FOCUSR = -1
		Me._SubPagesMnu_3.BCOL = 12632256
		Me._SubPagesMnu_3.BCOLO = 12632256
		Me._SubPagesMnu_3.FCOL = 0
		Me._SubPagesMnu_3.FCOLO = 0
		Me._SubPagesMnu_3.MCOL = 16777215
		Me._SubPagesMnu_3.MPTR = 1
		Me._SubPagesMnu_3.MICON = 0
		Me._SubPagesMnu_3.PICN = 0
		Me._SubPagesMnu_3.UMCOL = -1
		Me._SubPagesMnu_3.SOFT = 0
		Me._SubPagesMnu_3.PICPOS = 0
		Me._SubPagesMnu_3.NGREY = 0
		Me._SubPagesMnu_3.FX = 0
		Me._SubPagesMnu_3.HAND = 0
		Me._SubPagesMnu_3.CHECK = 0
		Me._SubPagesMnu_3.Name = "_SubPagesMnu_3"
		Me._SubPagesMnu_2.Size = New System.Drawing.Size(27, 27)
		Me._SubPagesMnu_2.Location = New System.Drawing.Point(2, 81)
		Me._SubPagesMnu_2.TabIndex = 5
		Me.ToolTip1.SetToolTip(Me._SubPagesMnu_2, "Note")
		Me._SubPagesMnu_2.TX = ""
		Me._SubPagesMnu_2.ENAB = -1
		Me._SubPagesMnu_2.COLTYPE = 1
		Me._SubPagesMnu_2.FOCUSR = -1
		Me._SubPagesMnu_2.BCOL = 12632256
		Me._SubPagesMnu_2.BCOLO = 12632256
		Me._SubPagesMnu_2.FCOL = 0
		Me._SubPagesMnu_2.FCOLO = 0
		Me._SubPagesMnu_2.MCOL = 16777215
		Me._SubPagesMnu_2.MPTR = 1
		Me._SubPagesMnu_2.MICON = 0
		Me._SubPagesMnu_2.PICN = 0
		Me._SubPagesMnu_2.UMCOL = -1
		Me._SubPagesMnu_2.SOFT = 0
		Me._SubPagesMnu_2.PICPOS = 0
		Me._SubPagesMnu_2.NGREY = 0
		Me._SubPagesMnu_2.FX = 0
		Me._SubPagesMnu_2.HAND = 0
		Me._SubPagesMnu_2.CHECK = 0
		Me._SubPagesMnu_2.Name = "_SubPagesMnu_2"
		Me._SubPagesMnu_1.Size = New System.Drawing.Size(27, 27)
		Me._SubPagesMnu_1.Location = New System.Drawing.Point(2, 53)
		Me._SubPagesMnu_1.TabIndex = 4
		Me.ToolTip1.SetToolTip(Me._SubPagesMnu_1, "Service & Merchandise")
		Me._SubPagesMnu_1.TX = ""
		Me._SubPagesMnu_1.ENAB = -1
		Me._SubPagesMnu_1.COLTYPE = 1
		Me._SubPagesMnu_1.FOCUSR = -1
		Me._SubPagesMnu_1.BCOL = 12632256
		Me._SubPagesMnu_1.BCOLO = 12632256
		Me._SubPagesMnu_1.FCOL = 0
		Me._SubPagesMnu_1.FCOLO = 0
		Me._SubPagesMnu_1.MCOL = 16777215
		Me._SubPagesMnu_1.MPTR = 1
		Me._SubPagesMnu_1.MICON = 0
		Me._SubPagesMnu_1.PICN = 0
		Me._SubPagesMnu_1.UMCOL = -1
		Me._SubPagesMnu_1.SOFT = 0
		Me._SubPagesMnu_1.PICPOS = 0
		Me._SubPagesMnu_1.NGREY = 0
		Me._SubPagesMnu_1.FX = 0
		Me._SubPagesMnu_1.HAND = 0
		Me._SubPagesMnu_1.CHECK = 0
		Me._SubPagesMnu_1.Name = "_SubPagesMnu_1"
		Me._SubPagesMnu_0.Size = New System.Drawing.Size(27, 27)
		Me._SubPagesMnu_0.Location = New System.Drawing.Point(2, 25)
		Me._SubPagesMnu_0.TabIndex = 3
		Me.ToolTip1.SetToolTip(Me._SubPagesMnu_0, "Information")
		Me._SubPagesMnu_0.TX = ""
		Me._SubPagesMnu_0.ENAB = -1
		Me._SubPagesMnu_0.COLTYPE = 1
		Me._SubPagesMnu_0.FOCUSR = -1
		Me._SubPagesMnu_0.BCOL = 12632256
		Me._SubPagesMnu_0.BCOLO = 12632256
		Me._SubPagesMnu_0.FCOL = 0
		Me._SubPagesMnu_0.FCOLO = 0
		Me._SubPagesMnu_0.MCOL = 16777215
		Me._SubPagesMnu_0.MPTR = 1
		Me._SubPagesMnu_0.MICON = 0
		Me._SubPagesMnu_0.PICN = 0
		Me._SubPagesMnu_0.UMCOL = -1
		Me._SubPagesMnu_0.SOFT = 0
		Me._SubPagesMnu_0.PICPOS = 0
		Me._SubPagesMnu_0.NGREY = 0
		Me._SubPagesMnu_0.FX = 0
		Me._SubPagesMnu_0.HAND = 0
		Me._SubPagesMnu_0.CHECK = 0
		Me._SubPagesMnu_0.Name = "_SubPagesMnu_0"
		Me._SubPages_1.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me._SubPages_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._SubPages_1.Size = New System.Drawing.Size(644, 171)
		Me._SubPages_1.Location = New System.Drawing.Point(38, 24)
		Me._SubPages_1.TabIndex = 22
		Me._SubPages_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._SubPages_1.Dock = System.Windows.Forms.DockStyle.None
		Me._SubPages_1.CausesValidation = True
		Me._SubPages_1.Enabled = True
		Me._SubPages_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._SubPages_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._SubPages_1.TabStop = True
		Me._SubPages_1.Visible = True
		Me._SubPages_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._SubPages_1.Name = "_SubPages_1"
		Me.SerTxtJumlah.AutoSize = False
		Me.SerTxtJumlah.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.SerTxtJumlah.BackColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me.SerTxtJumlah.Font = New System.Drawing.Font("Endless Showroom", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.SerTxtJumlah.ForeColor = System.Drawing.Color.Green
		Me.SerTxtJumlah.Size = New System.Drawing.Size(91, 25)
		Me.SerTxtJumlah.Location = New System.Drawing.Point(531, 8)
		Me.SerTxtJumlah.ReadOnly = True
		Me.SerTxtJumlah.TabIndex = 38
		Me.SerTxtJumlah.TabStop = False
		Me.SerTxtJumlah.AcceptsReturn = True
		Me.SerTxtJumlah.CausesValidation = True
		Me.SerTxtJumlah.Enabled = True
		Me.SerTxtJumlah.HideSelection = True
		Me.SerTxtJumlah.Maxlength = 0
		Me.SerTxtJumlah.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.SerTxtJumlah.MultiLine = False
		Me.SerTxtJumlah.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.SerTxtJumlah.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.SerTxtJumlah.Visible = True
		Me.SerTxtJumlah.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.SerTxtJumlah.Name = "SerTxtJumlah"
		Me.SerTxtQty.AutoSize = False
		Me.SerTxtQty.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.SerTxtQty.BackColor = System.Drawing.Color.White
		Me.SerTxtQty.Enabled = False
		Me.SerTxtQty.Size = New System.Drawing.Size(36, 21)
		Me.SerTxtQty.Location = New System.Drawing.Point(85, 75)
		Me.SerTxtQty.TabIndex = 28
		Me.SerTxtQty.Text = "1"
		Me.SerTxtQty.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.SerTxtQty.AcceptsReturn = True
		Me.SerTxtQty.CausesValidation = True
		Me.SerTxtQty.ForeColor = System.Drawing.SystemColors.WindowText
		Me.SerTxtQty.HideSelection = True
		Me.SerTxtQty.ReadOnly = False
		Me.SerTxtQty.Maxlength = 0
		Me.SerTxtQty.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.SerTxtQty.MultiLine = False
		Me.SerTxtQty.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.SerTxtQty.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.SerTxtQty.TabStop = True
		Me.SerTxtQty.Visible = True
		Me.SerTxtQty.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.SerTxtQty.Name = "SerTxtQty"
		Me.SerTxtBaki.AutoSize = False
		Me.SerTxtBaki.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.SerTxtBaki.BackColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me.SerTxtBaki.Font = New System.Drawing.Font("Endless Showroom", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.SerTxtBaki.ForeColor = System.Drawing.Color.Green
		Me.SerTxtBaki.Size = New System.Drawing.Size(91, 25)
		Me.SerTxtBaki.Location = New System.Drawing.Point(531, 44)
		Me.SerTxtBaki.ReadOnly = True
		Me.SerTxtBaki.TabIndex = 40
		Me.SerTxtBaki.TabStop = False
		Me.SerTxtBaki.AcceptsReturn = True
		Me.SerTxtBaki.CausesValidation = True
		Me.SerTxtBaki.Enabled = True
		Me.SerTxtBaki.HideSelection = True
		Me.SerTxtBaki.Maxlength = 0
		Me.SerTxtBaki.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.SerTxtBaki.MultiLine = False
		Me.SerTxtBaki.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.SerTxtBaki.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.SerTxtBaki.Visible = True
		Me.SerTxtBaki.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.SerTxtBaki.Name = "SerTxtBaki"
		Me.SerTxtBayar.AutoSize = False
		Me.SerTxtBayar.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.SerTxtBayar.BackColor = System.Drawing.Color.White
		Me.SerTxtBayar.Font = New System.Drawing.Font("Endless Showroom", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.SerTxtBayar.ForeColor = System.Drawing.Color.Green
		Me.SerTxtBayar.Size = New System.Drawing.Size(91, 25)
		Me.SerTxtBayar.Location = New System.Drawing.Point(531, 81)
		Me.SerTxtBayar.TabIndex = 42
		Me.SerTxtBayar.TabStop = False
		Me.SerTxtBayar.AcceptsReturn = True
		Me.SerTxtBayar.CausesValidation = True
		Me.SerTxtBayar.Enabled = True
		Me.SerTxtBayar.HideSelection = True
		Me.SerTxtBayar.ReadOnly = False
		Me.SerTxtBayar.Maxlength = 0
		Me.SerTxtBayar.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.SerTxtBayar.MultiLine = False
		Me.SerTxtBayar.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.SerTxtBayar.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.SerTxtBayar.Visible = True
		Me.SerTxtBayar.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.SerTxtBayar.Name = "SerTxtBayar"
		Me.SerScroll1.Size = New System.Drawing.Size(11, 22)
		Me.SerScroll1.Location = New System.Drawing.Point(127, 75)
		Me.SerScroll1.Maximum = 999
		Me.SerScroll1.Minimum = 1
		Me.SerScroll1.TabIndex = 29
		Me.SerScroll1.Value = 999
		Me.SerScroll1.CausesValidation = True
		Me.SerScroll1.Enabled = True
		Me.SerScroll1.LargeChange = 1
		Me.SerScroll1.Cursor = System.Windows.Forms.Cursors.Default
		Me.SerScroll1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.SerScroll1.SmallChange = 1
		Me.SerScroll1.TabStop = True
		Me.SerScroll1.Visible = True
		Me.SerScroll1.Name = "SerScroll1"
		Me.SerTxtTotalItm.AutoSize = False
		Me.SerTxtTotalItm.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.SerTxtTotalItm.BackColor = System.Drawing.Color.FromARGB(192, 255, 192)
		Me.SerTxtTotalItm.Size = New System.Drawing.Size(116, 21)
		Me.SerTxtTotalItm.Location = New System.Drawing.Point(85, 143)
		Me.SerTxtTotalItm.TabIndex = 35
		Me.SerTxtTotalItm.TabStop = False
		Me.SerTxtTotalItm.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.SerTxtTotalItm.AcceptsReturn = True
		Me.SerTxtTotalItm.CausesValidation = True
		Me.SerTxtTotalItm.Enabled = True
		Me.SerTxtTotalItm.ForeColor = System.Drawing.SystemColors.WindowText
		Me.SerTxtTotalItm.HideSelection = True
		Me.SerTxtTotalItm.ReadOnly = False
		Me.SerTxtTotalItm.Maxlength = 0
		Me.SerTxtTotalItm.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.SerTxtTotalItm.MultiLine = False
		Me.SerTxtTotalItm.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.SerTxtTotalItm.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.SerTxtTotalItm.Visible = True
		Me.SerTxtTotalItm.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.SerTxtTotalItm.Name = "SerTxtTotalItm"
		Me.SerTxtPriItm.AutoSize = False
		Me.SerTxtPriItm.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.SerTxtPriItm.BackColor = System.Drawing.Color.White
		Me.SerTxtPriItm.Size = New System.Drawing.Size(116, 21)
		Me.SerTxtPriItm.Location = New System.Drawing.Point(85, 109)
		Me.SerTxtPriItm.TabIndex = 33
		Me.SerTxtPriItm.TabStop = False
		Me.SerTxtPriItm.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.SerTxtPriItm.AcceptsReturn = True
		Me.SerTxtPriItm.CausesValidation = True
		Me.SerTxtPriItm.Enabled = True
		Me.SerTxtPriItm.ForeColor = System.Drawing.SystemColors.WindowText
		Me.SerTxtPriItm.HideSelection = True
		Me.SerTxtPriItm.ReadOnly = False
		Me.SerTxtPriItm.Maxlength = 0
		Me.SerTxtPriItm.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.SerTxtPriItm.MultiLine = False
		Me.SerTxtPriItm.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.SerTxtPriItm.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.SerTxtPriItm.Visible = True
		Me.SerTxtPriItm.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.SerTxtPriItm.Name = "SerTxtPriItm"
		SerImgCb2.OcxState = CType(resources.GetObject("SerImgCb2.OcxState"), System.Windows.Forms.AxHost.State)
		Me.SerImgCb2.Size = New System.Drawing.Size(116, 22)
		Me.SerImgCb2.Location = New System.Drawing.Point(85, 40)
		Me.SerImgCb2.TabIndex = 26
		Me.SerImgCb2.Name = "SerImgCb2"
		SerImgCb1.OcxState = CType(resources.GetObject("SerImgCb1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.SerImgCb1.Size = New System.Drawing.Size(116, 22)
		Me.SerImgCb1.Location = New System.Drawing.Point(85, 5)
		Me.SerImgCb1.TabIndex = 24
		Me.SerImgCb1.Name = "SerImgCb1"
		SerLv1.OcxState = CType(resources.GetObject("SerLv1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.SerLv1.Size = New System.Drawing.Size(218, 161)
		Me.SerLv1.Location = New System.Drawing.Point(210, 4)
		Me.SerLv1.TabIndex = 36
		Me.SerLv1.Name = "SerLv1"
		Me.SerAddBtn.Size = New System.Drawing.Size(42, 40)
		Me.SerAddBtn.Location = New System.Drawing.Point(583, 128)
		Me.SerAddBtn.TabIndex = 43
		Me.ToolTip1.SetToolTip(Me.SerAddBtn, "Confirm Transaction")
		Me.SerAddBtn.TX = ""
		Me.SerAddBtn.ENAB = -1
		Me.SerAddBtn.COLTYPE = 1
		Me.SerAddBtn.FOCUSR = -1
		Me.SerAddBtn.BCOL = 12632256
		Me.SerAddBtn.BCOLO = 12632256
		Me.SerAddBtn.FCOL = 0
		Me.SerAddBtn.FCOLO = 0
		Me.SerAddBtn.MCOL = 16777215
		Me.SerAddBtn.MPTR = 1
		Me.SerAddBtn.MICON = 0
		Me.SerAddBtn.PICN = 0
		Me.SerAddBtn.UMCOL = -1
		Me.SerAddBtn.SOFT = 0
		Me.SerAddBtn.PICPOS = 0
		Me.SerAddBtn.NGREY = 0
		Me.SerAddBtn.FX = 0
		Me.SerAddBtn.HAND = 0
		Me.SerAddBtn.CHECK = 0
		Me.SerAddBtn.Name = "SerAddBtn"
		Me._SerBtn_0.Size = New System.Drawing.Size(27, 23)
		Me._SerBtn_0.Location = New System.Drawing.Point(171, 75)
		Me._SerBtn_0.TabIndex = 30
		Me._SerBtn_0.TX = ""
		Me._SerBtn_0.ENAB = -1
		Me._SerBtn_0.COLTYPE = 1
		Me._SerBtn_0.FOCUSR = -1
		Me._SerBtn_0.BCOL = 12632256
		Me._SerBtn_0.BCOLO = 12632256
		Me._SerBtn_0.FCOL = 0
		Me._SerBtn_0.FCOLO = 0
		Me._SerBtn_0.MCOL = 16777215
		Me._SerBtn_0.MPTR = 1
		Me._SerBtn_0.MICON = 0
		Me._SerBtn_0.PICN = 0
		Me._SerBtn_0.UMCOL = -1
		Me._SerBtn_0.SOFT = 0
		Me._SerBtn_0.PICPOS = 0
		Me._SerBtn_0.NGREY = 0
		Me._SerBtn_0.FX = 0
		Me._SerBtn_0.HAND = 0
		Me._SerBtn_0.CHECK = 0
		Me._SerBtn_0.Name = "_SerBtn_0"
		Me._SerBtn_1.Size = New System.Drawing.Size(27, 23)
		Me._SerBtn_1.Location = New System.Drawing.Point(143, 75)
		Me._SerBtn_1.TabIndex = 31
		Me._SerBtn_1.TX = ""
		Me._SerBtn_1.ENAB = -1
		Me._SerBtn_1.COLTYPE = 1
		Me._SerBtn_1.FOCUSR = -1
		Me._SerBtn_1.BCOL = 12632256
		Me._SerBtn_1.BCOLO = 12632256
		Me._SerBtn_1.FCOL = 0
		Me._SerBtn_1.FCOLO = 0
		Me._SerBtn_1.MCOL = 16777215
		Me._SerBtn_1.MPTR = 1
		Me._SerBtn_1.MICON = 0
		Me._SerBtn_1.PICN = 0
		Me._SerBtn_1.UMCOL = -1
		Me._SerBtn_1.SOFT = 0
		Me._SerBtn_1.PICPOS = 0
		Me._SerBtn_1.NGREY = 0
		Me._SerBtn_1.FX = 0
		Me._SerBtn_1.HAND = 0
		Me._SerBtn_1.CHECK = 0
		Me._SerBtn_1.Name = "_SerBtn_1"
		Me._SerLbl_3.Text = "Total :"
		Me._SerLbl_3.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._SerLbl_3.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me._SerLbl_3.Size = New System.Drawing.Size(45, 16)
		Me._SerLbl_3.Location = New System.Drawing.Point(439, 12)
		Me._SerLbl_3.TabIndex = 37
		Me._SerLbl_3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._SerLbl_3.BackColor = System.Drawing.Color.Transparent
		Me._SerLbl_3.Enabled = True
		Me._SerLbl_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._SerLbl_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._SerLbl_3.UseMnemonic = True
		Me._SerLbl_3.Visible = True
		Me._SerLbl_3.AutoSize = True
		Me._SerLbl_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._SerLbl_3.Name = "_SerLbl_3"
		Me._SerLbl_2.Text = "Quantity :"
		Me._SerLbl_2.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me._SerLbl_2.Size = New System.Drawing.Size(57, 13)
		Me._SerLbl_2.Location = New System.Drawing.Point(5, 75)
		Me._SerLbl_2.TabIndex = 27
		Me._SerLbl_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._SerLbl_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._SerLbl_2.BackColor = System.Drawing.Color.Transparent
		Me._SerLbl_2.Enabled = True
		Me._SerLbl_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._SerLbl_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._SerLbl_2.UseMnemonic = True
		Me._SerLbl_2.Visible = True
		Me._SerLbl_2.AutoSize = True
		Me._SerLbl_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._SerLbl_2.Name = "_SerLbl_2"
		Me._SerLbl_1.Text = "Items :"
		Me._SerLbl_1.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me._SerLbl_1.Size = New System.Drawing.Size(42, 13)
		Me._SerLbl_1.Location = New System.Drawing.Point(5, 41)
		Me._SerLbl_1.TabIndex = 25
		Me._SerLbl_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._SerLbl_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._SerLbl_1.BackColor = System.Drawing.Color.Transparent
		Me._SerLbl_1.Enabled = True
		Me._SerLbl_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._SerLbl_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._SerLbl_1.UseMnemonic = True
		Me._SerLbl_1.Visible = True
		Me._SerLbl_1.AutoSize = True
		Me._SerLbl_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._SerLbl_1.Name = "_SerLbl_1"
		Me._SerLbl_0.Text = "Category :"
		Me._SerLbl_0.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me._SerLbl_0.Size = New System.Drawing.Size(62, 13)
		Me._SerLbl_0.Location = New System.Drawing.Point(5, 7)
		Me._SerLbl_0.TabIndex = 23
		Me._SerLbl_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._SerLbl_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._SerLbl_0.BackColor = System.Drawing.Color.Transparent
		Me._SerLbl_0.Enabled = True
		Me._SerLbl_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._SerLbl_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._SerLbl_0.UseMnemonic = True
		Me._SerLbl_0.Visible = True
		Me._SerLbl_0.AutoSize = True
		Me._SerLbl_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._SerLbl_0.Name = "_SerLbl_0"
		Me._SerLbl_4.Text = "Balanced :"
		Me._SerLbl_4.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._SerLbl_4.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me._SerLbl_4.Size = New System.Drawing.Size(76, 16)
		Me._SerLbl_4.Location = New System.Drawing.Point(439, 49)
		Me._SerLbl_4.TabIndex = 39
		Me._SerLbl_4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._SerLbl_4.BackColor = System.Drawing.Color.Transparent
		Me._SerLbl_4.Enabled = True
		Me._SerLbl_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._SerLbl_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._SerLbl_4.UseMnemonic = True
		Me._SerLbl_4.Visible = True
		Me._SerLbl_4.AutoSize = True
		Me._SerLbl_4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._SerLbl_4.Name = "_SerLbl_4"
		Me._SerLbl_5.Text = "Received :"
		Me._SerLbl_5.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._SerLbl_5.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me._SerLbl_5.Size = New System.Drawing.Size(75, 16)
		Me._SerLbl_5.Location = New System.Drawing.Point(439, 85)
		Me._SerLbl_5.TabIndex = 41
		Me._SerLbl_5.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._SerLbl_5.BackColor = System.Drawing.Color.Transparent
		Me._SerLbl_5.Enabled = True
		Me._SerLbl_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._SerLbl_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._SerLbl_5.UseMnemonic = True
		Me._SerLbl_5.Visible = True
		Me._SerLbl_5.AutoSize = True
		Me._SerLbl_5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._SerLbl_5.Name = "_SerLbl_5"
		Me._SerLbl_6.Text = "Items Price :"
		Me._SerLbl_6.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me._SerLbl_6.Size = New System.Drawing.Size(74, 13)
		Me._SerLbl_6.Location = New System.Drawing.Point(5, 144)
		Me._SerLbl_6.TabIndex = 34
		Me._SerLbl_6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._SerLbl_6.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._SerLbl_6.BackColor = System.Drawing.Color.Transparent
		Me._SerLbl_6.Enabled = True
		Me._SerLbl_6.Cursor = System.Windows.Forms.Cursors.Default
		Me._SerLbl_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._SerLbl_6.UseMnemonic = True
		Me._SerLbl_6.Visible = True
		Me._SerLbl_6.AutoSize = True
		Me._SerLbl_6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._SerLbl_6.Name = "_SerLbl_6"
		Me._SerLbl_7.Text = "Price :"
		Me._SerLbl_7.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me._SerLbl_7.Size = New System.Drawing.Size(37, 13)
		Me._SerLbl_7.Location = New System.Drawing.Point(5, 110)
		Me._SerLbl_7.TabIndex = 32
		Me._SerLbl_7.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._SerLbl_7.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._SerLbl_7.BackColor = System.Drawing.Color.Transparent
		Me._SerLbl_7.Enabled = True
		Me._SerLbl_7.Cursor = System.Windows.Forms.Cursors.Default
		Me._SerLbl_7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._SerLbl_7.UseMnemonic = True
		Me._SerLbl_7.Visible = True
		Me._SerLbl_7.AutoSize = True
		Me._SerLbl_7.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._SerLbl_7.Name = "_SerLbl_7"
		Me._SubPages_3.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me._SubPages_3.ForeColor = System.Drawing.SystemColors.WindowText
		Me._SubPages_3.Size = New System.Drawing.Size(644, 171)
		Me._SubPages_3.Location = New System.Drawing.Point(38, 24)
		Me._SubPages_3.TabIndex = 48
		Me._SubPages_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._SubPages_3.Dock = System.Windows.Forms.DockStyle.None
		Me._SubPages_3.CausesValidation = True
		Me._SubPages_3.Enabled = True
		Me._SubPages_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._SubPages_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._SubPages_3.TabStop = True
		Me._SubPages_3.Visible = True
		Me._SubPages_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._SubPages_3.Name = "_SubPages_3"
		Me.MainLog.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.MainLog.BackColor = System.Drawing.Color.FromARGB(224, 224, 224)
		Me.MainLog.Size = New System.Drawing.Size(516, 161)
		Me.MainLog.Location = New System.Drawing.Point(4, 5)
		Me.MainLog.TabIndex = 49
		Me.MainLog.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.MainLog.CausesValidation = True
		Me.MainLog.Enabled = True
		Me.MainLog.ForeColor = System.Drawing.SystemColors.WindowText
		Me.MainLog.IntegralHeight = True
		Me.MainLog.Cursor = System.Windows.Forms.Cursors.Default
		Me.MainLog.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.MainLog.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.MainLog.Sorted = False
		Me.MainLog.TabStop = True
		Me.MainLog.Visible = True
		Me.MainLog.MultiColumn = False
		Me.MainLog.Name = "MainLog"
		Me._SubPages_2.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me._SubPages_2.ForeColor = System.Drawing.SystemColors.WindowText
		Me._SubPages_2.Size = New System.Drawing.Size(644, 171)
		Me._SubPages_2.Location = New System.Drawing.Point(38, 24)
		Me._SubPages_2.TabIndex = 44
		Me._SubPages_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._SubPages_2.Dock = System.Windows.Forms.DockStyle.None
		Me._SubPages_2.CausesValidation = True
		Me._SubPages_2.Enabled = True
		Me._SubPages_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._SubPages_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._SubPages_2.TabStop = True
		Me._SubPages_2.Visible = True
		Me._SubPages_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._SubPages_2.Name = "_SubPages_2"
		Me.MainNote.AutoSize = False
		Me.MainNote.Size = New System.Drawing.Size(479, 162)
		Me.MainNote.Location = New System.Drawing.Point(3, 4)
		Me.MainNote.MultiLine = True
		Me.MainNote.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
		Me.MainNote.TabIndex = 46
		Me.MainNote.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.MainNote.AcceptsReturn = True
		Me.MainNote.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.MainNote.BackColor = System.Drawing.SystemColors.Window
		Me.MainNote.CausesValidation = True
		Me.MainNote.Enabled = True
		Me.MainNote.ForeColor = System.Drawing.SystemColors.WindowText
		Me.MainNote.HideSelection = True
		Me.MainNote.ReadOnly = False
		Me.MainNote.Maxlength = 0
		Me.MainNote.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.MainNote.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.MainNote.TabStop = True
		Me.MainNote.Visible = True
		Me.MainNote.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.MainNote.Name = "MainNote"
		Me._MainNoteBtn_0.Size = New System.Drawing.Size(24, 21)
		Me._MainNoteBtn_0.Location = New System.Drawing.Point(485, 3)
		Me._MainNoteBtn_0.TabIndex = 45
		Me._MainNoteBtn_0.TX = ""
		Me._MainNoteBtn_0.ENAB = -1
		Me._MainNoteBtn_0.COLTYPE = 1
		Me._MainNoteBtn_0.FOCUSR = -1
		Me._MainNoteBtn_0.BCOL = 12632256
		Me._MainNoteBtn_0.BCOLO = 12632256
		Me._MainNoteBtn_0.FCOL = 0
		Me._MainNoteBtn_0.FCOLO = 0
		Me._MainNoteBtn_0.MCOL = 16777215
		Me._MainNoteBtn_0.MPTR = 1
		Me._MainNoteBtn_0.MICON = 0
		Me._MainNoteBtn_0.PICN = 0
		Me._MainNoteBtn_0.UMCOL = -1
		Me._MainNoteBtn_0.SOFT = 0
		Me._MainNoteBtn_0.PICPOS = 0
		Me._MainNoteBtn_0.NGREY = 0
		Me._MainNoteBtn_0.FX = 0
		Me._MainNoteBtn_0.HAND = 0
		Me._MainNoteBtn_0.CHECK = 0
		Me._MainNoteBtn_0.Name = "_MainNoteBtn_0"
		Me._MainNoteBtn_1.Size = New System.Drawing.Size(24, 21)
		Me._MainNoteBtn_1.Location = New System.Drawing.Point(485, 25)
		Me._MainNoteBtn_1.TabIndex = 47
		Me._MainNoteBtn_1.TX = ""
		Me._MainNoteBtn_1.ENAB = -1
		Me._MainNoteBtn_1.COLTYPE = 1
		Me._MainNoteBtn_1.FOCUSR = -1
		Me._MainNoteBtn_1.BCOL = 12632256
		Me._MainNoteBtn_1.BCOLO = 12632256
		Me._MainNoteBtn_1.FCOL = 0
		Me._MainNoteBtn_1.FCOLO = 0
		Me._MainNoteBtn_1.MCOL = 16777215
		Me._MainNoteBtn_1.MPTR = 1
		Me._MainNoteBtn_1.MICON = 0
		Me._MainNoteBtn_1.PICN = 0
		Me._MainNoteBtn_1.UMCOL = -1
		Me._MainNoteBtn_1.SOFT = 0
		Me._MainNoteBtn_1.PICPOS = 0
		Me._MainNoteBtn_1.NGREY = 0
		Me._MainNoteBtn_1.FX = 0
		Me._MainNoteBtn_1.HAND = 0
		Me._MainNoteBtn_1.CHECK = 0
		Me._MainNoteBtn_1.Name = "_MainNoteBtn_1"
		Me._SubPages_0.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me._SubPages_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._SubPages_0.Size = New System.Drawing.Size(644, 171)
		Me._SubPages_0.Location = New System.Drawing.Point(38, 24)
		Me._SubPages_0.TabIndex = 8
		Me._SubPages_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._SubPages_0.Dock = System.Windows.Forms.DockStyle.None
		Me._SubPages_0.CausesValidation = True
		Me._SubPages_0.Enabled = True
		Me._SubPages_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._SubPages_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._SubPages_0.TabStop = True
		Me._SubPages_0.Visible = True
		Me._SubPages_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._SubPages_0.Name = "_SubPages_0"
		Me.uLine3D2.Size = New System.Drawing.Size(525, 3)
		Me.uLine3D2.Location = New System.Drawing.Point(0, 80)
		Me.uLine3D2.TabIndex = 17
		Me.uLine3D2.horizon = -1
		Me.uLine3D2.Name = "uLine3D2"
		Me._SpgInfoLblD_3.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._SpgInfoLblD_3.BackColor = System.Drawing.Color.FromARGB(224, 224, 224)
		Me._SpgInfoLblD_3.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me._SpgInfoLblD_3.Size = New System.Drawing.Size(141, 19)
		Me._SpgInfoLblD_3.Location = New System.Drawing.Point(379, 52)
		Me._SpgInfoLblD_3.TabIndex = 16
		Me._SpgInfoLblD_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._SpgInfoLblD_3.Enabled = True
		Me._SpgInfoLblD_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._SpgInfoLblD_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._SpgInfoLblD_3.UseMnemonic = True
		Me._SpgInfoLblD_3.Visible = True
		Me._SpgInfoLblD_3.AutoSize = False
		Me._SpgInfoLblD_3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._SpgInfoLblD_3.Name = "_SpgInfoLblD_3"
		Me._SpgInfoLblD_2.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._SpgInfoLblD_2.BackColor = System.Drawing.Color.FromARGB(224, 224, 224)
		Me._SpgInfoLblD_2.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me._SpgInfoLblD_2.Size = New System.Drawing.Size(141, 19)
		Me._SpgInfoLblD_2.Location = New System.Drawing.Point(379, 31)
		Me._SpgInfoLblD_2.TabIndex = 14
		Me._SpgInfoLblD_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._SpgInfoLblD_2.Enabled = True
		Me._SpgInfoLblD_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._SpgInfoLblD_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._SpgInfoLblD_2.UseMnemonic = True
		Me._SpgInfoLblD_2.Visible = True
		Me._SpgInfoLblD_2.AutoSize = False
		Me._SpgInfoLblD_2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._SpgInfoLblD_2.Name = "_SpgInfoLblD_2"
		Me._SpgInfoLblD_1.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._SpgInfoLblD_1.BackColor = System.Drawing.Color.FromARGB(224, 224, 224)
		Me._SpgInfoLblD_1.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me._SpgInfoLblD_1.Size = New System.Drawing.Size(141, 19)
		Me._SpgInfoLblD_1.Location = New System.Drawing.Point(129, 52)
		Me._SpgInfoLblD_1.TabIndex = 13
		Me._SpgInfoLblD_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._SpgInfoLblD_1.Enabled = True
		Me._SpgInfoLblD_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._SpgInfoLblD_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._SpgInfoLblD_1.UseMnemonic = True
		Me._SpgInfoLblD_1.Visible = True
		Me._SpgInfoLblD_1.AutoSize = False
		Me._SpgInfoLblD_1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._SpgInfoLblD_1.Name = "_SpgInfoLblD_1"
		Me._SpgInfoLblD_0.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._SpgInfoLblD_0.BackColor = System.Drawing.Color.FromARGB(224, 224, 224)
		Me._SpgInfoLblD_0.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me._SpgInfoLblD_0.Size = New System.Drawing.Size(141, 19)
		Me._SpgInfoLblD_0.Location = New System.Drawing.Point(129, 31)
		Me._SpgInfoLblD_0.TabIndex = 10
		Me._SpgInfoLblD_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._SpgInfoLblD_0.Enabled = True
		Me._SpgInfoLblD_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._SpgInfoLblD_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._SpgInfoLblD_0.UseMnemonic = True
		Me._SpgInfoLblD_0.Visible = True
		Me._SpgInfoLblD_0.AutoSize = False
		Me._SpgInfoLblD_0.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._SpgInfoLblD_0.Name = "_SpgInfoLblD_0"
		Me._SpgInfoLblC_3.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._SpgInfoLblC_3.Text = "Current Used :"
		Me._SpgInfoLblC_3.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._SpgInfoLblC_3.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me._SpgInfoLblC_3.Size = New System.Drawing.Size(93, 13)
		Me._SpgInfoLblC_3.Location = New System.Drawing.Point(275, 53)
		Me._SpgInfoLblC_3.TabIndex = 15
		Me._SpgInfoLblC_3.BackColor = System.Drawing.Color.Transparent
		Me._SpgInfoLblC_3.Enabled = True
		Me._SpgInfoLblC_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._SpgInfoLblC_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._SpgInfoLblC_3.UseMnemonic = True
		Me._SpgInfoLblC_3.Visible = True
		Me._SpgInfoLblC_3.AutoSize = True
		Me._SpgInfoLblC_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._SpgInfoLblC_3.Name = "_SpgInfoLblC_3"
		Me._SpgInfoLblC_2.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._SpgInfoLblC_2.Text = "MAC Address :"
		Me._SpgInfoLblC_2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._SpgInfoLblC_2.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me._SpgInfoLblC_2.Size = New System.Drawing.Size(92, 13)
		Me._SpgInfoLblC_2.Location = New System.Drawing.Point(276, 32)
		Me._SpgInfoLblC_2.TabIndex = 12
		Me._SpgInfoLblC_2.BackColor = System.Drawing.Color.Transparent
		Me._SpgInfoLblC_2.Enabled = True
		Me._SpgInfoLblC_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._SpgInfoLblC_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._SpgInfoLblC_2.UseMnemonic = True
		Me._SpgInfoLblC_2.Visible = True
		Me._SpgInfoLblC_2.AutoSize = True
		Me._SpgInfoLblC_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._SpgInfoLblC_2.Name = "_SpgInfoLblC_2"
		Me._SpgInfoLblC_1.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._SpgInfoLblC_1.Text = "IP Address :"
		Me._SpgInfoLblC_1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._SpgInfoLblC_1.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me._SpgInfoLblC_1.Size = New System.Drawing.Size(79, 13)
		Me._SpgInfoLblC_1.Location = New System.Drawing.Point(39, 53)
		Me._SpgInfoLblC_1.TabIndex = 11
		Me._SpgInfoLblC_1.BackColor = System.Drawing.Color.Transparent
		Me._SpgInfoLblC_1.Enabled = True
		Me._SpgInfoLblC_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._SpgInfoLblC_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._SpgInfoLblC_1.UseMnemonic = True
		Me._SpgInfoLblC_1.Visible = True
		Me._SpgInfoLblC_1.AutoSize = True
		Me._SpgInfoLblC_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._SpgInfoLblC_1.Name = "_SpgInfoLblC_1"
		Me._SpgInfoLblC_0.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._SpgInfoLblC_0.Text = "Connected At :"
		Me._SpgInfoLblC_0.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._SpgInfoLblC_0.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me._SpgInfoLblC_0.Size = New System.Drawing.Size(94, 13)
		Me._SpgInfoLblC_0.Location = New System.Drawing.Point(24, 32)
		Me._SpgInfoLblC_0.TabIndex = 9
		Me._SpgInfoLblC_0.BackColor = System.Drawing.Color.Transparent
		Me._SpgInfoLblC_0.Enabled = True
		Me._SpgInfoLblC_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._SpgInfoLblC_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._SpgInfoLblC_0.UseMnemonic = True
		Me._SpgInfoLblC_0.Visible = True
		Me._SpgInfoLblC_0.AutoSize = True
		Me._SpgInfoLblC_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._SpgInfoLblC_0.Name = "_SpgInfoLblC_0"
		Me._SpgInfoHdr_1.Size = New System.Drawing.Size(215, 20)
		Me._SpgInfoHdr_1.Location = New System.Drawing.Point(6, 4)
		Me._SpgInfoHdr_1.Image = CType(resources.GetObject("_SpgInfoHdr_1.Image"), System.Drawing.Image)
		Me._SpgInfoHdr_1.Enabled = True
		Me._SpgInfoHdr_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._SpgInfoHdr_1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._SpgInfoHdr_1.Visible = True
		Me._SpgInfoHdr_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._SpgInfoHdr_1.Name = "_SpgInfoHdr_1"
		Me._SpgInfoHdr_0.Size = New System.Drawing.Size(160, 18)
		Me._SpgInfoHdr_0.Location = New System.Drawing.Point(5, 94)
		Me._SpgInfoHdr_0.Image = CType(resources.GetObject("_SpgInfoHdr_0.Image"), System.Drawing.Image)
		Me._SpgInfoHdr_0.Enabled = True
		Me._SpgInfoHdr_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._SpgInfoHdr_0.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._SpgInfoHdr_0.Visible = True
		Me._SpgInfoHdr_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._SpgInfoHdr_0.Name = "_SpgInfoHdr_0"
		Me._SpgInfoLblB_1.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._SpgInfoLblB_1.BackColor = System.Drawing.Color.FromARGB(224, 224, 224)
		Me._SpgInfoLblB_1.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me._SpgInfoLblB_1.Size = New System.Drawing.Size(59, 19)
		Me._SpgInfoLblB_1.Location = New System.Drawing.Point(153, 141)
		Me._SpgInfoLblB_1.TabIndex = 21
		Me._SpgInfoLblB_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._SpgInfoLblB_1.Enabled = True
		Me._SpgInfoLblB_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._SpgInfoLblB_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._SpgInfoLblB_1.UseMnemonic = True
		Me._SpgInfoLblB_1.Visible = True
		Me._SpgInfoLblB_1.AutoSize = False
		Me._SpgInfoLblB_1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._SpgInfoLblB_1.Name = "_SpgInfoLblB_1"
		Me._SpgInfoLblA_1.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._SpgInfoLblA_1.Text = "Unused Station :"
		Me._SpgInfoLblA_1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._SpgInfoLblA_1.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me._SpgInfoLblA_1.Size = New System.Drawing.Size(106, 13)
		Me._SpgInfoLblA_1.Location = New System.Drawing.Point(36, 142)
		Me._SpgInfoLblA_1.TabIndex = 20
		Me._SpgInfoLblA_1.BackColor = System.Drawing.Color.Transparent
		Me._SpgInfoLblA_1.Enabled = True
		Me._SpgInfoLblA_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._SpgInfoLblA_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._SpgInfoLblA_1.UseMnemonic = True
		Me._SpgInfoLblA_1.Visible = True
		Me._SpgInfoLblA_1.AutoSize = True
		Me._SpgInfoLblA_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._SpgInfoLblA_1.Name = "_SpgInfoLblA_1"
		Me._SpgInfoLblB_0.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._SpgInfoLblB_0.BackColor = System.Drawing.Color.FromARGB(224, 224, 224)
		Me._SpgInfoLblB_0.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me._SpgInfoLblB_0.Size = New System.Drawing.Size(59, 19)
		Me._SpgInfoLblB_0.Location = New System.Drawing.Point(153, 120)
		Me._SpgInfoLblB_0.TabIndex = 19
		Me._SpgInfoLblB_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._SpgInfoLblB_0.Enabled = True
		Me._SpgInfoLblB_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._SpgInfoLblB_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._SpgInfoLblB_0.UseMnemonic = True
		Me._SpgInfoLblB_0.Visible = True
		Me._SpgInfoLblB_0.AutoSize = False
		Me._SpgInfoLblB_0.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._SpgInfoLblB_0.Name = "_SpgInfoLblB_0"
		Me._SpgInfoLblA_0.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._SpgInfoLblA_0.Text = "Connected Agent :"
		Me._SpgInfoLblA_0.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._SpgInfoLblA_0.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me._SpgInfoLblA_0.Size = New System.Drawing.Size(118, 13)
		Me._SpgInfoLblA_0.Location = New System.Drawing.Point(24, 121)
		Me._SpgInfoLblA_0.TabIndex = 18
		Me._SpgInfoLblA_0.BackColor = System.Drawing.Color.Transparent
		Me._SpgInfoLblA_0.Enabled = True
		Me._SpgInfoLblA_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._SpgInfoLblA_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._SpgInfoLblA_0.UseMnemonic = True
		Me._SpgInfoLblA_0.Visible = True
		Me._SpgInfoLblA_0.AutoSize = True
		Me._SpgInfoLblA_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._SpgInfoLblA_0.Name = "_SpgInfoLblA_0"
		Lv1.OcxState = CType(resources.GetObject("Lv1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.Lv1.Size = New System.Drawing.Size(677, 306)
		Me.Lv1.Location = New System.Drawing.Point(0, 0)
		Me.Lv1.TabIndex = 51
		Me.Lv1.Name = "Lv1"
		Me.menu1.Text = "Menu"
		Me.menu1.Checked = False
		Me.menu1.Enabled = True
		Me.menu1.Visible = True
		Me.menu1.MDIList = False
		Me.menu1penalaan.Text = "Setting"
		Me.menu1penalaan.Checked = False
		Me.menu1penalaan.Enabled = True
		Me.menu1penalaan.Visible = True
		Me.menu1penalaan.MDIList = False
		Me.menu1sep1.Text = "-"
		Me.menu1sep1.Checked = False
		Me.menu1sep1.Enabled = True
		Me.menu1sep1.Visible = True
		Me.menu1sep1.MDIList = False
		Me.menu1logout.Text = "Logout"
		Me.menu1logout.Checked = False
		Me.menu1logout.Enabled = True
		Me.menu1logout.Visible = True
		Me.menu1logout.MDIList = False
		Me.menu2keluar.Text = "Close"
		Me.menu2keluar.Checked = False
		Me.menu2keluar.Enabled = True
		Me.menu2keluar.Visible = True
		Me.menu2keluar.MDIList = False
		Me.menu3.Text = "Station"
		Me.menu3.Checked = False
		Me.menu3.Enabled = True
		Me.menu3.Visible = True
		Me.menu3.MDIList = False
		Me.menu3announce.Text = "Broadcast"
		Me.menu3announce.Checked = False
		Me.menu3announce.Enabled = True
		Me.menu3announce.Visible = True
		Me.menu3announce.MDIList = False
		Me.menu3mesej.Text = "Message"
		Me.menu3mesej.Checked = False
		Me.menu3mesej.Enabled = True
		Me.menu3mesej.Visible = True
		Me.menu3mesej.MDIList = False
		Me.menu3tiker.Text = "Ticker"
		Me.menu3tiker.Checked = False
		Me.menu3tiker.Enabled = True
		Me.menu3tiker.Visible = True
		Me.menu3tiker.MDIList = False
		Me.menu3ctl.Text = "Control"
		Me.menu3ctl.Checked = False
		Me.menu3ctl.Enabled = True
		Me.menu3ctl.Visible = True
		Me.menu3ctl.MDIList = False
		Me._menu3ctllock_0.Text = "Lock All"
		Me._menu3ctllock_0.Checked = False
		Me._menu3ctllock_0.Enabled = True
		Me._menu3ctllock_0.Visible = True
		Me._menu3ctllock_0.MDIList = False
		Me._menu3ctllock_1.Text = "Lock Unused"
		Me._menu3ctllock_1.Checked = False
		Me._menu3ctllock_1.Enabled = True
		Me._menu3ctllock_1.Visible = True
		Me._menu3ctllock_1.MDIList = False
		Me._menu3ctllock_2.Text = "Unlock All"
		Me._menu3ctllock_2.Checked = False
		Me._menu3ctllock_2.Enabled = True
		Me._menu3ctllock_2.Visible = True
		Me._menu3ctllock_2.MDIList = False
		Me.menu3sep2.Text = "-"
		Me.menu3sep2.Checked = False
		Me.menu3sep2.Enabled = True
		Me.menu3sep2.Visible = True
		Me.menu3sep2.MDIList = False
		Me._menu3ctlwinexit_0.Text = "Shutdown All"
		Me._menu3ctlwinexit_0.Checked = False
		Me._menu3ctlwinexit_0.Enabled = True
		Me._menu3ctlwinexit_0.Visible = True
		Me._menu3ctlwinexit_0.MDIList = False
		Me._menu3ctlwinexit_1.Text = "Shutdown Unused"
		Me._menu3ctlwinexit_1.Checked = False
		Me._menu3ctlwinexit_1.Enabled = True
		Me._menu3ctlwinexit_1.Visible = True
		Me._menu3ctlwinexit_1.MDIList = False
		Me._menu3ctlwinexit_2.Text = "Reboot All"
		Me._menu3ctlwinexit_2.Checked = False
		Me._menu3ctlwinexit_2.Enabled = True
		Me._menu3ctlwinexit_2.Visible = True
		Me._menu3ctlwinexit_2.MDIList = False
		Me._menu3ctlwinexit_3.Text = "Reboot Unused"
		Me._menu3ctlwinexit_3.Checked = False
		Me._menu3ctlwinexit_3.Enabled = True
		Me._menu3ctlwinexit_3.Visible = True
		Me._menu3ctlwinexit_3.MDIList = False
		Me.menu3cln.Text = "Cleaning"
		Me.menu3cln.Checked = False
		Me.menu3cln.Enabled = True
		Me.menu3cln.Visible = True
		Me.menu3cln.MDIList = False
		Me._menu3clnsub_0.Text = "All"
		Me._menu3clnsub_0.Checked = False
		Me._menu3clnsub_0.Enabled = True
		Me._menu3clnsub_0.Visible = True
		Me._menu3clnsub_0.MDIList = False
		Me._menu3clnsub_1.Text = "-"
		Me._menu3clnsub_1.Checked = False
		Me._menu3clnsub_1.Enabled = True
		Me._menu3clnsub_1.Visible = True
		Me._menu3clnsub_1.MDIList = False
		Me._menu3clnsub_2.Text = "Temp Folder"
		Me._menu3clnsub_2.Checked = False
		Me._menu3clnsub_2.Enabled = True
		Me._menu3clnsub_2.Visible = True
		Me._menu3clnsub_2.MDIList = False
		Me._menu3clnsub_3.Text = "Recycle Bin"
		Me._menu3clnsub_3.Checked = False
		Me._menu3clnsub_3.Enabled = True
		Me._menu3clnsub_3.Visible = True
		Me._menu3clnsub_3.MDIList = False
		Me._menu3clnsub_4.Text = "Internet History"
		Me._menu3clnsub_4.Checked = False
		Me._menu3clnsub_4.Enabled = True
		Me._menu3clnsub_4.Visible = True
		Me._menu3clnsub_4.MDIList = False
		Me._menu3clnsub_5.Text = "Recent Docs"
		Me._menu3clnsub_5.Checked = False
		Me._menu3clnsub_5.Enabled = True
		Me._menu3clnsub_5.Visible = True
		Me._menu3clnsub_5.MDIList = False
		Me.menu3sep1.Text = "-"
		Me.menu3sep1.Checked = False
		Me.menu3sep1.Enabled = True
		Me.menu3sep1.Visible = True
		Me.menu3sep1.MDIList = False
		Me.menu3AgMgr.Text = "Agent Manager"
		Me.menu3AgMgr.Checked = False
		Me.menu3AgMgr.Enabled = True
		Me.menu3AgMgr.Visible = True
		Me.menu3AgMgr.MDIList = False
		Me.menu4.Text = "View"
		Me.menu4.Checked = False
		Me.menu4.Enabled = True
		Me.menu4.Visible = True
		Me.menu4.MDIList = False
		Me._menu4mon_0.Text = "Printing Monitoring"
		Me._menu4mon_0.Checked = False
		Me._menu4mon_0.Enabled = True
		Me._menu4mon_0.Visible = True
		Me._menu4mon_0.MDIList = False
		Me._menu4mon_1.Text = "Resource Monitoring"
		Me._menu4mon_1.Checked = False
		Me._menu4mon_1.Enabled = True
		Me._menu4mon_1.Visible = True
		Me._menu4mon_1.MDIList = False
		Me._menu4mon_2.Text = "Application Monitoring"
		Me._menu4mon_2.Checked = False
		Me._menu4mon_2.Enabled = True
		Me._menu4mon_2.Visible = True
		Me._menu4mon_2.MDIList = False
		Me._menu4mon_3.Text = "Traffic Monitoring"
		Me._menu4mon_3.Checked = False
		Me._menu4mon_3.Enabled = True
		Me._menu4mon_3.Visible = True
		Me._menu4mon_3.MDIList = False
		Me.menu4sep1.Text = "-"
		Me.menu4sep1.Checked = False
		Me.menu4sep1.Enabled = True
		Me.menu4sep1.Visible = True
		Me.menu4sep1.MDIList = False
		Me.menu4env.Text = "Enviroment"
		Me.menu4env.Checked = False
		Me.menu4env.Enabled = True
		Me.menu4env.Visible = True
		Me.menu4env.MDIList = False
		Me._menu4envsub_0.Text = "Toolbox"
		Me._menu4envsub_0.Checked = True
		Me._menu4envsub_0.Enabled = True
		Me._menu4envsub_0.Visible = True
		Me._menu4envsub_0.MDIList = False
		Me._menu4envsub_1.Text = "Menu Bar"
		Me._menu4envsub_1.Checked = True
		Me._menu4envsub_1.Enabled = True
		Me._menu4envsub_1.Visible = True
		Me._menu4envsub_1.MDIList = False
		Me.menu5.Text = "Tools"
		Me.menu5.Checked = False
		Me.menu5.Enabled = True
		Me.menu5.Visible = True
		Me.menu5.MDIList = False
		Me.menu4PosMgr.Text = "S&&M Manager"
		Me.menu4PosMgr.Checked = False
		Me.menu4PosMgr.Enabled = True
		Me.menu4PosMgr.Visible = True
		Me.menu4PosMgr.MDIList = False
		Me.menu4Stat.Text = "Statistic System"
		Me.menu4Stat.Checked = False
		Me.menu4Stat.Enabled = True
		Me.menu4Stat.Visible = True
		Me.menu4Stat.MDIList = False
		Me.menu5Console.Text = "Console System"
		Me.menu5Console.Checked = False
		Me.menu5Console.Enabled = True
		Me.menu5Console.Visible = True
		Me.menu5Console.MDIList = False
		Me.menu2.Text = "Info"
		Me.menu2.Checked = False
		Me.menu2.Enabled = True
		Me.menu2.Visible = True
		Me.menu2.MDIList = False
		Me.menu2bantuan.Text = "Help"
		Me.menu2bantuan.Checked = False
		Me.menu2bantuan.Enabled = True
		Me.menu2bantuan.Visible = True
		Me.menu2bantuan.MDIList = False
		Me.menu2sep1.Text = "-"
		Me.menu2sep1.Checked = False
		Me.menu2sep1.Enabled = True
		Me.menu2sep1.Visible = True
		Me.menu2sep1.MDIList = False
		Me.menu2aplikasi.Text = "About.."
		Me.menu2aplikasi.Checked = False
		Me.menu2aplikasi.Enabled = True
		Me.menu2aplikasi.Visible = True
		Me.menu2aplikasi.MDIList = False
		Me.popmenu1.Text = "<popmenu1>"
		Me.popmenu1.Visible = False
		Me.popmenu1.Checked = False
		Me.popmenu1.Enabled = True
		Me.popmenu1.MDIList = False
		Me.pmenu1flog.Text = "Fast Login"
		Me.pmenu1flog.Checked = False
		Me.pmenu1flog.Enabled = True
		Me.pmenu1flog.Visible = True
		Me.pmenu1flog.MDIList = False
		Me.pmenu1flout.Text = "Fast Logout"
		Me.pmenu1flout.Checked = False
		Me.pmenu1flout.Enabled = True
		Me.pmenu1flout.Visible = True
		Me.pmenu1flout.MDIList = False
		Me.psep2.Text = "-"
		Me.psep2.Checked = False
		Me.psep2.Enabled = True
		Me.psep2.Visible = True
		Me.psep2.MDIList = False
		Me.pmenu1cancel.Text = "Cancel User"
		Me.pmenu1cancel.Checked = False
		Me.pmenu1cancel.Enabled = True
		Me.pmenu1cancel.Visible = True
		Me.pmenu1cancel.MDIList = False
		Me.pmenu1trans.Text = "Transfer PC"
		Me.pmenu1trans.Checked = False
		Me.pmenu1trans.Enabled = True
		Me.pmenu1trans.Visible = True
		Me.pmenu1trans.MDIList = False
		Me.pmenu1terminal.Text = "Terminal"
		Me.pmenu1terminal.Checked = False
		Me.pmenu1terminal.Enabled = True
		Me.pmenu1terminal.Visible = True
		Me.pmenu1terminal.MDIList = False
		Me.psep1.Text = "-"
		Me.psep1.Checked = False
		Me.psep1.Enabled = True
		Me.psep1.Visible = True
		Me.psep1.MDIList = False
		Me.pmenu1cln.Text = "Cleaning"
		Me.pmenu1cln.Checked = False
		Me.pmenu1cln.Enabled = True
		Me.pmenu1cln.Visible = True
		Me.pmenu1cln.MDIList = False
		Me._pmenu1clnsub_0.Text = "All"
		Me._pmenu1clnsub_0.Checked = False
		Me._pmenu1clnsub_0.Enabled = True
		Me._pmenu1clnsub_0.Visible = True
		Me._pmenu1clnsub_0.MDIList = False
		Me._pmenu1clnsub_1.Text = "-"
		Me._pmenu1clnsub_1.Checked = False
		Me._pmenu1clnsub_1.Enabled = True
		Me._pmenu1clnsub_1.Visible = True
		Me._pmenu1clnsub_1.MDIList = False
		Me._pmenu1clnsub_2.Text = "Temp Folder"
		Me._pmenu1clnsub_2.Checked = False
		Me._pmenu1clnsub_2.Enabled = True
		Me._pmenu1clnsub_2.Visible = True
		Me._pmenu1clnsub_2.MDIList = False
		Me._pmenu1clnsub_3.Text = "Recycle Bin"
		Me._pmenu1clnsub_3.Checked = False
		Me._pmenu1clnsub_3.Enabled = True
		Me._pmenu1clnsub_3.Visible = True
		Me._pmenu1clnsub_3.MDIList = False
		Me._pmenu1clnsub_4.Text = "Internet History"
		Me._pmenu1clnsub_4.Checked = False
		Me._pmenu1clnsub_4.Enabled = True
		Me._pmenu1clnsub_4.Visible = True
		Me._pmenu1clnsub_4.MDIList = False
		Me._pmenu1clnsub_5.Text = "Recent Docs"
		Me._pmenu1clnsub_5.Checked = False
		Me._pmenu1clnsub_5.Enabled = True
		Me._pmenu1clnsub_5.Visible = True
		Me._pmenu1clnsub_5.MDIList = False
		Me.pmenu1ctl.Text = "Control"
		Me.pmenu1ctl.Checked = False
		Me.pmenu1ctl.Enabled = True
		Me.pmenu1ctl.Visible = True
		Me.pmenu1ctl.MDIList = False
		Me._pmenu1ctlsub_0.Text = "Lock Computer"
		Me._pmenu1ctlsub_0.Checked = False
		Me._pmenu1ctlsub_0.Enabled = True
		Me._pmenu1ctlsub_0.Visible = True
		Me._pmenu1ctlsub_0.MDIList = False
		Me._pmenu1ctlsub_1.Text = "Unlock Computer"
		Me._pmenu1ctlsub_1.Checked = False
		Me._pmenu1ctlsub_1.Enabled = True
		Me._pmenu1ctlsub_1.Visible = True
		Me._pmenu1ctlsub_1.MDIList = False
		Me._pmenu1ctlsub_2.Text = "Reboot Computer"
		Me._pmenu1ctlsub_2.Checked = False
		Me._pmenu1ctlsub_2.Enabled = True
		Me._pmenu1ctlsub_2.Visible = True
		Me._pmenu1ctlsub_2.MDIList = False
		Me._pmenu1ctlsub_3.Text = "Shutdown Computer"
		Me._pmenu1ctlsub_3.Checked = False
		Me._pmenu1ctlsub_3.Enabled = True
		Me._pmenu1ctlsub_3.Visible = True
		Me._pmenu1ctlsub_3.MDIList = False
		Me.Controls.Add(MainPdock)
		Me.Controls.Add(ImgList32)
		Me.Controls.Add(ImgList16)
		Me.Controls.Add(ImgListSnm)
		Me.Controls.Add(MainSbar)
		Me.Controls.Add(MainPhold)
		Me.Controls.Add(Lv1)
		Me.MainPdock.Controls.Add(DockBar)
		Me.DockBar.Controls.Add(_MainDbtn_5)
		Me.DockBar.Controls.Add(_MainDbtn_4)
		Me.DockBar.Controls.Add(_MainDbtn_3)
		Me.DockBar.Controls.Add(_MainDbtn_2)
		Me.DockBar.Controls.Add(_MainDbtn_1)
		Me.DockBar.Controls.Add(_MainDbtn_0)
		Me.MainPhold.Controls.Add(uLine3D1)
		Me.MainPhold.Controls.Add(_SubPagesMnu_3)
		Me.MainPhold.Controls.Add(_SubPagesMnu_2)
		Me.MainPhold.Controls.Add(_SubPagesMnu_1)
		Me.MainPhold.Controls.Add(_SubPagesMnu_0)
		Me.MainPhold.Controls.Add(_SubPages_1)
		Me.MainPhold.Controls.Add(_SubPages_3)
		Me.MainPhold.Controls.Add(_SubPages_2)
		Me.MainPhold.Controls.Add(_SubPages_0)
		Me._SubPages_1.Controls.Add(SerTxtJumlah)
		Me._SubPages_1.Controls.Add(SerTxtQty)
		Me._SubPages_1.Controls.Add(SerTxtBaki)
		Me._SubPages_1.Controls.Add(SerTxtBayar)
		Me._SubPages_1.Controls.Add(SerScroll1)
		Me._SubPages_1.Controls.Add(SerTxtTotalItm)
		Me._SubPages_1.Controls.Add(SerTxtPriItm)
		Me._SubPages_1.Controls.Add(SerImgCb2)
		Me._SubPages_1.Controls.Add(SerImgCb1)
		Me._SubPages_1.Controls.Add(SerLv1)
		Me._SubPages_1.Controls.Add(SerAddBtn)
		Me._SubPages_1.Controls.Add(_SerBtn_0)
		Me._SubPages_1.Controls.Add(_SerBtn_1)
		Me._SubPages_1.Controls.Add(_SerLbl_3)
		Me._SubPages_1.Controls.Add(_SerLbl_2)
		Me._SubPages_1.Controls.Add(_SerLbl_1)
		Me._SubPages_1.Controls.Add(_SerLbl_0)
		Me._SubPages_1.Controls.Add(_SerLbl_4)
		Me._SubPages_1.Controls.Add(_SerLbl_5)
		Me._SubPages_1.Controls.Add(_SerLbl_6)
		Me._SubPages_1.Controls.Add(_SerLbl_7)
		Me._SubPages_3.Controls.Add(MainLog)
		Me._SubPages_2.Controls.Add(MainNote)
		Me._SubPages_2.Controls.Add(_MainNoteBtn_0)
		Me._SubPages_2.Controls.Add(_MainNoteBtn_1)
		Me._SubPages_0.Controls.Add(uLine3D2)
		Me._SubPages_0.Controls.Add(_SpgInfoLblD_3)
		Me._SubPages_0.Controls.Add(_SpgInfoLblD_2)
		Me._SubPages_0.Controls.Add(_SpgInfoLblD_1)
		Me._SubPages_0.Controls.Add(_SpgInfoLblD_0)
		Me._SubPages_0.Controls.Add(_SpgInfoLblC_3)
		Me._SubPages_0.Controls.Add(_SpgInfoLblC_2)
		Me._SubPages_0.Controls.Add(_SpgInfoLblC_1)
		Me._SubPages_0.Controls.Add(_SpgInfoLblC_0)
		Me._SubPages_0.Controls.Add(_SpgInfoHdr_1)
		Me._SubPages_0.Controls.Add(_SpgInfoHdr_0)
		Me._SubPages_0.Controls.Add(_SpgInfoLblB_1)
		Me._SubPages_0.Controls.Add(_SpgInfoLblA_1)
		Me._SubPages_0.Controls.Add(_SpgInfoLblB_0)
		Me._SubPages_0.Controls.Add(_SpgInfoLblA_0)
		Me.MainDbtn.SetIndex(_MainDbtn_5, CType(5, Short))
		Me.MainDbtn.SetIndex(_MainDbtn_4, CType(4, Short))
		Me.MainDbtn.SetIndex(_MainDbtn_3, CType(3, Short))
		Me.MainDbtn.SetIndex(_MainDbtn_2, CType(2, Short))
		Me.MainDbtn.SetIndex(_MainDbtn_1, CType(1, Short))
		Me.MainDbtn.SetIndex(_MainDbtn_0, CType(0, Short))
		Me.MainNoteBtn.SetIndex(_MainNoteBtn_0, CType(0, Short))
		Me.MainNoteBtn.SetIndex(_MainNoteBtn_1, CType(1, Short))
		Me.SerBtn.SetIndex(_SerBtn_0, CType(0, Short))
		Me.SerBtn.SetIndex(_SerBtn_1, CType(1, Short))
		Me.SerLbl.SetIndex(_SerLbl_3, CType(3, Short))
		Me.SerLbl.SetIndex(_SerLbl_2, CType(2, Short))
		Me.SerLbl.SetIndex(_SerLbl_1, CType(1, Short))
		Me.SerLbl.SetIndex(_SerLbl_0, CType(0, Short))
		Me.SerLbl.SetIndex(_SerLbl_4, CType(4, Short))
		Me.SerLbl.SetIndex(_SerLbl_5, CType(5, Short))
		Me.SerLbl.SetIndex(_SerLbl_6, CType(6, Short))
		Me.SerLbl.SetIndex(_SerLbl_7, CType(7, Short))
		Me.SpgInfoHdr.SetIndex(_SpgInfoHdr_1, CType(1, Short))
		Me.SpgInfoHdr.SetIndex(_SpgInfoHdr_0, CType(0, Short))
		Me.SpgInfoLblA.SetIndex(_SpgInfoLblA_1, CType(1, Short))
		Me.SpgInfoLblA.SetIndex(_SpgInfoLblA_0, CType(0, Short))
		Me.SpgInfoLblB.SetIndex(_SpgInfoLblB_1, CType(1, Short))
		Me.SpgInfoLblB.SetIndex(_SpgInfoLblB_0, CType(0, Short))
		Me.SpgInfoLblC.SetIndex(_SpgInfoLblC_3, CType(3, Short))
		Me.SpgInfoLblC.SetIndex(_SpgInfoLblC_2, CType(2, Short))
		Me.SpgInfoLblC.SetIndex(_SpgInfoLblC_1, CType(1, Short))
		Me.SpgInfoLblC.SetIndex(_SpgInfoLblC_0, CType(0, Short))
		Me.SpgInfoLblD.SetIndex(_SpgInfoLblD_3, CType(3, Short))
		Me.SpgInfoLblD.SetIndex(_SpgInfoLblD_2, CType(2, Short))
		Me.SpgInfoLblD.SetIndex(_SpgInfoLblD_1, CType(1, Short))
		Me.SpgInfoLblD.SetIndex(_SpgInfoLblD_0, CType(0, Short))
		Me.SubPages.SetIndex(_SubPages_1, CType(1, Short))
		Me.SubPages.SetIndex(_SubPages_3, CType(3, Short))
		Me.SubPages.SetIndex(_SubPages_2, CType(2, Short))
		Me.SubPages.SetIndex(_SubPages_0, CType(0, Short))
		Me.SubPagesMnu.SetIndex(_SubPagesMnu_3, CType(3, Short))
		Me.SubPagesMnu.SetIndex(_SubPagesMnu_2, CType(2, Short))
		Me.SubPagesMnu.SetIndex(_SubPagesMnu_1, CType(1, Short))
		Me.SubPagesMnu.SetIndex(_SubPagesMnu_0, CType(0, Short))
		Me.menu3clnsub.SetIndex(_menu3clnsub_0, CType(0, Short))
		Me.menu3clnsub.SetIndex(_menu3clnsub_1, CType(1, Short))
		Me.menu3clnsub.SetIndex(_menu3clnsub_2, CType(2, Short))
		Me.menu3clnsub.SetIndex(_menu3clnsub_3, CType(3, Short))
		Me.menu3clnsub.SetIndex(_menu3clnsub_4, CType(4, Short))
		Me.menu3clnsub.SetIndex(_menu3clnsub_5, CType(5, Short))
		Me.menu3ctllock.SetIndex(_menu3ctllock_0, CType(0, Short))
		Me.menu3ctllock.SetIndex(_menu3ctllock_1, CType(1, Short))
		Me.menu3ctllock.SetIndex(_menu3ctllock_2, CType(2, Short))
		Me.menu3ctlwinexit.SetIndex(_menu3ctlwinexit_0, CType(0, Short))
		Me.menu3ctlwinexit.SetIndex(_menu3ctlwinexit_1, CType(1, Short))
		Me.menu3ctlwinexit.SetIndex(_menu3ctlwinexit_2, CType(2, Short))
		Me.menu3ctlwinexit.SetIndex(_menu3ctlwinexit_3, CType(3, Short))
		Me.menu4envsub.SetIndex(_menu4envsub_0, CType(0, Short))
		Me.menu4envsub.SetIndex(_menu4envsub_1, CType(1, Short))
		Me.menu4mon.SetIndex(_menu4mon_0, CType(0, Short))
		Me.menu4mon.SetIndex(_menu4mon_1, CType(1, Short))
		Me.menu4mon.SetIndex(_menu4mon_2, CType(2, Short))
		Me.menu4mon.SetIndex(_menu4mon_3, CType(3, Short))
		Me.pmenu1clnsub.SetIndex(_pmenu1clnsub_0, CType(0, Short))
		Me.pmenu1clnsub.SetIndex(_pmenu1clnsub_1, CType(1, Short))
		Me.pmenu1clnsub.SetIndex(_pmenu1clnsub_2, CType(2, Short))
		Me.pmenu1clnsub.SetIndex(_pmenu1clnsub_3, CType(3, Short))
		Me.pmenu1clnsub.SetIndex(_pmenu1clnsub_4, CType(4, Short))
		Me.pmenu1clnsub.SetIndex(_pmenu1clnsub_5, CType(5, Short))
		Me.pmenu1ctlsub.SetIndex(_pmenu1ctlsub_0, CType(0, Short))
		Me.pmenu1ctlsub.SetIndex(_pmenu1ctlsub_1, CType(1, Short))
		Me.pmenu1ctlsub.SetIndex(_pmenu1ctlsub_2, CType(2, Short))
		Me.pmenu1ctlsub.SetIndex(_pmenu1ctlsub_3, CType(3, Short))
		CType(Me.pmenu1ctlsub, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.pmenu1clnsub, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.menu4mon, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.menu4envsub, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.menu3ctlwinexit, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.menu3ctllock, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.menu3clnsub, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.SubPagesMnu, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.SubPages, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.SpgInfoLblD, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.SpgInfoLblC, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.SpgInfoLblB, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.SpgInfoLblA, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.SpgInfoHdr, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.SerLbl, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.SerBtn, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.MainNoteBtn, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.MainDbtn, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Lv1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.SerLv1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.SerImgCb1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.SerImgCb2, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.MainSbar, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.ImgListSnm, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.ImgList16, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.ImgList32, System.ComponentModel.ISupportInitialize).EndInit()
		Me.menu1.Index = 0
		Me.menu3.Index = 1
		Me.menu4.Index = 2
		Me.menu5.Index = 3
		Me.menu2.Index = 4
		Me.popmenu1.Index = 5
		MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem(){Me.menu1, Me.menu3, Me.menu4, Me.menu5, Me.menu2, Me.popmenu1})
		Me.menu1penalaan.Index = 0
		Me.menu1sep1.Index = 1
		Me.menu1logout.Index = 2
		Me.menu2keluar.Index = 3
		menu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem(){Me.menu1penalaan, Me.menu1sep1, Me.menu1logout, Me.menu2keluar})
		Me.menu3announce.Index = 0
		Me.menu3ctl.Index = 1
		Me.menu3cln.Index = 2
		Me.menu3sep1.Index = 3
		Me.menu3AgMgr.Index = 4
		menu3.MenuItems.AddRange(New System.Windows.Forms.MenuItem(){Me.menu3announce, Me.menu3ctl, Me.menu3cln, Me.menu3sep1, Me.menu3AgMgr})
		Me.menu3mesej.Index = 0
		Me.menu3tiker.Index = 1
		menu3announce.MenuItems.AddRange(New System.Windows.Forms.MenuItem(){Me.menu3mesej, Me.menu3tiker})
		Me._menu3ctllock_0.Index = 0
		Me._menu3ctllock_1.Index = 1
		Me._menu3ctllock_2.Index = 2
		Me.menu3sep2.Index = 3
		Me._menu3ctlwinexit_0.Index = 4
		Me._menu3ctlwinexit_1.Index = 5
		Me._menu3ctlwinexit_2.Index = 6
		Me._menu3ctlwinexit_3.Index = 7
		menu3ctl.MenuItems.AddRange(New System.Windows.Forms.MenuItem(){Me._menu3ctllock_0, Me._menu3ctllock_1, Me._menu3ctllock_2, Me.menu3sep2, Me._menu3ctlwinexit_0, Me._menu3ctlwinexit_1, Me._menu3ctlwinexit_2, Me._menu3ctlwinexit_3})
		Me._menu3clnsub_0.Index = 0
		Me._menu3clnsub_1.Index = 1
		Me._menu3clnsub_2.Index = 2
		Me._menu3clnsub_3.Index = 3
		Me._menu3clnsub_4.Index = 4
		Me._menu3clnsub_5.Index = 5
		menu3cln.MenuItems.AddRange(New System.Windows.Forms.MenuItem(){Me._menu3clnsub_0, Me._menu3clnsub_1, Me._menu3clnsub_2, Me._menu3clnsub_3, Me._menu3clnsub_4, Me._menu3clnsub_5})
		Me._menu4mon_0.Index = 0
		Me._menu4mon_1.Index = 1
		Me._menu4mon_2.Index = 2
		Me._menu4mon_3.Index = 3
		Me.menu4sep1.Index = 4
		Me.menu4env.Index = 5
		menu4.MenuItems.AddRange(New System.Windows.Forms.MenuItem(){Me._menu4mon_0, Me._menu4mon_1, Me._menu4mon_2, Me._menu4mon_3, Me.menu4sep1, Me.menu4env})
		Me._menu4envsub_0.Index = 0
		Me._menu4envsub_1.Index = 1
		menu4env.MenuItems.AddRange(New System.Windows.Forms.MenuItem(){Me._menu4envsub_0, Me._menu4envsub_1})
		Me.menu4PosMgr.Index = 0
		Me.menu4Stat.Index = 1
		Me.menu5Console.Index = 2
		menu5.MenuItems.AddRange(New System.Windows.Forms.MenuItem(){Me.menu4PosMgr, Me.menu4Stat, Me.menu5Console})
		Me.menu2bantuan.Index = 0
		Me.menu2sep1.Index = 1
		Me.menu2aplikasi.Index = 2
		menu2.MenuItems.AddRange(New System.Windows.Forms.MenuItem(){Me.menu2bantuan, Me.menu2sep1, Me.menu2aplikasi})
		Me.pmenu1flog.Index = 0
		Me.pmenu1flout.Index = 1
		Me.psep2.Index = 2
		Me.pmenu1cancel.Index = 3
		Me.pmenu1trans.Index = 4
		Me.pmenu1terminal.Index = 5
		Me.psep1.Index = 6
		Me.pmenu1cln.Index = 7
		Me.pmenu1ctl.Index = 8
		popmenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem(){Me.pmenu1flog, Me.pmenu1flout, Me.psep2, Me.pmenu1cancel, Me.pmenu1trans, Me.pmenu1terminal, Me.psep1, Me.pmenu1cln, Me.pmenu1ctl})
		Me._pmenu1clnsub_0.Index = 0
		Me._pmenu1clnsub_1.Index = 1
		Me._pmenu1clnsub_2.Index = 2
		Me._pmenu1clnsub_3.Index = 3
		Me._pmenu1clnsub_4.Index = 4
		Me._pmenu1clnsub_5.Index = 5
		pmenu1cln.MenuItems.AddRange(New System.Windows.Forms.MenuItem(){Me._pmenu1clnsub_0, Me._pmenu1clnsub_1, Me._pmenu1clnsub_2, Me._pmenu1clnsub_3, Me._pmenu1clnsub_4, Me._pmenu1clnsub_5})
		Me._pmenu1ctlsub_0.Index = 0
		Me._pmenu1ctlsub_1.Index = 1
		Me._pmenu1ctlsub_2.Index = 2
		Me._pmenu1ctlsub_3.Index = 3
		pmenu1ctl.MenuItems.AddRange(New System.Windows.Forms.MenuItem(){Me._pmenu1ctlsub_0, Me._pmenu1ctlsub_1, Me._pmenu1ctlsub_2, Me._pmenu1ctlsub_3})
		Me.Menu = MainMenu1
	End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As FrmMain
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As FrmMain
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New FrmMain()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	Public NegateX As Integer 'beza x saiz untuk Lv
	Public NegateY As Integer 'beza y saiz untuk Lv
	Public NegateXtmp As Integer
	Public NegateYtmp As Integer
	
	Private SerTotal As Double
	
	
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Form Initialize
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	'UPGRADE_NOTE: Form_Initialize was upgraded to Form_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
	Private Sub Form_Initialize_Renamed()
		FrmMain.DefInstance.Height = VB6.TwipsToPixelsY(8640)
		Call LayOutMeasure()
		Call CbFrmMetricLoad(FrmMain.DefInstance)
	End Sub
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Form Makeup
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Private Sub FrmMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'UPGRADE_WARNING: Couldn't resolve default property of object Lv1.Icons. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Lv1.Icons = ImgList32.GetOCX
		'UPGRADE_WARNING: Couldn't resolve default property of object Lv1.SmallIcons. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Lv1.SmallIcons = ImgList16.GetOCX
		'UPGRADE_WARNING: Couldn't resolve default property of object Lv1.ColumnHeaderIcons. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Lv1.ColumnHeaderIcons = ImgList16.GetOCX
		Call LoadIconic()
	End Sub
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Form Query Unload
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	'UPGRADE_WARNING: Form event FrmMain.QueryUnload has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2065"'
	Private Sub FrmMain_Closing(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
		Dim Cancel As Short = eventArgs.Cancel
		'UPGRADE_ISSUE: Event parameter UnloadMode was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1057"'
		If UnloadMode = 0 Then
			Cancel = 1
			Call menu2keluar_Click(menu2keluar, New System.EventArgs())
		End If
		eventArgs.Cancel = Cancel
	End Sub
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Form Query Unload
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	'UPGRADE_NOTE: Form_Terminate was upgraded to Form_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
	'UPGRADE_WARNING: FrmMain event Form.Terminate has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2065"'
	Private Sub Form_Terminate_Renamed()
		Call UnloadIconic()
	End Sub
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Form Resizing
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	'UPGRADE_WARNING: Event FrmMain.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2075"'
	Private Sub FrmMain_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		On Error Resume Next
		
		If VB6.PixelsToTwipsY(FrmMain.DefInstance.Height) < 7000 Then FrmMain.DefInstance.Height = VB6.TwipsToPixelsY(8000)
		If VB6.PixelsToTwipsX(FrmMain.DefInstance.Width) < 11400 Then FrmMain.DefInstance.Width = VB6.TwipsToPixelsX(11400)
		Call LayOutSize()
	End Sub
	
	
	Private Sub MainDbtn_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MainDbtn.Click
		Dim Index As Short = MainDbtn.GetIndex(eventSender)
		Select Case Index
			Case 0
				Accessing((mSecurity.eCbAccessTo.Configuration))
			Case 1
				Accessing((mSecurity.eCbAccessTo.Statistic))
			Case 2
				'ConstructPage 1
			Case 3
				'ConstructPage 2
			Case 4
				'ConstructPage 0
			Case 5
				Call menu2keluar_Click(menu2keluar, New System.EventArgs())
		End Select
	End Sub
	
	'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
	' MAIN LAYOUT
	'
	'
	'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
	'[ PAGE HOLDER ]'
	Private Sub MainPhold_PageFlip(ByVal Sender As System.Object, ByVal e As PageHolder.PageFlipEventArgs) Handles MainPhold.PageFlip
		Dim Collapse As Boolean = e.Collapse
		If Collapse = True Then
			NegateYtmp = NegateY
			Lv1.Height = VB6.ToPixelsUserHeight(VB6.PixelsToTwipsY(MainPhold.Top), 7950, 0)
			NegateY = VB6.PixelsToTwipsY(FrmMain.DefInstance.Height) - VB6.FromPixelsUserHeight(Lv1.Height, 7950, 0)
		Else
			NegateY = NegateYtmp
			Lv1.Height = VB6.ToPixelsUserHeight(VB6.PixelsToTwipsY(FrmMain.DefInstance.Height) - NegateY, 7950, 0)
		End If
		menu4envsub(0).Checked = Collapse Xor True
		Call LayOutSize()
		SetSimpan("tooltab", CStr(FrmMain.DefInstance.menu4envsub(0).Checked))
	End Sub
	'[ PAGE DOCK ]'
	Private Sub MainPdock_PageFliped(ByVal Sender As System.Object, ByVal e As PageDock.PageFlipedEventArgs) Handles MainPdock.PageFliped
		Dim Flipped As Boolean = e.Flipped
		If Flipped = True Then
			NegateXtmp = NegateX
			Lv1.Width = VB6.ToPixelsUserWidth(VB6.PixelsToTwipsX(MainPdock.Left), 11340, 0)
			NegateX = VB6.PixelsToTwipsX(FrmMain.DefInstance.Width) - VB6.FromPixelsUserWidth(Lv1.Width, 11340, 0)
		Else
			NegateX = NegateXtmp
			Lv1.Width = VB6.ToPixelsUserWidth(VB6.PixelsToTwipsX(FrmMain.DefInstance.Width) - NegateX, 11340, 0)
		End If
		menu4envsub(1).Checked = Flipped Xor True
		Call LayOutSize()
		SetSimpan("dockbar", CStr(FrmMain.DefInstance.menu4envsub(1).Checked))
	End Sub
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Main Page - Note
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Private Sub MainNoteBtn_Click(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles MainNoteBtn.Click
		Dim Index As Short = MainNoteBtn.GetIndex(Sender)
		Select Case Index
			Case 0
				SetSimpan("mainnote", MainNote.Text)
			Case 1
				MainNote.Text = ""
				SetSimpan("mainnote", " ")
		End Select
	End Sub
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' DockBar | Menu
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	
	
	Public Sub menu5Console_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menu5Console.Popup
		menu5Console_Click(eventSender, eventArgs)
	End Sub
	Public Sub menu5Console_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menu5Console.Click
		FrmSysConsole.DefInstance.Show()
	End Sub
	
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' SubPages | Menu
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Private Sub SubPagesMnu_Click(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles SubPagesMnu.Click
		Dim Index As Short = SubPagesMnu.GetIndex(Sender)
		SubPages(Index).BringToFront()
	End Sub
	
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' ListView | DoubleClick
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Private Sub Lv1_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Lv1.DblClick
		If UniAgents.Count = 0 Then Exit Sub
		'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1039"'
		Load(FrmLogout)
		
		If SelSubItm(1) = VS(4) Then FrmLogout.DefInstance.ViewMode((0))
		If SelSubItm(1) = VS(3) Then FrmLogout.DefInstance.ViewMode((1))
		If SelSubItm(1) = VS(5) Then FrmLogout.DefInstance.ViewMode((2))
		FrmLogout.DefInstance.ShowDialog()
	End Sub
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' ListView | ItemClick
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Private Sub Lv1_ItemClick(ByVal eventSender As System.Object, ByVal eventArgs As AxMSComctlLib.ListViewEvents_ItemClickEvent) Handles Lv1.ItemClick
		Call UpdatePanel((eventArgs.Item.Text))
		Call UpdateStat(eventArgs.Item)
	End Sub
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' ListView | When the mouse goes up
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Private Sub Lv1_MouseUpEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSComctlLib.ListViewEvents_MouseUpEvent) Handles Lv1.MouseUpEvent
		If UniAgents.Count = 0 Then Exit Sub
		If eventArgs.Button = 2 Then
			'UPGRADE_ISSUE: Form method FrmMain.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			FrmMain.PopupMenu(popmenu1,  , eventArgs.x, eventArgs.y) : Exit Sub
		End If
	End Sub
	
	
	
	'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
	' SERVICES & MERCHANDISE
	'
	'
	'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
	
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Add Transaction
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Private Sub SerAddBtn_Click(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles SerAddBtn.Click
		Dim ret As Object
		Dim d As Object
		Dim Msg As Object
		Dim sItm As MSComctlLib.ListItem
		
		If SerTotal = 0 Then
			MsgBox(MB(14), MsgBoxStyle.Information, CbMsgWarn)
			Exit Sub
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Msg. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Msg = "Confirm transaction of " & Crnc & " " & VB6.Format(SerTotal, "#0.00")
		'UPGRADE_WARNING: Couldn't resolve default property of object Msg. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Msg = Msg & vbCrLf & "for the below items :-" & vbCrLf & vbCrLf
		For d = 1 To SerLv1.ListItems.Count
			sItm = SerLv1.ListItems(d)
			'UPGRADE_WARNING: Couldn't resolve default property of object d. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Msg. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			Msg = Msg & "  " & d & ". " & sItm.Text & vbTab & " =   " & sItm.SubItems(2) & vbCrLf
		Next d
		'UPGRADE_WARNING: Couldn't resolve default property of object ret. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ret = MsgBox(Msg, MsgBoxStyle.OKCancel, CbMsgApp)
		If ret = MsgBoxResult.OK Then
			SavePosTrans(SerLv1.ListItems)
			Call SerControls()
		End If
	End Sub
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Services Bayar - Change
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	'UPGRADE_WARNING: Event SerTxtBayar.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2075"'
	Private Sub SerTxtBayar_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles SerTxtBayar.TextChanged
		If Trim(SerTxtBayar.Text) <> "" And IsNumeric(SerTxtBayar.Text) = True Then
			SerTxtBaki.Text = Crnc & VB6.Format(CDbl(SerTxtBayar.Text) - SerTotal, "#0.00")
		Else
			SerTxtBaki.Text = ""
		End If
	End Sub
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Services Price PerItem - KeyUp
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Private Sub SerTxtPriItm_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles SerTxtPriItm.KeyUp
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		If KeyCode = System.Windows.Forms.Keys.Return Then
			If SerTxtPriItm.Text = "" Then Exit Sub
			If IsNumeric(SerTxtPriItm.Text) = False Then Exit Sub
			SerLv1.SelectedItem.SubItems(1) = VB6.Format(SerTxtPriItm.Text, "#0.00")
		End If
	End Sub
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Services ImageCombo1 - Change
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Private Sub SerImgCb1_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles SerImgCb1.Change
		If SerImgCb1.Text <> VS(1) Then
			Call SerControls(True, True)
			Call LoadPosItmCB(SerImgCb2, Mid(SerImgCb1.SelectedItem.Key, 2), ImgListSnm)
		Else
			Call SerControls(False, True)
		End If
	End Sub
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Services ImageCombo1 - Click
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Private Sub SerImgCb1_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles SerImgCb1.ClickEvent
		Call SerImgCb1_Change(SerImgCb1, New System.EventArgs())
	End Sub
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Services Lv1 - Total Item Price
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Private Sub SerLv1_ItemClick(ByVal eventSender As System.Object, ByVal eventArgs As AxMSComctlLib.ListViewEvents_ItemClickEvent) Handles SerLv1.ItemClick
		Dim ItemTotal As Double
		
		ItemTotal = CDbl(eventArgs.Item.SubItems(1)) * CShort(eventArgs.Item.SubItems(2))
		SerTxtJumlah.Text = Crnc & " " & VB6.Format(SerTotal, "#0.00")
		SerTxtTotalItm.Text = Crnc & " " & VB6.Format(ItemTotal, "#0.00")
		SerTxtPriItm.Text = eventArgs.Item.SubItems(1)
	End Sub
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Quantity scroller
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	'UPGRADE_NOTE: SerScroll1.Change was changed from an event to a procedure. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2010"'
	'UPGRADE_WARNING: VScrollBar event SerScroll1.Change has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2065"'
	Private Sub SerScroll1_Change(ByVal newScrollValue As Integer)
		SerTxtQty.Text = CStr(1000 - newScrollValue)
	End Sub
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	'
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Private Sub SerBtn_Click(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles SerBtn.Click
		Dim Index As Short = SerBtn.GetIndex(Sender)
		Dim g As Object
		Dim SCbItm As MSComctlLib.ComboItem
		Dim lvItm, fItm As MSComctlLib.ListItem
		
		Select Case Index
			Case 1
				SCbItm = SerImgCb2.SelectedItem
				If CDbl(SerTxtQty.Text) > 0 Then
					fItm = SerLv1.FindItem(SCbItm.Text)
					
					If fItm Is Nothing Then
						lvItm = SerLv1.ListItems.Add( , SCbItm.Key, SCbItm.Text)
						'UPGRADE_WARNING: Couldn't resolve default property of object SCbItm.Tag. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
						lvItm.SubItems(1) = SCbItm.Tag
						lvItm.SubItems(2) = SerTxtQty.Text
					Else
						fItm.SubItems(2) = SerTxtQty.Text
					End If
				Else
					MsgBox("Please enter quantity !", MsgBoxStyle.OKOnly, CbMsgWarn)
					SerTxtQty.SelectionStart = 1
					SerTxtQty.SelectionLength = Len(SerTxtQty.Text)
					Exit Sub
				End If
			Case 0
				If SerLv1.ListItems.Count = 0 Then Exit Sub
				SerLv1.ListItems.Remove((SerLv1.SelectedItem.Index))
		End Select
		
		'recalculate total
		SerTotal = 0
		For g = 1 To SerLv1.ListItems.Count
			SerTotal = SerTotal + (CDbl(SerLv1.ListItems(g).SubItems(1)) * CShort(SerLv1.ListItems(g).SubItems(2)))
		Next g
		SerTxtJumlah.Text = Crnc & " " & VB6.Format(SerTotal, "#0.00")
	End Sub
	
	
	
	'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
	' Section untuk Menu - SubMenu Object
	'
	'
	'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
	'[ menu penalaan ]
	Public Sub menu1penalaan_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menu1penalaan.Popup
		menu1penalaan_Click(eventSender, eventArgs)
	End Sub
	Public Sub menu1penalaan_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menu1penalaan.Click
		If Mid(CbUserAccess, 1, 1) = "0" Then MsgBox(MB(10), MsgBoxStyle.OKOnly, CbMsgWarn) : Exit Sub
		LogWorker(SL(5)) '((security log))
		FrmSet.DefInstance.ShowDialog()
	End Sub
	'[ menu about ]
	Public Sub menu2aplikasi_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menu2aplikasi.Popup
		menu2aplikasi_Click(eventSender, eventArgs)
	End Sub
	Public Sub menu2aplikasi_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menu2aplikasi.Click
		FrmAbout.DefInstance.ShowDialog()
	End Sub
	'[ menu LogOut ]
	Public Sub menu1logout_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menu1logout.Popup
		menu1logout_Click(eventSender, eventArgs)
	End Sub
	Public Sub menu1logout_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menu1logout.Click
		LogWorker(SL(2)) '((security log))
		FrmMain.DefInstance.Hide()
		FrmPass.DefInstance.Show()
	End Sub
	'[ menu Bantuan ]'
	Public Sub menu2bantuan_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menu2bantuan.Popup
		menu2bantuan_Click(eventSender, eventArgs)
	End Sub
	Public Sub menu2bantuan_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menu2bantuan.Click
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
		If Len(Dir(VB6.GetPath & "\help.htm", FileAttribute.Normal)) = 0 Then
			Call ShellExecute(Me.Handle.ToInt32, "open", "http://www.nematix.net", vbNullString, vbNullString, SW_NORMAL)
			Exit Sub
		End If
		Call ShellExecute(Me.Handle.ToInt32, "open", VB6.GetPath & "\help.htm", vbNullString, vbNullString, SW_NORMAL)
	End Sub
	'[ menu keluar ]'
	Public Sub menu2keluar_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menu2keluar.Popup
		menu2keluar_Click(eventSender, eventArgs)
	End Sub
	Public Sub menu2keluar_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menu2keluar.Click
		If CbDemoMode = True Then FrmSysDemo.DefInstance.ShowDialog()
		Call Keluar()
	End Sub
	
	'[ menu kunci ]'
	Public Sub menu3ctllock_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menu3ctllock.Popup
		menu3ctllock_Click(eventSender, eventArgs)
	End Sub
	Public Sub menu3ctllock_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menu3ctllock.Click
		Dim Index As Short = menu3ctllock.GetIndex(eventSender)
		Dim j As Short
		For j = 1 To UniAgents.Count
			Select Case Index
				Case 0
					UniAgents.Agents(j).NetSend("//kunci:1")
				Case 1
					If UniAgents.Agents(j).AgentStatus = VS(4) Then UniAgents.Agents(j).NetSend("//kunci:1")
				Case 2
					UniAgents.Agents(j).NetSend("//kunci:0")
			End Select
		Next j
	End Sub
	'[ menu mass reboot/shutdown ]'
	Public Sub menu3ctlwinexit_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menu3ctlwinexit.Popup
		menu3ctlwinexit_Click(eventSender, eventArgs)
	End Sub
	Public Sub menu3ctlwinexit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menu3ctlwinexit.Click
		Dim Index As Short = menu3ctlwinexit.GetIndex(eventSender)
		Dim uA As clsAgent
		Dim j As Integer
		
		For j = 1 To UniAgents.Count
			uA = UniAgents.Agents(j)
			Select Case Index
				Case 0
					uA.NetSend("//sdown:2")
				Case 1
					If uA.AgentStatus = VS(4) Then uA.NetSend("//sdown:2")
				Case 2
					uA.NetSend("//sdown:3")
				Case 3
					If uA.AgentStatus = VS(4) Then uA.NetSend("//sdown:3")
			End Select
		Next j
		Select Case Index
			Case 0 : LogWorker(SL(11)) '((security log))
			Case 2 : LogWorker(SL(10)) '((security log))
		End Select
	End Sub
	'[ menu pengumuman ]'
	Public Sub menu3mesej_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menu3mesej.Popup
		menu3mesej_Click(eventSender, eventArgs)
	End Sub
	Public Sub menu3mesej_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menu3mesej.Click
		Dim s_msg As String
		s_msg = MgoInpt.GetInput("Sila masukkan pengumuman anda", VisualSuite1.eButStyle.BtnClose)
		If Trim(s_msg) <> "" Then UniAgents.SendCommand("mesej:Server:" & s_msg)
	End Sub
	'[ menu hantar tiker ]'
	Public Sub menu3tiker_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menu3tiker.Popup
		menu3tiker_Click(eventSender, eventArgs)
	End Sub
	Public Sub menu3tiker_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menu3tiker.Click
		Dim Msg As Object
		Dim s_msg As String
		s_msg = MgoInpt.GetInput("Sila masukkan mesej tiker anda", VisualSuite1.eButStyle.BtnClose)
		'UPGRADE_WARNING: Couldn't resolve default property of object Msg. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If Trim(s_msg) <> "" Then UniAgents.SendCommand("tiker:" & Msg)
	End Sub
	'[ menu agent manager ]'
	Public Sub menu3AgMgr_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menu3AgMgr.Popup
		menu3AgMgr_Click(eventSender, eventArgs)
	End Sub
	Public Sub menu3AgMgr_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menu3AgMgr.Click
		FrmAgnMgr.DefInstance.Show()
	End Sub
	'[ menu monitoring ]'
	Public Sub menu4mon_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menu4mon.Popup
		menu4mon_Click(eventSender, eventArgs)
	End Sub
	Public Sub menu4mon_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menu4mon.Click
		Dim Index As Short = menu4mon.GetIndex(eventSender)
		ConstructPage(Index + 1)
	End Sub
	'[ menu view type ]
	Public Sub menu4envsub_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menu4envsub.Popup
		menu4envsub_Click(eventSender, eventArgs)
	End Sub
	Public Sub menu4envsub_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menu4envsub.Click
		Dim Index As Short = menu4envsub.GetIndex(eventSender)
		Select Case Index
			Case 0
				menu4envsub(0).Checked = menu4envsub(0).Checked Xor True
				MainPhold.PageCollapse = menu4envsub(0).Checked
			Case 1
				MainPdock.PageFlip = menu4envsub(1).Checked
				menu4envsub(1).Checked = menu4envsub(1).Checked Xor True
		End Select
	End Sub
	'[ menu open pos manager ]'
	Public Sub menu4posmgr_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menu4posmgr.Popup
		menu4posmgr_Click(eventSender, eventArgs)
	End Sub
	Public Sub menu4posmgr_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menu4posmgr.Click
		'FrmPosMg.Show
		Call LoadModule(mApplication.EnuModule.CafeSnmMgr)
	End Sub
	'[ menu open pos statistic ]'
	Public Sub menu4stat_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menu4stat.Popup
		menu4stat_Click(eventSender, eventArgs)
	End Sub
	Public Sub menu4stat_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menu4stat.Click
		Call MainDbtn_Click(MainDbtn.Item(1), New System.EventArgs())
	End Sub
	
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' POPUPMENU
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	'[ popmenu fast login ]'
	Public Sub pmenu1flog_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles pmenu1flog.Popup
		pmenu1flog_Click(eventSender, eventArgs)
	End Sub
	Public Sub pmenu1flog_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles pmenu1flog.Click
		If UniAgents.Count = 0 Then Exit Sub
		' Get station status
		If SelSubItm(1) = VS(3) Or SelSubItm(1) = VS(5) Then
			MsgBox(MB(21), MsgBoxStyle.Information, CbMsgWarn)
			Exit Sub
		End If
		FrmLogin.DefInstance.FastLogin()
		FrmLogin.DefInstance.Close()
	End Sub
	'[ popmenu fast logout ]'
	Public Sub pmenu1flout_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles pmenu1flout.Popup
		pmenu1flout_Click(eventSender, eventArgs)
	End Sub
	Public Sub pmenu1flout_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles pmenu1flout.Click
		If UniAgents.Count = 0 Then Exit Sub
		If SelSubItm(1) = VS(3) Then Call FrmLogout.DefInstance.ViewMode(1) : FrmLogout.DefInstance.Show()
		If SelSubItm(1) = VS(5) Then Call FrmLogout.DefInstance.ViewMode(2) : FrmLogout.DefInstance.Show()
	End Sub
	'[ popmenu transfer ]'
	Public Sub pmenu1trans_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles pmenu1trans.Popup
		pmenu1trans_Click(eventSender, eventArgs)
	End Sub
	Public Sub pmenu1trans_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles pmenu1trans.Click
		If AgentSel.AgentStatus = VS(4) Then Exit Sub
		If AgentSel.AgentStatus = VS(5) Then Exit Sub
		FrmAgnTrans.DefInstance.Show()
	End Sub
	'[ popmenu terminal ]'
	Public Sub pmenu1terminal_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles pmenu1terminal.Popup
		pmenu1terminal_Click(eventSender, eventArgs)
	End Sub
	Public Sub pmenu1terminal_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles pmenu1terminal.Click
		If UniAgents.Count = 0 Then Exit Sub
		FrmSysConsole.DefInstance.Show()
		FrmSysConsole.DefInstance.Activate()
	End Sub
	'[ popmenu cancel ]'
	Public Sub pmenu1cancel_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles pmenu1cancel.Popup
		pmenu1cancel_Click(eventSender, eventArgs)
	End Sub
	Public Sub pmenu1cancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles pmenu1cancel.Click
		Dim ret As Object
		Dim l_UsedTime As Integer
		l_UsedTime = AgentSel.CusGetTimeUseEx
		
		If UniAgents.Count = 0 Then Exit Sub
		If AgentSel.CustomerName = "" Then Exit Sub
		If l_UsedTime > 10 Then MsgBox(MB(23), MsgBoxStyle.Critical, CbMsgApp) : Exit Sub
		
		'UPGRADE_WARNING: Couldn't resolve default property of object ret. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ret = MsgBox(MB(22), MsgBoxStyle.OKCancel, CbMsgApp)
		LogWorker(SL(7)) '((security log))
		
		If ret = MsgBoxResult.OK Then
			AgentSel.CusStop()
			If SelTag <> "dump" Then
				AgentSel.NetSend("//kunci:1")
			End If
		End If
	End Sub
	'[ popmenu cleaning ]'
	Public Sub pmenu1clnsub_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles pmenu1clnsub.Popup
		pmenu1clnsub_Click(eventSender, eventArgs)
	End Sub
	Public Sub pmenu1clnsub_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles pmenu1clnsub.Click
		Dim Index As Short = pmenu1clnsub.GetIndex(eventSender)
		If UniAgents.Count = 0 Then Exit Sub
		If Index = 0 Then
			AgentSel.NetSend("//cleand:0")
		Else
			AgentSel.NetSend("//cleand:" & (Index - 1))
		End If
		AgentSel.AgentSmallIcon = "TerminalClean"
	End Sub
	'[ popmenu controlling ]'
	Public Sub pmenu1ctlsub_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles pmenu1ctlsub.Popup
		pmenu1ctlsub_Click(eventSender, eventArgs)
	End Sub
	Public Sub pmenu1ctlsub_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles pmenu1ctlsub.Click
		Dim Index As Short = pmenu1ctlsub.GetIndex(eventSender)
		If UniAgents.Count = 0 Then Exit Sub
		Select Case Index
			Case 0
				AgentSel.NetSend("//kunci:1")
			Case 1
				If AgentSel.AgentStatus = VS(3) Then
					AgentSel.NetSend("//kunci:0")
					LogWorker(SL(4)) '((security log))
				ElseIf AgentSel.AgentStatus = VS(4) Then 
					If CDbl(Mid(CbUserAccess, 3, 1)) = 0 Then
						MsgBox(MB(10), MsgBoxStyle.OKOnly, CbMsgWarn)
					Else
						AgentSel.NetSend("//kunci:0")
						LogWorker(SL(4)) '((security log))
					End If
				End If
			Case 2
				AgentSel.NetSend("//sdown:3")
				LogWorker(SL(8)) '((security log))
			Case 3
				AgentSel.NetSend("//sdown:2")
				LogWorker(SL(9)) '((security log))
		End Select
	End Sub
	
	
	
	'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
	' FUNCTION
	'
	'
	'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Layout Measuring
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Sub LayOutMeasure()
		NegateX = KiraBezaSaizX(FrmMain.DefInstance, Lv1)
		NegateY = KiraBezaSaizY(FrmMain.DefInstance, Lv1)
	End Sub
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Layout Resizing
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Sub LayOutSize()
		MainPhold.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(FrmMain.DefInstance.Height) - NegateY)
		MainPhold.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(FrmMain.DefInstance.Width) - NegateX)
		
		Lv1.Width = VB6.ToPixelsUserWidth(VB6.PixelsToTwipsX(FrmMain.DefInstance.Width) - NegateX, 11340, 0)
		Lv1.Height = VB6.ToPixelsUserHeight(VB6.PixelsToTwipsY(FrmMain.DefInstance.Height) - NegateY, 7950, 0)
	End Sub
	
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Service & Merchandise Reset
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Private Sub SerControls(Optional ByRef EnableCtl As Boolean = False, Optional ByRef AvoidChange As Boolean = False)
		If AvoidChange = False Then
			SerImgCb1.ComboItems.Item(VS(1)).Selected = True
			SerImgCb2.Text = VS(1)
		End If
		
		SerImgCb2.Enabled = EnableCtl
		SerTxtQty.Enabled = EnableCtl
		SerBtn(0).Enabled = EnableCtl
		SerBtn(1).Enabled = EnableCtl
		SerLv1.Enabled = EnableCtl
		SerScroll1.Enabled = EnableCtl
		
		SerLv1.ListItems.Clear()
		SerScroll1.Value = 999
		SerTxtQty.Text = CStr(1)
		SerTotal = 0
		SerTxtJumlah.Text = ""
		SerTxtBaki.Text = ""
		SerTxtBayar.Text = ""
		SerTxtTotalItm.Text = ""
		SerTxtPriItm.Text = ""
	End Sub
	Private Sub SerScroll1_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.ScrollEventArgs) Handles SerScroll1.Scroll
		Select Case eventArgs.type
			Case System.Windows.Forms.ScrollEventType.EndScroll
				SerScroll1_Change(eventArgs.newValue)
		End Select
	End Sub
End Class