Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class FrmSet
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
	Public WithEvents _SetLbl_1 As Label3D
	Public WithEvents _SetLbl_0 As Label3D
	Public WithEvents _MpassTxt_0 As System.Windows.Forms.TextBox
	Public WithEvents _MpassTxt_1 As System.Windows.Forms.TextBox
	Public WithEvents _MpassTxt_2 As System.Windows.Forms.TextBox
	Public WithEvents _MpassLbl_0 As System.Windows.Forms.Label
	Public WithEvents _MpassLbl_1 As System.Windows.Forms.Label
	Public WithEvents _MpassLbl_2 As System.Windows.Forms.Label
	Public WithEvents MpassFrame As System.Windows.Forms.GroupBox
	Public WithEvents _InfoTxt_1 As System.Windows.Forms.TextBox
	Public WithEvents _InfoTxt_0 As System.Windows.Forms.TextBox
	Public WithEvents _InfoTxt_2 As System.Windows.Forms.TextBox
	Public WithEvents _InfoLbl_1 As System.Windows.Forms.Label
	Public WithEvents _InfoLbl_0 As System.Windows.Forms.Label
	Public WithEvents _InfoLbl_2 As System.Windows.Forms.Label
	Public WithEvents InfoFrame As System.Windows.Forms.GroupBox
	Public WithEvents NetPortTxt As System.Windows.Forms.TextBox
	Public WithEvents _NetPassTxt_0 As System.Windows.Forms.TextBox
	Public WithEvents _NetPassTxt_1 As System.Windows.Forms.TextBox
	Public WithEvents _NetLbl_0 As System.Windows.Forms.Label
	Public WithEvents _NetLbl_1 As System.Windows.Forms.Label
	Public WithEvents _NetLbl_2 As System.Windows.Forms.Label
	Public WithEvents NetFrame As System.Windows.Forms.GroupBox
	Public WithEvents TxtNama As System.Windows.Forms.TextBox
	Public WithEvents TxtNombor As System.Windows.Forms.TextBox
	Public WithEvents RegBtn As XpButton
	Public WithEvents _RegLbl_0 As System.Windows.Forms.Label
	Public WithEvents _RegLbl_1 As System.Windows.Forms.Label
	Public WithEvents IDFrame As System.Windows.Forms.GroupBox
	Public WithEvents _SSTab1_TabPage0 As System.Windows.Forms.TabPage
	Public WithEvents _OvhTxt_2 As System.Windows.Forms.TextBox
	Public WithEvents _OvhTxt_1 As System.Windows.Forms.TextBox
	Public WithEvents _OvhTxt_0 As System.Windows.Forms.TextBox
	Public WithEvents _OverSesTxt_0 As System.Windows.Forms.TextBox
	Public WithEvents _OverSesTxt_1 As System.Windows.Forms.TextBox
	Public WithEvents OverCmb1 As System.Windows.Forms.ComboBox
	Public WithEvents OverLb3 As System.Windows.Forms.Label
	Public WithEvents OverLb2 As System.Windows.Forms.Label
	Public WithEvents OverLb1 As System.Windows.Forms.Label
	Public WithEvents OverHdr1 As System.Windows.Forms.Label
	Public WithEvents OverHdr2 As System.Windows.Forms.Label
	Public WithEvents OverLb4 As System.Windows.Forms.Label
	Public WithEvents _Label1_0 As System.Windows.Forms.Label
	Public WithEvents OverFrame As System.Windows.Forms.GroupBox
	Public WithEvents PriChk1 As System.Windows.Forms.CheckBox
	Public WithEvents _PriTxt1_0 As System.Windows.Forms.TextBox
	Public WithEvents _PriTxt1_1 As System.Windows.Forms.TextBox
	Public WithEvents _PriPmTxt_1 As System.Windows.Forms.TextBox
	Public WithEvents _PriPmTxt_0 As System.Windows.Forms.TextBox
	Public WithEvents _PriBtn_0 As XpButton
	Public WithEvents PriuLine1 As Line3D
	Public WithEvents PriLV1 As AxMSComctlLib.AxListView
	Public WithEvents _PriBtn_1 As XpButton
	Public WithEvents PriHrd2 As System.Windows.Forms.Label
	Public WithEvents PriHdr1 As System.Windows.Forms.Label
	Public WithEvents PriLBL1 As System.Windows.Forms.Label
	Public WithEvents PriLBL2 As System.Windows.Forms.Label
	Public WithEvents Label9 As System.Windows.Forms.Label
	Public WithEvents Label8 As System.Windows.Forms.Label
	Public WithEvents HargaFrame As System.Windows.Forms.GroupBox
	Public WithEvents _SSTab1_TabPage1 As System.Windows.Forms.TabPage
	Public WithEvents _EmpTxt_0 As System.Windows.Forms.TextBox
	Public WithEvents _EmpTxt_1 As System.Windows.Forms.TextBox
	Public WithEvents _EmpTxt_2 As System.Windows.Forms.TextBox
	Public WithEvents _EmpTxt_3 As System.Windows.Forms.TextBox
	Public WithEvents _Opt1_2 As System.Windows.Forms.CheckBox
	Public WithEvents _Opt1_1 As System.Windows.Forms.CheckBox
	Public WithEvents _Opt1_0 As System.Windows.Forms.CheckBox
	Public WithEvents Lv1 As AxMSComctlLib.AxListView
	Public WithEvents _EmpBtn_0 As XpButton
	Public WithEvents _EmpBtn_1 As XpButton
	Public WithEvents _EmpBtn_2 As XpButton
	Public WithEvents _EmpLbl_0 As System.Windows.Forms.Label
	Public WithEvents _EmpLbl_1 As System.Windows.Forms.Label
	Public WithEvents _EmpLbl_2 As System.Windows.Forms.Label
	Public WithEvents _EmpLbl_3 As System.Windows.Forms.Label
	Public WithEvents _EmpHdr_1 As System.Windows.Forms.Label
	Public WithEvents _EmpHdr_0 As System.Windows.Forms.Label
	Public WithEvents _SSTab1_TabPage2 As System.Windows.Forms.TabPage
	Public WithEvents _GwsOpt1_2 As System.Windows.Forms.CheckBox
	Public WithEvents _GwsOpt1_1 As System.Windows.Forms.CheckBox
	Public WithEvents _GwsOpt1_0 As System.Windows.Forms.CheckBox
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents _SSTab1_TabPage3 As System.Windows.Forms.TabPage
	Public WithEvents SSTab1 As System.Windows.Forms.TabControl
	Public WithEvents _SetBtn_1 As XpButton
	Public WithEvents _SetBtn_0 As XpButton
	Public WithEvents Image1 As System.Windows.Forms.PictureBox
	Public WithEvents EmpBtn As XpButtonArray
	Public WithEvents EmpHdr As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents EmpLbl As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents EmpTxt As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	Public WithEvents GwsOpt1 As Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray
	Public WithEvents InfoLbl As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents InfoTxt As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents MpassLbl As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents MpassTxt As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	Public WithEvents NetLbl As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents NetPassTxt As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	Public WithEvents Opt1 As Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray
	Public WithEvents OverSesTxt As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	Public WithEvents OvhTxt As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	Public WithEvents PriBtn As XpButtonArray
	Public WithEvents PriPmTxt As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	Public WithEvents PriTxt1 As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	Public WithEvents RegLbl As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents SetBtn As XpButtonArray
	Public WithEvents SetLbl As Label3DArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmSet))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me._SetLbl_1 = New Label3D
		Me._SetLbl_0 = New Label3D
		Me.SSTab1 = New System.Windows.Forms.TabControl
		Me._SSTab1_TabPage0 = New System.Windows.Forms.TabPage
		Me.MpassFrame = New System.Windows.Forms.GroupBox
		Me._MpassTxt_0 = New System.Windows.Forms.TextBox
		Me._MpassTxt_1 = New System.Windows.Forms.TextBox
		Me._MpassTxt_2 = New System.Windows.Forms.TextBox
		Me._MpassLbl_0 = New System.Windows.Forms.Label
		Me._MpassLbl_1 = New System.Windows.Forms.Label
		Me._MpassLbl_2 = New System.Windows.Forms.Label
		Me.InfoFrame = New System.Windows.Forms.GroupBox
		Me._InfoTxt_1 = New System.Windows.Forms.TextBox
		Me._InfoTxt_0 = New System.Windows.Forms.TextBox
		Me._InfoTxt_2 = New System.Windows.Forms.TextBox
		Me._InfoLbl_1 = New System.Windows.Forms.Label
		Me._InfoLbl_0 = New System.Windows.Forms.Label
		Me._InfoLbl_2 = New System.Windows.Forms.Label
		Me.NetFrame = New System.Windows.Forms.GroupBox
		Me.NetPortTxt = New System.Windows.Forms.TextBox
		Me._NetPassTxt_0 = New System.Windows.Forms.TextBox
		Me._NetPassTxt_1 = New System.Windows.Forms.TextBox
		Me._NetLbl_0 = New System.Windows.Forms.Label
		Me._NetLbl_1 = New System.Windows.Forms.Label
		Me._NetLbl_2 = New System.Windows.Forms.Label
		Me.IDFrame = New System.Windows.Forms.GroupBox
		Me.TxtNama = New System.Windows.Forms.TextBox
		Me.TxtNombor = New System.Windows.Forms.TextBox
		Me.RegBtn = New XpButton
		Me._RegLbl_0 = New System.Windows.Forms.Label
		Me._RegLbl_1 = New System.Windows.Forms.Label
		Me._SSTab1_TabPage1 = New System.Windows.Forms.TabPage
		Me.OverFrame = New System.Windows.Forms.GroupBox
		Me._OvhTxt_2 = New System.Windows.Forms.TextBox
		Me._OvhTxt_1 = New System.Windows.Forms.TextBox
		Me._OvhTxt_0 = New System.Windows.Forms.TextBox
		Me._OverSesTxt_0 = New System.Windows.Forms.TextBox
		Me._OverSesTxt_1 = New System.Windows.Forms.TextBox
		Me.OverCmb1 = New System.Windows.Forms.ComboBox
		Me.OverLb3 = New System.Windows.Forms.Label
		Me.OverLb2 = New System.Windows.Forms.Label
		Me.OverLb1 = New System.Windows.Forms.Label
		Me.OverHdr1 = New System.Windows.Forms.Label
		Me.OverHdr2 = New System.Windows.Forms.Label
		Me.OverLb4 = New System.Windows.Forms.Label
		Me._Label1_0 = New System.Windows.Forms.Label
		Me.HargaFrame = New System.Windows.Forms.GroupBox
		Me.PriChk1 = New System.Windows.Forms.CheckBox
		Me._PriTxt1_0 = New System.Windows.Forms.TextBox
		Me._PriTxt1_1 = New System.Windows.Forms.TextBox
		Me._PriPmTxt_1 = New System.Windows.Forms.TextBox
		Me._PriPmTxt_0 = New System.Windows.Forms.TextBox
		Me._PriBtn_0 = New XpButton
		Me.PriuLine1 = New Line3D
		Me.PriLV1 = New AxMSComctlLib.AxListView
		Me._PriBtn_1 = New XpButton
		Me.PriHrd2 = New System.Windows.Forms.Label
		Me.PriHdr1 = New System.Windows.Forms.Label
		Me.PriLBL1 = New System.Windows.Forms.Label
		Me.PriLBL2 = New System.Windows.Forms.Label
		Me.Label9 = New System.Windows.Forms.Label
		Me.Label8 = New System.Windows.Forms.Label
		Me._SSTab1_TabPage2 = New System.Windows.Forms.TabPage
		Me._EmpTxt_0 = New System.Windows.Forms.TextBox
		Me._EmpTxt_1 = New System.Windows.Forms.TextBox
		Me._EmpTxt_2 = New System.Windows.Forms.TextBox
		Me._EmpTxt_3 = New System.Windows.Forms.TextBox
		Me._Opt1_2 = New System.Windows.Forms.CheckBox
		Me._Opt1_1 = New System.Windows.Forms.CheckBox
		Me._Opt1_0 = New System.Windows.Forms.CheckBox
		Me.Lv1 = New AxMSComctlLib.AxListView
		Me._EmpBtn_0 = New XpButton
		Me._EmpBtn_1 = New XpButton
		Me._EmpBtn_2 = New XpButton
		Me._EmpLbl_0 = New System.Windows.Forms.Label
		Me._EmpLbl_1 = New System.Windows.Forms.Label
		Me._EmpLbl_2 = New System.Windows.Forms.Label
		Me._EmpLbl_3 = New System.Windows.Forms.Label
		Me._EmpHdr_1 = New System.Windows.Forms.Label
		Me._EmpHdr_0 = New System.Windows.Forms.Label
		Me._SSTab1_TabPage3 = New System.Windows.Forms.TabPage
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me._GwsOpt1_2 = New System.Windows.Forms.CheckBox
		Me._GwsOpt1_1 = New System.Windows.Forms.CheckBox
		Me._GwsOpt1_0 = New System.Windows.Forms.CheckBox
		Me._SetBtn_1 = New XpButton
		Me._SetBtn_0 = New XpButton
		Me.Image1 = New System.Windows.Forms.PictureBox
		Me.EmpBtn = New XpButtonArray(components)
		Me.EmpHdr = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.EmpLbl = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.EmpTxt = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(components)
		Me.GwsOpt1 = New Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray(components)
		Me.InfoLbl = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.InfoTxt = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(components)
		Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.MpassLbl = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.MpassTxt = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(components)
		Me.NetLbl = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.NetPassTxt = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(components)
		Me.Opt1 = New Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray(components)
		Me.OverSesTxt = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(components)
		Me.OvhTxt = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(components)
		Me.PriBtn = New XpButtonArray(components)
		Me.PriPmTxt = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(components)
		Me.PriTxt1 = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(components)
		Me.RegLbl = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.SetBtn = New XpButtonArray(components)
		Me.SetLbl = New Label3DArray(components)
		CType(Me.PriLV1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Lv1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.EmpBtn, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.EmpHdr, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.EmpLbl, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.EmpTxt, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.GwsOpt1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.InfoLbl, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.InfoTxt, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.MpassLbl, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.MpassTxt, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.NetLbl, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.NetPassTxt, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Opt1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.OverSesTxt, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.OvhTxt, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.PriBtn, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.PriPmTxt, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.PriTxt1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.RegLbl, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.SetBtn, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.SetLbl, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "CafeBonzer - Setting"
		Me.ClientSize = New System.Drawing.Size(556, 491)
		Me.Location = New System.Drawing.Point(3, 22)
		Me.ControlBox = False
		Me.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Icon = CType(resources.GetObject("FrmSet.Icon"), System.Drawing.Icon)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "FrmSet"
		Me._SetLbl_1.Size = New System.Drawing.Size(289, 15)
		Me._SetLbl_1.Location = New System.Drawing.Point(47, 467)
		Me._SetLbl_1.TabIndex = 80
		Me._SetLbl_1.Name = "_SetLbl_1"
		Me._SetLbl_0.Size = New System.Drawing.Size(122, 14)
		Me._SetLbl_0.Location = New System.Drawing.Point(46, 449)
		Me._SetLbl_0.TabIndex = 79
		Me._SetLbl_0.Name = "_SetLbl_0"
		Me.SSTab1.Size = New System.Drawing.Size(548, 440)
		Me.SSTab1.Location = New System.Drawing.Point(4, 3)
		Me.SSTab1.TabIndex = 0
		Me.SSTab1.Appearance = System.Windows.Forms.TabAppearance.FlatButtons
		Me.SSTab1.ItemSize = New System.Drawing.Size(42, 15)
		Me.SSTab1.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me.SSTab1.ForeColor = System.Drawing.Color.FromARGB(0, 0, 255)
		Me.SSTab1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Underline Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.SSTab1.Name = "SSTab1"
		Me._SSTab1_TabPage0.Text = "General"
		Me.MpassFrame.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me.MpassFrame.Text = "Master Password :"
		Me.MpassFrame.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.MpassFrame.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me.MpassFrame.Size = New System.Drawing.Size(248, 119)
		Me.MpassFrame.Location = New System.Drawing.Point(11, 307)
		Me.MpassFrame.TabIndex = 72
		Me.MpassFrame.Enabled = True
		Me.MpassFrame.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.MpassFrame.Visible = True
		Me.MpassFrame.Name = "MpassFrame"
		Me._MpassTxt_0.AutoSize = False
		Me._MpassTxt_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me._MpassTxt_0.BackColor = System.Drawing.Color.FromARGB(255, 224, 192)
		Me._MpassTxt_0.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._MpassTxt_0.Size = New System.Drawing.Size(131, 21)
		Me._MpassTxt_0.IMEMode = System.Windows.Forms.ImeMode.Disable
		Me._MpassTxt_0.Location = New System.Drawing.Point(103, 19)
		Me._MpassTxt_0.Maxlength = 20
		Me._MpassTxt_0.TabIndex = 74
		Me._MpassTxt_0.AcceptsReturn = True
		Me._MpassTxt_0.CausesValidation = True
		Me._MpassTxt_0.Enabled = True
		Me._MpassTxt_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._MpassTxt_0.HideSelection = True
		Me._MpassTxt_0.ReadOnly = False
		Me._MpassTxt_0.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._MpassTxt_0.MultiLine = False
		Me._MpassTxt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._MpassTxt_0.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._MpassTxt_0.TabStop = True
		Me._MpassTxt_0.Visible = True
		Me._MpassTxt_0.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._MpassTxt_0.Name = "_MpassTxt_0"
		Me._MpassTxt_1.AutoSize = False
		Me._MpassTxt_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me._MpassTxt_1.BackColor = System.Drawing.Color.FromARGB(255, 224, 192)
		Me._MpassTxt_1.Font = New System.Drawing.Font("Wingdings", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
		Me._MpassTxt_1.Size = New System.Drawing.Size(131, 21)
		Me._MpassTxt_1.IMEMode = System.Windows.Forms.ImeMode.Disable
		Me._MpassTxt_1.Location = New System.Drawing.Point(103, 51)
		Me._MpassTxt_1.Maxlength = 20
		Me._MpassTxt_1.PasswordChar = ChrW(108)
		Me._MpassTxt_1.TabIndex = 76
		Me._MpassTxt_1.AcceptsReturn = True
		Me._MpassTxt_1.CausesValidation = True
		Me._MpassTxt_1.Enabled = True
		Me._MpassTxt_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._MpassTxt_1.HideSelection = True
		Me._MpassTxt_1.ReadOnly = False
		Me._MpassTxt_1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._MpassTxt_1.MultiLine = False
		Me._MpassTxt_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._MpassTxt_1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._MpassTxt_1.TabStop = True
		Me._MpassTxt_1.Visible = True
		Me._MpassTxt_1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._MpassTxt_1.Name = "_MpassTxt_1"
		Me._MpassTxt_2.AutoSize = False
		Me._MpassTxt_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me._MpassTxt_2.BackColor = System.Drawing.Color.FromARGB(255, 224, 192)
		Me._MpassTxt_2.Font = New System.Drawing.Font("Wingdings", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
		Me._MpassTxt_2.Size = New System.Drawing.Size(131, 21)
		Me._MpassTxt_2.IMEMode = System.Windows.Forms.ImeMode.Disable
		Me._MpassTxt_2.Location = New System.Drawing.Point(103, 83)
		Me._MpassTxt_2.Maxlength = 20
		Me._MpassTxt_2.PasswordChar = ChrW(108)
		Me._MpassTxt_2.TabIndex = 78
		Me._MpassTxt_2.AcceptsReturn = True
		Me._MpassTxt_2.CausesValidation = True
		Me._MpassTxt_2.Enabled = True
		Me._MpassTxt_2.ForeColor = System.Drawing.SystemColors.WindowText
		Me._MpassTxt_2.HideSelection = True
		Me._MpassTxt_2.ReadOnly = False
		Me._MpassTxt_2.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._MpassTxt_2.MultiLine = False
		Me._MpassTxt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._MpassTxt_2.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._MpassTxt_2.TabStop = True
		Me._MpassTxt_2.Visible = True
		Me._MpassTxt_2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._MpassTxt_2.Name = "_MpassTxt_2"
		Me._MpassLbl_0.Text = "Username :"
		Me._MpassLbl_0.Size = New System.Drawing.Size(67, 13)
		Me._MpassLbl_0.Location = New System.Drawing.Point(14, 22)
		Me._MpassLbl_0.TabIndex = 73
		Me._MpassLbl_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._MpassLbl_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._MpassLbl_0.BackColor = System.Drawing.Color.Transparent
		Me._MpassLbl_0.Enabled = True
		Me._MpassLbl_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._MpassLbl_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._MpassLbl_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._MpassLbl_0.UseMnemonic = True
		Me._MpassLbl_0.Visible = True
		Me._MpassLbl_0.AutoSize = True
		Me._MpassLbl_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._MpassLbl_0.Name = "_MpassLbl_0"
		Me._MpassLbl_1.Text = "Password :"
		Me._MpassLbl_1.Size = New System.Drawing.Size(63, 13)
		Me._MpassLbl_1.Location = New System.Drawing.Point(14, 53)
		Me._MpassLbl_1.TabIndex = 75
		Me._MpassLbl_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._MpassLbl_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._MpassLbl_1.BackColor = System.Drawing.Color.Transparent
		Me._MpassLbl_1.Enabled = True
		Me._MpassLbl_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._MpassLbl_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._MpassLbl_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._MpassLbl_1.UseMnemonic = True
		Me._MpassLbl_1.Visible = True
		Me._MpassLbl_1.AutoSize = True
		Me._MpassLbl_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._MpassLbl_1.Name = "_MpassLbl_1"
		Me._MpassLbl_2.Text = "Retype :"
		Me._MpassLbl_2.Size = New System.Drawing.Size(49, 13)
		Me._MpassLbl_2.Location = New System.Drawing.Point(14, 85)
		Me._MpassLbl_2.TabIndex = 77
		Me._MpassLbl_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._MpassLbl_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._MpassLbl_2.BackColor = System.Drawing.Color.Transparent
		Me._MpassLbl_2.Enabled = True
		Me._MpassLbl_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._MpassLbl_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._MpassLbl_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._MpassLbl_2.UseMnemonic = True
		Me._MpassLbl_2.Visible = True
		Me._MpassLbl_2.AutoSize = True
		Me._MpassLbl_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._MpassLbl_2.Name = "_MpassLbl_2"
		Me.InfoFrame.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me.InfoFrame.Text = "Information :"
		Me.InfoFrame.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.InfoFrame.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me.InfoFrame.Size = New System.Drawing.Size(249, 169)
		Me.InfoFrame.Location = New System.Drawing.Point(11, 132)
		Me.InfoFrame.TabIndex = 57
		Me.InfoFrame.Enabled = True
		Me.InfoFrame.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.InfoFrame.Visible = True
		Me.InfoFrame.Name = "InfoFrame"
		Me._InfoTxt_1.AutoSize = False
		Me._InfoTxt_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me._InfoTxt_1.BackColor = System.Drawing.Color.White
		Me._InfoTxt_1.Size = New System.Drawing.Size(191, 21)
		Me._InfoTxt_1.Location = New System.Drawing.Point(46, 84)
		Me._InfoTxt_1.TabIndex = 61
		Me._InfoTxt_1.Text = "owner@cybercafe.com"
		Me._InfoTxt_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._InfoTxt_1.AcceptsReturn = True
		Me._InfoTxt_1.CausesValidation = True
		Me._InfoTxt_1.Enabled = True
		Me._InfoTxt_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._InfoTxt_1.HideSelection = True
		Me._InfoTxt_1.ReadOnly = False
		Me._InfoTxt_1.Maxlength = 0
		Me._InfoTxt_1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._InfoTxt_1.MultiLine = False
		Me._InfoTxt_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._InfoTxt_1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._InfoTxt_1.TabStop = True
		Me._InfoTxt_1.Visible = True
		Me._InfoTxt_1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._InfoTxt_1.Name = "_InfoTxt_1"
		Me._InfoTxt_0.AutoSize = False
		Me._InfoTxt_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me._InfoTxt_0.BackColor = System.Drawing.Color.White
		Me._InfoTxt_0.Size = New System.Drawing.Size(191, 21)
		Me._InfoTxt_0.Location = New System.Drawing.Point(46, 37)
		Me._InfoTxt_0.TabIndex = 59
		Me._InfoTxt_0.Text = "good cybercafe"
		Me._InfoTxt_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._InfoTxt_0.AcceptsReturn = True
		Me._InfoTxt_0.CausesValidation = True
		Me._InfoTxt_0.Enabled = True
		Me._InfoTxt_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._InfoTxt_0.HideSelection = True
		Me._InfoTxt_0.ReadOnly = False
		Me._InfoTxt_0.Maxlength = 0
		Me._InfoTxt_0.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._InfoTxt_0.MultiLine = False
		Me._InfoTxt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._InfoTxt_0.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._InfoTxt_0.TabStop = True
		Me._InfoTxt_0.Visible = True
		Me._InfoTxt_0.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._InfoTxt_0.Name = "_InfoTxt_0"
		Me._InfoTxt_2.AutoSize = False
		Me._InfoTxt_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me._InfoTxt_2.BackColor = System.Drawing.Color.White
		Me._InfoTxt_2.Size = New System.Drawing.Size(191, 21)
		Me._InfoTxt_2.Location = New System.Drawing.Point(45, 131)
		Me._InfoTxt_2.TabIndex = 63
		Me._InfoTxt_2.Text = "keep our pc clean"
		Me._InfoTxt_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._InfoTxt_2.AcceptsReturn = True
		Me._InfoTxt_2.CausesValidation = True
		Me._InfoTxt_2.Enabled = True
		Me._InfoTxt_2.ForeColor = System.Drawing.SystemColors.WindowText
		Me._InfoTxt_2.HideSelection = True
		Me._InfoTxt_2.ReadOnly = False
		Me._InfoTxt_2.Maxlength = 0
		Me._InfoTxt_2.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._InfoTxt_2.MultiLine = False
		Me._InfoTxt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._InfoTxt_2.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._InfoTxt_2.TabStop = True
		Me._InfoTxt_2.Visible = True
		Me._InfoTxt_2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._InfoTxt_2.Name = "_InfoTxt_2"
		Me._InfoLbl_1.Text = "E-mail :"
		Me._InfoLbl_1.Size = New System.Drawing.Size(45, 13)
		Me._InfoLbl_1.Location = New System.Drawing.Point(15, 67)
		Me._InfoLbl_1.TabIndex = 60
		Me._InfoLbl_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._InfoLbl_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._InfoLbl_1.BackColor = System.Drawing.Color.Transparent
		Me._InfoLbl_1.Enabled = True
		Me._InfoLbl_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._InfoLbl_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._InfoLbl_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._InfoLbl_1.UseMnemonic = True
		Me._InfoLbl_1.Visible = True
		Me._InfoLbl_1.AutoSize = True
		Me._InfoLbl_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._InfoLbl_1.Name = "_InfoLbl_1"
		Me._InfoLbl_0.Text = "Cybercafes Name :"
		Me._InfoLbl_0.Size = New System.Drawing.Size(111, 13)
		Me._InfoLbl_0.Location = New System.Drawing.Point(16, 19)
		Me._InfoLbl_0.TabIndex = 58
		Me._InfoLbl_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._InfoLbl_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._InfoLbl_0.BackColor = System.Drawing.Color.Transparent
		Me._InfoLbl_0.Enabled = True
		Me._InfoLbl_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._InfoLbl_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._InfoLbl_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._InfoLbl_0.UseMnemonic = True
		Me._InfoLbl_0.Visible = True
		Me._InfoLbl_0.AutoSize = True
		Me._InfoLbl_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._InfoLbl_0.Name = "_InfoLbl_0"
		Me._InfoLbl_2.Text = "Motto :"
		Me._InfoLbl_2.Size = New System.Drawing.Size(40, 13)
		Me._InfoLbl_2.Location = New System.Drawing.Point(14, 113)
		Me._InfoLbl_2.TabIndex = 62
		Me._InfoLbl_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._InfoLbl_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._InfoLbl_2.BackColor = System.Drawing.Color.Transparent
		Me._InfoLbl_2.Enabled = True
		Me._InfoLbl_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._InfoLbl_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._InfoLbl_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._InfoLbl_2.UseMnemonic = True
		Me._InfoLbl_2.Visible = True
		Me._InfoLbl_2.AutoSize = True
		Me._InfoLbl_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._InfoLbl_2.Name = "_InfoLbl_2"
		Me.NetFrame.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me.NetFrame.Text = "Networking :"
		Me.NetFrame.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.NetFrame.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me.NetFrame.Size = New System.Drawing.Size(269, 124)
		Me.NetFrame.Location = New System.Drawing.Point(268, 26)
		Me.NetFrame.TabIndex = 25
		Me.NetFrame.Enabled = True
		Me.NetFrame.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.NetFrame.Visible = True
		Me.NetFrame.Name = "NetFrame"
		Me.NetPortTxt.AutoSize = False
		Me.NetPortTxt.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.NetPortTxt.BackColor = System.Drawing.Color.White
		Me.NetPortTxt.Size = New System.Drawing.Size(105, 21)
		Me.NetPortTxt.Location = New System.Drawing.Point(153, 18)
		Me.NetPortTxt.Maxlength = 5
		Me.NetPortTxt.TabIndex = 27
		Me.NetPortTxt.Text = "56266"
		Me.ToolTip1.SetToolTip(Me.NetPortTxt, "Communication port between the server and the client, 56266 is the default value.")
		Me.NetPortTxt.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.NetPortTxt.AcceptsReturn = True
		Me.NetPortTxt.CausesValidation = True
		Me.NetPortTxt.Enabled = True
		Me.NetPortTxt.ForeColor = System.Drawing.SystemColors.WindowText
		Me.NetPortTxt.HideSelection = True
		Me.NetPortTxt.ReadOnly = False
		Me.NetPortTxt.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.NetPortTxt.MultiLine = False
		Me.NetPortTxt.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.NetPortTxt.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.NetPortTxt.TabStop = True
		Me.NetPortTxt.Visible = True
		Me.NetPortTxt.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.NetPortTxt.Name = "NetPortTxt"
		Me._NetPassTxt_0.AutoSize = False
		Me._NetPassTxt_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me._NetPassTxt_0.BackColor = System.Drawing.Color.White
		Me._NetPassTxt_0.Font = New System.Drawing.Font("Wingdings", 9!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
		Me._NetPassTxt_0.Size = New System.Drawing.Size(122, 21)
		Me._NetPassTxt_0.IMEMode = System.Windows.Forms.ImeMode.Disable
		Me._NetPassTxt_0.Location = New System.Drawing.Point(137, 52)
		Me._NetPassTxt_0.PasswordChar = ChrW(108)
		Me._NetPassTxt_0.TabIndex = 29
		Me.ToolTip1.SetToolTip(Me._NetPassTxt_0, "Default password for clients, if this field left empty, it will be set same as master password.")
		Me._NetPassTxt_0.AcceptsReturn = True
		Me._NetPassTxt_0.CausesValidation = True
		Me._NetPassTxt_0.Enabled = True
		Me._NetPassTxt_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._NetPassTxt_0.HideSelection = True
		Me._NetPassTxt_0.ReadOnly = False
		Me._NetPassTxt_0.Maxlength = 0
		Me._NetPassTxt_0.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._NetPassTxt_0.MultiLine = False
		Me._NetPassTxt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._NetPassTxt_0.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._NetPassTxt_0.TabStop = True
		Me._NetPassTxt_0.Visible = True
		Me._NetPassTxt_0.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._NetPassTxt_0.Name = "_NetPassTxt_0"
		Me._NetPassTxt_1.AutoSize = False
		Me._NetPassTxt_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me._NetPassTxt_1.BackColor = System.Drawing.Color.White
		Me._NetPassTxt_1.Font = New System.Drawing.Font("Wingdings", 9!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
		Me._NetPassTxt_1.Size = New System.Drawing.Size(122, 21)
		Me._NetPassTxt_1.IMEMode = System.Windows.Forms.ImeMode.Disable
		Me._NetPassTxt_1.Location = New System.Drawing.Point(137, 87)
		Me._NetPassTxt_1.PasswordChar = ChrW(108)
		Me._NetPassTxt_1.TabIndex = 31
		Me.ToolTip1.SetToolTip(Me._NetPassTxt_1, "Communication port between the server and the client, 56266 is the default value.")
		Me._NetPassTxt_1.AcceptsReturn = True
		Me._NetPassTxt_1.CausesValidation = True
		Me._NetPassTxt_1.Enabled = True
		Me._NetPassTxt_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._NetPassTxt_1.HideSelection = True
		Me._NetPassTxt_1.ReadOnly = False
		Me._NetPassTxt_1.Maxlength = 0
		Me._NetPassTxt_1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._NetPassTxt_1.MultiLine = False
		Me._NetPassTxt_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._NetPassTxt_1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._NetPassTxt_1.TabStop = True
		Me._NetPassTxt_1.Visible = True
		Me._NetPassTxt_1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._NetPassTxt_1.Name = "_NetPassTxt_1"
		Me._NetLbl_0.Text = "Local Port :"
		Me._NetLbl_0.Size = New System.Drawing.Size(65, 13)
		Me._NetLbl_0.Location = New System.Drawing.Point(14, 21)
		Me._NetLbl_0.TabIndex = 26
		Me._NetLbl_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._NetLbl_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._NetLbl_0.BackColor = System.Drawing.Color.Transparent
		Me._NetLbl_0.Enabled = True
		Me._NetLbl_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._NetLbl_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._NetLbl_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._NetLbl_0.UseMnemonic = True
		Me._NetLbl_0.Visible = True
		Me._NetLbl_0.AutoSize = True
		Me._NetLbl_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._NetLbl_0.Name = "_NetLbl_0"
		Me._NetLbl_1.Text = "Default Client Password :"
		Me._NetLbl_1.Size = New System.Drawing.Size(92, 28)
		Me._NetLbl_1.Location = New System.Drawing.Point(14, 50)
		Me._NetLbl_1.TabIndex = 28
		Me._NetLbl_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._NetLbl_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._NetLbl_1.BackColor = System.Drawing.Color.Transparent
		Me._NetLbl_1.Enabled = True
		Me._NetLbl_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._NetLbl_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._NetLbl_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._NetLbl_1.UseMnemonic = True
		Me._NetLbl_1.Visible = True
		Me._NetLbl_1.AutoSize = False
		Me._NetLbl_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._NetLbl_1.Name = "_NetLbl_1"
		Me._NetLbl_2.Text = "Retype Password :"
		Me._NetLbl_2.Size = New System.Drawing.Size(107, 13)
		Me._NetLbl_2.Location = New System.Drawing.Point(14, 90)
		Me._NetLbl_2.TabIndex = 30
		Me._NetLbl_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._NetLbl_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._NetLbl_2.BackColor = System.Drawing.Color.Transparent
		Me._NetLbl_2.Enabled = True
		Me._NetLbl_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._NetLbl_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._NetLbl_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._NetLbl_2.UseMnemonic = True
		Me._NetLbl_2.Visible = True
		Me._NetLbl_2.AutoSize = True
		Me._NetLbl_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._NetLbl_2.Name = "_NetLbl_2"
		Me.IDFrame.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me.IDFrame.Text = "Registered To :"
		Me.IDFrame.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.IDFrame.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me.IDFrame.Size = New System.Drawing.Size(248, 102)
		Me.IDFrame.Location = New System.Drawing.Point(12, 26)
		Me.IDFrame.TabIndex = 19
		Me.IDFrame.Enabled = True
		Me.IDFrame.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.IDFrame.Visible = True
		Me.IDFrame.Name = "IDFrame"
		Me.TxtNama.AutoSize = False
		Me.TxtNama.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.TxtNama.BackColor = System.Drawing.Color.FromARGB(255, 192, 192)
		Me.TxtNama.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.TxtNama.Size = New System.Drawing.Size(148, 21)
		Me.TxtNama.Location = New System.Drawing.Point(89, 16)
		Me.TxtNama.TabIndex = 21
		Me.TxtNama.AcceptsReturn = True
		Me.TxtNama.CausesValidation = True
		Me.TxtNama.Enabled = True
		Me.TxtNama.ForeColor = System.Drawing.SystemColors.WindowText
		Me.TxtNama.HideSelection = True
		Me.TxtNama.ReadOnly = False
		Me.TxtNama.Maxlength = 0
		Me.TxtNama.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.TxtNama.MultiLine = False
		Me.TxtNama.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.TxtNama.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.TxtNama.TabStop = True
		Me.TxtNama.Visible = True
		Me.TxtNama.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.TxtNama.Name = "TxtNama"
		Me.TxtNombor.AutoSize = False
		Me.TxtNombor.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.TxtNombor.BackColor = System.Drawing.Color.FromARGB(255, 192, 192)
		Me.TxtNombor.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.TxtNombor.Size = New System.Drawing.Size(148, 21)
		Me.TxtNombor.Location = New System.Drawing.Point(89, 43)
		Me.TxtNombor.TabIndex = 23
		Me.TxtNombor.AcceptsReturn = True
		Me.TxtNombor.CausesValidation = True
		Me.TxtNombor.Enabled = True
		Me.TxtNombor.ForeColor = System.Drawing.SystemColors.WindowText
		Me.TxtNombor.HideSelection = True
		Me.TxtNombor.ReadOnly = False
		Me.TxtNombor.Maxlength = 0
		Me.TxtNombor.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.TxtNombor.MultiLine = False
		Me.TxtNombor.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.TxtNombor.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.TxtNombor.TabStop = True
		Me.TxtNombor.Visible = True
		Me.TxtNombor.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.TxtNombor.Name = "TxtNombor"
		Me.RegBtn.Size = New System.Drawing.Size(72, 26)
		Me.RegBtn.Location = New System.Drawing.Point(165, 69)
		Me.RegBtn.TabIndex = 24
		Me.RegBtn.TX = "Verify"
		Me.RegBtn.ENAB = -1
		Me.RegBtn.COLTYPE = 1
		Me.RegBtn.FOCUSR = -1
		Me.RegBtn.BCOL = 12632256
		Me.RegBtn.BCOLO = 12632256
		Me.RegBtn.FCOL = 0
		Me.RegBtn.FCOLO = 0
		Me.RegBtn.MCOL = 16777215
		Me.RegBtn.MPTR = 1
		Me.RegBtn.MICON = 0
		Me.RegBtn.PICN = 0
		Me.RegBtn.UMCOL = -1
		Me.RegBtn.SOFT = 0
		Me.RegBtn.PICPOS = 0
		Me.RegBtn.NGREY = 0
		Me.RegBtn.FX = 0
		Me.RegBtn.HAND = 0
		Me.RegBtn.CHECK = 0
		Me.RegBtn.Name = "RegBtn"
		Me._RegLbl_0.Text = "Name :"
		Me._RegLbl_0.ForeColor = System.Drawing.Color.FromARGB(192, 0, 0)
		Me._RegLbl_0.Size = New System.Drawing.Size(42, 13)
		Me._RegLbl_0.Location = New System.Drawing.Point(12, 19)
		Me._RegLbl_0.TabIndex = 20
		Me._RegLbl_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._RegLbl_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._RegLbl_0.BackColor = System.Drawing.Color.Transparent
		Me._RegLbl_0.Enabled = True
		Me._RegLbl_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._RegLbl_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._RegLbl_0.UseMnemonic = True
		Me._RegLbl_0.Visible = True
		Me._RegLbl_0.AutoSize = True
		Me._RegLbl_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._RegLbl_0.Name = "_RegLbl_0"
		Me._RegLbl_1.Text = "Liscence  :"
		Me._RegLbl_1.ForeColor = System.Drawing.Color.FromARGB(192, 0, 0)
		Me._RegLbl_1.Size = New System.Drawing.Size(61, 13)
		Me._RegLbl_1.Location = New System.Drawing.Point(12, 46)
		Me._RegLbl_1.TabIndex = 22
		Me._RegLbl_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._RegLbl_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._RegLbl_1.BackColor = System.Drawing.Color.Transparent
		Me._RegLbl_1.Enabled = True
		Me._RegLbl_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._RegLbl_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._RegLbl_1.UseMnemonic = True
		Me._RegLbl_1.Visible = True
		Me._RegLbl_1.AutoSize = True
		Me._RegLbl_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._RegLbl_1.Name = "_RegLbl_1"
		Me._SSTab1_TabPage1.Text = "Financial"
		Me.OverFrame.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me.OverFrame.Text = "Overhead && Account Setting :"
		Me.OverFrame.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.OverFrame.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me.OverFrame.Size = New System.Drawing.Size(244, 403)
		Me.OverFrame.Location = New System.Drawing.Point(11, 26)
		Me.OverFrame.TabIndex = 5
		Me.OverFrame.Enabled = True
		Me.OverFrame.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.OverFrame.Visible = True
		Me.OverFrame.Name = "OverFrame"
		Me._OvhTxt_2.AutoSize = False
		Me._OvhTxt_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me._OvhTxt_2.BackColor = System.Drawing.Color.White
		Me._OvhTxt_2.Size = New System.Drawing.Size(62, 21)
		Me._OvhTxt_2.Location = New System.Drawing.Point(140, 97)
		Me._OvhTxt_2.TabIndex = 12
		Me._OvhTxt_2.Text = "20"
		Me._OvhTxt_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._OvhTxt_2.AcceptsReturn = True
		Me._OvhTxt_2.CausesValidation = True
		Me._OvhTxt_2.Enabled = True
		Me._OvhTxt_2.ForeColor = System.Drawing.SystemColors.WindowText
		Me._OvhTxt_2.HideSelection = True
		Me._OvhTxt_2.ReadOnly = False
		Me._OvhTxt_2.Maxlength = 0
		Me._OvhTxt_2.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._OvhTxt_2.MultiLine = False
		Me._OvhTxt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._OvhTxt_2.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._OvhTxt_2.TabStop = True
		Me._OvhTxt_2.Visible = True
		Me._OvhTxt_2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._OvhTxt_2.Name = "_OvhTxt_2"
		Me._OvhTxt_1.AutoSize = False
		Me._OvhTxt_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me._OvhTxt_1.BackColor = System.Drawing.Color.White
		Me._OvhTxt_1.Size = New System.Drawing.Size(62, 21)
		Me._OvhTxt_1.Location = New System.Drawing.Point(140, 71)
		Me._OvhTxt_1.TabIndex = 10
		Me._OvhTxt_1.Text = "350"
		Me._OvhTxt_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._OvhTxt_1.AcceptsReturn = True
		Me._OvhTxt_1.CausesValidation = True
		Me._OvhTxt_1.Enabled = True
		Me._OvhTxt_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._OvhTxt_1.HideSelection = True
		Me._OvhTxt_1.ReadOnly = False
		Me._OvhTxt_1.Maxlength = 0
		Me._OvhTxt_1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._OvhTxt_1.MultiLine = False
		Me._OvhTxt_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._OvhTxt_1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._OvhTxt_1.TabStop = True
		Me._OvhTxt_1.Visible = True
		Me._OvhTxt_1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._OvhTxt_1.Name = "_OvhTxt_1"
		Me._OvhTxt_0.AutoSize = False
		Me._OvhTxt_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me._OvhTxt_0.BackColor = System.Drawing.Color.White
		Me._OvhTxt_0.Size = New System.Drawing.Size(62, 21)
		Me._OvhTxt_0.Location = New System.Drawing.Point(140, 45)
		Me._OvhTxt_0.TabIndex = 8
		Me._OvhTxt_0.Text = "800"
		Me._OvhTxt_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._OvhTxt_0.AcceptsReturn = True
		Me._OvhTxt_0.CausesValidation = True
		Me._OvhTxt_0.Enabled = True
		Me._OvhTxt_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._OvhTxt_0.HideSelection = True
		Me._OvhTxt_0.ReadOnly = False
		Me._OvhTxt_0.Maxlength = 0
		Me._OvhTxt_0.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._OvhTxt_0.MultiLine = False
		Me._OvhTxt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._OvhTxt_0.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._OvhTxt_0.TabStop = True
		Me._OvhTxt_0.Visible = True
		Me._OvhTxt_0.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._OvhTxt_0.Name = "_OvhTxt_0"
		Me._OverSesTxt_0.AutoSize = False
		Me._OverSesTxt_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me._OverSesTxt_0.BackColor = System.Drawing.Color.White
		Me._OverSesTxt_0.Size = New System.Drawing.Size(30, 21)
		Me._OverSesTxt_0.Location = New System.Drawing.Point(77, 195)
		Me._OverSesTxt_0.Maxlength = 2
		Me._OverSesTxt_0.TabIndex = 15
		Me._OverSesTxt_0.Text = "12"
		Me._OverSesTxt_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._OverSesTxt_0.AcceptsReturn = True
		Me._OverSesTxt_0.CausesValidation = True
		Me._OverSesTxt_0.Enabled = True
		Me._OverSesTxt_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._OverSesTxt_0.HideSelection = True
		Me._OverSesTxt_0.ReadOnly = False
		Me._OverSesTxt_0.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._OverSesTxt_0.MultiLine = False
		Me._OverSesTxt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._OverSesTxt_0.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._OverSesTxt_0.TabStop = True
		Me._OverSesTxt_0.Visible = True
		Me._OverSesTxt_0.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._OverSesTxt_0.Name = "_OverSesTxt_0"
		Me._OverSesTxt_1.AutoSize = False
		Me._OverSesTxt_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me._OverSesTxt_1.BackColor = System.Drawing.Color.White
		Me._OverSesTxt_1.Size = New System.Drawing.Size(30, 21)
		Me._OverSesTxt_1.Location = New System.Drawing.Point(123, 195)
		Me._OverSesTxt_1.Maxlength = 2
		Me._OverSesTxt_1.TabIndex = 17
		Me._OverSesTxt_1.Text = "30"
		Me._OverSesTxt_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._OverSesTxt_1.AcceptsReturn = True
		Me._OverSesTxt_1.CausesValidation = True
		Me._OverSesTxt_1.Enabled = True
		Me._OverSesTxt_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._OverSesTxt_1.HideSelection = True
		Me._OverSesTxt_1.ReadOnly = False
		Me._OverSesTxt_1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._OverSesTxt_1.MultiLine = False
		Me._OverSesTxt_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._OverSesTxt_1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._OverSesTxt_1.TabStop = True
		Me._OverSesTxt_1.Visible = True
		Me._OverSesTxt_1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._OverSesTxt_1.Name = "_OverSesTxt_1"
		Me.OverCmb1.BackColor = System.Drawing.Color.White
		Me.OverCmb1.Size = New System.Drawing.Size(40, 21)
		Me.OverCmb1.Location = New System.Drawing.Point(162, 195)
		Me.OverCmb1.Items.AddRange(New Object(){"AM", "PM"})
		Me.OverCmb1.TabIndex = 18
		Me.OverCmb1.Text = "AM"
		Me.OverCmb1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.OverCmb1.CausesValidation = True
		Me.OverCmb1.Enabled = True
		Me.OverCmb1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.OverCmb1.IntegralHeight = True
		Me.OverCmb1.Cursor = System.Windows.Forms.Cursors.Default
		Me.OverCmb1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.OverCmb1.Sorted = False
		Me.OverCmb1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.OverCmb1.TabStop = True
		Me.OverCmb1.Visible = True
		Me.OverCmb1.Name = "OverCmb1"
		Me.OverLb3.Text = "Others :"
		Me.OverLb3.Size = New System.Drawing.Size(47, 13)
		Me.OverLb3.Location = New System.Drawing.Point(18, 98)
		Me.OverLb3.TabIndex = 11
		Me.OverLb3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.OverLb3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.OverLb3.BackColor = System.Drawing.Color.Transparent
		Me.OverLb3.Enabled = True
		Me.OverLb3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.OverLb3.Cursor = System.Windows.Forms.Cursors.Default
		Me.OverLb3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.OverLb3.UseMnemonic = True
		Me.OverLb3.Visible = True
		Me.OverLb3.AutoSize = True
		Me.OverLb3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.OverLb3.Name = "OverLb3"
		Me.OverLb2.Text = "Electric/Water Bills :"
		Me.OverLb2.Size = New System.Drawing.Size(116, 13)
		Me.OverLb2.Location = New System.Drawing.Point(17, 73)
		Me.OverLb2.TabIndex = 9
		Me.OverLb2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.OverLb2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.OverLb2.BackColor = System.Drawing.Color.Transparent
		Me.OverLb2.Enabled = True
		Me.OverLb2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.OverLb2.Cursor = System.Windows.Forms.Cursors.Default
		Me.OverLb2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.OverLb2.UseMnemonic = True
		Me.OverLb2.Visible = True
		Me.OverLb2.AutoSize = True
		Me.OverLb2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.OverLb2.Name = "OverLb2"
		Me.OverLb1.Text = "Premise Rental :"
		Me.OverLb1.Size = New System.Drawing.Size(95, 13)
		Me.OverLb1.Location = New System.Drawing.Point(17, 47)
		Me.OverLb1.TabIndex = 7
		Me.OverLb1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.OverLb1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.OverLb1.BackColor = System.Drawing.Color.Transparent
		Me.OverLb1.Enabled = True
		Me.OverLb1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.OverLb1.Cursor = System.Windows.Forms.Cursors.Default
		Me.OverLb1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.OverLb1.UseMnemonic = True
		Me.OverLb1.Visible = True
		Me.OverLb1.AutoSize = True
		Me.OverLb1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.OverLb1.Name = "OverLb1"
		Me.OverHdr1.BackColor = System.Drawing.Color.FromARGB(224, 224, 224)
		Me.OverHdr1.Text = " Monthly Overhead"
		Me.OverHdr1.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me.OverHdr1.Size = New System.Drawing.Size(225, 17)
		Me.OverHdr1.Location = New System.Drawing.Point(10, 20)
		Me.OverHdr1.TabIndex = 6
		Me.OverHdr1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.OverHdr1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.OverHdr1.Enabled = True
		Me.OverHdr1.Cursor = System.Windows.Forms.Cursors.Default
		Me.OverHdr1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.OverHdr1.UseMnemonic = True
		Me.OverHdr1.Visible = True
		Me.OverHdr1.AutoSize = False
		Me.OverHdr1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.OverHdr1.Name = "OverHdr1"
		Me.OverHdr2.BackColor = System.Drawing.Color.FromARGB(224, 224, 224)
		Me.OverHdr2.Text = " Account Settings"
		Me.OverHdr2.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me.OverHdr2.Size = New System.Drawing.Size(223, 17)
		Me.OverHdr2.Location = New System.Drawing.Point(10, 141)
		Me.OverHdr2.TabIndex = 13
		Me.OverHdr2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.OverHdr2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.OverHdr2.Enabled = True
		Me.OverHdr2.Cursor = System.Windows.Forms.Cursors.Default
		Me.OverHdr2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.OverHdr2.UseMnemonic = True
		Me.OverHdr2.Visible = True
		Me.OverHdr2.AutoSize = False
		Me.OverHdr2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.OverHdr2.Name = "OverHdr2"
		Me.OverLb4.Text = "Auto Session Close On :"
		Me.OverLb4.Size = New System.Drawing.Size(139, 13)
		Me.OverLb4.Location = New System.Drawing.Point(19, 172)
		Me.OverLb4.TabIndex = 14
		Me.OverLb4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.OverLb4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.OverLb4.BackColor = System.Drawing.Color.Transparent
		Me.OverLb4.Enabled = True
		Me.OverLb4.ForeColor = System.Drawing.SystemColors.ControlText
		Me.OverLb4.Cursor = System.Windows.Forms.Cursors.Default
		Me.OverLb4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.OverLb4.UseMnemonic = True
		Me.OverLb4.Visible = True
		Me.OverLb4.AutoSize = True
		Me.OverLb4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.OverLb4.Name = "OverLb4"
		Me._Label1_0.Text = ":"
		Me._Label1_0.Size = New System.Drawing.Size(5, 13)
		Me._Label1_0.Location = New System.Drawing.Point(112, 197)
		Me._Label1_0.TabIndex = 16
		Me._Label1_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_0.BackColor = System.Drawing.Color.Transparent
		Me._Label1_0.Enabled = True
		Me._Label1_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_0.UseMnemonic = True
		Me._Label1_0.Visible = True
		Me._Label1_0.AutoSize = True
		Me._Label1_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_0.Name = "_Label1_0"
		Me.HargaFrame.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me.HargaFrame.Text = "Pricing :"
		Me.HargaFrame.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.HargaFrame.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me.HargaFrame.Size = New System.Drawing.Size(275, 404)
		Me.HargaFrame.Location = New System.Drawing.Point(263, 26)
		Me.HargaFrame.TabIndex = 41
		Me.HargaFrame.Enabled = True
		Me.HargaFrame.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HargaFrame.Visible = True
		Me.HargaFrame.Name = "HargaFrame"
		Me.PriChk1.Text = "Round Up Price"
		Me.PriChk1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.PriChk1.Size = New System.Drawing.Size(148, 18)
		Me.PriChk1.Location = New System.Drawing.Point(9, 311)
		Me.PriChk1.TabIndex = 49
		Me.PriChk1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.PriChk1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.PriChk1.BackColor = System.Drawing.SystemColors.Control
		Me.PriChk1.CausesValidation = True
		Me.PriChk1.Enabled = True
		Me.PriChk1.Cursor = System.Windows.Forms.Cursors.Default
		Me.PriChk1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.PriChk1.Appearance = System.Windows.Forms.Appearance.Normal
		Me.PriChk1.TabStop = True
		Me.PriChk1.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.PriChk1.Visible = True
		Me.PriChk1.Name = "PriChk1"
		Me._PriTxt1_0.AutoSize = False
		Me._PriTxt1_0.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._PriTxt1_0.Size = New System.Drawing.Size(82, 21)
		Me._PriTxt1_0.Location = New System.Drawing.Point(118, 345)
		Me._PriTxt1_0.TabIndex = 52
		Me._PriTxt1_0.AcceptsReturn = True
		Me._PriTxt1_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._PriTxt1_0.BackColor = System.Drawing.SystemColors.Window
		Me._PriTxt1_0.CausesValidation = True
		Me._PriTxt1_0.Enabled = True
		Me._PriTxt1_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._PriTxt1_0.HideSelection = True
		Me._PriTxt1_0.ReadOnly = False
		Me._PriTxt1_0.Maxlength = 0
		Me._PriTxt1_0.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._PriTxt1_0.MultiLine = False
		Me._PriTxt1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._PriTxt1_0.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._PriTxt1_0.TabStop = True
		Me._PriTxt1_0.Visible = True
		Me._PriTxt1_0.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._PriTxt1_0.Name = "_PriTxt1_0"
		Me._PriTxt1_1.AutoSize = False
		Me._PriTxt1_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._PriTxt1_1.Size = New System.Drawing.Size(82, 21)
		Me._PriTxt1_1.Location = New System.Drawing.Point(118, 375)
		Me._PriTxt1_1.TabIndex = 54
		Me._PriTxt1_1.AcceptsReturn = True
		Me._PriTxt1_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._PriTxt1_1.BackColor = System.Drawing.SystemColors.Window
		Me._PriTxt1_1.CausesValidation = True
		Me._PriTxt1_1.Enabled = True
		Me._PriTxt1_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._PriTxt1_1.HideSelection = True
		Me._PriTxt1_1.ReadOnly = False
		Me._PriTxt1_1.Maxlength = 0
		Me._PriTxt1_1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._PriTxt1_1.MultiLine = False
		Me._PriTxt1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._PriTxt1_1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._PriTxt1_1.TabStop = True
		Me._PriTxt1_1.Visible = True
		Me._PriTxt1_1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._PriTxt1_1.Name = "_PriTxt1_1"
		Me._PriPmTxt_1.AutoSize = False
		Me._PriPmTxt_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me._PriPmTxt_1.BackColor = System.Drawing.Color.White
		Me._PriPmTxt_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._PriPmTxt_1.Size = New System.Drawing.Size(49, 21)
		Me._PriPmTxt_1.Location = New System.Drawing.Point(173, 43)
		Me._PriPmTxt_1.Maxlength = 5
		Me._PriPmTxt_1.TabIndex = 46
		Me.ToolTip1.SetToolTip(Me._PriPmTxt_1, "Initial price.")
		Me._PriPmTxt_1.AcceptsReturn = True
		Me._PriPmTxt_1.CausesValidation = True
		Me._PriPmTxt_1.Enabled = True
		Me._PriPmTxt_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._PriPmTxt_1.HideSelection = True
		Me._PriPmTxt_1.ReadOnly = False
		Me._PriPmTxt_1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._PriPmTxt_1.MultiLine = False
		Me._PriPmTxt_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._PriPmTxt_1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._PriPmTxt_1.TabStop = True
		Me._PriPmTxt_1.Visible = True
		Me._PriPmTxt_1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._PriPmTxt_1.Name = "_PriPmTxt_1"
		Me._PriPmTxt_0.AutoSize = False
		Me._PriPmTxt_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me._PriPmTxt_0.BackColor = System.Drawing.Color.White
		Me._PriPmTxt_0.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._PriPmTxt_0.Size = New System.Drawing.Size(38, 21)
		Me._PriPmTxt_0.Location = New System.Drawing.Point(57, 43)
		Me._PriPmTxt_0.Maxlength = 5
		Me._PriPmTxt_0.TabIndex = 44
		Me.ToolTip1.SetToolTip(Me._PriPmTxt_0, "Example : 0.03 for 3 cent per minute.")
		Me._PriPmTxt_0.AcceptsReturn = True
		Me._PriPmTxt_0.CausesValidation = True
		Me._PriPmTxt_0.Enabled = True
		Me._PriPmTxt_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._PriPmTxt_0.HideSelection = True
		Me._PriPmTxt_0.ReadOnly = False
		Me._PriPmTxt_0.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._PriPmTxt_0.MultiLine = False
		Me._PriPmTxt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._PriPmTxt_0.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._PriPmTxt_0.TabStop = True
		Me._PriPmTxt_0.Visible = True
		Me._PriPmTxt_0.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._PriPmTxt_0.Name = "_PriPmTxt_0"
		Me._PriBtn_0.Size = New System.Drawing.Size(28, 24)
		Me._PriBtn_0.Location = New System.Drawing.Point(241, 373)
		Me._PriBtn_0.TabIndex = 55
		Me._PriBtn_0.TX = ""
		Me._PriBtn_0.ENAB = -1
		Me._PriBtn_0.COLTYPE = 1
		Me._PriBtn_0.FOCUSR = -1
		Me._PriBtn_0.BCOL = 12632256
		Me._PriBtn_0.BCOLO = 12632256
		Me._PriBtn_0.FCOL = 0
		Me._PriBtn_0.FCOLO = 0
		Me._PriBtn_0.MCOL = 16777215
		Me._PriBtn_0.MPTR = 1
		Me._PriBtn_0.MICON = 0
		Me._PriBtn_0.PICN = 0
		Me._PriBtn_0.UMCOL = -1
		Me._PriBtn_0.SOFT = 0
		Me._PriBtn_0.PICPOS = 0
		Me._PriBtn_0.NGREY = 0
		Me._PriBtn_0.FX = 0
		Me._PriBtn_0.HAND = 0
		Me._PriBtn_0.CHECK = 0
		Me._PriBtn_0.Name = "_PriBtn_0"
		Me.PriuLine1.Size = New System.Drawing.Size(266, 3)
		Me.PriuLine1.Location = New System.Drawing.Point(4, 334)
		Me.PriuLine1.TabIndex = 50
		Me.PriuLine1.horizon = -1
		Me.PriuLine1.Name = "PriuLine1"
		PriLV1.OcxState = CType(resources.GetObject("PriLV1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.PriLV1.Size = New System.Drawing.Size(243, 117)
		Me.PriLV1.Location = New System.Drawing.Point(17, 117)
		Me.PriLV1.TabIndex = 48
		Me.PriLV1.Name = "PriLV1"
		Me._PriBtn_1.Size = New System.Drawing.Size(28, 24)
		Me._PriBtn_1.Location = New System.Drawing.Point(212, 373)
		Me._PriBtn_1.TabIndex = 56
		Me._PriBtn_1.TX = ""
		Me._PriBtn_1.ENAB = -1
		Me._PriBtn_1.COLTYPE = 1
		Me._PriBtn_1.FOCUSR = -1
		Me._PriBtn_1.BCOL = 12632256
		Me._PriBtn_1.BCOLO = 12632256
		Me._PriBtn_1.FCOL = 0
		Me._PriBtn_1.FCOLO = 0
		Me._PriBtn_1.MCOL = 16777215
		Me._PriBtn_1.MPTR = 1
		Me._PriBtn_1.MICON = 0
		Me._PriBtn_1.PICN = 0
		Me._PriBtn_1.UMCOL = -1
		Me._PriBtn_1.SOFT = 0
		Me._PriBtn_1.PICPOS = 0
		Me._PriBtn_1.NGREY = 0
		Me._PriBtn_1.FX = 0
		Me._PriBtn_1.HAND = 0
		Me._PriBtn_1.CHECK = 0
		Me._PriBtn_1.Name = "_PriBtn_1"
		Me.PriHrd2.BackColor = System.Drawing.Color.FromARGB(224, 224, 224)
		Me.PriHrd2.Text = " Additional Pricing"
		Me.PriHrd2.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me.PriHrd2.Size = New System.Drawing.Size(255, 17)
		Me.PriHrd2.Location = New System.Drawing.Point(10, 89)
		Me.PriHrd2.TabIndex = 47
		Me.PriHrd2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.PriHrd2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.PriHrd2.Enabled = True
		Me.PriHrd2.Cursor = System.Windows.Forms.Cursors.Default
		Me.PriHrd2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.PriHrd2.UseMnemonic = True
		Me.PriHrd2.Visible = True
		Me.PriHrd2.AutoSize = False
		Me.PriHrd2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.PriHrd2.Name = "PriHrd2"
		Me.PriHdr1.BackColor = System.Drawing.Color.FromARGB(224, 224, 224)
		Me.PriHdr1.Text = " Normal Pricing"
		Me.PriHdr1.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me.PriHdr1.Size = New System.Drawing.Size(255, 17)
		Me.PriHdr1.Location = New System.Drawing.Point(10, 17)
		Me.PriHdr1.TabIndex = 42
		Me.PriHdr1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.PriHdr1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.PriHdr1.Enabled = True
		Me.PriHdr1.Cursor = System.Windows.Forms.Cursors.Default
		Me.PriHdr1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.PriHdr1.UseMnemonic = True
		Me.PriHdr1.Visible = True
		Me.PriHdr1.AutoSize = False
		Me.PriHdr1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.PriHdr1.Name = "PriHdr1"
		Me.PriLBL1.Text = "Scheme Name :"
		Me.PriLBL1.Size = New System.Drawing.Size(93, 17)
		Me.PriLBL1.Location = New System.Drawing.Point(21, 347)
		Me.PriLBL1.TabIndex = 51
		Me.PriLBL1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.PriLBL1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.PriLBL1.BackColor = System.Drawing.Color.Transparent
		Me.PriLBL1.Enabled = True
		Me.PriLBL1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.PriLBL1.Cursor = System.Windows.Forms.Cursors.Default
		Me.PriLBL1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.PriLBL1.UseMnemonic = True
		Me.PriLBL1.Visible = True
		Me.PriLBL1.AutoSize = False
		Me.PriLBL1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.PriLBL1.Name = "PriLBL1"
		Me.PriLBL2.Text = "Price per Minute :"
		Me.PriLBL2.Size = New System.Drawing.Size(102, 17)
		Me.PriLBL2.Location = New System.Drawing.Point(11, 376)
		Me.PriLBL2.TabIndex = 53
		Me.PriLBL2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.PriLBL2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.PriLBL2.BackColor = System.Drawing.Color.Transparent
		Me.PriLBL2.Enabled = True
		Me.PriLBL2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.PriLBL2.Cursor = System.Windows.Forms.Cursors.Default
		Me.PriLBL2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.PriLBL2.UseMnemonic = True
		Me.PriLBL2.Visible = True
		Me.PriLBL2.AutoSize = False
		Me.PriLBL2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.PriLBL2.Name = "PriLBL2"
		Me.Label9.Text = "/Minute  +"
		Me.Label9.Size = New System.Drawing.Size(59, 13)
		Me.Label9.Location = New System.Drawing.Point(100, 46)
		Me.Label9.TabIndex = 45
		Me.Label9.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label9.BackColor = System.Drawing.Color.Transparent
		Me.Label9.Enabled = True
		Me.Label9.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label9.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label9.UseMnemonic = True
		Me.Label9.Visible = True
		Me.Label9.AutoSize = True
		Me.Label9.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label9.Name = "Label9"
		Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label8.Text = "RM :"
		Me.Label8.Size = New System.Drawing.Size(26, 13)
		Me.Label8.Location = New System.Drawing.Point(26, 46)
		Me.Label8.TabIndex = 43
		Me.Label8.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label8.BackColor = System.Drawing.Color.Transparent
		Me.Label8.Enabled = True
		Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label8.UseMnemonic = True
		Me.Label8.Visible = True
		Me.Label8.AutoSize = True
		Me.Label8.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label8.Name = "Label8"
		Me._SSTab1_TabPage2.Text = "Employee"
		Me._EmpTxt_0.AutoSize = False
		Me._EmpTxt_0.Size = New System.Drawing.Size(140, 21)
		Me._EmpTxt_0.Location = New System.Drawing.Point(129, 249)
		Me._EmpTxt_0.TabIndex = 35
		Me.ToolTip1.SetToolTip(Me._EmpTxt_0, "Enter worker name.")
		Me._EmpTxt_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._EmpTxt_0.AcceptsReturn = True
		Me._EmpTxt_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._EmpTxt_0.BackColor = System.Drawing.SystemColors.Window
		Me._EmpTxt_0.CausesValidation = True
		Me._EmpTxt_0.Enabled = True
		Me._EmpTxt_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._EmpTxt_0.HideSelection = True
		Me._EmpTxt_0.ReadOnly = False
		Me._EmpTxt_0.Maxlength = 0
		Me._EmpTxt_0.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._EmpTxt_0.MultiLine = False
		Me._EmpTxt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._EmpTxt_0.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._EmpTxt_0.TabStop = True
		Me._EmpTxt_0.Visible = True
		Me._EmpTxt_0.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._EmpTxt_0.Name = "_EmpTxt_0"
		Me._EmpTxt_1.AutoSize = False
		Me._EmpTxt_1.Size = New System.Drawing.Size(140, 21)
		Me._EmpTxt_1.Location = New System.Drawing.Point(129, 275)
		Me._EmpTxt_1.TabIndex = 39
		Me.ToolTip1.SetToolTip(Me._EmpTxt_1, "Enter the worker nick name.")
		Me._EmpTxt_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._EmpTxt_1.AcceptsReturn = True
		Me._EmpTxt_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._EmpTxt_1.BackColor = System.Drawing.SystemColors.Window
		Me._EmpTxt_1.CausesValidation = True
		Me._EmpTxt_1.Enabled = True
		Me._EmpTxt_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._EmpTxt_1.HideSelection = True
		Me._EmpTxt_1.ReadOnly = False
		Me._EmpTxt_1.Maxlength = 0
		Me._EmpTxt_1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._EmpTxt_1.MultiLine = False
		Me._EmpTxt_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._EmpTxt_1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._EmpTxt_1.TabStop = True
		Me._EmpTxt_1.Visible = True
		Me._EmpTxt_1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._EmpTxt_1.Name = "_EmpTxt_1"
		Me._EmpTxt_2.AutoSize = False
		Me._EmpTxt_2.Size = New System.Drawing.Size(140, 21)
		Me._EmpTxt_2.Location = New System.Drawing.Point(129, 302)
		Me._EmpTxt_2.TabIndex = 65
		Me.ToolTip1.SetToolTip(Me._EmpTxt_2, "Please enter the monthly salary.")
		Me._EmpTxt_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._EmpTxt_2.AcceptsReturn = True
		Me._EmpTxt_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._EmpTxt_2.BackColor = System.Drawing.SystemColors.Window
		Me._EmpTxt_2.CausesValidation = True
		Me._EmpTxt_2.Enabled = True
		Me._EmpTxt_2.ForeColor = System.Drawing.SystemColors.WindowText
		Me._EmpTxt_2.HideSelection = True
		Me._EmpTxt_2.ReadOnly = False
		Me._EmpTxt_2.Maxlength = 0
		Me._EmpTxt_2.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._EmpTxt_2.MultiLine = False
		Me._EmpTxt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._EmpTxt_2.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._EmpTxt_2.TabStop = True
		Me._EmpTxt_2.Visible = True
		Me._EmpTxt_2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._EmpTxt_2.Name = "_EmpTxt_2"
		Me._EmpTxt_3.AutoSize = False
		Me._EmpTxt_3.Size = New System.Drawing.Size(140, 21)
		Me._EmpTxt_3.Location = New System.Drawing.Point(129, 329)
		Me._EmpTxt_3.TabIndex = 66
		Me.ToolTip1.SetToolTip(Me._EmpTxt_3, "Enter the password for this worker.")
		Me._EmpTxt_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._EmpTxt_3.AcceptsReturn = True
		Me._EmpTxt_3.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._EmpTxt_3.BackColor = System.Drawing.SystemColors.Window
		Me._EmpTxt_3.CausesValidation = True
		Me._EmpTxt_3.Enabled = True
		Me._EmpTxt_3.ForeColor = System.Drawing.SystemColors.WindowText
		Me._EmpTxt_3.HideSelection = True
		Me._EmpTxt_3.ReadOnly = False
		Me._EmpTxt_3.Maxlength = 0
		Me._EmpTxt_3.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._EmpTxt_3.MultiLine = False
		Me._EmpTxt_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._EmpTxt_3.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._EmpTxt_3.TabStop = True
		Me._EmpTxt_3.Visible = True
		Me._EmpTxt_3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._EmpTxt_3.Name = "_EmpTxt_3"
		Me._Opt1_2.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me._Opt1_2.Text = "Allow unlock Client."
		Me._Opt1_2.ForeColor = System.Drawing.SystemColors.WindowText
		Me._Opt1_2.Size = New System.Drawing.Size(182, 15)
		Me._Opt1_2.Location = New System.Drawing.Point(313, 291)
		Me._Opt1_2.TabIndex = 64
		Me._Opt1_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Opt1_2.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._Opt1_2.CausesValidation = True
		Me._Opt1_2.Enabled = True
		Me._Opt1_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._Opt1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Opt1_2.Appearance = System.Windows.Forms.Appearance.Normal
		Me._Opt1_2.TabStop = True
		Me._Opt1_2.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._Opt1_2.Visible = True
		Me._Opt1_2.Name = "_Opt1_2"
		Me._Opt1_1.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me._Opt1_1.Text = "Allow access to Statistic."
		Me._Opt1_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._Opt1_1.Size = New System.Drawing.Size(182, 15)
		Me._Opt1_1.Location = New System.Drawing.Point(313, 269)
		Me._Opt1_1.TabIndex = 38
		Me._Opt1_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Opt1_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._Opt1_1.CausesValidation = True
		Me._Opt1_1.Enabled = True
		Me._Opt1_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Opt1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Opt1_1.Appearance = System.Windows.Forms.Appearance.Normal
		Me._Opt1_1.TabStop = True
		Me._Opt1_1.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._Opt1_1.Visible = True
		Me._Opt1_1.Name = "_Opt1_1"
		Me._Opt1_0.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me._Opt1_0.Text = "Allow access to Settings."
		Me._Opt1_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._Opt1_0.Size = New System.Drawing.Size(175, 15)
		Me._Opt1_0.Location = New System.Drawing.Point(313, 249)
		Me._Opt1_0.TabIndex = 36
		Me._Opt1_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Opt1_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._Opt1_0.CausesValidation = True
		Me._Opt1_0.Enabled = True
		Me._Opt1_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._Opt1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Opt1_0.Appearance = System.Windows.Forms.Appearance.Normal
		Me._Opt1_0.TabStop = True
		Me._Opt1_0.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._Opt1_0.Visible = True
		Me._Opt1_0.Name = "_Opt1_0"
		Lv1.OcxState = CType(resources.GetObject("Lv1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.Lv1.Size = New System.Drawing.Size(529, 181)
		Me.Lv1.Location = New System.Drawing.Point(9, 31)
		Me.Lv1.TabIndex = 32
		Me.Lv1.Name = "Lv1"
		Me._EmpBtn_0.Size = New System.Drawing.Size(28, 24)
		Me._EmpBtn_0.Location = New System.Drawing.Point(69, 362)
		Me._EmpBtn_0.TabIndex = 69
		Me.ToolTip1.SetToolTip(Me._EmpBtn_0, "Add new employee.")
		Me._EmpBtn_0.TX = ""
		Me._EmpBtn_0.ENAB = -1
		Me._EmpBtn_0.COLTYPE = 1
		Me._EmpBtn_0.FOCUSR = -1
		Me._EmpBtn_0.BCOL = 12632256
		Me._EmpBtn_0.BCOLO = 12632256
		Me._EmpBtn_0.FCOL = 0
		Me._EmpBtn_0.FCOLO = 0
		Me._EmpBtn_0.MCOL = 16777215
		Me._EmpBtn_0.MPTR = 1
		Me._EmpBtn_0.MICON = 0
		Me._EmpBtn_0.PICN = 0
		Me._EmpBtn_0.UMCOL = -1
		Me._EmpBtn_0.SOFT = 0
		Me._EmpBtn_0.PICPOS = 0
		Me._EmpBtn_0.NGREY = 0
		Me._EmpBtn_0.FX = 0
		Me._EmpBtn_0.HAND = 0
		Me._EmpBtn_0.CHECK = 0
		Me._EmpBtn_0.Name = "_EmpBtn_0"
		Me._EmpBtn_1.Size = New System.Drawing.Size(28, 24)
		Me._EmpBtn_1.Location = New System.Drawing.Point(40, 362)
		Me._EmpBtn_1.TabIndex = 70
		Me.ToolTip1.SetToolTip(Me._EmpBtn_1, "Delete selected employee.")
		Me._EmpBtn_1.TX = ""
		Me._EmpBtn_1.ENAB = -1
		Me._EmpBtn_1.COLTYPE = 1
		Me._EmpBtn_1.FOCUSR = -1
		Me._EmpBtn_1.BCOL = 12632256
		Me._EmpBtn_1.BCOLO = 12632256
		Me._EmpBtn_1.FCOL = 0
		Me._EmpBtn_1.FCOLO = 0
		Me._EmpBtn_1.MCOL = 16777215
		Me._EmpBtn_1.MPTR = 1
		Me._EmpBtn_1.MICON = 0
		Me._EmpBtn_1.PICN = 0
		Me._EmpBtn_1.UMCOL = -1
		Me._EmpBtn_1.SOFT = 0
		Me._EmpBtn_1.PICPOS = 0
		Me._EmpBtn_1.NGREY = 0
		Me._EmpBtn_1.FX = 0
		Me._EmpBtn_1.HAND = 0
		Me._EmpBtn_1.CHECK = 0
		Me._EmpBtn_1.Name = "_EmpBtn_1"
		Me._EmpBtn_2.Size = New System.Drawing.Size(28, 24)
		Me._EmpBtn_2.Location = New System.Drawing.Point(11, 362)
		Me._EmpBtn_2.TabIndex = 71
		Me.ToolTip1.SetToolTip(Me._EmpBtn_2, "Delete selected employee.")
		Me._EmpBtn_2.TX = ""
		Me._EmpBtn_2.ENAB = -1
		Me._EmpBtn_2.COLTYPE = 1
		Me._EmpBtn_2.FOCUSR = -1
		Me._EmpBtn_2.BCOL = 12632256
		Me._EmpBtn_2.BCOLO = 12632256
		Me._EmpBtn_2.FCOL = 0
		Me._EmpBtn_2.FCOLO = 0
		Me._EmpBtn_2.MCOL = 16777215
		Me._EmpBtn_2.MPTR = 1
		Me._EmpBtn_2.MICON = 0
		Me._EmpBtn_2.PICN = 0
		Me._EmpBtn_2.UMCOL = -1
		Me._EmpBtn_2.SOFT = 0
		Me._EmpBtn_2.PICPOS = 0
		Me._EmpBtn_2.NGREY = 0
		Me._EmpBtn_2.FX = 0
		Me._EmpBtn_2.HAND = 0
		Me._EmpBtn_2.CHECK = 0
		Me._EmpBtn_2.Name = "_EmpBtn_2"
		Me._EmpLbl_0.Text = "Worker Name :"
		Me._EmpLbl_0.Size = New System.Drawing.Size(96, 20)
		Me._EmpLbl_0.Location = New System.Drawing.Point(30, 252)
		Me._EmpLbl_0.TabIndex = 37
		Me._EmpLbl_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._EmpLbl_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._EmpLbl_0.BackColor = System.Drawing.Color.Transparent
		Me._EmpLbl_0.Enabled = True
		Me._EmpLbl_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._EmpLbl_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._EmpLbl_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._EmpLbl_0.UseMnemonic = True
		Me._EmpLbl_0.Visible = True
		Me._EmpLbl_0.AutoSize = False
		Me._EmpLbl_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._EmpLbl_0.Name = "_EmpLbl_0"
		Me._EmpLbl_1.Text = "UserName :"
		Me._EmpLbl_1.Size = New System.Drawing.Size(96, 20)
		Me._EmpLbl_1.Location = New System.Drawing.Point(32, 278)
		Me._EmpLbl_1.TabIndex = 40
		Me._EmpLbl_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._EmpLbl_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._EmpLbl_1.BackColor = System.Drawing.Color.Transparent
		Me._EmpLbl_1.Enabled = True
		Me._EmpLbl_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._EmpLbl_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._EmpLbl_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._EmpLbl_1.UseMnemonic = True
		Me._EmpLbl_1.Visible = True
		Me._EmpLbl_1.AutoSize = False
		Me._EmpLbl_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._EmpLbl_1.Name = "_EmpLbl_1"
		Me._EmpLbl_2.Text = "Salary :"
		Me._EmpLbl_2.Size = New System.Drawing.Size(96, 20)
		Me._EmpLbl_2.Location = New System.Drawing.Point(31, 331)
		Me._EmpLbl_2.TabIndex = 67
		Me._EmpLbl_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._EmpLbl_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._EmpLbl_2.BackColor = System.Drawing.Color.Transparent
		Me._EmpLbl_2.Enabled = True
		Me._EmpLbl_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._EmpLbl_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._EmpLbl_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._EmpLbl_2.UseMnemonic = True
		Me._EmpLbl_2.Visible = True
		Me._EmpLbl_2.AutoSize = False
		Me._EmpLbl_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._EmpLbl_2.Name = "_EmpLbl_2"
		Me._EmpLbl_3.Text = "Password :"
		Me._EmpLbl_3.Size = New System.Drawing.Size(96, 20)
		Me._EmpLbl_3.Location = New System.Drawing.Point(32, 305)
		Me._EmpLbl_3.TabIndex = 68
		Me._EmpLbl_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._EmpLbl_3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._EmpLbl_3.BackColor = System.Drawing.Color.Transparent
		Me._EmpLbl_3.Enabled = True
		Me._EmpLbl_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._EmpLbl_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._EmpLbl_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._EmpLbl_3.UseMnemonic = True
		Me._EmpLbl_3.Visible = True
		Me._EmpLbl_3.AutoSize = False
		Me._EmpLbl_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._EmpLbl_3.Name = "_EmpLbl_3"
		Me._EmpHdr_1.BackColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me._EmpHdr_1.Text = " Employee Information"
		Me._EmpHdr_1.ForeColor = System.Drawing.Color.White
		Me._EmpHdr_1.Size = New System.Drawing.Size(277, 17)
		Me._EmpHdr_1.Location = New System.Drawing.Point(9, 223)
		Me._EmpHdr_1.TabIndex = 34
		Me._EmpHdr_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._EmpHdr_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._EmpHdr_1.Enabled = True
		Me._EmpHdr_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._EmpHdr_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._EmpHdr_1.UseMnemonic = True
		Me._EmpHdr_1.Visible = True
		Me._EmpHdr_1.AutoSize = False
		Me._EmpHdr_1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._EmpHdr_1.Name = "_EmpHdr_1"
		Me._EmpHdr_0.BackColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me._EmpHdr_0.Text = " Employee Security Settings"
		Me._EmpHdr_0.ForeColor = System.Drawing.Color.White
		Me._EmpHdr_0.Size = New System.Drawing.Size(244, 17)
		Me._EmpHdr_0.Location = New System.Drawing.Point(294, 223)
		Me._EmpHdr_0.TabIndex = 33
		Me._EmpHdr_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._EmpHdr_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._EmpHdr_0.Enabled = True
		Me._EmpHdr_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._EmpHdr_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._EmpHdr_0.UseMnemonic = True
		Me._EmpHdr_0.Visible = True
		Me._EmpHdr_0.AutoSize = False
		Me._EmpHdr_0.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._EmpHdr_0.Name = "_EmpHdr_0"
		Me._SSTab1_TabPage3.Text = "Security"
		Me.Frame1.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me.Frame1.Text = "Global Employee Setting :"
		Me.Frame1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.ForeColor = System.Drawing.Color.FromARGB(0, 0, 128)
		Me.Frame1.Size = New System.Drawing.Size(532, 144)
		Me.Frame1.Location = New System.Drawing.Point(8, 24)
		Me.Frame1.TabIndex = 1
		Me.Frame1.Enabled = True
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Name = "Frame1"
		Me._GwsOpt1_2.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me._GwsOpt1_2.Text = "Allow cancel after 10 sec."
		Me._GwsOpt1_2.ForeColor = System.Drawing.SystemColors.WindowText
		Me._GwsOpt1_2.Size = New System.Drawing.Size(175, 15)
		Me._GwsOpt1_2.Location = New System.Drawing.Point(11, 63)
		Me._GwsOpt1_2.TabIndex = 4
		Me._GwsOpt1_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GwsOpt1_2.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._GwsOpt1_2.CausesValidation = True
		Me._GwsOpt1_2.Enabled = True
		Me._GwsOpt1_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._GwsOpt1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GwsOpt1_2.Appearance = System.Windows.Forms.Appearance.Normal
		Me._GwsOpt1_2.TabStop = True
		Me._GwsOpt1_2.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._GwsOpt1_2.Visible = True
		Me._GwsOpt1_2.Name = "_GwsOpt1_2"
		Me._GwsOpt1_1.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me._GwsOpt1_1.Text = "Allow changes price."
		Me._GwsOpt1_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._GwsOpt1_1.Size = New System.Drawing.Size(175, 15)
		Me._GwsOpt1_1.Location = New System.Drawing.Point(11, 41)
		Me._GwsOpt1_1.TabIndex = 3
		Me._GwsOpt1_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GwsOpt1_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._GwsOpt1_1.CausesValidation = True
		Me._GwsOpt1_1.Enabled = True
		Me._GwsOpt1_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._GwsOpt1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GwsOpt1_1.Appearance = System.Windows.Forms.Appearance.Normal
		Me._GwsOpt1_1.TabStop = True
		Me._GwsOpt1_1.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._GwsOpt1_1.Visible = True
		Me._GwsOpt1_1.Name = "_GwsOpt1_1"
		Me._GwsOpt1_0.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me._GwsOpt1_0.Text = "Log workers activities."
		Me._GwsOpt1_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._GwsOpt1_0.Size = New System.Drawing.Size(175, 15)
		Me._GwsOpt1_0.Location = New System.Drawing.Point(11, 21)
		Me._GwsOpt1_0.TabIndex = 2
		Me._GwsOpt1_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._GwsOpt1_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._GwsOpt1_0.CausesValidation = True
		Me._GwsOpt1_0.Enabled = True
		Me._GwsOpt1_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._GwsOpt1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._GwsOpt1_0.Appearance = System.Windows.Forms.Appearance.Normal
		Me._GwsOpt1_0.TabStop = True
		Me._GwsOpt1_0.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._GwsOpt1_0.Visible = True
		Me._GwsOpt1_0.Name = "_GwsOpt1_0"
		Me._SetBtn_1.Size = New System.Drawing.Size(36, 29)
		Me._SetBtn_1.Location = New System.Drawing.Point(515, 457)
		Me._SetBtn_1.TabIndex = 82
		Me.ToolTip1.SetToolTip(Me._SetBtn_1, "Save settings and exit.")
		Me._SetBtn_1.TX = ""
		Me._SetBtn_1.ENAB = -1
		Me._SetBtn_1.COLTYPE = 1
		Me._SetBtn_1.FOCUSR = -1
		Me._SetBtn_1.BCOL = 12632256
		Me._SetBtn_1.BCOLO = 12632256
		Me._SetBtn_1.FCOL = 0
		Me._SetBtn_1.FCOLO = 0
		Me._SetBtn_1.MCOL = 16777215
		Me._SetBtn_1.MPTR = 1
		Me._SetBtn_1.MICON = 0
		Me._SetBtn_1.PICN = 0
		Me._SetBtn_1.UMCOL = -1
		Me._SetBtn_1.SOFT = 0
		Me._SetBtn_1.PICPOS = 0
		Me._SetBtn_1.NGREY = 0
		Me._SetBtn_1.FX = 0
		Me._SetBtn_1.HAND = 0
		Me._SetBtn_1.CHECK = 0
		Me._SetBtn_1.Name = "_SetBtn_1"
		Me._SetBtn_0.Size = New System.Drawing.Size(36, 29)
		Me._SetBtn_0.Location = New System.Drawing.Point(478, 457)
		Me._SetBtn_0.TabIndex = 81
		Me.ToolTip1.SetToolTip(Me._SetBtn_0, "Cancel all settings and exit.")
		Me._SetBtn_0.TX = ""
		Me._SetBtn_0.ENAB = -1
		Me._SetBtn_0.COLTYPE = 1
		Me._SetBtn_0.FOCUSR = -1
		Me._SetBtn_0.BCOL = 12632256
		Me._SetBtn_0.BCOLO = 12632256
		Me._SetBtn_0.FCOL = 0
		Me._SetBtn_0.FCOLO = 0
		Me._SetBtn_0.MCOL = 16777215
		Me._SetBtn_0.MPTR = 1
		Me._SetBtn_0.MICON = 0
		Me._SetBtn_0.PICN = 0
		Me._SetBtn_0.UMCOL = -1
		Me._SetBtn_0.SOFT = 0
		Me._SetBtn_0.PICPOS = 0
		Me._SetBtn_0.NGREY = 0
		Me._SetBtn_0.FX = 0
		Me._SetBtn_0.HAND = 0
		Me._SetBtn_0.CHECK = 0
		Me._SetBtn_0.Name = "_SetBtn_0"
		Me.Image1.Size = New System.Drawing.Size(32, 32)
		Me.Image1.Location = New System.Drawing.Point(6, 450)
		Me.Image1.Image = CType(resources.GetObject("Image1.Image"), System.Drawing.Image)
		Me.Image1.Enabled = True
		Me.Image1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Image1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.Image1.Visible = True
		Me.Image1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Image1.Name = "Image1"
		Me.Controls.Add(_SetLbl_1)
		Me.Controls.Add(_SetLbl_0)
		Me.Controls.Add(SSTab1)
		Me.Controls.Add(_SetBtn_1)
		Me.Controls.Add(_SetBtn_0)
		Me.Controls.Add(Image1)
		Me.SSTab1.Controls.Add(_SSTab1_TabPage0)
		Me.SSTab1.Controls.Add(_SSTab1_TabPage1)
		Me.SSTab1.Controls.Add(_SSTab1_TabPage2)
		Me.SSTab1.Controls.Add(_SSTab1_TabPage3)
		Me._SSTab1_TabPage0.Controls.Add(MpassFrame)
		Me._SSTab1_TabPage0.Controls.Add(InfoFrame)
		Me._SSTab1_TabPage0.Controls.Add(NetFrame)
		Me._SSTab1_TabPage0.Controls.Add(IDFrame)
		Me.MpassFrame.Controls.Add(_MpassTxt_0)
		Me.MpassFrame.Controls.Add(_MpassTxt_1)
		Me.MpassFrame.Controls.Add(_MpassTxt_2)
		Me.MpassFrame.Controls.Add(_MpassLbl_0)
		Me.MpassFrame.Controls.Add(_MpassLbl_1)
		Me.MpassFrame.Controls.Add(_MpassLbl_2)
		Me.InfoFrame.Controls.Add(_InfoTxt_1)
		Me.InfoFrame.Controls.Add(_InfoTxt_0)
		Me.InfoFrame.Controls.Add(_InfoTxt_2)
		Me.InfoFrame.Controls.Add(_InfoLbl_1)
		Me.InfoFrame.Controls.Add(_InfoLbl_0)
		Me.InfoFrame.Controls.Add(_InfoLbl_2)
		Me.NetFrame.Controls.Add(NetPortTxt)
		Me.NetFrame.Controls.Add(_NetPassTxt_0)
		Me.NetFrame.Controls.Add(_NetPassTxt_1)
		Me.NetFrame.Controls.Add(_NetLbl_0)
		Me.NetFrame.Controls.Add(_NetLbl_1)
		Me.NetFrame.Controls.Add(_NetLbl_2)
		Me.IDFrame.Controls.Add(TxtNama)
		Me.IDFrame.Controls.Add(TxtNombor)
		Me.IDFrame.Controls.Add(RegBtn)
		Me.IDFrame.Controls.Add(_RegLbl_0)
		Me.IDFrame.Controls.Add(_RegLbl_1)
		Me._SSTab1_TabPage1.Controls.Add(OverFrame)
		Me._SSTab1_TabPage1.Controls.Add(HargaFrame)
		Me.OverFrame.Controls.Add(_OvhTxt_2)
		Me.OverFrame.Controls.Add(_OvhTxt_1)
		Me.OverFrame.Controls.Add(_OvhTxt_0)
		Me.OverFrame.Controls.Add(_OverSesTxt_0)
		Me.OverFrame.Controls.Add(_OverSesTxt_1)
		Me.OverFrame.Controls.Add(OverCmb1)
		Me.OverFrame.Controls.Add(OverLb3)
		Me.OverFrame.Controls.Add(OverLb2)
		Me.OverFrame.Controls.Add(OverLb1)
		Me.OverFrame.Controls.Add(OverHdr1)
		Me.OverFrame.Controls.Add(OverHdr2)
		Me.OverFrame.Controls.Add(OverLb4)
		Me.OverFrame.Controls.Add(_Label1_0)
		Me.HargaFrame.Controls.Add(PriChk1)
		Me.HargaFrame.Controls.Add(_PriTxt1_0)
		Me.HargaFrame.Controls.Add(_PriTxt1_1)
		Me.HargaFrame.Controls.Add(_PriPmTxt_1)
		Me.HargaFrame.Controls.Add(_PriPmTxt_0)
		Me.HargaFrame.Controls.Add(_PriBtn_0)
		Me.HargaFrame.Controls.Add(PriuLine1)
		Me.HargaFrame.Controls.Add(PriLV1)
		Me.HargaFrame.Controls.Add(_PriBtn_1)
		Me.HargaFrame.Controls.Add(PriHrd2)
		Me.HargaFrame.Controls.Add(PriHdr1)
		Me.HargaFrame.Controls.Add(PriLBL1)
		Me.HargaFrame.Controls.Add(PriLBL2)
		Me.HargaFrame.Controls.Add(Label9)
		Me.HargaFrame.Controls.Add(Label8)
		Me._SSTab1_TabPage2.Controls.Add(_EmpTxt_0)
		Me._SSTab1_TabPage2.Controls.Add(_EmpTxt_1)
		Me._SSTab1_TabPage2.Controls.Add(_EmpTxt_2)
		Me._SSTab1_TabPage2.Controls.Add(_EmpTxt_3)
		Me._SSTab1_TabPage2.Controls.Add(_Opt1_2)
		Me._SSTab1_TabPage2.Controls.Add(_Opt1_1)
		Me._SSTab1_TabPage2.Controls.Add(_Opt1_0)
		Me._SSTab1_TabPage2.Controls.Add(Lv1)
		Me._SSTab1_TabPage2.Controls.Add(_EmpBtn_0)
		Me._SSTab1_TabPage2.Controls.Add(_EmpBtn_1)
		Me._SSTab1_TabPage2.Controls.Add(_EmpBtn_2)
		Me._SSTab1_TabPage2.Controls.Add(_EmpLbl_0)
		Me._SSTab1_TabPage2.Controls.Add(_EmpLbl_1)
		Me._SSTab1_TabPage2.Controls.Add(_EmpLbl_2)
		Me._SSTab1_TabPage2.Controls.Add(_EmpLbl_3)
		Me._SSTab1_TabPage2.Controls.Add(_EmpHdr_1)
		Me._SSTab1_TabPage2.Controls.Add(_EmpHdr_0)
		Me._SSTab1_TabPage3.Controls.Add(Frame1)
		Me.Frame1.Controls.Add(_GwsOpt1_2)
		Me.Frame1.Controls.Add(_GwsOpt1_1)
		Me.Frame1.Controls.Add(_GwsOpt1_0)
		Me.EmpBtn.SetIndex(_EmpBtn_0, CType(0, Short))
		Me.EmpBtn.SetIndex(_EmpBtn_1, CType(1, Short))
		Me.EmpBtn.SetIndex(_EmpBtn_2, CType(2, Short))
		Me.EmpHdr.SetIndex(_EmpHdr_1, CType(1, Short))
		Me.EmpHdr.SetIndex(_EmpHdr_0, CType(0, Short))
		Me.EmpLbl.SetIndex(_EmpLbl_0, CType(0, Short))
		Me.EmpLbl.SetIndex(_EmpLbl_1, CType(1, Short))
		Me.EmpLbl.SetIndex(_EmpLbl_2, CType(2, Short))
		Me.EmpLbl.SetIndex(_EmpLbl_3, CType(3, Short))
		Me.EmpTxt.SetIndex(_EmpTxt_0, CType(0, Short))
		Me.EmpTxt.SetIndex(_EmpTxt_1, CType(1, Short))
		Me.EmpTxt.SetIndex(_EmpTxt_2, CType(2, Short))
		Me.EmpTxt.SetIndex(_EmpTxt_3, CType(3, Short))
		Me.GwsOpt1.SetIndex(_GwsOpt1_2, CType(2, Short))
		Me.GwsOpt1.SetIndex(_GwsOpt1_1, CType(1, Short))
		Me.GwsOpt1.SetIndex(_GwsOpt1_0, CType(0, Short))
		Me.InfoLbl.SetIndex(_InfoLbl_1, CType(1, Short))
		Me.InfoLbl.SetIndex(_InfoLbl_0, CType(0, Short))
		Me.InfoLbl.SetIndex(_InfoLbl_2, CType(2, Short))
		Me.InfoTxt.SetIndex(_InfoTxt_1, CType(1, Short))
		Me.InfoTxt.SetIndex(_InfoTxt_0, CType(0, Short))
		Me.InfoTxt.SetIndex(_InfoTxt_2, CType(2, Short))
		Me.Label1.SetIndex(_Label1_0, CType(0, Short))
		Me.MpassLbl.SetIndex(_MpassLbl_0, CType(0, Short))
		Me.MpassLbl.SetIndex(_MpassLbl_1, CType(1, Short))
		Me.MpassLbl.SetIndex(_MpassLbl_2, CType(2, Short))
		Me.MpassTxt.SetIndex(_MpassTxt_0, CType(0, Short))
		Me.MpassTxt.SetIndex(_MpassTxt_1, CType(1, Short))
		Me.MpassTxt.SetIndex(_MpassTxt_2, CType(2, Short))
		Me.NetLbl.SetIndex(_NetLbl_0, CType(0, Short))
		Me.NetLbl.SetIndex(_NetLbl_1, CType(1, Short))
		Me.NetLbl.SetIndex(_NetLbl_2, CType(2, Short))
		Me.NetPassTxt.SetIndex(_NetPassTxt_0, CType(0, Short))
		Me.NetPassTxt.SetIndex(_NetPassTxt_1, CType(1, Short))
		Me.Opt1.SetIndex(_Opt1_2, CType(2, Short))
		Me.Opt1.SetIndex(_Opt1_1, CType(1, Short))
		Me.Opt1.SetIndex(_Opt1_0, CType(0, Short))
		Me.OverSesTxt.SetIndex(_OverSesTxt_0, CType(0, Short))
		Me.OverSesTxt.SetIndex(_OverSesTxt_1, CType(1, Short))
		Me.OvhTxt.SetIndex(_OvhTxt_2, CType(2, Short))
		Me.OvhTxt.SetIndex(_OvhTxt_1, CType(1, Short))
		Me.OvhTxt.SetIndex(_OvhTxt_0, CType(0, Short))
		Me.PriBtn.SetIndex(_PriBtn_0, CType(0, Short))
		Me.PriBtn.SetIndex(_PriBtn_1, CType(1, Short))
		Me.PriPmTxt.SetIndex(_PriPmTxt_1, CType(1, Short))
		Me.PriPmTxt.SetIndex(_PriPmTxt_0, CType(0, Short))
		Me.PriTxt1.SetIndex(_PriTxt1_0, CType(0, Short))
		Me.PriTxt1.SetIndex(_PriTxt1_1, CType(1, Short))
		Me.RegLbl.SetIndex(_RegLbl_0, CType(0, Short))
		Me.RegLbl.SetIndex(_RegLbl_1, CType(1, Short))
		Me.SetBtn.SetIndex(_SetBtn_1, CType(1, Short))
		Me.SetBtn.SetIndex(_SetBtn_0, CType(0, Short))
		Me.SetLbl.SetIndex(_SetLbl_1, CType(1, Short))
		Me.SetLbl.SetIndex(_SetLbl_0, CType(0, Short))
		CType(Me.SetLbl, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.SetBtn, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.RegLbl, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.PriTxt1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.PriPmTxt, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.PriBtn, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.OvhTxt, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.OverSesTxt, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Opt1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.NetPassTxt, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.NetLbl, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.MpassTxt, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.MpassLbl, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.InfoTxt, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.InfoLbl, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.GwsOpt1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.EmpTxt, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.EmpLbl, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.EmpHdr, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.EmpBtn, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Lv1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.PriLV1, System.ComponentModel.ISupportInitialize).EndInit()
	End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As FrmSet
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As FrmSet
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New FrmSet()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	Dim DemoMode As Boolean
	
	
	Private Sub FrmSet_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim e As Object
		Dim hg As Object
		Dim sk As Object
		Dim n As Object
		On Error GoTo ErrInt
		Dim lItm As MSComctlLib.ListItem
		Dim acsTime, strNm As String
		
		NumOnly(NetPortTxt)
		NumOnly(OvhTxt(0))
		NumOnly(OvhTxt(1))
		NumOnly(OvhTxt(2))
		NumOnly(OverSesTxt(0))
		NumOnly(OverSesTxt(1))
		
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		TxtNama.Text = SetAmbil("namadaftar") 'ambil data nama bagi pengguna
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		TxtNombor.Text = SetAmbil("nombordaftar")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(mu, admin). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		MpassTxt(0).Text = SetAmbil("mu", "admin")
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(mp). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		MpassTxt(1).Text = SetAmbil("mp")
		MpassTxt(2).Text = MpassTxt(1).Text
		
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(namacc). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		InfoTxt(0).Text = SetAmbil("namacc") 'nama kedai cc
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(emailpengguna). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		InfoTxt(1).Text = SetAmbil("emailpengguna") 'email pengguna
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(tajukatas). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		InfoTxt(2).Text = SetAmbil("tajukatas") 'tajuk atas
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		NetPortTxt.Text = SetAmbil("porttempatan", 8180) 'port tempatan
		
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(harga, ). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		PriPmTxt(0).Text = SetAmbil("harga", 0.03)
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(hargaex, 0). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		PriPmTxt(1).Text = SetAmbil("hargaex", 0)
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		PriChk1.CheckState = SetAmbil("roundup", System.Windows.Forms.CheckState.Checked)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(sewa, 900). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		OvhTxt(0).Text = SetAmbil("sewa", 900)
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(bil, 230). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		OvhTxt(1).Text = SetAmbil("bil", 230)
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(billain, 90). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		OvhTxt(2).Text = SetAmbil("billain", 90)
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(logaktiviti, Checked). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		GwsOpt1(0).CheckState = SetAmbil("logaktiviti", System.Windows.Forms.CheckState.Checked)
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(tukarharga, Unchecked). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		GwsOpt1(1).CheckState = SetAmbil("tukarharga", System.Windows.Forms.CheckState.UnChecked)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		acsTime = SetAmbil("autocloses", "12:30:00 AM")
		OverSesTxt(0).Text = CStr(Hour(CDate(acsTime)))
		OverSesTxt(1).Text = CStr(Minute(CDate(acsTime)))
		OverCmb1.Text = VB.Right(acsTime, 2)
		If CDbl(OverSesTxt(0).Text) = 0 Then OverSesTxt(0).Text = CStr(12)
		
		If TxtNama.Text = "" Then TxtNama.Text = "demo"
		If TxtNombor.Text = "" Then TxtNombor.Text = "demo"
		
		'UPGRADE_WARNING: Couldn't resolve default property of object uSDBe.DataCount(skema). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		For n = 0 To uSDBe.DataCount("skema") - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object uSDBe.DataGet(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Couldn't resolve default property of object sk. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			sk = uSDBe.DataGet("skema", "skema", n)
			'UPGRADE_WARNING: Couldn't resolve default property of object uSDBe.DataGet(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Couldn't resolve default property of object hg. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			hg = uSDBe.DataGet("skema", "harga", n)
			lItm = PriLV1.ListItems.Add( ,  , sk)
			'UPGRADE_WARNING: Couldn't resolve default property of object hg. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			lItm.SubItems(1) = hg
		Next n
		
		'loading workers
		'UPGRADE_WARNING: Couldn't resolve default property of object Lv1.SmallIcons. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Lv1.SmallIcons = FrmMain.DefInstance.ImgListSnm.GetOCX
		'UPGRADE_WARNING: Couldn't resolve default property of object uSDBe.DataCount(pekerja-list). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		For e = 0 To uSDBe.DataCount("pekerja-list") - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object uSDBe.DataGet(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			strNm = uSDBe.DataGet("pekerja-list", "nama", e)
			lItm = Lv1.ListItems.Add( ,  , strNm,  , "user")
			'UPGRADE_WARNING: Couldn't resolve default property of object uSDBe.DataGet(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			lItm.SubItems(1) = uSDBe.DataGet("pekerja-list", "gaji", e)
			'UPGRADE_WARNING: Couldn't resolve default property of object uSDBe.DataGet(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			lItm.SubItems(2) = uSDBe.DataGet("pekerja-list", "nick", e)
			'UPGRADE_WARNING: Couldn't resolve default property of object uSDBe.DataGet(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			lItm.SubItems(3) = uSDBe.DataGet("pekerja-list", "password", e)
			'UPGRADE_WARNING: Couldn't resolve default property of object uSDBe.DataGet(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			lItm.SubItems(4) = uSDBe.DataGet("pekerja-list", "akses", e)
			lItm.let_Tag(strNm)
		Next e
		
		Exit Sub
ErrInt: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, CbMsgWarn)
	End Sub
	
	
	Private Sub EmpBtn_Click(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles EmpBtn.Click
		Dim Index As Short = EmpBtn.GetIndex(Sender)
		Dim g As Object
		Dim lret As Object
		Dim nm As Object
		Dim TmpAk As String
		Select Case Index
			Case 0
				Call EmployeeAdd()
			Case 1
				If Lv1.ListItems.Count = 0 Then Exit Sub
				If Lv1.SelectedItem.Text = "" Then Exit Sub
				'UPGRADE_WARNING: Couldn't resolve default property of object Lv1.SelectedItem.Tag. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				'UPGRADE_WARNING: Couldn't resolve default property of object nm. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				nm = Lv1.SelectedItem.Tag
				'UPGRADE_WARNING: Couldn't resolve default property of object lret. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				lret = MsgBox("Delete " & Lv1.SelectedItem.Text & " ?", MsgBoxStyle.OKCancel, CbMsgWarn)
				If lret = MsgBoxResult.Cancel Then Exit Sub
				uSDBe.DataRemove("pekerja-list", "nama", nm)
				Lv1.ListItems.Remove((Lv1.SelectedItem.Index))
			Case 2
				If Lv1.ListItems.Count = 0 Then Exit Sub
				If Lv1.SelectedItem.Text = "" Then Exit Sub
				
				For g = 1 To 3
					'UPGRADE_WARNING: Couldn't resolve default property of object g. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					If Opt1(g - 1).CheckState = 1 Then
						TmpAk = TmpAk & "1"
					Else
						TmpAk = TmpAk & "0"
					End If
				Next g
				Lv1.SelectedItem.SubItems(4) = TmpAk
				uSDBe.DataEdit("pekerja-list", "akses", "nick", Lv1.SelectedItem.SubItems(2), TmpAk, True, True)
		End Select
	End Sub
	
	
	Private Sub Lv1_ItemClick(ByVal eventSender As System.Object, ByVal eventArgs As AxMSComctlLib.ListViewEvents_ItemClickEvent) Handles Lv1.ItemClick
		Dim d As Object
		Dim strAccess As String
		strAccess = eventArgs.Item.SubItems(4)
		
		EmpTxt(0).Text = eventArgs.Item.Text
		EmpTxt(1).Text = eventArgs.Item.SubItems(2)
		EmpTxt(2).Text = eventArgs.Item.SubItems(3)
		EmpTxt(3).Text = eventArgs.Item.SubItems(1)
		For d = 1 To 3
			'UPGRADE_WARNING: Couldn't resolve default property of object d. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			Opt1(d - 1).CheckState = System.Windows.Forms.CheckState.Unchecked
			'UPGRADE_WARNING: Couldn't resolve default property of object d. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			If Mid(strAccess, d, 1) = "1" Then Opt1(d - 1).CheckState = System.Windows.Forms.CheckState.Checked
		Next d
	End Sub
	
	
	Private Sub PriBtn_Click(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles PriBtn.Click
		Dim Index As Short = PriBtn.GetIndex(Sender)
		Dim lret As Object
		Dim nm As Object
		Dim lItm As MSComctlLib.ListItem
		Dim itmfind As MSComctlLib.ListItem
		
		Select Case Index
			Case 0
				If PriTxt1(0).Text = "" Then Exit Sub
				If PriTxt1(1).Text = "" Then Exit Sub
				
				itmfind = PriLV1.FindItem(PriTxt1(0).Text)
				If itmfind Is Nothing Then
					lItm = PriLV1.ListItems.Add( ,  , PriTxt1(0))
					lItm.SubItems(1) = PriTxt1(1).Text
					'tambah dalam database
					uSDBe.DataSave("skema", "skema", PriTxt1(0), True, False)
					uSDBe.DataSave("Skema", "harga", PriTxt1(1), False, True)
				Else
					MsgBox(MB(1), MsgBoxStyle.OKOnly, CbMsgWarn) : Exit Sub
				End If
				
				PriTxt1(0).Text = ""
				PriTxt1(1).Text = ""
			Case 1
				If PriLV1.ListItems.Count = 0 Then Exit Sub
				If PriLV1.SelectedItem.Text = "" Then Exit Sub
				
				'UPGRADE_WARNING: Couldn't resolve default property of object nm. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				nm = PriLV1.SelectedItem.Text
				'UPGRADE_WARNING: Couldn't resolve default property of object nm. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				'UPGRADE_WARNING: Couldn't resolve default property of object lret. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				lret = MsgBox(MB(2) & " " & nm & " ?", MsgBoxStyle.OKCancel, CbMsgWarn)
				If lret = MsgBoxResult.Cancel Then Exit Sub
				PriLV1.ListItems.Remove((PriLV1.SelectedItem.Index))
				uSDBe.DataRemove("skema", "skema", nm)
		End Select
	End Sub
	
	Private Sub RegBtn_Click(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles RegBtn.Click
		Dim sRet As String
		
		If CbDrvStr = "" Then CbDrvStr = "a:"
		If Trim(TxtNama.Text) = "" And Trim(TxtNombor.Text) = "" Then GoTo Register
		If LCase(TxtNama.Text) = "demo" And LCase(TxtNombor.Text) = "demo" Then GoTo Register
		
		If TxtNama.Text <> "" And TxtNombor.Text <> "" Then
			sRet = CStr(MsgBox(MB(5), MsgBoxStyle.OKCancel + MsgBoxStyle.Information, CbMsgWarn))
			If sRet = CStr(MsgBoxResult.OK) Then
				If CreateDiskKey(TxtNama.Text, TxtNombor.Text, "a:") = True Then
					SetSimpan("demo", CStr(True))
					SetSimpan("demoday", CStr(10))
					SetSimpan("namadaftar", "demo")
					SetSimpan("nombordaftar", "demo")
					CbDemoMode = True
					DemoMode = True
				End If
			End If
		End If
		Exit Sub
		
Register: 
		If ValidateDisk(CbDrvStr) = True Then
			TxtNama.Text = GetName(CbDrvStr)
			TxtNombor.Text = GetKey(CbDrvStr)
			If InitReg = True Then
				SetSimpan("namadaftar", TxtNama.Text)
				SetSimpan("nombordaftar", TxtNombor.Text)
				DemoMode = False
				CbDemoMode = False
				SetSimpan("demo", CStr(DemoMode))
				MsgBox(MB(6), MsgBoxStyle.OKOnly, "CafeBonzer")
			End If
		End If
	End Sub
	
	Private Sub SetBtn_Click(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles SetBtn.Click
		Dim Index As Short = SetBtn.GetIndex(Sender)
		Dim Text5 As Object
		Dim Text7 As Object
		Dim Text6 As Object
		DemoMode = False
		
		Select Case Index
			Case 0
				If LCase(TxtNama.Text) = "demo" And LCase(TxtNombor.Text) = "demo" Then DemoMode = True
				'************************************
				'* simpan pada data yang program
				'* telah di buka untuk pertama kali
				'************************************
				'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(pertamakali). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				If InitReg = False And SetAmbil("pertamakali") <> "tidak" Then
					SetSimpan("pertamakali", "ya")
					FrmSet.DefInstance.Hide()
					Keluar(False)
					End
				End If
				
				Call CloseFrm(FrmSet.DefInstance)
				FrmMain.DefInstance.Show()
				
			Case 1
				'************************************
				'* Periksa katalaluan utama dan client
				'************************************
				If MpassTxt(0).Text = "" Or MpassTxt(1).Text = "" Then
					MsgBox(MB(8), MsgBoxStyle.Information, CbMsgWarn)
					MpassTxt(0).Focus()
					Exit Sub
				ElseIf MpassTxt(1).Text <> MpassTxt(2).Text Then 
					MpassTxt(1).Focus()
					MsgBox(MB(9), MsgBoxStyle.Information, CbMsgWarn)
					Exit Sub
				End If
				If NetPassTxt(0).Text <> NetPassTxt(1).Text Then
					NetPassTxt(0).Focus()
					MsgBox(MB(9), MsgBoxStyle.Information, CbMsgWarn)
					Exit Sub
				End If
				
				'************************************
				'* Periksa txtbox lain bagi kesalahan
				'************************************
				If NetPortTxt.Text = "" Then NetPortTxt.Text = CStr(56266)
				'UPGRADE_WARNING: Couldn't resolve default property of object Text6.Text. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				If PriPmTxt(0).Text = "" Or IsNumeric(Text6) = False Then Text6.Text = 0.05
				'UPGRADE_WARNING: Couldn't resolve default property of object Text7.Text. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				If PriPmTxt(1).Text = "" Or IsNumeric(Text7) = False Then Text7.Text = 0
				If OvhTxt(0).Text = "" Then OvhTxt(0).Text = CStr(800)
				If OvhTxt(1).Text = "" Then OvhTxt(1).Text = CStr(350)
				If OvhTxt(2).Text = "" Then OvhTxt(2).Text = CStr(20)
				
				'************************************
				'* Saving all settings
				'************************************
				SetSimpan("mu", MpassTxt(0).Text) 'simpan master username..
				SetSimpan("mp", MpassTxt(1).Text) 'simpan master password
				
				SetSimpan("namadaftar", TxtNama.Text) 'simpan no. daftar
				SetSimpan("nombordaftar", TxtNombor.Text) 'simpan nama daftar
				SetSimpan("namacc", InfoTxt(0).Text) 'simpan nama cc
				SetSimpan("emailpengguna", InfoTxt(1).Text) 'email pengguna
				SetSimpan("tajukatas", InfoTxt(2).Text) 'simpan tajuk atas
				
				SetSimpan("porttempatan", NetPortTxt.Text) 'simpan no. port tempatan
				SetSimpan("netcpwd", NetPassTxt(0).Text)
				
				SetSimpan("sewa", OvhTxt(0).Text)
				SetSimpan("bil", OvhTxt(1).Text)
				SetSimpan("billain", OvhTxt(2).Text)
				SetSimpan("autocloses", (OverSesTxt(0).Text & ":" & OverSesTxt(1).Text & ":00 " & OverCmb1.Text))
				
				SetSimpan("harga", PriPmTxt(0).Text) 'simpan harga per/minit
				SetSimpan("hargaex", PriPmTxt(1).Text)
				SetSimpan("roundup", CStr(PriChk1.CheckState))
				
				SetSimpan("logaktiviti", CStr(GwsOpt1(0).CheckState))
				SetSimpan("tukarharga", CStr(GwsOpt1(1).CheckState))
				
				SetSimpan("pertamakali", "tidak")
				
				'UPGRADE_WARNING: Couldn't resolve default property of object Text5. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				FrmMain.DefInstance.Text = "CafeBonzer - " & Text5
				'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(demo). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				If SetAmbil("demo") = "True" Then FrmMain.DefInstance.Text = FrmMain.DefInstance.Text & " UNREGISTERED" : CbDemoMode = True
				Call CloseFrm(FrmSet.DefInstance)
		End Select
		
		Exit Sub
EnterNumeric: 
		MsgBox(MB(4), MsgBoxStyle.OKOnly, "CafeBonzer")
	End Sub
	
	
	Function InitReg() As Boolean
		Dim idstr As Object
		Dim genstr As Object
		Dim a6 As Object
		Dim a5 As Object
		Dim a4 As Object
		Dim a3 As Object
		Dim a2 As Object
		Dim a1 As Object
		'Menyediakan algoritma bagi pengiraan kod daftar
		'UPGRADE_WARNING: Couldn't resolve default property of object a1. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		a1 = Len(TxtNama.Text) 'ambil panjang nama
		'UPGRADE_WARNING: Couldn't resolve default property of object a1. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object a2. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		a2 = a1 * 5 'panjang nama darab dengan 5
		'UPGRADE_WARNING: Couldn't resolve default property of object a1. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object a2. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object a3. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		a3 = a2 * a1 'hasil darab, di darabkan balik dengan panjang nama
		'UPGRADE_WARNING: Couldn't resolve default property of object a4. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		a4 = VB.Left(TxtNama.Text, 1) 'ambil huruf paling hujung sebelah kiri
		'UPGRADE_WARNING: Couldn't resolve default property of object a5. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		a5 = VB.Right(TxtNama.Text, 1) 'ambil huruf paling hujung sebelah kanan
		'UPGRADE_WARNING: Couldn't resolve default property of object a6. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		a6 = "0v10" 'untuk versi = versi 0.10
		
		'menambahkan dan mengaturkan algoritma kod daftar
		'UPGRADE_WARNING: Couldn't resolve default property of object a6. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object a5. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object a4. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object a3. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object a2. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object a1. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object genstr. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		genstr = a1 & a2 & a3 & a4 & a5 & a6
		'UPGRADE_WARNING: Couldn't resolve default property of object idstr. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		idstr = TxtNombor.Text
		
		'membandingkan kod daftar dengan nama
		'UPGRADE_WARNING: Couldn't resolve default property of object idstr. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object genstr. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If LCase(Trim(genstr)) = LCase(Trim(idstr)) Then
			InitReg = True
		Else
			InitReg = False
		End If
	End Function
	
	
	Private Sub EmployeeAdd()
		Dim nItm As MSComctlLib.ListItem
		If EmpTxt(0).Text = "" Then Exit Sub
		If EmpTxt(1).Text = "" Then Exit Sub
		If EmpTxt(2).Text = "" Then Exit Sub
		If EmpTxt(3).Text = "" Then Exit Sub
		
		nItm = Lv1.ListItems.Add( ,  , EmpTxt(1),  , "user")
		nItm.SubItems(1) = EmpTxt(2).Text
		nItm.SubItems(2) = EmpTxt(1).Text
		nItm.SubItems(3) = EmpTxt(3).Text
		nItm.SubItems(4) = "000"
		
		uSDBe.DataSave("pekerja-list", "nama", EmpTxt(0), True, False)
		uSDBe.DataSave("pekerja-list", "nick", EmpTxt(1), False, False)
		uSDBe.DataSave("pekerja-list", "gaji", EmpTxt(2), False, False)
		uSDBe.DataSave("pekerja-list", "akses", "000", False, False)
		uSDBe.DataSave("pekerja-list", "password", EmpTxt(3), False, True)
	End Sub
End Class