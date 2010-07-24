Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class FrmLogout
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
	Public WithEvents PayBal As System.Windows.Forms.TextBox
	Public WithEvents PayTotal As System.Windows.Forms.TextBox
	Public WithEvents _PayLbl_3 As System.Windows.Forms.Label
	Public WithEvents _PayLbl_2 As System.Windows.Forms.Label
	Public WithEvents _PayBox_0 As System.Windows.Forms.Panel
	Public WithEvents PayTime As Label3D
	Public WithEvents _PayBox_1 As System.Windows.Forms.Panel
	Public WithEvents PayPrice As System.Windows.Forms.TextBox
	Public WithEvents _PayLbl_0 As System.Windows.Forms.Label
	Public WithEvents _PayLbl_1 As System.Windows.Forms.Label
	Public WithEvents _PayImg_0 As System.Windows.Forms.PictureBox
	Public WithEvents _PayImg_1 As System.Windows.Forms.PictureBox
	Public WithEvents Dock2 As PageDock
	Public WithEvents Line3D2 As Line3D
	Public WithEvents _PayBtn_1 As XpButton
	Public WithEvents PayRcv As System.Windows.Forms.TextBox
	Public WithEvents _PayBtn_0 As XpButton
	Public WithEvents _MnuABtn_3 As XpButton
	Public WithEvents _MnuABtn_2 As XpButton
	Public WithEvents _MnuABtn_1 As XpButton
	Public WithEvents _MnuABtn_0 As XpButton
	Public WithEvents Dock1 As PageDock
	Public WithEvents _Line3D1_0 As Line3D
	Public WithEvents _MnuBBtn_0 As XpButton
	Public WithEvents _MnuBBtn_1 As XpButton
	Public WithEvents _MnuBBtn_2 As XpButton
	Public WithEvents _MnuBBtn_3 As XpButton
	Public WithEvents _SerICmb_1 As AxMSComctlLib.AxImageCombo
	Public WithEvents _SerICmb_0 As AxMSComctlLib.AxImageCombo
	Public WithEvents SerLv As AxMSComctlLib.AxListView
	Public WithEvents _SerBtn_0 As XpButton
	Public WithEvents _SerBtn_1 As XpButton
	Public WithEvents SerQty As System.Windows.Forms.TextBox
	Public WithEvents SerScroll1 As System.Windows.Forms.VScrollBar
	Public WithEvents SerJumlah As System.Windows.Forms.TextBox
	Public WithEvents _SerLbl_3 As System.Windows.Forms.Label
	Public WithEvents _SerLbl_2 As System.Windows.Forms.Label
	Public WithEvents _SerLbl_1 As System.Windows.Forms.Label
	Public WithEvents _SerLbl_0 As System.Windows.Forms.Label
	Public WithEvents Line3D1 As Line3DArray
	Public WithEvents MnuABtn As XpButtonArray
	Public WithEvents MnuBBtn As XpButtonArray
	Public WithEvents PayBox As Microsoft.VisualBasic.Compatibility.VB6.PanelArray
	Public WithEvents PayBtn As XpButtonArray
	Public WithEvents PayImg As Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray
	Public WithEvents PayLbl As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents SerBtn As XpButtonArray
	Public WithEvents SerICmb As AxImageComboArray.AxImageComboArray
	Public WithEvents SerLbl As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmLogout))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.Dock2 = New PageDock
		Me._PayBox_0 = New System.Windows.Forms.Panel
		Me.PayBal = New System.Windows.Forms.TextBox
		Me.PayTotal = New System.Windows.Forms.TextBox
		Me._PayLbl_3 = New System.Windows.Forms.Label
		Me._PayLbl_2 = New System.Windows.Forms.Label
		Me._PayBox_1 = New System.Windows.Forms.Panel
		Me.PayTime = New Label3D
		Me.PayPrice = New System.Windows.Forms.TextBox
		Me._PayLbl_0 = New System.Windows.Forms.Label
		Me._PayLbl_1 = New System.Windows.Forms.Label
		Me._PayImg_0 = New System.Windows.Forms.PictureBox
		Me._PayImg_1 = New System.Windows.Forms.PictureBox
		Me.Dock1 = New PageDock
		Me.Line3D2 = New Line3D
		Me._PayBtn_1 = New XpButton
		Me.PayRcv = New System.Windows.Forms.TextBox
		Me._PayBtn_0 = New XpButton
		Me._MnuABtn_3 = New XpButton
		Me._MnuABtn_2 = New XpButton
		Me._MnuABtn_1 = New XpButton
		Me._MnuABtn_0 = New XpButton
		Me._Line3D1_0 = New Line3D
		Me._MnuBBtn_0 = New XpButton
		Me._MnuBBtn_1 = New XpButton
		Me._MnuBBtn_2 = New XpButton
		Me._MnuBBtn_3 = New XpButton
		Me._SerICmb_1 = New AxMSComctlLib.AxImageCombo
		Me._SerICmb_0 = New AxMSComctlLib.AxImageCombo
		Me.SerLv = New AxMSComctlLib.AxListView
		Me._SerBtn_0 = New XpButton
		Me._SerBtn_1 = New XpButton
		Me.SerQty = New System.Windows.Forms.TextBox
		Me.SerScroll1 = New System.Windows.Forms.VScrollBar
		Me.SerJumlah = New System.Windows.Forms.TextBox
		Me._SerLbl_3 = New System.Windows.Forms.Label
		Me._SerLbl_2 = New System.Windows.Forms.Label
		Me._SerLbl_1 = New System.Windows.Forms.Label
		Me._SerLbl_0 = New System.Windows.Forms.Label
		Me.Line3D1 = New Line3DArray(components)
		Me.MnuABtn = New XpButtonArray(components)
		Me.MnuBBtn = New XpButtonArray(components)
		Me.PayBox = New Microsoft.VisualBasic.Compatibility.VB6.PanelArray(components)
		Me.PayBtn = New XpButtonArray(components)
		Me.PayImg = New Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray(components)
		Me.PayLbl = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.SerBtn = New XpButtonArray(components)
		Me.SerICmb = New AxImageComboArray.AxImageComboArray(components)
		Me.SerLbl = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		CType(Me._SerICmb_1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me._SerICmb_0, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.SerLv, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Line3D1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.MnuABtn, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.MnuBBtn, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.PayBox, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.PayBtn, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.PayImg, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.PayLbl, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.SerBtn, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.SerICmb, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.SerLbl, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.ControlBox = False
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.ClientSize = New System.Drawing.Size(441, 173)
		Me.Location = New System.Drawing.Point(17, 94)
		Me.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.KeyPreview = True
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
		Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.Enabled = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "FrmLogout"
		Me.Dock2.Size = New System.Drawing.Size(441, 115)
		Me.Dock2.Location = New System.Drawing.Point(0, 58)
		Me.Dock2.TabIndex = 14
		Me.Dock2.HldrBtnPos = 0
		Me.Dock2.HldrLne = -1
		Me.Dock2.PageState = 0
		Me.Dock2.PageWidth = 6615
		Me.Dock2.Name = "Dock2"
		Me._PayBox_0.BackColor = System.Drawing.Color.Black
		Me._PayBox_0.Size = New System.Drawing.Size(183, 95)
		Me._PayBox_0.Location = New System.Drawing.Point(251, 10)
		Me._PayBox_0.TabIndex = 20
		Me._PayBox_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._PayBox_0.Dock = System.Windows.Forms.DockStyle.None
		Me._PayBox_0.CausesValidation = True
		Me._PayBox_0.Enabled = True
		Me._PayBox_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._PayBox_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._PayBox_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._PayBox_0.TabStop = True
		Me._PayBox_0.Visible = True
		Me._PayBox_0.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._PayBox_0.Name = "_PayBox_0"
		Me.PayBal.AutoSize = False
		Me.PayBal.BackColor = System.Drawing.Color.Black
		Me.PayBal.Font = New System.Drawing.Font("Endless Showroom", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.PayBal.ForeColor = System.Drawing.Color.Green
		Me.PayBal.Size = New System.Drawing.Size(83, 37)
		Me.PayBal.Location = New System.Drawing.Point(84, 49)
		Me.PayBal.ReadOnly = True
		Me.PayBal.TabIndex = 23
		Me.PayBal.TabStop = False
		Me.PayBal.Text = "0.00"
		Me.ToolTip1.SetToolTip(Me.PayBal, "Just press Enter if the value same as above.")
		Me.PayBal.AcceptsReturn = True
		Me.PayBal.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.PayBal.CausesValidation = True
		Me.PayBal.Enabled = True
		Me.PayBal.HideSelection = True
		Me.PayBal.Maxlength = 0
		Me.PayBal.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.PayBal.MultiLine = False
		Me.PayBal.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.PayBal.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.PayBal.Visible = True
		Me.PayBal.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.PayBal.Name = "PayBal"
		Me.PayTotal.AutoSize = False
		Me.PayTotal.BackColor = System.Drawing.Color.Black
		Me.PayTotal.Font = New System.Drawing.Font("Endless Showroom", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.PayTotal.ForeColor = System.Drawing.Color.Green
		Me.PayTotal.Size = New System.Drawing.Size(84, 37)
		Me.PayTotal.Location = New System.Drawing.Point(83, 7)
		Me.PayTotal.ReadOnly = True
		Me.PayTotal.TabIndex = 22
		Me.PayTotal.TabStop = False
		Me.PayTotal.Text = "0.00"
		Me.PayTotal.AcceptsReturn = True
		Me.PayTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.PayTotal.CausesValidation = True
		Me.PayTotal.Enabled = True
		Me.PayTotal.HideSelection = True
		Me.PayTotal.Maxlength = 0
		Me.PayTotal.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.PayTotal.MultiLine = False
		Me.PayTotal.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.PayTotal.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.PayTotal.Visible = True
		Me.PayTotal.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.PayTotal.Name = "PayTotal"
		Me._PayLbl_3.Text = "Balance :"
		Me._PayLbl_3.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._PayLbl_3.ForeColor = System.Drawing.Color.FromARGB(224, 224, 224)
		Me._PayLbl_3.Size = New System.Drawing.Size(81, 16)
		Me._PayLbl_3.Location = New System.Drawing.Point(9, 54)
		Me._PayLbl_3.TabIndex = 24
		Me._PayLbl_3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._PayLbl_3.BackColor = System.Drawing.Color.Transparent
		Me._PayLbl_3.Enabled = True
		Me._PayLbl_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._PayLbl_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._PayLbl_3.UseMnemonic = True
		Me._PayLbl_3.Visible = True
		Me._PayLbl_3.AutoSize = False
		Me._PayLbl_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._PayLbl_3.Name = "_PayLbl_3"
		Me._PayLbl_2.Text = "Total :"
		Me._PayLbl_2.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._PayLbl_2.ForeColor = System.Drawing.Color.FromARGB(224, 224, 224)
		Me._PayLbl_2.Size = New System.Drawing.Size(54, 16)
		Me._PayLbl_2.Location = New System.Drawing.Point(8, 11)
		Me._PayLbl_2.TabIndex = 21
		Me._PayLbl_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._PayLbl_2.BackColor = System.Drawing.Color.Transparent
		Me._PayLbl_2.Enabled = True
		Me._PayLbl_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._PayLbl_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._PayLbl_2.UseMnemonic = True
		Me._PayLbl_2.Visible = True
		Me._PayLbl_2.AutoSize = False
		Me._PayLbl_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._PayLbl_2.Name = "_PayLbl_2"
		Me._PayBox_1.BackColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me._PayBox_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._PayBox_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._PayBox_1.Size = New System.Drawing.Size(123, 24)
		Me._PayBox_1.Location = New System.Drawing.Point(109, 12)
		Me._PayBox_1.TabIndex = 15
		Me._PayBox_1.TabStop = False
		Me._PayBox_1.Dock = System.Windows.Forms.DockStyle.None
		Me._PayBox_1.CausesValidation = True
		Me._PayBox_1.Enabled = True
		Me._PayBox_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._PayBox_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._PayBox_1.Visible = True
		Me._PayBox_1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._PayBox_1.Name = "_PayBox_1"
		Me.PayTime.Size = New System.Drawing.Size(115, 14)
		Me.PayTime.Location = New System.Drawing.Point(3, 3)
		Me.PayTime.TabIndex = 16
		Me.PayTime.TabStop = 0
		Me.PayTime.Name = "PayTime"
		Me.PayPrice.AutoSize = False
		Me.PayPrice.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.PayPrice.BackColor = System.Drawing.Color.White
		Me.PayPrice.Font = New System.Drawing.Font("Endless Showroom", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.PayPrice.ForeColor = System.Drawing.Color.Black
		Me.PayPrice.Size = New System.Drawing.Size(124, 25)
		Me.PayPrice.Location = New System.Drawing.Point(108, 67)
		Me.PayPrice.TabIndex = 18
		Me.PayPrice.Text = "0.00"
		Me.PayPrice.AcceptsReturn = True
		Me.PayPrice.CausesValidation = True
		Me.PayPrice.Enabled = True
		Me.PayPrice.HideSelection = True
		Me.PayPrice.ReadOnly = False
		Me.PayPrice.Maxlength = 0
		Me.PayPrice.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.PayPrice.MultiLine = False
		Me.PayPrice.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.PayPrice.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.PayPrice.TabStop = True
		Me.PayPrice.Visible = True
		Me.PayPrice.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.PayPrice.Name = "PayPrice"
		Me._PayLbl_0.Text = "Time :"
		Me._PayLbl_0.Font = New System.Drawing.Font("Verdana", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._PayLbl_0.ForeColor = System.Drawing.Color.Blue
		Me._PayLbl_0.Size = New System.Drawing.Size(54, 21)
		Me._PayLbl_0.Location = New System.Drawing.Point(55, 15)
		Me._PayLbl_0.TabIndex = 17
		Me._PayLbl_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._PayLbl_0.BackColor = System.Drawing.Color.Transparent
		Me._PayLbl_0.Enabled = True
		Me._PayLbl_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._PayLbl_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._PayLbl_0.UseMnemonic = True
		Me._PayLbl_0.Visible = True
		Me._PayLbl_0.AutoSize = False
		Me._PayLbl_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._PayLbl_0.Name = "_PayLbl_0"
		Me._PayLbl_1.Text = "Price :"
		Me._PayLbl_1.Font = New System.Drawing.Font("Verdana", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._PayLbl_1.ForeColor = System.Drawing.Color.Blue
		Me._PayLbl_1.Size = New System.Drawing.Size(54, 21)
		Me._PayLbl_1.Location = New System.Drawing.Point(55, 70)
		Me._PayLbl_1.TabIndex = 19
		Me._PayLbl_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._PayLbl_1.BackColor = System.Drawing.Color.Transparent
		Me._PayLbl_1.Enabled = True
		Me._PayLbl_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._PayLbl_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._PayLbl_1.UseMnemonic = True
		Me._PayLbl_1.Visible = True
		Me._PayLbl_1.AutoSize = False
		Me._PayLbl_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._PayLbl_1.Name = "_PayLbl_1"
		Me._PayImg_0.Size = New System.Drawing.Size(16, 16)
		Me._PayImg_0.Location = New System.Drawing.Point(34, 15)
		Me._PayImg_0.Image = CType(resources.GetObject("_PayImg_0.Image"), System.Drawing.Image)
		Me._PayImg_0.Enabled = True
		Me._PayImg_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._PayImg_0.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._PayImg_0.Visible = True
		Me._PayImg_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._PayImg_0.Name = "_PayImg_0"
		Me._PayImg_1.Size = New System.Drawing.Size(16, 16)
		Me._PayImg_1.Location = New System.Drawing.Point(33, 70)
		Me._PayImg_1.Image = CType(resources.GetObject("_PayImg_1.Image"), System.Drawing.Image)
		Me._PayImg_1.Enabled = True
		Me._PayImg_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._PayImg_1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._PayImg_1.Visible = True
		Me._PayImg_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._PayImg_1.Name = "_PayImg_1"
		Me.Dock1.Size = New System.Drawing.Size(441, 56)
		Me.Dock1.Location = New System.Drawing.Point(0, 0)
		Me.Dock1.TabIndex = 0
		Me.Dock1.HldrBtnPos = 0
		Me.Dock1.HldrLne = -1
		Me.Dock1.PageState = 0
		Me.Dock1.PageWidth = 6615
		Me.Dock1.Name = "Dock1"
		Me.Line3D2.Size = New System.Drawing.Size(3, 55)
		Me.Line3D2.Location = New System.Drawing.Point(220, 0)
		Me.Line3D2.TabIndex = 5
		Me.Line3D2.horizon = 0
		Me.Line3D2.Name = "Line3D2"
		Me._PayBtn_1.Size = New System.Drawing.Size(42, 40)
		Me._PayBtn_1.Location = New System.Drawing.Point(350, 8)
		Me._PayBtn_1.TabIndex = 8
		Me.ToolTip1.SetToolTip(Me._PayBtn_1, "Add Services or Merchandise")
		Me._PayBtn_1.TX = ""
		Me._PayBtn_1.ENAB = 0
		Me._PayBtn_1.COLTYPE = 1
		Me._PayBtn_1.FOCUSR = -1
		Me._PayBtn_1.BCOL = 12632256
		Me._PayBtn_1.BCOLO = 12632256
		Me._PayBtn_1.FCOL = 0
		Me._PayBtn_1.FCOLO = 0
		Me._PayBtn_1.MCOL = 16777215
		Me._PayBtn_1.MPTR = 1
		Me._PayBtn_1.MICON = 0
		Me._PayBtn_1.PICN = 0
		Me._PayBtn_1.UMCOL = -1
		Me._PayBtn_1.SOFT = 0
		Me._PayBtn_1.PICPOS = 0
		Me._PayBtn_1.NGREY = 0
		Me._PayBtn_1.FX = 0
		Me._PayBtn_1.HAND = 0
		Me._PayBtn_1.CHECK = 0
		Me._PayBtn_1.Name = "_PayBtn_1"
		Me.PayRcv.AutoSize = False
		Me.PayRcv.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.PayRcv.BackColor = System.Drawing.Color.White
		Me.PayRcv.Enabled = False
		Me.PayRcv.Font = New System.Drawing.Font("Endless Showroom", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.PayRcv.ForeColor = System.Drawing.Color.Black
		Me.PayRcv.Size = New System.Drawing.Size(103, 37)
		Me.PayRcv.Location = New System.Drawing.Point(239, 9)
		Me.PayRcv.TabIndex = 6
		Me.ToolTip1.SetToolTip(Me.PayRcv, "Just press Enter if the value same as above.")
		Me.PayRcv.AcceptsReturn = True
		Me.PayRcv.CausesValidation = True
		Me.PayRcv.HideSelection = True
		Me.PayRcv.ReadOnly = False
		Me.PayRcv.Maxlength = 0
		Me.PayRcv.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.PayRcv.MultiLine = False
		Me.PayRcv.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.PayRcv.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.PayRcv.TabStop = True
		Me.PayRcv.Visible = True
		Me.PayRcv.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.PayRcv.Name = "PayRcv"
		Me._PayBtn_0.Size = New System.Drawing.Size(42, 40)
		Me._PayBtn_0.Location = New System.Drawing.Point(394, 8)
		Me._PayBtn_0.TabIndex = 7
		Me.ToolTip1.SetToolTip(Me._PayBtn_0, "Confirm Transaction")
		Me._PayBtn_0.TX = ""
		Me._PayBtn_0.ENAB = 0
		Me._PayBtn_0.COLTYPE = 1
		Me._PayBtn_0.FOCUSR = -1
		Me._PayBtn_0.BCOL = 12632256
		Me._PayBtn_0.BCOLO = 12632256
		Me._PayBtn_0.FCOL = 0
		Me._PayBtn_0.FCOLO = 0
		Me._PayBtn_0.MCOL = 16777215
		Me._PayBtn_0.MPTR = 1
		Me._PayBtn_0.MICON = 0
		Me._PayBtn_0.PICN = 0
		Me._PayBtn_0.UMCOL = -1
		Me._PayBtn_0.SOFT = 0
		Me._PayBtn_0.PICPOS = 0
		Me._PayBtn_0.NGREY = 0
		Me._PayBtn_0.FX = 0
		Me._PayBtn_0.HAND = 0
		Me._PayBtn_0.CHECK = 0
		Me._PayBtn_0.Name = "_PayBtn_0"
		Me._MnuABtn_3.Size = New System.Drawing.Size(42, 42)
		Me._MnuABtn_3.Location = New System.Drawing.Point(161, 7)
		Me._MnuABtn_3.TabIndex = 4
		Me.ToolTip1.SetToolTip(Me._MnuABtn_3, "Cancel")
		Me._MnuABtn_3.TX = ""
		Me._MnuABtn_3.ENAB = -1
		Me._MnuABtn_3.COLTYPE = 1
		Me._MnuABtn_3.FOCUSR = -1
		Me._MnuABtn_3.BCOL = 12632256
		Me._MnuABtn_3.BCOLO = 12632256
		Me._MnuABtn_3.FCOL = 0
		Me._MnuABtn_3.FCOLO = 0
		Me._MnuABtn_3.MCOL = 16777215
		Me._MnuABtn_3.MPTR = 1
		Me._MnuABtn_3.MICON = 0
		Me._MnuABtn_3.PICN = 0
		Me._MnuABtn_3.UMCOL = -1
		Me._MnuABtn_3.SOFT = 0
		Me._MnuABtn_3.PICPOS = 0
		Me._MnuABtn_3.NGREY = 0
		Me._MnuABtn_3.FX = 0
		Me._MnuABtn_3.HAND = 0
		Me._MnuABtn_3.CHECK = 0
		Me._MnuABtn_3.Name = "_MnuABtn_3"
		Me._MnuABtn_2.Size = New System.Drawing.Size(42, 42)
		Me._MnuABtn_2.Location = New System.Drawing.Point(117, 7)
		Me._MnuABtn_2.TabIndex = 3
		Me.ToolTip1.SetToolTip(Me._MnuABtn_2, "Continue Terminal Usage")
		Me._MnuABtn_2.TX = ""
		Me._MnuABtn_2.ENAB = -1
		Me._MnuABtn_2.COLTYPE = 1
		Me._MnuABtn_2.FOCUSR = -1
		Me._MnuABtn_2.BCOL = 12632256
		Me._MnuABtn_2.BCOLO = 12632256
		Me._MnuABtn_2.FCOL = 0
		Me._MnuABtn_2.FCOLO = 0
		Me._MnuABtn_2.MCOL = 16777215
		Me._MnuABtn_2.MPTR = 1
		Me._MnuABtn_2.MICON = 0
		Me._MnuABtn_2.PICN = 0
		Me._MnuABtn_2.UMCOL = -1
		Me._MnuABtn_2.SOFT = 0
		Me._MnuABtn_2.PICPOS = 0
		Me._MnuABtn_2.NGREY = 0
		Me._MnuABtn_2.FX = 0
		Me._MnuABtn_2.HAND = 0
		Me._MnuABtn_2.CHECK = 0
		Me._MnuABtn_2.Name = "_MnuABtn_2"
		Me._MnuABtn_1.Size = New System.Drawing.Size(42, 42)
		Me._MnuABtn_1.Location = New System.Drawing.Point(73, 7)
		Me._MnuABtn_1.TabIndex = 2
		Me.ToolTip1.SetToolTip(Me._MnuABtn_1, "Logout Customer")
		Me._MnuABtn_1.TX = ""
		Me._MnuABtn_1.ENAB = -1
		Me._MnuABtn_1.COLTYPE = 1
		Me._MnuABtn_1.FOCUSR = -1
		Me._MnuABtn_1.BCOL = 12632256
		Me._MnuABtn_1.BCOLO = 12632256
		Me._MnuABtn_1.FCOL = 0
		Me._MnuABtn_1.FCOLO = 0
		Me._MnuABtn_1.MCOL = 16777215
		Me._MnuABtn_1.MPTR = 1
		Me._MnuABtn_1.MICON = 0
		Me._MnuABtn_1.PICN = 0
		Me._MnuABtn_1.UMCOL = -1
		Me._MnuABtn_1.SOFT = 0
		Me._MnuABtn_1.PICPOS = 0
		Me._MnuABtn_1.NGREY = 0
		Me._MnuABtn_1.FX = 0
		Me._MnuABtn_1.HAND = 0
		Me._MnuABtn_1.CHECK = 0
		Me._MnuABtn_1.Name = "_MnuABtn_1"
		Me._MnuABtn_0.Size = New System.Drawing.Size(42, 42)
		Me._MnuABtn_0.Location = New System.Drawing.Point(29, 7)
		Me._MnuABtn_0.TabIndex = 1
		Me.ToolTip1.SetToolTip(Me._MnuABtn_0, "Login Customer")
		Me._MnuABtn_0.TX = ""
		Me._MnuABtn_0.ENAB = -1
		Me._MnuABtn_0.COLTYPE = 1
		Me._MnuABtn_0.FOCUSR = -1
		Me._MnuABtn_0.BCOL = 12632256
		Me._MnuABtn_0.BCOLO = 12632256
		Me._MnuABtn_0.FCOL = 0
		Me._MnuABtn_0.FCOLO = 0
		Me._MnuABtn_0.MCOL = 16777215
		Me._MnuABtn_0.MPTR = 1
		Me._MnuABtn_0.MICON = 0
		Me._MnuABtn_0.PICN = 0
		Me._MnuABtn_0.UMCOL = -1
		Me._MnuABtn_0.SOFT = 0
		Me._MnuABtn_0.PICPOS = 0
		Me._MnuABtn_0.NGREY = 0
		Me._MnuABtn_0.FX = 0
		Me._MnuABtn_0.HAND = 0
		Me._MnuABtn_0.CHECK = 0
		Me._MnuABtn_0.Name = "_MnuABtn_0"
		Me._Line3D1_0.Size = New System.Drawing.Size(440, 3)
		Me._Line3D1_0.Location = New System.Drawing.Point(1, 55)
		Me._Line3D1_0.TabIndex = 13
		Me._Line3D1_0.horizon = -1
		Me._Line3D1_0.Name = "_Line3D1_0"
		Me._MnuBBtn_0.Size = New System.Drawing.Size(42, 42)
		Me._MnuBBtn_0.Location = New System.Drawing.Point(8, 7)
		Me._MnuBBtn_0.TabIndex = 9
		Me._MnuBBtn_0.TX = ""
		Me._MnuBBtn_0.ENAB = -1
		Me._MnuBBtn_0.COLTYPE = 1
		Me._MnuBBtn_0.FOCUSR = -1
		Me._MnuBBtn_0.BCOL = 12632256
		Me._MnuBBtn_0.BCOLO = 12632256
		Me._MnuBBtn_0.FCOL = 0
		Me._MnuBBtn_0.FCOLO = 0
		Me._MnuBBtn_0.MCOL = 16777215
		Me._MnuBBtn_0.MPTR = 1
		Me._MnuBBtn_0.MICON = 0
		Me._MnuBBtn_0.PICN = 0
		Me._MnuBBtn_0.UMCOL = -1
		Me._MnuBBtn_0.SOFT = 0
		Me._MnuBBtn_0.PICPOS = 0
		Me._MnuBBtn_0.NGREY = 0
		Me._MnuBBtn_0.FX = 0
		Me._MnuBBtn_0.HAND = 0
		Me._MnuBBtn_0.CHECK = 0
		Me._MnuBBtn_0.Name = "_MnuBBtn_0"
		Me._MnuBBtn_1.Size = New System.Drawing.Size(42, 42)
		Me._MnuBBtn_1.Location = New System.Drawing.Point(52, 7)
		Me._MnuBBtn_1.TabIndex = 10
		Me._MnuBBtn_1.TX = ""
		Me._MnuBBtn_1.ENAB = -1
		Me._MnuBBtn_1.COLTYPE = 1
		Me._MnuBBtn_1.FOCUSR = -1
		Me._MnuBBtn_1.BCOL = 12632256
		Me._MnuBBtn_1.BCOLO = 12632256
		Me._MnuBBtn_1.FCOL = 0
		Me._MnuBBtn_1.FCOLO = 0
		Me._MnuBBtn_1.MCOL = 16777215
		Me._MnuBBtn_1.MPTR = 1
		Me._MnuBBtn_1.MICON = 0
		Me._MnuBBtn_1.PICN = 0
		Me._MnuBBtn_1.UMCOL = -1
		Me._MnuBBtn_1.SOFT = 0
		Me._MnuBBtn_1.PICPOS = 0
		Me._MnuBBtn_1.NGREY = 0
		Me._MnuBBtn_1.FX = 0
		Me._MnuBBtn_1.HAND = 0
		Me._MnuBBtn_1.CHECK = 0
		Me._MnuBBtn_1.Name = "_MnuBBtn_1"
		Me._MnuBBtn_2.Size = New System.Drawing.Size(42, 42)
		Me._MnuBBtn_2.Location = New System.Drawing.Point(96, 7)
		Me._MnuBBtn_2.TabIndex = 11
		Me._MnuBBtn_2.TX = ""
		Me._MnuBBtn_2.ENAB = -1
		Me._MnuBBtn_2.COLTYPE = 1
		Me._MnuBBtn_2.FOCUSR = -1
		Me._MnuBBtn_2.BCOL = 12632256
		Me._MnuBBtn_2.BCOLO = 12632256
		Me._MnuBBtn_2.FCOL = 0
		Me._MnuBBtn_2.FCOLO = 0
		Me._MnuBBtn_2.MCOL = 16777215
		Me._MnuBBtn_2.MPTR = 1
		Me._MnuBBtn_2.MICON = 0
		Me._MnuBBtn_2.PICN = 0
		Me._MnuBBtn_2.UMCOL = -1
		Me._MnuBBtn_2.SOFT = 0
		Me._MnuBBtn_2.PICPOS = 0
		Me._MnuBBtn_2.NGREY = 0
		Me._MnuBBtn_2.FX = 0
		Me._MnuBBtn_2.HAND = 0
		Me._MnuBBtn_2.CHECK = 0
		Me._MnuBBtn_2.Name = "_MnuBBtn_2"
		Me._MnuBBtn_3.Size = New System.Drawing.Size(42, 42)
		Me._MnuBBtn_3.Location = New System.Drawing.Point(140, 7)
		Me._MnuBBtn_3.TabIndex = 12
		Me._MnuBBtn_3.TX = ""
		Me._MnuBBtn_3.ENAB = -1
		Me._MnuBBtn_3.COLTYPE = 1
		Me._MnuBBtn_3.FOCUSR = -1
		Me._MnuBBtn_3.BCOL = 12632256
		Me._MnuBBtn_3.BCOLO = 12632256
		Me._MnuBBtn_3.FCOL = 0
		Me._MnuBBtn_3.FCOLO = 0
		Me._MnuBBtn_3.MCOL = 16777215
		Me._MnuBBtn_3.MPTR = 1
		Me._MnuBBtn_3.MICON = 0
		Me._MnuBBtn_3.PICN = 0
		Me._MnuBBtn_3.UMCOL = -1
		Me._MnuBBtn_3.SOFT = 0
		Me._MnuBBtn_3.PICPOS = 0
		Me._MnuBBtn_3.NGREY = 0
		Me._MnuBBtn_3.FX = 0
		Me._MnuBBtn_3.HAND = 0
		Me._MnuBBtn_3.CHECK = 0
		Me._MnuBBtn_3.Name = "_MnuBBtn_3"
		_SerICmb_1.OcxState = CType(resources.GetObject("_SerICmb_1.OcxState"), System.Windows.Forms.AxHost.State)
		Me._SerICmb_1.Size = New System.Drawing.Size(124, 22)
		Me._SerICmb_1.Location = New System.Drawing.Point(77, 102)
		Me._SerICmb_1.TabIndex = 27
		Me._SerICmb_1.Name = "_SerICmb_1"
		_SerICmb_0.OcxState = CType(resources.GetObject("_SerICmb_0.OcxState"), System.Windows.Forms.AxHost.State)
		Me._SerICmb_0.Size = New System.Drawing.Size(124, 22)
		Me._SerICmb_0.Location = New System.Drawing.Point(77, 70)
		Me._SerICmb_0.TabIndex = 25
		Me._SerICmb_0.Name = "_SerICmb_0"
		SerLv.OcxState = CType(resources.GetObject("SerLv.OcxState"), System.Windows.Forms.AxHost.State)
		Me.SerLv.Size = New System.Drawing.Size(208, 71)
		Me.SerLv.Location = New System.Drawing.Point(210, 70)
		Me.SerLv.TabIndex = 29
		Me.SerLv.Name = "SerLv"
		Me._SerBtn_0.Size = New System.Drawing.Size(27, 23)
		Me._SerBtn_0.Location = New System.Drawing.Point(173, 137)
		Me._SerBtn_0.TabIndex = 32
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
		Me._SerBtn_1.Location = New System.Drawing.Point(145, 137)
		Me._SerBtn_1.TabIndex = 33
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
		Me.SerQty.AutoSize = False
		Me.SerQty.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.SerQty.BackColor = System.Drawing.Color.White
		Me.SerQty.Enabled = False
		Me.SerQty.Size = New System.Drawing.Size(52, 21)
		Me.SerQty.Location = New System.Drawing.Point(76, 138)
		Me.SerQty.TabIndex = 31
		Me.SerQty.Text = "1"
		Me.SerQty.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.SerQty.AcceptsReturn = True
		Me.SerQty.CausesValidation = True
		Me.SerQty.ForeColor = System.Drawing.SystemColors.WindowText
		Me.SerQty.HideSelection = True
		Me.SerQty.ReadOnly = False
		Me.SerQty.Maxlength = 0
		Me.SerQty.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.SerQty.MultiLine = False
		Me.SerQty.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.SerQty.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.SerQty.TabStop = True
		Me.SerQty.Visible = True
		Me.SerQty.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.SerQty.Name = "SerQty"
		Me.SerScroll1.Size = New System.Drawing.Size(11, 22)
		Me.SerScroll1.Location = New System.Drawing.Point(131, 137)
		Me.SerScroll1.Maximum = 999
		Me.SerScroll1.Minimum = 1
		Me.SerScroll1.TabIndex = 30
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
		Me.SerJumlah.AutoSize = False
		Me.SerJumlah.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.SerJumlah.BackColor = System.Drawing.Color.FromARGB(224, 224, 224)
		Me.SerJumlah.Size = New System.Drawing.Size(137, 20)
		Me.SerJumlah.Location = New System.Drawing.Point(280, 147)
		Me.SerJumlah.TabIndex = 36
		Me.SerJumlah.TabStop = False
		Me.SerJumlah.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.SerJumlah.AcceptsReturn = True
		Me.SerJumlah.CausesValidation = True
		Me.SerJumlah.Enabled = True
		Me.SerJumlah.ForeColor = System.Drawing.SystemColors.WindowText
		Me.SerJumlah.HideSelection = True
		Me.SerJumlah.ReadOnly = False
		Me.SerJumlah.Maxlength = 0
		Me.SerJumlah.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.SerJumlah.MultiLine = False
		Me.SerJumlah.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.SerJumlah.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.SerJumlah.Visible = True
		Me.SerJumlah.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.SerJumlah.Name = "SerJumlah"
		Me._SerLbl_3.Text = "Total :"
		Me._SerLbl_3.Font = New System.Drawing.Font("Verdana", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._SerLbl_3.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me._SerLbl_3.Size = New System.Drawing.Size(43, 15)
		Me._SerLbl_3.Location = New System.Drawing.Point(235, 149)
		Me._SerLbl_3.TabIndex = 35
		Me._SerLbl_3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._SerLbl_3.BackColor = System.Drawing.Color.Transparent
		Me._SerLbl_3.Enabled = True
		Me._SerLbl_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._SerLbl_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._SerLbl_3.UseMnemonic = True
		Me._SerLbl_3.Visible = True
		Me._SerLbl_3.AutoSize = False
		Me._SerLbl_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._SerLbl_3.Name = "_SerLbl_3"
		Me._SerLbl_2.Text = "Quantity :"
		Me._SerLbl_2.ForeColor = System.Drawing.Color.Black
		Me._SerLbl_2.Size = New System.Drawing.Size(64, 21)
		Me._SerLbl_2.Location = New System.Drawing.Point(5, 140)
		Me._SerLbl_2.TabIndex = 34
		Me._SerLbl_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._SerLbl_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._SerLbl_2.BackColor = System.Drawing.Color.Transparent
		Me._SerLbl_2.Enabled = True
		Me._SerLbl_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._SerLbl_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._SerLbl_2.UseMnemonic = True
		Me._SerLbl_2.Visible = True
		Me._SerLbl_2.AutoSize = False
		Me._SerLbl_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._SerLbl_2.Name = "_SerLbl_2"
		Me._SerLbl_1.Text = "Items :"
		Me._SerLbl_1.ForeColor = System.Drawing.Color.Black
		Me._SerLbl_1.Size = New System.Drawing.Size(64, 21)
		Me._SerLbl_1.Location = New System.Drawing.Point(5, 107)
		Me._SerLbl_1.TabIndex = 28
		Me._SerLbl_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._SerLbl_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._SerLbl_1.BackColor = System.Drawing.Color.Transparent
		Me._SerLbl_1.Enabled = True
		Me._SerLbl_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._SerLbl_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._SerLbl_1.UseMnemonic = True
		Me._SerLbl_1.Visible = True
		Me._SerLbl_1.AutoSize = False
		Me._SerLbl_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._SerLbl_1.Name = "_SerLbl_1"
		Me._SerLbl_0.Text = "Category :"
		Me._SerLbl_0.ForeColor = System.Drawing.Color.Black
		Me._SerLbl_0.Size = New System.Drawing.Size(67, 21)
		Me._SerLbl_0.Location = New System.Drawing.Point(5, 72)
		Me._SerLbl_0.TabIndex = 26
		Me._SerLbl_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._SerLbl_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._SerLbl_0.BackColor = System.Drawing.Color.Transparent
		Me._SerLbl_0.Enabled = True
		Me._SerLbl_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._SerLbl_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._SerLbl_0.UseMnemonic = True
		Me._SerLbl_0.Visible = True
		Me._SerLbl_0.AutoSize = False
		Me._SerLbl_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._SerLbl_0.Name = "_SerLbl_0"
		Me.Controls.Add(Dock2)
		Me.Controls.Add(Dock1)
		Me.Controls.Add(_Line3D1_0)
		Me.Controls.Add(_MnuBBtn_0)
		Me.Controls.Add(_MnuBBtn_1)
		Me.Controls.Add(_MnuBBtn_2)
		Me.Controls.Add(_MnuBBtn_3)
		Me.Controls.Add(_SerICmb_1)
		Me.Controls.Add(_SerICmb_0)
		Me.Controls.Add(SerLv)
		Me.Controls.Add(_SerBtn_0)
		Me.Controls.Add(_SerBtn_1)
		Me.Controls.Add(SerQty)
		Me.Controls.Add(SerScroll1)
		Me.Controls.Add(SerJumlah)
		Me.Controls.Add(_SerLbl_3)
		Me.Controls.Add(_SerLbl_2)
		Me.Controls.Add(_SerLbl_1)
		Me.Controls.Add(_SerLbl_0)
		Me.Dock2.Controls.Add(_PayBox_0)
		Me.Dock2.Controls.Add(_PayBox_1)
		Me.Dock2.Controls.Add(PayPrice)
		Me.Dock2.Controls.Add(_PayLbl_0)
		Me.Dock2.Controls.Add(_PayLbl_1)
		Me.Dock2.Controls.Add(_PayImg_0)
		Me.Dock2.Controls.Add(_PayImg_1)
		Me._PayBox_0.Controls.Add(PayBal)
		Me._PayBox_0.Controls.Add(PayTotal)
		Me._PayBox_0.Controls.Add(_PayLbl_3)
		Me._PayBox_0.Controls.Add(_PayLbl_2)
		Me._PayBox_1.Controls.Add(PayTime)
		Me.Dock1.Controls.Add(Line3D2)
		Me.Dock1.Controls.Add(_PayBtn_1)
		Me.Dock1.Controls.Add(PayRcv)
		Me.Dock1.Controls.Add(_PayBtn_0)
		Me.Dock1.Controls.Add(_MnuABtn_3)
		Me.Dock1.Controls.Add(_MnuABtn_2)
		Me.Dock1.Controls.Add(_MnuABtn_1)
		Me.Dock1.Controls.Add(_MnuABtn_0)
		Me.Line3D1.SetIndex(_Line3D1_0, CType(0, Short))
		Me.MnuABtn.SetIndex(_MnuABtn_3, CType(3, Short))
		Me.MnuABtn.SetIndex(_MnuABtn_2, CType(2, Short))
		Me.MnuABtn.SetIndex(_MnuABtn_1, CType(1, Short))
		Me.MnuABtn.SetIndex(_MnuABtn_0, CType(0, Short))
		Me.MnuBBtn.SetIndex(_MnuBBtn_0, CType(0, Short))
		Me.MnuBBtn.SetIndex(_MnuBBtn_1, CType(1, Short))
		Me.MnuBBtn.SetIndex(_MnuBBtn_2, CType(2, Short))
		Me.MnuBBtn.SetIndex(_MnuBBtn_3, CType(3, Short))
		Me.PayBox.SetIndex(_PayBox_0, CType(0, Short))
		Me.PayBox.SetIndex(_PayBox_1, CType(1, Short))
		Me.PayBtn.SetIndex(_PayBtn_1, CType(1, Short))
		Me.PayBtn.SetIndex(_PayBtn_0, CType(0, Short))
		Me.PayImg.SetIndex(_PayImg_0, CType(0, Short))
		Me.PayImg.SetIndex(_PayImg_1, CType(1, Short))
		Me.PayLbl.SetIndex(_PayLbl_3, CType(3, Short))
		Me.PayLbl.SetIndex(_PayLbl_2, CType(2, Short))
		Me.PayLbl.SetIndex(_PayLbl_0, CType(0, Short))
		Me.PayLbl.SetIndex(_PayLbl_1, CType(1, Short))
		Me.SerBtn.SetIndex(_SerBtn_0, CType(0, Short))
		Me.SerBtn.SetIndex(_SerBtn_1, CType(1, Short))
		Me.SerICmb.SetIndex(_SerICmb_1, CType(1, Short))
		Me.SerICmb.SetIndex(_SerICmb_0, CType(0, Short))
		Me.SerLbl.SetIndex(_SerLbl_3, CType(3, Short))
		Me.SerLbl.SetIndex(_SerLbl_2, CType(2, Short))
		Me.SerLbl.SetIndex(_SerLbl_1, CType(1, Short))
		Me.SerLbl.SetIndex(_SerLbl_0, CType(0, Short))
		CType(Me.SerLbl, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.SerICmb, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.SerBtn, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.PayLbl, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.PayImg, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.PayBtn, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.PayBox, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.MnuBBtn, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.MnuABtn, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Line3D1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.SerLv, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me._SerICmb_0, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me._SerICmb_1, System.ComponentModel.ISupportInitialize).EndInit()
	End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As FrmLogout
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As FrmLogout
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New FrmLogout()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	'WIDTH
	' Expanded = 6710 x 2690
	' Normal = 6710 x 930
	
	Private Rs As DAO.Recordset
	Private LoginTrue As Boolean
	
	Public PrePaid As Boolean
	Public PcName As String
	Public pcCusName As String
	Public pcInTime As String
	Public pcOutTime As String
	Public pcTotalTime As String
	Public pcPaid As String
	Public SerTotal As Double
	
	
	'UPGRADE_WARNING: Event PayRcv.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2075"'
	Private Sub PayRcv_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles PayRcv.TextChanged
		If Trim(PayRcv.Text) <> "" And IsNumeric(PayRcv.Text) = True Then
			PayBal.Text = VB6.Format(CDbl(PayRcv.Text) - (CDbl(pcPaid) + SerTotal), "#0.00")
		Else
			PayBal.Text = ""
		End If
	End Sub
	
	Private Sub PayRcv_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles PayRcv.KeyUp
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		If KeyCode = System.Windows.Forms.Keys.Return Then
			Call LogUsage()
			Me.Close()
		End If
	End Sub
	
	Private Sub PayBtn_Click(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles PayBtn.Click
		Dim Index As Short = PayBtn.GetIndex(Sender)
		Select Case Index
			Case 0
				Call PayPrice_KeyUp(PayPrice, New System.Windows.Forms.KeyEventArgs(13 Or 1 * &H10000))
				Call LogUsage()
				Call UpdatePanel(SelText)
				AgentSel.AgnRecoverRemove()
				Me.Close()
			Case 1
				Dock2.PageFlip = True Xor Dock2.PageFlip
		End Select
	End Sub
	
	Private Sub PayPrice_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles PayPrice.KeyUp
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Harga As Object
		Dim ret As Object
		If KeyCode = 13 Then
			If PayPrice.Text = "" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object ret. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				ret = MsgBox("Void this PC usage ?", MsgBoxStyle.OKCancel, CbMsgWarn)
				If ret = MsgBoxResult.OK Then
					Me.Close()
				Else
					PayPrice.Text = pcPaid
					PayPrice.SelectionStart = 1
					PayPrice.SelectionLength = Len(PayPrice.Text)
					Exit Sub
				End If
			End If
			If IsNumeric(Harga) = False Then
				MsgBox(MB(4), MsgBoxStyle.Information, CbMsgWarn)
				PayPrice.Focus()
				PayPrice.Text = pcPaid
				Exit Sub
			End If
			
			pcPaid = PayPrice.Text
			PayTotal.Text = VB6.Format(CDbl(pcPaid) + SerTotal, "#0.00")
			PayPrice.Focus()
		End If
	End Sub
	
	
	Private Sub MnuABtn_Click(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles MnuABtn.Click
		Dim Index As Short = MnuABtn.GetIndex(Sender)
		Dim cItm As MSComctlLib.ListItem
		cItm = FrmMain.DefInstance.Lv1.SelectedItem
		
		Select Case Index
			Case 0 'Login
				Me.Close()
				FrmLogin.DefInstance.Show()
			Case 1 'Logout
				Call ViewMode(10)
			Case 2 'Renew
				FrmLogin.DefInstance.Sambung = True
				Me.Close()
				FrmLogin.DefInstance.Show()
			Case 3 'cancel
				Me.Close()
		End Select
	End Sub
	
	Private Sub MnuBBtn_Click(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles MnuBBtn.Click
		Dim Index As Short = MnuBBtn.GetIndex(Sender)
		Dim strMsg As String
		Select Case Index
			Case 0
				If CDbl(Mid(CbUserAccess, 3, 1)) = 0 Then MsgBox(MB(10), MsgBoxStyle.OKOnly, CbMsgWarn) : Exit Sub
				AgentSel.NetSend("//kunci:0")
				Me.Close()
			Case 1
				AgentSel.NetSend("//kunci:1")
				Me.Close()
			Case 2
				Me.Close()
				FrmAgnMsg.DefInstance.Show()
			Case 3
				Me.Close()
				strMsg = MgoInpt.GetInput("Sila masukkan mesej tiker anda", VisualSuite1.eButStyle.BtnClose)
				If Trim(strMsg) <> "" Then AgentSel.NetSend("tiker:" & strMsg)
		End Select
	End Sub
	
	'UPGRADE_NOTE: SerScroll1.Change was changed from an event to a procedure. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2010"'
	'UPGRADE_WARNING: VScrollBar event SerScroll1.Change has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2065"'
	Private Sub SerScroll1_Change(ByVal newScrollValue As Integer)
		Dim QtySer As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object QtySer. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		QtySer = 1000 - newScrollValue
	End Sub
	
	Private Sub SerLv_ItemClick(ByVal eventSender As System.Object, ByVal eventArgs As AxMSComctlLib.ListViewEvents_ItemClickEvent) Handles SerLv.ItemClick
		Dim ItemTotal As Double
		
		ItemTotal = CDbl(eventArgs.Item.SubItems(1)) * CShort(eventArgs.Item.SubItems(2))
		SerTotal = CDbl(Crnc & " " & ItemTotal & " / " & Crnc & " " & SerTotal)
	End Sub
	
	Private Sub SerBtn_Click(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles SerBtn.Click
		Dim Index As Short = SerBtn.GetIndex(Sender)
		Dim g As Object
		Dim SCbItm As MSComctlLib.ComboItem
		Dim lvItm, fItm As MSComctlLib.ListItem
		
		Select Case Index
			Case 1
				SCbItm = SerICmb(1).SelectedItem
				If CDbl(SerQty.Text) > 0 Then
					fItm = SerLv.FindItem(SCbItm.Text)
					
					If fItm Is Nothing Then
						lvItm = SerLv.ListItems.Add( , SCbItm.Key, SCbItm.Text)
						'UPGRADE_WARNING: Couldn't resolve default property of object SCbItm.Tag. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
						lvItm.SubItems(1) = SCbItm.Tag
						lvItm.SubItems(2) = SerQty.Text
					Else
						fItm.SubItems(2) = SerQty.Text
					End If
				Else
					MsgBox("Please enter quantity !", MsgBoxStyle.OKOnly, CbMsgWarn)
					SerQty.SelectionStart = 1
					SerQty.SelectionLength = Len(SerQty.Text)
					Exit Sub
				End If
			Case 0
				If SerLv.ListItems.Count = 0 Then Exit Sub
				SerLv.ListItems.Remove((SerLv.SelectedItem.Index))
		End Select
		
		'recalculate total
		SerTotal = 0
		For g = 1 To SerLv.ListItems.Count
			SerTotal = SerTotal + (CDbl(SerLv.ListItems(g).SubItems(1)) * CShort(SerLv.ListItems(g).SubItems(2)))
		Next g
		SerJumlah.Text = Crnc & " " & VB6.Format(SerTotal, "#0.00")
		
		'recalculate overall total
		PayTotal.Text = VB6.Format(CDbl(pcPaid) + SerTotal, "#0.00")
	End Sub
	
	Private Sub SerICmb_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles SerICmb.ClickEvent
		Dim Index As Short = SerICmb.GetIndex(eventSender)
		Select Case Index
			Case 0
				Call SerControl(True)
				Call SerLoadItems()
		End Select
	End Sub
	
	
	'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
	' FUNCTION
	'
	'
	'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Interface View Mode
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Public Sub ViewMode(ByRef Mode As Short)
		' Mode
		'  0 = Normal, button only, login function
		'  1 = Normal, button only, logout function
		'  2 = Normal, button only, continue,prepaid function
		'  10 = Expanded, button + price, log user function
		Select Case Mode
			Case 0
				Me.Height = VB6.TwipsToPixelsY(930)
				LoginTrue = True
				MnuABtn(0).Enabled = True
				MnuABtn(1).Enabled = False
				MnuABtn(2).Enabled = False
			Case 1
				Me.Height = VB6.TwipsToPixelsY(930)
				LoginTrue = False
				MnuABtn(0).Enabled = False
				MnuABtn(1).Enabled = True
				MnuABtn(2).Enabled = False
				Call LoadUsage()
			Case 2
				Me.Height = VB6.TwipsToPixelsY(930)
				LoginTrue = False
				MnuABtn(0).Enabled = False
				MnuABtn(1).Enabled = True
				MnuABtn(2).Enabled = True
				Call LoadUsage()
			Case 10
				Me.Height = VB6.TwipsToPixelsY(2690)
				PayRcv.Enabled = True
				PayBtn(0).Enabled = True
				PayBtn(1).Enabled = True
				'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(tukarharga). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				If SetAmbil("tukarharga") = 0 Then PayPrice.ReadOnly = True
				'UPGRADE_WARNING: Couldn't resolve default property of object SerICmb().ImageList. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				SerICmb(0).ImageList = FrmMain.DefInstance.ImgListSnm
				SerICmb(0).ComboItems.Clear()
				Call LoadPosCatCB(SerICmb(0), (FrmMain.DefInstance.ImgListSnm))
		End Select
	End Sub
	
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Service Controls Enable\Disable
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	'UPGRADE_NOTE: Enabled was upgraded to Enabled_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
	Private Sub SerControl(ByRef Enabled_Renamed As Boolean)
		If Enabled_Renamed = True Then
			SerICmb(1).Enabled = True
			SerBtn(0).Enabled = True
			SerBtn(1).Enabled = True
			SerQty.Enabled = True
			SerLv.Enabled = True
			SerScroll1.Enabled = True
		Else
			SerICmb(1).Text = VS(1)
			SerICmb(1).Enabled = False
			SerBtn(0).Enabled = False
			SerBtn(1).Enabled = False
			SerQty.Enabled = False
			SerLv.Enabled = False
			SerScroll1.Enabled = False
		End If
	End Sub
	
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Service Load Items
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Private Sub SerLoadItems()
		Dim ret As Object
		Dim tInt As Object
		Dim tWd As Object
		Dim g As Object
		Dim Rss As DAO.Recordset
		Dim CbItm As MSComctlLib.ComboItem
		Rss = uSDB.OpenRecordset("pos-items", DAO.RecordsetTypeEnum.dbOpenSnapshot)
		
		SerICmb(1).ComboItems.Clear()
		Rss.Filter = "groupid = '" & Mid(SerICmb(0).SelectedItem.Key, 2) & "'"
		
		Rs = Rss.OpenRecordset
		
		If Rs.BOF = True Then Exit Sub
		With Rs
			.MoveFirst()
			Do Until .EOF = True
				CbItm = SerICmb(1).ComboItems.Add( , .Fields("id"), .Fields("Nama"))
				CbItm.let_Tag(.Fields("Harga"))
				.MoveNext()
			Loop 
		End With
		
		'resize the list width
		For g = 1 To SerICmb(1).ComboItems.Count
			'UPGRADE_ISSUE: Form method FrmLogout.TextWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			'UPGRADE_WARNING: Couldn't resolve default property of object tWd. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			tWd = TextWidth(SerICmb(1).ComboItems(g).Text)
			'UPGRADE_WARNING: Couldn't resolve default property of object tInt. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Couldn't resolve default property of object tWd. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Couldn't resolve default property of object tWd. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Couldn't resolve default property of object tInt. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			If tWd > tInt Then tInt = tWd
		Next g
		'UPGRADE_WARNING: Couldn't resolve default property of object tInt. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		tInt = (tInt / 15) + 40
		'UPGRADE_WARNING: Couldn't resolve default property of object tInt. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object ret. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ret = SendMessage(SerICmb(1).Handle.ToInt32, CB_SETDROPPEDWIDTH, tInt, 0)
		
		'default selection
		SerICmb(1).ComboItems(1).Selected = True
	End Sub
	
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Load Customer Usage
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Private Sub LoadUsage()
		Dim tMinit, tJam, uMin As Short
		Dim mHo As String
		Dim hO, mIno As Double
		Dim s_cOutTime, s_cInTime As String
		
		If AgentSel.CustomerFlag = "g" Then
			hO = CDbl(AgentSel.CusGetPrice)
			mHo = AgentSel.CusGetTimeUse
			mIno = CDbl(AgentSel.CusGetTimeUse(True))
			
			s_cInTime = CStr(TimeValue(AgentSel.CustomerTimeIn))
			s_cOutTime = CStr(TimeOfDay)
		Else
			'pengiraan harga dan masa
			If VB.Left(AgentSel.CustomerFlag, 1) = "p" Then
				hO = CDbl(Mid(AgentSel.CustomerFlag, 2))
				'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				mIno = hO / CDbl(SetAmbil("harga"))
			ElseIf VB.Left(AgentSel.CustomerFlag, 1) = "f" Then 
				uMin = CDbl(Mid(AgentSel.CustomerFlag, 2))
				'UPGRADE_WARNING: Couldn't resolve default property of object SetAmbil(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				hO = uMin * CDbl(SetAmbil("harga"))
				mIno = uMin
			End If
			tJam = mIno \ 60
			'UPGRADE_WARNING: Mod has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
			tMinit = mIno Mod 60
			mHo = tJam & " " & VS(7) & ", " & tMinit & " " & VS(8)
			s_cInTime = CStr(TimeValue(AgentSel.CustomerTimeIn))
			s_cOutTime = CStr(TimeValue(AgentSel.CustomerTimeOut))
		End If
		
		'masukkan kedalam variable Form
		PcName = AgentSel.AgentName
		pcCusName = AgentSel.CustomerName
		pcInTime = s_cInTime
		pcOutTime = s_cOutTime
		pcTotalTime = CStr(mIno)
		pcPaid = CStr(hO)
		
		'prepare for output
		PayTime.Caption = mHo
		PayPrice.Text = VB6.Format(hO, "#0.00")
		PayTotal.Text = VB6.Format(hO, "#0.00")
	End Sub
	
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' LogUsage Customer
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Private Sub LogUsage()
		Dim dDay As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object dDay. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		dDay = WeekDay(Today)
		'UPGRADE_WARNING: Couldn't resolve default property of object dDay. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Choose(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		dDay = Choose(dDay, "Ahad", "Isnin", "Selasa", "Rabu", "Khamis", "Jumaat", "Sabtu")
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2065"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		
		'{ save dalam table usage }'
		SavePcUsage(PcName, pcCusName, pcInTime, pcOutTime, pcPaid)
		'{ save dalam table bulanan }'
		SavePcBulanan(PcName, pcTotalTime, pcPaid)
		'{ rekod jualan harian }'
		SavePcHarian(pcPaid)
		'{ save dalam table graf-mingguan(untuk mengira hari graf hari) }'
		SavePcMingguan(dDay, pcPaid)
		'{ save dalam table pelanggan (!!!tolong pikir skit untuk CusID tuh) }'
		SavePelanggan(pcCusName, "", pcTotalTime, pcPaid)
		'{ save transaksi POS }'
		SavePosTrans(SerLv.ListItems)
		
		'reset agent, main ui..
		AgentSel.CusStop()
		AgentSel.NetSend("//logout")
		
		SerTotal = 0
		pcPaid = CStr(0)
		pcInTime = ""
		pcOutTime = ""
		pcTotalTime = CStr(0)
		PcName = ""
		pcCusName = ""
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2065"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
	End Sub
	Private Sub SerScroll1_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.ScrollEventArgs) Handles SerScroll1.Scroll
		Select Case eventArgs.type
			Case System.Windows.Forms.ScrollEventType.EndScroll
				SerScroll1_Change(eventArgs.newValue)
		End Select
	End Sub
End Class