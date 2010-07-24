Option Strict Off
Option Explicit On
Friend Class FrmLogin
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
	Public WithEvents BtnOk As XpButton
	Public WithEvents BtnKo As XpButton
	Public WithEvents Picture5 As System.Windows.Forms.Panel
	Public WithEvents uLine3D1 As Line3D
	Public WithEvents cbPaid As System.Windows.Forms.ComboBox
	Public WithEvents _Opt1_3 As System.Windows.Forms.RadioButton
	Public WithEvents _Opt1_2 As System.Windows.Forms.RadioButton
	Public WithEvents _Opt1_1 As System.Windows.Forms.RadioButton
	Public WithEvents Combo3 As System.Windows.Forms.ComboBox
	Public WithEvents NamaPc As Label3D
	Public WithEvents Combo2 As System.Windows.Forms.ComboBox
	Public WithEvents Combo1 As System.Windows.Forms.ComboBox
	Public WithEvents Text1 As System.Windows.Forms.TextBox
	Public WithEvents LblCrnc As System.Windows.Forms.Label
	Public WithEvents _Lbl1_1 As System.Windows.Forms.Label
	Public WithEvents _Lbl1_3 As System.Windows.Forms.Label
	Public WithEvents _Lbl1_2 As System.Windows.Forms.Label
	Public WithEvents _Lbl1_0 As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents Lbl1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents Opt1 As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmLogin))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.Picture5 = New System.Windows.Forms.Panel
		Me.BtnOk = New XpButton
		Me.BtnKo = New XpButton
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.uLine3D1 = New Line3D
		Me.cbPaid = New System.Windows.Forms.ComboBox
		Me._Opt1_3 = New System.Windows.Forms.RadioButton
		Me._Opt1_2 = New System.Windows.Forms.RadioButton
		Me._Opt1_1 = New System.Windows.Forms.RadioButton
		Me.Combo3 = New System.Windows.Forms.ComboBox
		Me.NamaPc = New Label3D
		Me.Combo2 = New System.Windows.Forms.ComboBox
		Me.Combo1 = New System.Windows.Forms.ComboBox
		Me.Text1 = New System.Windows.Forms.TextBox
		Me.LblCrnc = New System.Windows.Forms.Label
		Me._Lbl1_1 = New System.Windows.Forms.Label
		Me._Lbl1_3 = New System.Windows.Forms.Label
		Me._Lbl1_2 = New System.Windows.Forms.Label
		Me._Lbl1_0 = New System.Windows.Forms.Label
		Me.Lbl1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.Opt1 = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(components)
		CType(Me.Lbl1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Opt1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.ControlBox = False
		Me.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.ClientSize = New System.Drawing.Size(288, 228)
		Me.Location = New System.Drawing.Point(3, 3)
		Me.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.AutoScaleBaseSize = New System.Drawing.Size(7, 14)
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "FrmLogin"
		Me.Picture5.BackColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me.Picture5.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Picture5.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Picture5.Size = New System.Drawing.Size(24, 228)
		Me.Picture5.Location = New System.Drawing.Point(264, 0)
		Me.Picture5.TabIndex = 14
		Me.Picture5.Dock = System.Windows.Forms.DockStyle.None
		Me.Picture5.CausesValidation = True
		Me.Picture5.Enabled = True
		Me.Picture5.Cursor = System.Windows.Forms.Cursors.Default
		Me.Picture5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Picture5.TabStop = True
		Me.Picture5.Visible = True
		Me.Picture5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Picture5.Name = "Picture5"
		Me.BtnOk.Size = New System.Drawing.Size(24, 23)
		Me.BtnOk.Location = New System.Drawing.Point(0, 205)
		Me.BtnOk.TabIndex = 17
		Me.ToolTip1.SetToolTip(Me.BtnOk, "Add new employee.")
		Me.BtnOk.TX = ""
		Me.BtnOk.ENAB = -1
		Me.BtnOk.COLTYPE = 1
		Me.BtnOk.FOCUSR = -1
		Me.BtnOk.BCOL = 12632256
		Me.BtnOk.BCOLO = 12632256
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
		Me.BtnKo.Size = New System.Drawing.Size(24, 23)
		Me.BtnKo.Location = New System.Drawing.Point(0, 182)
		Me.BtnKo.TabIndex = 18
		Me.ToolTip1.SetToolTip(Me.BtnKo, "Add new employee.")
		Me.BtnKo.TX = ""
		Me.BtnKo.ENAB = -1
		Me.BtnKo.COLTYPE = 1
		Me.BtnKo.FOCUSR = -1
		Me.BtnKo.BCOL = 12632256
		Me.BtnKo.BCOLO = 12632256
		Me.BtnKo.FCOL = 0
		Me.BtnKo.FCOLO = 0
		Me.BtnKo.MCOL = 16777215
		Me.BtnKo.MPTR = 1
		Me.BtnKo.MICON = 0
		Me.BtnKo.PICN = 0
		Me.BtnKo.UMCOL = -1
		Me.BtnKo.SOFT = 0
		Me.BtnKo.PICPOS = 0
		Me.BtnKo.NGREY = 0
		Me.BtnKo.FX = 0
		Me.BtnKo.HAND = 0
		Me.BtnKo.CHECK = 0
		Me.BtnKo.Name = "BtnKo"
		Me.Frame1.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.Size = New System.Drawing.Size(260, 231)
		Me.Frame1.Location = New System.Drawing.Point(2, -5)
		Me.Frame1.TabIndex = 8
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Name = "Frame1"
		Me.uLine3D1.Size = New System.Drawing.Size(250, 3)
		Me.uLine3D1.Location = New System.Drawing.Point(5, 100)
		Me.uLine3D1.TabIndex = 16
		Me.uLine3D1.horizon = -1
		Me.uLine3D1.Name = "uLine3D1"
		Me.cbPaid.BackColor = System.Drawing.Color.White
		Me.cbPaid.Enabled = False
		Me.cbPaid.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cbPaid.Size = New System.Drawing.Size(72, 22)
		Me.cbPaid.Location = New System.Drawing.Point(75, 150)
		Me.cbPaid.Items.AddRange(New Object(){"1.00", "1.50", "2.00", "2.50", "3.00"})
		Me.cbPaid.TabIndex = 4
		Me.cbPaid.Text = "1.00"
		Me.cbPaid.CausesValidation = True
		Me.cbPaid.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cbPaid.IntegralHeight = True
		Me.cbPaid.Cursor = System.Windows.Forms.Cursors.Default
		Me.cbPaid.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cbPaid.Sorted = False
		Me.cbPaid.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cbPaid.TabStop = True
		Me.cbPaid.Visible = True
		Me.cbPaid.Name = "cbPaid"
		Me._Opt1_3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._Opt1_3.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me._Opt1_3.Text = "Fixed Time"
		Me._Opt1_3.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Opt1_3.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me._Opt1_3.Size = New System.Drawing.Size(114, 18)
		Me._Opt1_3.Location = New System.Drawing.Point(19, 177)
		Me._Opt1_3.TabIndex = 5
		Me._Opt1_3.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._Opt1_3.CausesValidation = True
		Me._Opt1_3.Enabled = True
		Me._Opt1_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._Opt1_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Opt1_3.Appearance = System.Windows.Forms.Appearance.Normal
		Me._Opt1_3.TabStop = True
		Me._Opt1_3.Checked = False
		Me._Opt1_3.Visible = True
		Me._Opt1_3.Name = "_Opt1_3"
		Me._Opt1_2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._Opt1_2.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me._Opt1_2.Text = "Pre Paid"
		Me._Opt1_2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Opt1_2.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me._Opt1_2.Size = New System.Drawing.Size(114, 18)
		Me._Opt1_2.Location = New System.Drawing.Point(19, 129)
		Me._Opt1_2.TabIndex = 3
		Me._Opt1_2.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._Opt1_2.CausesValidation = True
		Me._Opt1_2.Enabled = True
		Me._Opt1_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._Opt1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Opt1_2.Appearance = System.Windows.Forms.Appearance.Normal
		Me._Opt1_2.TabStop = True
		Me._Opt1_2.Checked = False
		Me._Opt1_2.Visible = True
		Me._Opt1_2.Name = "_Opt1_2"
		Me._Opt1_1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._Opt1_1.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me._Opt1_1.Text = "Pay As U Go"
		Me._Opt1_1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Opt1_1.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me._Opt1_1.Size = New System.Drawing.Size(114, 18)
		Me._Opt1_1.Location = New System.Drawing.Point(19, 107)
		Me._Opt1_1.TabIndex = 2
		Me._Opt1_1.Checked = True
		Me._Opt1_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._Opt1_1.CausesValidation = True
		Me._Opt1_1.Enabled = True
		Me._Opt1_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Opt1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Opt1_1.Appearance = System.Windows.Forms.Appearance.Normal
		Me._Opt1_1.TabStop = True
		Me._Opt1_1.Visible = True
		Me._Opt1_1.Name = "_Opt1_1"
		Me.Combo3.BackColor = System.Drawing.Color.White
		Me.Combo3.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Combo3.Size = New System.Drawing.Size(119, 21)
		Me.Combo3.Location = New System.Drawing.Point(127, 69)
		Me.Combo3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.Combo3.TabIndex = 1
		Me.Combo3.CausesValidation = True
		Me.Combo3.Enabled = True
		Me.Combo3.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Combo3.IntegralHeight = True
		Me.Combo3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Combo3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Combo3.Sorted = False
		Me.Combo3.TabStop = True
		Me.Combo3.Visible = True
		Me.Combo3.Name = "Combo3"
		Me.NamaPc.Size = New System.Drawing.Size(234, 20)
		Me.NamaPc.Location = New System.Drawing.Point(13, 14)
		Me.NamaPc.TabIndex = 12
		Me.ToolTip1.SetToolTip(Me.NamaPc, "Nama Pc")
		Me.NamaPc.Name = "NamaPc"
		Me.Combo2.BackColor = System.Drawing.Color.White
		Me.Combo2.Enabled = False
		Me.Combo2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Combo2.Size = New System.Drawing.Size(56, 22)
		Me.Combo2.Location = New System.Drawing.Point(133, 201)
		Me.Combo2.Items.AddRange(New Object(){"5", "10", "20", "30", "40", "50"})
		Me.Combo2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.Combo2.TabIndex = 7
		Me.Combo2.CausesValidation = True
		Me.Combo2.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Combo2.IntegralHeight = True
		Me.Combo2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Combo2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Combo2.Sorted = False
		Me.Combo2.TabStop = True
		Me.Combo2.Visible = True
		Me.Combo2.Name = "Combo2"
		Me.Combo1.BackColor = System.Drawing.Color.White
		Me.Combo1.Enabled = False
		Me.Combo1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Combo1.Size = New System.Drawing.Size(56, 22)
		Me.Combo1.Location = New System.Drawing.Point(43, 201)
		Me.Combo1.Items.AddRange(New Object(){"1", "2", "3", "4", "5", "6", "7", "8", "9", "10"})
		Me.Combo1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.Combo1.TabIndex = 6
		Me.Combo1.CausesValidation = True
		Me.Combo1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Combo1.IntegralHeight = True
		Me.Combo1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Combo1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Combo1.Sorted = False
		Me.Combo1.TabStop = True
		Me.Combo1.Visible = True
		Me.Combo1.Name = "Combo1"
		Me.Text1.AutoSize = False
		Me.Text1.BackColor = System.Drawing.Color.FromARGB(255, 224, 192)
		Me.Text1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text1.Size = New System.Drawing.Size(170, 20)
		Me.Text1.Location = New System.Drawing.Point(73, 41)
		Me.Text1.TabIndex = 0
		Me.Text1.Text = "User"
		Me.Text1.AcceptsReturn = True
		Me.Text1.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text1.CausesValidation = True
		Me.Text1.Enabled = True
		Me.Text1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text1.HideSelection = True
		Me.Text1.ReadOnly = False
		Me.Text1.Maxlength = 0
		Me.Text1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text1.MultiLine = False
		Me.Text1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text1.TabStop = True
		Me.Text1.Visible = True
		Me.Text1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.Text1.Name = "Text1"
		Me.LblCrnc.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.LblCrnc.Text = "RM"
		Me.LblCrnc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.LblCrnc.Size = New System.Drawing.Size(33, 14)
		Me.LblCrnc.Location = New System.Drawing.Point(32, 154)
		Me.LblCrnc.TabIndex = 15
		Me.LblCrnc.BackColor = System.Drawing.Color.Transparent
		Me.LblCrnc.Enabled = True
		Me.LblCrnc.ForeColor = System.Drawing.SystemColors.ControlText
		Me.LblCrnc.Cursor = System.Windows.Forms.Cursors.Default
		Me.LblCrnc.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.LblCrnc.UseMnemonic = True
		Me.LblCrnc.Visible = True
		Me.LblCrnc.AutoSize = True
		Me.LblCrnc.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.LblCrnc.Name = "LblCrnc"
		Me._Lbl1_1.Text = "Customer Type :"
		Me._Lbl1_1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Lbl1_1.Size = New System.Drawing.Size(97, 13)
		Me._Lbl1_1.Location = New System.Drawing.Point(19, 72)
		Me._Lbl1_1.TabIndex = 13
		Me._Lbl1_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Lbl1_1.BackColor = System.Drawing.Color.Transparent
		Me._Lbl1_1.Enabled = True
		Me._Lbl1_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Lbl1_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Lbl1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Lbl1_1.UseMnemonic = True
		Me._Lbl1_1.Visible = True
		Me._Lbl1_1.AutoSize = True
		Me._Lbl1_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Lbl1_1.Name = "_Lbl1_1"
		Me._Lbl1_3.Text = "Minute"
		Me._Lbl1_3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Lbl1_3.Size = New System.Drawing.Size(39, 14)
		Me._Lbl1_3.Location = New System.Drawing.Point(193, 206)
		Me._Lbl1_3.TabIndex = 11
		Me._Lbl1_3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Lbl1_3.BackColor = System.Drawing.Color.Transparent
		Me._Lbl1_3.Enabled = True
		Me._Lbl1_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Lbl1_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._Lbl1_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Lbl1_3.UseMnemonic = True
		Me._Lbl1_3.Visible = True
		Me._Lbl1_3.AutoSize = True
		Me._Lbl1_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Lbl1_3.Name = "_Lbl1_3"
		Me._Lbl1_2.Text = "Hour"
		Me._Lbl1_2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Lbl1_2.Size = New System.Drawing.Size(26, 14)
		Me._Lbl1_2.Location = New System.Drawing.Point(102, 204)
		Me._Lbl1_2.TabIndex = 10
		Me._Lbl1_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Lbl1_2.BackColor = System.Drawing.Color.Transparent
		Me._Lbl1_2.Enabled = True
		Me._Lbl1_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Lbl1_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._Lbl1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Lbl1_2.UseMnemonic = True
		Me._Lbl1_2.Visible = True
		Me._Lbl1_2.AutoSize = True
		Me._Lbl1_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Lbl1_2.Name = "_Lbl1_2"
		Me._Lbl1_0.Text = "Name :"
		Me._Lbl1_0.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Lbl1_0.Size = New System.Drawing.Size(42, 13)
		Me._Lbl1_0.Location = New System.Drawing.Point(19, 43)
		Me._Lbl1_0.TabIndex = 9
		Me._Lbl1_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Lbl1_0.BackColor = System.Drawing.Color.Transparent
		Me._Lbl1_0.Enabled = True
		Me._Lbl1_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Lbl1_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._Lbl1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Lbl1_0.UseMnemonic = True
		Me._Lbl1_0.Visible = True
		Me._Lbl1_0.AutoSize = True
		Me._Lbl1_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Lbl1_0.Name = "_Lbl1_0"
		Me.Controls.Add(Picture5)
		Me.Controls.Add(Frame1)
		Me.Picture5.Controls.Add(BtnOk)
		Me.Picture5.Controls.Add(BtnKo)
		Me.Frame1.Controls.Add(uLine3D1)
		Me.Frame1.Controls.Add(cbPaid)
		Me.Frame1.Controls.Add(_Opt1_3)
		Me.Frame1.Controls.Add(_Opt1_2)
		Me.Frame1.Controls.Add(_Opt1_1)
		Me.Frame1.Controls.Add(Combo3)
		Me.Frame1.Controls.Add(NamaPc)
		Me.Frame1.Controls.Add(Combo2)
		Me.Frame1.Controls.Add(Combo1)
		Me.Frame1.Controls.Add(Text1)
		Me.Frame1.Controls.Add(LblCrnc)
		Me.Frame1.Controls.Add(_Lbl1_1)
		Me.Frame1.Controls.Add(_Lbl1_3)
		Me.Frame1.Controls.Add(_Lbl1_2)
		Me.Frame1.Controls.Add(_Lbl1_0)
		Me.Lbl1.SetIndex(_Lbl1_1, CType(1, Short))
		Me.Lbl1.SetIndex(_Lbl1_3, CType(3, Short))
		Me.Lbl1.SetIndex(_Lbl1_2, CType(2, Short))
		Me.Lbl1.SetIndex(_Lbl1_0, CType(0, Short))
		Me.Opt1.SetIndex(_Opt1_3, CType(3, Short))
		Me.Opt1.SetIndex(_Opt1_2, CType(2, Short))
		Me.Opt1.SetIndex(_Opt1_1, CType(1, Short))
		CType(Me.Opt1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Lbl1, System.ComponentModel.ISupportInitialize).EndInit()
	End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As FrmLogin
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As FrmLogin
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New FrmLogin()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	Public Sambung As Boolean
	
	
	Private Sub BtnOk_Click(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles BtnOk.Click
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Flag Dalam Tag
		'
		' Flag digunakan bagi mengenal pasti jenis dan keadaan
		' pelanggan, samada prepaid,payg dan fixed time.. dan
		' juga untuk mengenalpasti sama keadaan sambung adalah benar..
		' semua flag dan nilai.. akan diletakkan dalam subitems(1)
		' Bagi prepaid, nilai wang yang telah dibayar akan masukkan ke
		' dalam tag bersama flag, bagi fixed time pula.. jumlah masa yang
		' telah ditetapkan akan dimaukkan bersama flag juga, dan untuk payg hanya
		' flag g sahaja..
		'
		'   Contoh :
		'       f10 = Fixed Time dan 10 minit
		'       p0.5 = Prepaid dan 0.5 sen
		'
		'   g - Pay As You Go
		'   P - PrePaid
		'   f - Fixed Time
		'
		
		'------------- variable declaration -------------
		Dim Cb1, Cb2 As System.Windows.Forms.ComboBox
		Dim CsmerType As String
		Dim gMin, gJam As Short
		
		'------------- defining -------------
		Cb1 = Combo1
		Cb2 = Combo2
		
		'------------- assigning & checking value -------------
		CsmerType = Combo3.Text 'ambil jenis pelanggan
		If Text1.Text = "" Then Exit Sub 'jika nama pelanggan kosong.. keluar..
		If Opt1(2).Checked = True And IsNumeric(cbPaid.Text) = False Then Exit Sub 'jika nilai dibayar bukan nombor.. keluar
		If Opt1(3).Checked = True And (Cb1.Text = "" And Cb2.Text = "") Then Exit Sub 'jika nilai minit atau tidak dipilih.. keluar
		If CsmerType = "" Then CsmerType = VS(2) 'jika tiada jenis pelanggan dipilih.. guna jenis biasa
		
		
		' an option ! yes its an option .. -------------
		'------- PAY AS YOU GO --------------------------------------
		If Opt1(1).Checked = True Then
			AgentSel.CusStartPAYG(UCase(Text1.Text), CsmerType)
			GoTo UnloadAll
			
			'------- PREPAID --------------------------------------------
		ElseIf Opt1(2).Checked = True Then 
			AgentSel.CusStartPPAID(UCase(Text1.Text), CsmerType, CDbl(cbPaid.Text), Sambung)
			GoTo UnloadAll
			
			'------ FIXED TIME -----------------------------------------
		ElseIf Opt1(3).Checked = True Then 
			gJam = IIf(Cb1.Text = "", 0, Cb1.Text) 'jika jam = "" return 0
			gMin = IIf(Cb2.Text = "", 0, Cb2.Text) 'jika minit = "" return 0
			AgentSel.CusStartTIME(UCase(Text1.Text), CsmerType, gJam, gMin)
			GoTo UnloadAll
		End If
		Exit Sub
		'---------------- End point of algorithm... -------------------
		
UnloadAll: 
		
		Me.Hide()
		If Sambung = True Then Sambung = False
		If SelTag <> "dump" Then AgentSel.NetSend("//kunci:0")
		Call UpdatePanel(SelText)
		Me.Close()
	End Sub
	
	Private Sub BtnKo_Click(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles BtnKo.Click
		Me.Close()
	End Sub
	
	Private Sub FrmLogin_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim s As Object
		NamaPc.Caption = SelText
		LblCrnc.Text = Crnc
		VB6.SetItemString(Combo3, 0, VS(2))
		Combo3.Text = VS(2)
		
		If Combo3.Items.Count = 1 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object uSDBe.DataCount(skema). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			For s = 0 To uSDBe.DataCount("skema") - 1
				'UPGRADE_WARNING: Couldn't resolve default property of object uSDBe.DataGet(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				Combo3.Items.Add(uSDBe.DataGet("skema", "skema", s))
			Next s
		End If
		
		'untuk orang yang ingin sambung...
		If Sambung = True Then
			Text1.Text = AgentSel.CustomerName
			Combo3.Text = AgentSel.CustomerType
			Text1.Enabled = False
			BtnKo.Enabled = False
		End If
	End Sub
	
	'UPGRADE_WARNING: Event Opt1.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2075"'
	Private Sub Opt1_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Opt1.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = Opt1.GetIndex(eventSender)
			Select Case Index
				Case 1
					cbPaid.Enabled = False
					Combo1.Enabled = False
					Combo2.Enabled = False
				Case 2
					cbPaid.Enabled = True
					Combo1.Enabled = False
					Combo2.Enabled = False
				Case 3
					cbPaid.Enabled = False
					Combo1.Enabled = True
					Combo2.Enabled = True
			End Select
		End If
	End Sub
	
	Private Sub Picture5_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles Picture5.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		If Button = 1 Then MoveFrm(Me.Handle.ToInt32)
	End Sub
	
	Public Sub FastLogin()
		Call BtnOk_Click(BtnOk, Nothing)
	End Sub
End Class