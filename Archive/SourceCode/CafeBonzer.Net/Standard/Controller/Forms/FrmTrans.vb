Option Strict Off
Option Explicit On
Friend Class FrmAgnTrans
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
	Public WithEvents List1 As AxMSComctlLib.AxListView
	Public WithEvents BtnOk As XpButton
	Public WithEvents BtnKo As XpButton
	Public WithEvents Picture5 As System.Windows.Forms.Panel
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents TermName As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmAgnTrans))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.List1 = New AxMSComctlLib.AxListView
		Me.Picture5 = New System.Windows.Forms.Panel
		Me.BtnOk = New XpButton
		Me.BtnKo = New XpButton
		Me.Label3 = New System.Windows.Forms.Label
		Me.TermName = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		CType(Me.List1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.ControlBox = False
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.ClientSize = New System.Drawing.Size(240, 145)
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
		Me.Name = "FrmAgnTrans"
		List1.OcxState = CType(resources.GetObject("List1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.List1.Size = New System.Drawing.Size(206, 90)
		Me.List1.Location = New System.Drawing.Point(4, 51)
		Me.List1.TabIndex = 4
		Me.List1.Name = "List1"
		Me.Picture5.BackColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me.Picture5.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Picture5.Size = New System.Drawing.Size(24, 147)
		Me.Picture5.Location = New System.Drawing.Point(216, 0)
		Me.Picture5.TabIndex = 2
		Me.Picture5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
		Me.BtnOk.Location = New System.Drawing.Point(0, 122)
		Me.BtnOk.TabIndex = 5
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
		Me.BtnKo.Location = New System.Drawing.Point(0, 99)
		Me.BtnKo.TabIndex = 6
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
		Me.Label3.Text = "To :"
		Me.Label3.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.Size = New System.Drawing.Size(42, 18)
		Me.Label3.Location = New System.Drawing.Point(6, 32)
		Me.Label3.TabIndex = 3
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label3.BackColor = System.Drawing.Color.Transparent
		Me.Label3.Enabled = True
		Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label3.UseMnemonic = True
		Me.Label3.Visible = True
		Me.Label3.AutoSize = False
		Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label3.Name = "Label3"
		Me.TermName.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.TermName.BackColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me.TermName.Text = "Ais Krim Soda"
		Me.TermName.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.TermName.ForeColor = System.Drawing.Color.White
		Me.TermName.Size = New System.Drawing.Size(150, 22)
		Me.TermName.Location = New System.Drawing.Point(61, 6)
		Me.TermName.TabIndex = 1
		Me.TermName.Enabled = True
		Me.TermName.Cursor = System.Windows.Forms.Cursors.Default
		Me.TermName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.TermName.UseMnemonic = True
		Me.TermName.Visible = True
		Me.TermName.AutoSize = False
		Me.TermName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.TermName.Name = "TermName"
		Me.Label1.Text = "From :"
		Me.Label1.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Size = New System.Drawing.Size(47, 18)
		Me.Label1.Location = New System.Drawing.Point(6, 8)
		Me.Label1.TabIndex = 0
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
		Me.Controls.Add(List1)
		Me.Controls.Add(Picture5)
		Me.Controls.Add(Label3)
		Me.Controls.Add(TermName)
		Me.Controls.Add(Label1)
		Me.Picture5.Controls.Add(BtnOk)
		Me.Picture5.Controls.Add(BtnKo)
		CType(Me.List1, System.ComponentModel.ISupportInitialize).EndInit()
	End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As FrmAgnTrans
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As FrmAgnTrans
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New FrmAgnTrans()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	
	
	Private Sub BtnKo_Click(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles BtnKo.Click
		Me.Close()
	End Sub
	
	Private Sub BtnOk_Click(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles BtnOk.Click
		If List1.ListItems.Count = 0 Then Exit Sub
		
		If List1.SelectedItem.Text = "" Then
			FrmAgnTrans.DefInstance.Hide()
			FrmAgnTrans.DefInstance.Close()
		Else
			'hentikan timer sementara
			FrmSysHost.DefInstance.Timer1.Enabled = False
			FrmSysHost.DefInstance.Pinger.Enabled = False
			
			Call AgentSel.AgnTransfer(List1.SelectedItem.Tag)
			
			FrmSysHost.DefInstance.Timer1.Enabled = True
			FrmSysHost.DefInstance.Pinger.Enabled = True
			FrmAgnTrans.DefInstance.Close()
		End If
	End Sub
	
	Private Sub FrmAgnTrans_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim g As Object
		Dim tLv As AxMSComctlLib.AxListView
		Dim tLtm As MSComctlLib.ListItem
		tLv = FrmMain.DefInstance.Lv1
		
		'UPGRADE_WARNING: Couldn't resolve default property of object List1.SmallIcons. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		List1.SmallIcons = FrmMain.DefInstance.ImgList16.GetOCX
		TermName.Text = tLv.SelectedItem.Text
		
		For g = 1 To tLv.ListItems.Count
			If tLv.ListItems(g).Text <> TermName.Text Then
				If tLv.ListItems(g).SubItems(1) = VS(4) Then
					tLtm = List1.ListItems.Add( ,  , tLv.ListItems(g).Text,  , "TerminalOnline")
					tLtm.let_Tag(tLv.ListItems(g).Index)
				End If
			End If
		Next g
	End Sub
End Class