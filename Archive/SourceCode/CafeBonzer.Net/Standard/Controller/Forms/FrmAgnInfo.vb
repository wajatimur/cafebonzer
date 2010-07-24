Option Strict Off
Option Explicit On
Friend Class FrmAgnInfo
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
	Public WithEvents DynaCombo As System.Windows.Forms.ComboBox
	Public WithEvents _DynaPbar_0 As AxMSComctlLib.AxProgressBar
	Public WithEvents _DynaLv_0 As AxMSComctlLib.AxListView
	Public WithEvents _DynaLv_1 As AxMSComctlLib.AxListView
	Public WithEvents _Pages_1 As System.Windows.Forms.Panel
	Public WithEvents DynaLv As AxListViewArray.AxListViewArray
	Public WithEvents DynaPbar As AxProgressBarArray.AxProgressBarArray
	Public WithEvents Pages As Microsoft.VisualBasic.Compatibility.VB6.PanelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmAgnInfo))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me._Pages_1 = New System.Windows.Forms.Panel
		Me.DynaCombo = New System.Windows.Forms.ComboBox
		Me._DynaPbar_0 = New AxMSComctlLib.AxProgressBar
		Me._DynaLv_0 = New AxMSComctlLib.AxListView
		Me._DynaLv_1 = New AxMSComctlLib.AxListView
		Me.DynaLv = New AxListViewArray.AxListViewArray(components)
		Me.DynaPbar = New AxProgressBarArray.AxProgressBarArray(components)
		Me.Pages = New Microsoft.VisualBasic.Compatibility.VB6.PanelArray(components)
		CType(Me._DynaPbar_0, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me._DynaLv_0, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me._DynaLv_1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.DynaLv, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.DynaPbar, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Pages, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.Text = "Agent Information"
		Me.ClientSize = New System.Drawing.Size(704, 349)
		Me.Location = New System.Drawing.Point(4, 23)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.BackColor = System.Drawing.SystemColors.Control
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
		Me.Name = "FrmAgnInfo"
		Me._Pages_1.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me._Pages_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Pages_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._Pages_1.Size = New System.Drawing.Size(678, 307)
		Me._Pages_1.Location = New System.Drawing.Point(0, 0)
		Me._Pages_1.TabIndex = 0
		Me._Pages_1.Dock = System.Windows.Forms.DockStyle.None
		Me._Pages_1.CausesValidation = True
		Me._Pages_1.Enabled = True
		Me._Pages_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Pages_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Pages_1.TabStop = True
		Me._Pages_1.Visible = True
		Me._Pages_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Pages_1.Name = "_Pages_1"
		Me.DynaCombo.Font = New System.Drawing.Font("Verdana", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.DynaCombo.Size = New System.Drawing.Size(103, 20)
		Me.DynaCombo.Location = New System.Drawing.Point(3, 4)
		Me.DynaCombo.Items.AddRange(New Object(){"All Printer"})
		Me.DynaCombo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.DynaCombo.TabIndex = 2
		Me.DynaCombo.Visible = False
		Me.DynaCombo.BackColor = System.Drawing.SystemColors.Window
		Me.DynaCombo.CausesValidation = True
		Me.DynaCombo.Enabled = True
		Me.DynaCombo.ForeColor = System.Drawing.SystemColors.WindowText
		Me.DynaCombo.IntegralHeight = True
		Me.DynaCombo.Cursor = System.Windows.Forms.Cursors.Default
		Me.DynaCombo.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.DynaCombo.Sorted = False
		Me.DynaCombo.TabStop = True
		Me.DynaCombo.Name = "DynaCombo"
		_DynaPbar_0.OcxState = CType(resources.GetObject("_DynaPbar_0.OcxState"), System.Windows.Forms.AxHost.State)
		Me._DynaPbar_0.Size = New System.Drawing.Size(90, 16)
		Me._DynaPbar_0.Location = New System.Drawing.Point(5, 27)
		Me._DynaPbar_0.TabIndex = 1
		Me._DynaPbar_0.Visible = False
		Me._DynaPbar_0.Name = "_DynaPbar_0"
		_DynaLv_0.OcxState = CType(resources.GetObject("_DynaLv_0.OcxState"), System.Windows.Forms.AxHost.State)
		Me._DynaLv_0.Size = New System.Drawing.Size(246, 302)
		Me._DynaLv_0.Location = New System.Drawing.Point(0, 1)
		Me._DynaLv_0.TabIndex = 3
		Me._DynaLv_0.Name = "_DynaLv_0"
		_DynaLv_1.OcxState = CType(resources.GetObject("_DynaLv_1.OcxState"), System.Windows.Forms.AxHost.State)
		Me._DynaLv_1.Size = New System.Drawing.Size(431, 302)
		Me._DynaLv_1.Location = New System.Drawing.Point(247, 1)
		Me._DynaLv_1.TabIndex = 4
		Me._DynaLv_1.Name = "_DynaLv_1"
		Me.Controls.Add(_Pages_1)
		Me._Pages_1.Controls.Add(DynaCombo)
		Me._Pages_1.Controls.Add(_DynaPbar_0)
		Me._Pages_1.Controls.Add(_DynaLv_0)
		Me._Pages_1.Controls.Add(_DynaLv_1)
		Me.DynaLv.SetIndex(_DynaLv_0, CType(0, Short))
		Me.DynaLv.SetIndex(_DynaLv_1, CType(1, Short))
		Me.DynaPbar.SetIndex(_DynaPbar_0, CType(0, Short))
		Me.Pages.SetIndex(_Pages_1, CType(1, Short))
		CType(Me.Pages, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.DynaPbar, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.DynaLv, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me._DynaLv_1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me._DynaLv_0, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me._DynaPbar_0, System.ComponentModel.ISupportInitialize).EndInit()
	End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As FrmAgnInfo
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As FrmAgnInfo
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New FrmAgnInfo()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
	' Dynamic Page
	'
	'
	'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
	'UPGRADE_WARNING: Event DynaCombo.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2075"'
	Private Sub DynaCombo_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles DynaCombo.SelectedIndexChanged
		DynaLv(0).SelectedItem.SubItems(1) = DynaCombo.Text
		DynaCombo.Visible = False
	End Sub
	
	Private Sub DynaLv_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles DynaLv.DblClick
		Dim Index As Short = DynaLv.GetIndex(eventSender)
		Select Case Index
			Case 0
				
		End Select
	End Sub
	
	Private Sub DynaLv_ItemClick(ByVal eventSender As System.Object, ByVal eventArgs As AxMSComctlLib.ListViewEvents_ItemClickEvent) Handles DynaLv.ItemClick
		Dim Index As Short = DynaLv.GetIndex(eventSender)
		Select Case Index
			Case 0
				
			Case 1
				
		End Select
	End Sub
End Class