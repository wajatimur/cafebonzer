Option Strict Off
Option Explicit On
Friend Class TitleBar
	Inherits System.Windows.Forms.UserControl
#Region "Windows Form Designer generated code "
	Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			UserControl_Terminate()
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Friend WithEvents HolderLne As Line3D
	Friend WithEvents _SysBtn_2 As System.Windows.Forms.PictureBox
	Friend WithEvents _SysBtn_1 As System.Windows.Forms.PictureBox
	Friend WithEvents _SysBtn_0 As System.Windows.Forms.PictureBox
	Friend WithEvents HolderIcn As System.Windows.Forms.PictureBox
	Friend WithEvents HolderLbl As System.Windows.Forms.Label
	Friend WithEvents SysBtn As Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(TitleBar))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.HolderLne = New Line3D
		Me._SysBtn_2 = New System.Windows.Forms.PictureBox
		Me._SysBtn_1 = New System.Windows.Forms.PictureBox
		Me._SysBtn_0 = New System.Windows.Forms.PictureBox
		Me.HolderIcn = New System.Windows.Forms.PictureBox
		Me.HolderLbl = New System.Windows.Forms.Label
		Me.SysBtn = New Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray(components)
		CType(Me.SysBtn, System.ComponentModel.ISupportInitialize).BeginInit()
		MyBase.BackColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me.ClientSize = New System.Drawing.Size(392, 22)
		MyBase.Location = New System.Drawing.Point(0, 0)
		MyBase.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		MyBase.Name = "TitleBar"
		Me.HolderLne.Size = New System.Drawing.Size(392, 3)
		Me.HolderLne.Location = New System.Drawing.Point(0, 19)
		Me.HolderLne.TabIndex = 1
		Me.HolderLne.horizon = -1
		Me.HolderLne.Name = "HolderLne"
		Me._SysBtn_2.Size = New System.Drawing.Size(16, 16)
		Me._SysBtn_2.Location = New System.Drawing.Point(375, 1)
		Me._SysBtn_2.Image = CType(resources.GetObject("_SysBtn_2.Image"), System.Drawing.Image)
		Me._SysBtn_2.Enabled = True
		Me._SysBtn_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._SysBtn_2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._SysBtn_2.Visible = True
		Me._SysBtn_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._SysBtn_2.Name = "_SysBtn_2"
		Me._SysBtn_1.Size = New System.Drawing.Size(16, 16)
		Me._SysBtn_1.Location = New System.Drawing.Point(360, 1)
		Me._SysBtn_1.Image = CType(resources.GetObject("_SysBtn_1.Image"), System.Drawing.Image)
		Me._SysBtn_1.Enabled = True
		Me._SysBtn_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._SysBtn_1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._SysBtn_1.Visible = True
		Me._SysBtn_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._SysBtn_1.Name = "_SysBtn_1"
		Me._SysBtn_0.Size = New System.Drawing.Size(16, 16)
		Me._SysBtn_0.Location = New System.Drawing.Point(345, 1)
		Me._SysBtn_0.Image = CType(resources.GetObject("_SysBtn_0.Image"), System.Drawing.Image)
		Me._SysBtn_0.Enabled = True
		Me._SysBtn_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._SysBtn_0.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._SysBtn_0.Visible = True
		Me._SysBtn_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._SysBtn_0.Name = "_SysBtn_0"
		Me.HolderIcn.Size = New System.Drawing.Size(16, 16)
		Me.HolderIcn.Location = New System.Drawing.Point(1, 1)
		Me.HolderIcn.Image = CType(resources.GetObject("HolderIcn.Image"), System.Drawing.Image)
		Me.HolderIcn.Enabled = True
		Me.HolderIcn.Cursor = System.Windows.Forms.Cursors.Default
		Me.HolderIcn.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.HolderIcn.Visible = True
		Me.HolderIcn.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.HolderIcn.Name = "HolderIcn"
		Me.HolderLbl.Text = "Caption"
		Me.HolderLbl.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.HolderLbl.ForeColor = System.Drawing.Color.White
		Me.HolderLbl.Size = New System.Drawing.Size(49, 13)
		Me.HolderLbl.Location = New System.Drawing.Point(20, 2)
		Me.HolderLbl.TabIndex = 0
		Me.HolderLbl.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.HolderLbl.BackColor = System.Drawing.Color.Transparent
		Me.HolderLbl.Enabled = True
		Me.HolderLbl.Cursor = System.Windows.Forms.Cursors.Default
		Me.HolderLbl.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HolderLbl.UseMnemonic = True
		Me.HolderLbl.Visible = True
		Me.HolderLbl.AutoSize = True
		Me.HolderLbl.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.HolderLbl.Name = "HolderLbl"
		Me.Controls.Add(HolderLne)
		Me.Controls.Add(_SysBtn_2)
		Me.Controls.Add(_SysBtn_1)
		Me.Controls.Add(_SysBtn_0)
		Me.Controls.Add(HolderIcn)
		Me.Controls.Add(HolderLbl)
		Me.SysBtn.SetIndex(_SysBtn_2, CType(2, Short))
		Me.SysBtn.SetIndex(_SysBtn_1, CType(1, Short))
		Me.SysBtn.SetIndex(_SysBtn_0, CType(0, Short))
		CType(Me.SysBtn, System.ComponentModel.ISupportInitialize).EndInit()
	End Sub
#End Region 
	Private sHldrCap As String
	Private oHldrCapClr As System.Drawing.Color
	Private bSysBtn(2) As Boolean
	
	
	Private Sub UserControl_Terminate()
		sHldrCap = ""
		'UPGRADE_NOTE: Erase was upgraded to System.Array.Clear. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
		System.Array.Clear(bSysBtn, 0, bSysBtn.Length)
	End Sub
	
	'UPGRADE_WARNING: UserControl Event UserControl.InitProperties was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2050"'
	Private Sub UserControl_InitProperties()
		'UPGRADE_ISSUE: UserControl property TitleBar.Extender was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Extender.Name. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		sHldrCap = Extender.Name
		oHldrCapClr = System.Drawing.Color.White
		bSysBtn(0) = True
		bSysBtn(1) = True
		bSysBtn(2) = True
	End Sub
	
	Private Sub TitleBar_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		'UPGRADE_WARNING: Control property UserControl.Parent was upgraded to UserControl.FindForm which has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2074"'
		'UPGRADE_NOTE: The Following line was commented to give the same effect as VB6. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2041"'
		'MyBase.FindForm.FormBorderStyle = 2
		'UPGRADE_WARNING: Control property UserControl.Parent was upgraded to UserControl.FindForm which has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2074"'
		MyBase.FindForm.Text = ""
		
		'UPGRADE_ISSUE: UserControl property TitleBar.Extender was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Extender.Align. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Extender.Align = 1
		'UPGRADE_ISSUE: UserControl property TitleBar.Extender was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Extender.Height. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Extender.Height = 330
		'UPGRADE_ISSUE: UserControl property TitleBar.Extender was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Extender.Width. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		HolderLne.Width = VB6.TwipsToPixelsX(Extender.Width)
		Call zRedrawHolder()
	End Sub
	
	Private Sub TitleBar_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim lret As Object
		ReleaseCapture()
		'UPGRADE_WARNING: Control property .Parent was upgraded to .FindForm which has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2074"'
		'UPGRADE_WARNING: Couldn't resolve default property of object lret. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		lret = SendMessage(FindForm.Handle.ToInt32, WM_NCLBUTTONDOWN, HTCAPTION, 0)
	End Sub
	
	'UPGRADE_WARNING: PropertyBag object was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup6003"'
	'UPGRADE_WARNING: UserControl Event ReadProperties is not supported. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup6002"'
	Private Sub UserControl_ReadProperties(ByRef PropBag As PropertyBag)
		'UPGRADE_ISSUE: UserControl property TitleBar.Extender was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Extender.Name. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		sHldrCap = PropBag.ReadProperty("HldrCap", Extender.Name)
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		oHldrCapClr = System.Drawing.ColorTranslator.FromOle(PropBag.ReadProperty("HldrCapClr", oHldrCapClr))
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		bSysBtn(0) = PropBag.ReadProperty("SysBtnMin", True)
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		bSysBtn(1) = PropBag.ReadProperty("SysBtnMax", True)
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		bSysBtn(2) = PropBag.ReadProperty("SysBtnClose", True)
		Call zRedrawHolder()
	End Sub
	
	'UPGRADE_WARNING: PropertyBag object was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup6003"'
	'UPGRADE_WARNING: UserControl Event WriteProperties is not supported. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup6002"'
	Private Sub UserControl_WriteProperties(ByRef PropBag As PropertyBag)
		'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		PropBag.WriteProperty("HldrCap", sHldrCap)
		'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		PropBag.WriteProperty("HldrCapClr", oHldrCapClr)
		'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		PropBag.WriteProperty("SysBtnMin", bSysBtn(0))
		'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		PropBag.WriteProperty("SysBtnMax", bSysBtn(1))
		'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		PropBag.WriteProperty("SysBtnClose", bSysBtn(2))
	End Sub
	
	
	Private Sub SysBtn_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles SysBtn.Click
		Dim Index As Short = SysBtn.GetIndex(eventSender)
		Select Case Index
			Case 0
				'UPGRADE_WARNING: Control property .Parent was upgraded to .FindForm which has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2074"'
				FindForm.WindowState = System.Windows.Forms.FormWindowState.Minimized
			Case 1
				'UPGRADE_WARNING: Control property .Parent was upgraded to .FindForm which has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2074"'
				If FindForm.WindowState = 2 Then
					'UPGRADE_WARNING: Control property .Parent was upgraded to .FindForm which has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2074"'
					FindForm.WindowState = System.Windows.Forms.FormWindowState.Normal
				Else
					'UPGRADE_WARNING: Control property .Parent was upgraded to .FindForm which has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2074"'
					FindForm.WindowState = System.Windows.Forms.FormWindowState.Maximized
				End If
			Case 2
				'UPGRADE_WARNING: Control property .Parent was upgraded to .FindForm which has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2074"'
				FindForm.Close()
		End Select
	End Sub
	
	
	Public Property Caption() As String
		Get
			Caption = sHldrCap
		End Get
		Set(ByVal Value As String)
			sHldrCap = Value
			Call zRedrawHolder()
		End Set
	End Property
	
	Public Property CaptionColor() As System.Drawing.Color
		Get
			CaptionColor = oHldrCapClr
		End Get
		Set(ByVal Value As System.Drawing.Color)
			oHldrCapClr = Value
			Call zRedrawHolder()
		End Set
	End Property
	
	Public Property MinButton() As Boolean
		Get
			MinButton = bSysBtn(0)
		End Get
		Set(ByVal Value As Boolean)
			bSysBtn(0) = Value
			Call zRedrawHolder()
		End Set
	End Property
	
	Public Property MaxButton() As Boolean
		Get
			MaxButton = bSysBtn(1)
		End Get
		Set(ByVal Value As Boolean)
			bSysBtn(1) = Value
			Call zRedrawHolder()
		End Set
	End Property
	
	Public Property CloseButton() As Boolean
		Get
			CloseButton = bSysBtn(2)
		End Get
		Set(ByVal Value As Boolean)
			bSysBtn(2) = Value
			Call zRedrawHolder()
		End Set
	End Property
	
	
	Public Sub zRedrawHolder()
		Dim a As Short
		Dim l_Pos As Integer
		
		HolderLbl.Text = sHldrCap
		HolderLbl.ForeColor = oHldrCapClr
		
		For a = UBound(bSysBtn) To 0 Step -1
			SysBtn(a).Visible = False
			If bSysBtn(a) = True Then
				l_Pos = l_Pos + 1
				SysBtn(a).Visible = True
				SysBtn(a).Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(MyBase.Width) - (VB6.PixelsToTwipsX(SysBtn(a).Width) * l_Pos))
			End If
		Next a
	End Sub
End Class