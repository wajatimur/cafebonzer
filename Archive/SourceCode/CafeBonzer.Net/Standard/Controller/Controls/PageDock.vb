Option Strict Off
Option Explicit On
Friend Class PageDock
	Inherits System.Windows.Forms.ContainerControl
#Region "Windows Form Designer generated code "
	Public Sub New()
		MyBase.New()
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
	Friend WithEvents HolderLne As Line3D
	Friend WithEvents HolderBtn As System.Windows.Forms.PictureBox
	Friend WithEvents HolderIcn As System.Windows.Forms.PictureBox
	Friend WithEvents Holder As System.Windows.Forms.Panel
	Friend WithEvents _ImgCnt_0 As System.Windows.Forms.PictureBox
	Friend WithEvents _ImgCnt_1 As System.Windows.Forms.PictureBox
	Friend WithEvents ImgCnt As Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(PageDock))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.HolderLne = New Line3D
		Me.Holder = New System.Windows.Forms.Panel
		Me.HolderBtn = New System.Windows.Forms.PictureBox
		Me.HolderIcn = New System.Windows.Forms.PictureBox
		Me._ImgCnt_0 = New System.Windows.Forms.PictureBox
		Me._ImgCnt_1 = New System.Windows.Forms.PictureBox
		Me.ImgCnt = New Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray(components)
		CType(Me.ImgCnt, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.ClientSize = New System.Drawing.Size(132, 274)
		MyBase.Location = New System.Drawing.Point(0, 0)
		MyBase.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		MyBase.Name = "PageDock"
		Me.HolderLne.Size = New System.Drawing.Size(3, 273)
		Me.HolderLne.Location = New System.Drawing.Point(20, 0)
		Me.HolderLne.TabIndex = 1
		Me.HolderLne.horizon = 0
		Me.HolderLne.Name = "HolderLne"
		Me.Holder.BackColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me.Holder.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Holder.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Holder.Size = New System.Drawing.Size(19, 274)
		Me.Holder.Location = New System.Drawing.Point(1, 0)
		Me.Holder.TabIndex = 0
		Me.Holder.Dock = System.Windows.Forms.DockStyle.None
		Me.Holder.CausesValidation = True
		Me.Holder.Enabled = True
		Me.Holder.Cursor = System.Windows.Forms.Cursors.Default
		Me.Holder.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Holder.TabStop = True
		Me.Holder.Visible = True
		Me.Holder.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Holder.Name = "Holder"
		Me.HolderBtn.Size = New System.Drawing.Size(16, 16)
		Me.HolderBtn.Location = New System.Drawing.Point(1, 0)
		Me.HolderBtn.Image = CType(resources.GetObject("HolderBtn.Image"), System.Drawing.Image)
		Me.HolderBtn.Enabled = True
		Me.HolderBtn.Cursor = System.Windows.Forms.Cursors.Default
		Me.HolderBtn.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.HolderBtn.Visible = True
		Me.HolderBtn.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.HolderBtn.Name = "HolderBtn"
		Me.HolderIcn.Size = New System.Drawing.Size(16, 16)
		Me.HolderIcn.Location = New System.Drawing.Point(1, 257)
		Me.HolderIcn.Image = CType(resources.GetObject("HolderIcn.Image"), System.Drawing.Image)
		Me.HolderIcn.Enabled = True
		Me.HolderIcn.Cursor = System.Windows.Forms.Cursors.Default
		Me.HolderIcn.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.HolderIcn.Visible = True
		Me.HolderIcn.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.HolderIcn.Name = "HolderIcn"
		Me._ImgCnt_0.Size = New System.Drawing.Size(16, 16)
		Me._ImgCnt_0.Location = New System.Drawing.Point(94, 255)
		Me._ImgCnt_0.Image = CType(resources.GetObject("_ImgCnt_0.Image"), System.Drawing.Image)
		Me._ImgCnt_0.Visible = False
		Me._ImgCnt_0.Enabled = True
		Me._ImgCnt_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._ImgCnt_0.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._ImgCnt_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._ImgCnt_0.Name = "_ImgCnt_0"
		Me._ImgCnt_1.Size = New System.Drawing.Size(16, 16)
		Me._ImgCnt_1.Location = New System.Drawing.Point(112, 255)
		Me._ImgCnt_1.Image = CType(resources.GetObject("_ImgCnt_1.Image"), System.Drawing.Image)
		Me._ImgCnt_1.Visible = False
		Me._ImgCnt_1.Enabled = True
		Me._ImgCnt_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._ImgCnt_1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._ImgCnt_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._ImgCnt_1.Name = "_ImgCnt_1"
		Me.Controls.Add(HolderLne)
		Me.Controls.Add(Holder)
		Me.Controls.Add(_ImgCnt_0)
		Me.Controls.Add(_ImgCnt_1)
		Me.Holder.Controls.Add(HolderBtn)
		Me.Holder.Controls.Add(HolderIcn)
		Me.ImgCnt.SetIndex(_ImgCnt_0, CType(0, Short))
		Me.ImgCnt.SetIndex(_ImgCnt_1, CType(1, Short))
		CType(Me.ImgCnt, System.ComponentModel.ISupportInitialize).EndInit()
	End Sub
#End Region 
	Private lHldrBtnPos As Integer
	Private bHldrLne As Boolean
	
	Private lPageWidth As Integer
	Private lPageState As Integer
	
	Public Enum eHldrBtnPos
		Top = 0
		Bottom = 1
		Middle = 2
	End Enum
	
	Public Event HolderButtonClick(ByVal Sender As System.Object, ByVal e As System.EventArgs)
	<System.Runtime.InteropServices.ProgId("PageFlipedEventArgs_NET.PageFlipedEventArgs")> Public NotInheritable Class PageFlipedEventArgs
		Inherits System.EventArgs
		Public Flipped As Boolean
		Public Sub New(ByVal Flipped As Boolean)
			MyBase.New()
			Me.Flipped = Flipped
		End Sub
	End Class
	Public Event PageFliped(ByVal Sender As System.Object, ByVal e As PageFlipedEventArgs)
	
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' USERCONTROL
	'
	'
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'UPGRADE_WARNING: UserControl Event UserControl.InitProperties was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2050"'
	Private Sub UserControl_InitProperties()
		bHldrLne = True
		lPageState = 0
		lPageWidth = VB6.PixelsToTwipsX(MyBase.Width)
	End Sub
	
	Private Sub PageDock_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		Dim Control As Object
		'{ Resizer & Var }'
		Holder.Height = MyBase.Height
		HolderLne.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Holder.Height))
		If lPageState = 0 Then
			lPageWidth = VB6.PixelsToTwipsX(MyBase.Width)
		Else
			'UPGRADE_ISSUE: UserControl property PageDock.Extender was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Extender.Width. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			Extender.Width = VB6.PixelsToTwipsX(Holder.Width)
		End If
		
		'{ Icon & Button Position }'
		If lHldrBtnPos = 0 Then
			HolderIcn.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Holder.Height) - VB6.PixelsToTwipsY(HolderIcn.Height))
		ElseIf lHldrBtnPos = 1 Then 
			HolderBtn.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Holder.Height) - VB6.PixelsToTwipsY(HolderBtn.Height))
		End If
		
		'{ Smart Container Handler }'
		'UPGRADE_ISSUE: UserControl property UserControl.ContainedControls was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		For	Each Control In MyBase.ContainedControls
			'UPGRADE_WARNING: Couldn't resolve default property of object Control.Tag. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			If Control.Tag = "subcontainer" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Control.Height. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				'UPGRADE_ISSUE: UserControl property PageDock.Extender was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Extender.Height. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				Control.Height = Extender.Height
			End If
		Next Control
	End Sub
	
	Private Sub HolderBtn_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles HolderBtn.Click
		RaiseEvent HolderButtonClick(Me, Nothing)
		Me.PageFlip = CBool(lPageState) Xor True
		Call zRedrawHolder()
	End Sub
	
	
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' READ\WRITE PROPERTIES
	'
	'
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'UPGRADE_WARNING: PropertyBag object was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup6003"'
	'UPGRADE_WARNING: UserControl Event ReadProperties is not supported. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup6002"'
	Private Sub UserControl_ReadProperties(ByRef PropBag As PropertyBag)
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		lHldrBtnPos = PropBag.ReadProperty("HldrBtnPos", 0)
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		bHldrLne = PropBag.ReadProperty("HldrLne", True)
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		lPageState = PropBag.ReadProperty("PageState", 0)
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		lPageWidth = PropBag.ReadProperty("PageWidth", lPageWidth)
		Call zRedrawHolder()
	End Sub
	'UPGRADE_WARNING: PropertyBag object was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup6003"'
	'UPGRADE_WARNING: UserControl Event WriteProperties is not supported. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup6002"'
	Private Sub UserControl_WriteProperties(ByRef PropBag As PropertyBag)
		'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		PropBag.WriteProperty("HldrBtnPos", lHldrBtnPos)
		'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		PropBag.WriteProperty("HldrLne", bHldrLne)
		'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		PropBag.WriteProperty("PageState", lPageState)
		'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		PropBag.WriteProperty("PageWidth", lPageWidth)
	End Sub
	
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' PROPERTY SECTION
	'
	'
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Public Property HolderButtonPos() As eHldrBtnPos
		Get
			HolderButtonPos = lHldrBtnPos
		End Get
		Set(ByVal Value As eHldrBtnPos)
			lHldrBtnPos = Value
			Call zRedrawHolder()
			'UPGRADE_ISSUE: UserControl method PageDock.PropertyChanged was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Call PropertyChanged()
		End Set
	End Property
	
	Public Property HolderLine() As Boolean
		Get
			HolderLine = bHldrLne
		End Get
		Set(ByVal Value As Boolean)
			bHldrLne = Value
			Call zRedrawHolder()
			'UPGRADE_ISSUE: UserControl method PageDock.PropertyChanged was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Call PropertyChanged()
		End Set
	End Property
	
	Public Property PageFlip() As Boolean
		Get
			PageFlip = CBool(lPageState)
		End Get
		Set(ByVal Value As Boolean)
			lPageState = IIf(Value, 1, 0)
			Call zFlip(Value)
			Call zRedrawHolder()
			'UPGRADE_ISSUE: UserControl method PageDock.PropertyChanged was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Call PropertyChanged()
		End Set
	End Property
	
	
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' PRIVATE METHOD & FUNCTION
	'
	'
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Private Sub zRedrawHolder()
		Dim Mid_Renamed As Short
		Dim Mid_Renamed As Short
		HolderLne.Visible = bHldrLne
		Select Case lHldrBtnPos
			Case 0
				HolderBtn.Top = 0
				HolderBtn.Left = VB6.TwipsToPixelsX(15)
				HolderIcn.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Holder.Height) - VB6.PixelsToTwipsY(HolderIcn.Height))
				HolderIcn.Left = VB6.TwipsToPixelsX(15)
			Case 1
				HolderBtn.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Holder.Height) - VB6.PixelsToTwipsY(HolderBtn.Height))
				HolderBtn.Left = VB6.TwipsToPixelsX(15)
				HolderIcn.Top = 0
				HolderIcn.Left = VB6.TwipsToPixelsX(15)
			Case 2
				Mid_Renamed = (VB6.PixelsToTwipsY(Holder.Height) \ 2) - (VB6.PixelsToTwipsY(HolderBtn.Height) \ 2)
				'UPGRADE_WARNING: Untranslated statement in zRedrawHolder. Please check source code.
				HolderBtn.Left = 0
				HolderIcn.Top = VB6.TwipsToPixelsY(-500)
		End Select
		If lPageState = 1 Then HolderBtn.Image = ImgCnt(1).Image
	End Sub
	
	Private Sub zFlip(Optional ByRef Flip As Boolean = True)
		If Flip = True Then
			'UPGRADE_ISSUE: UserControl property PageDock.Extender was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Extender.Left. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Extender.Width. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			Extender.Left = (Extender.Left + Extender.Width) - VB6.PixelsToTwipsX(Holder.Width)
			'UPGRADE_ISSUE: UserControl property PageDock.Extender was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Extender.Width. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			Extender.Width = VB6.PixelsToTwipsX(Holder.Width)
			HolderBtn.Image = ImgCnt(1).Image
		Else
			'UPGRADE_ISSUE: UserControl property PageDock.Extender was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Extender.Left. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Extender.Width. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			Extender.Left = (Extender.Left + Extender.Width) - lPageWidth
			'UPGRADE_ISSUE: UserControl property PageDock.Extender was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Extender.Width. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			Extender.Width = lPageWidth
			HolderBtn.Image = ImgCnt(0).Image
		End If
		RaiseEvent PageFliped(Me, New PageFlipedEventArgs(Flip))
	End Sub
End Class