Option Strict Off
Option Explicit On
Friend Class PageHolder
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
	Friend WithEvents HolderLbl As System.Windows.Forms.Label
	Friend WithEvents HolderBtn As System.Windows.Forms.PictureBox
	Friend WithEvents HolderIcn As System.Windows.Forms.PictureBox
	Friend WithEvents Holder As System.Windows.Forms.Panel
	Friend WithEvents _ImgCnt_1 As System.Windows.Forms.PictureBox
	Friend WithEvents _ImgCnt_3 As System.Windows.Forms.PictureBox
	Friend WithEvents _ImgCnt_2 As System.Windows.Forms.PictureBox
	Friend WithEvents _ImgCnt_0 As System.Windows.Forms.PictureBox
	Friend WithEvents ImgCnt As Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(PageHolder))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.HolderLne = New Line3D
		Me.Holder = New System.Windows.Forms.Panel
		Me.HolderLbl = New System.Windows.Forms.Label
		Me.HolderBtn = New System.Windows.Forms.PictureBox
		Me.HolderIcn = New System.Windows.Forms.PictureBox
		Me._ImgCnt_1 = New System.Windows.Forms.PictureBox
		Me._ImgCnt_3 = New System.Windows.Forms.PictureBox
		Me._ImgCnt_2 = New System.Windows.Forms.PictureBox
		Me._ImgCnt_0 = New System.Windows.Forms.PictureBox
		Me.ImgCnt = New Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray(components)
		CType(Me.ImgCnt, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.ClientSize = New System.Drawing.Size(424, 120)
		MyBase.Location = New System.Drawing.Point(0, 0)
		MyBase.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		MyBase.Name = "PageHolder"
		Me.HolderLne.Size = New System.Drawing.Size(421, 3)
		Me.HolderLne.Location = New System.Drawing.Point(1, 19)
		Me.HolderLne.TabIndex = 2
		Me.HolderLne.horizon = -1
		Me.HolderLne.Name = "HolderLne"
		Me.Holder.BackColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me.Holder.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Holder.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Holder.Size = New System.Drawing.Size(422, 19)
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
		Me.HolderLbl.Text = "PageHolder"
		Me.HolderLbl.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.HolderLbl.ForeColor = System.Drawing.Color.White
		Me.HolderLbl.Size = New System.Drawing.Size(75, 13)
		Me.HolderLbl.Location = New System.Drawing.Point(35, 2)
		Me.HolderLbl.TabIndex = 1
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
		Me.HolderBtn.Size = New System.Drawing.Size(16, 16)
		Me.HolderBtn.Location = New System.Drawing.Point(17, 2)
		Me.HolderBtn.Image = CType(resources.GetObject("HolderBtn.Image"), System.Drawing.Image)
		Me.HolderBtn.Enabled = True
		Me.HolderBtn.Cursor = System.Windows.Forms.Cursors.Default
		Me.HolderBtn.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.HolderBtn.Visible = True
		Me.HolderBtn.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.HolderBtn.Name = "HolderBtn"
		Me.HolderIcn.Size = New System.Drawing.Size(16, 16)
		Me.HolderIcn.Location = New System.Drawing.Point(1, 2)
		Me.HolderIcn.Image = CType(resources.GetObject("HolderIcn.Image"), System.Drawing.Image)
		Me.HolderIcn.Enabled = True
		Me.HolderIcn.Cursor = System.Windows.Forms.Cursors.Default
		Me.HolderIcn.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.HolderIcn.Visible = True
		Me.HolderIcn.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.HolderIcn.Name = "HolderIcn"
		Me._ImgCnt_1.Size = New System.Drawing.Size(16, 16)
		Me._ImgCnt_1.Location = New System.Drawing.Point(359, 95)
		Me._ImgCnt_1.Image = CType(resources.GetObject("_ImgCnt_1.Image"), System.Drawing.Image)
		Me._ImgCnt_1.Visible = False
		Me._ImgCnt_1.Enabled = True
		Me._ImgCnt_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._ImgCnt_1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._ImgCnt_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._ImgCnt_1.Name = "_ImgCnt_1"
		Me._ImgCnt_3.Size = New System.Drawing.Size(16, 16)
		Me._ImgCnt_3.Location = New System.Drawing.Point(395, 95)
		Me._ImgCnt_3.Image = CType(resources.GetObject("_ImgCnt_3.Image"), System.Drawing.Image)
		Me._ImgCnt_3.Visible = False
		Me._ImgCnt_3.Enabled = True
		Me._ImgCnt_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._ImgCnt_3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._ImgCnt_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._ImgCnt_3.Name = "_ImgCnt_3"
		Me._ImgCnt_2.Size = New System.Drawing.Size(16, 16)
		Me._ImgCnt_2.Location = New System.Drawing.Point(377, 95)
		Me._ImgCnt_2.Image = CType(resources.GetObject("_ImgCnt_2.Image"), System.Drawing.Image)
		Me._ImgCnt_2.Visible = False
		Me._ImgCnt_2.Enabled = True
		Me._ImgCnt_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._ImgCnt_2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._ImgCnt_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._ImgCnt_2.Name = "_ImgCnt_2"
		Me._ImgCnt_0.Size = New System.Drawing.Size(16, 16)
		Me._ImgCnt_0.Location = New System.Drawing.Point(340, 95)
		Me._ImgCnt_0.Image = CType(resources.GetObject("_ImgCnt_0.Image"), System.Drawing.Image)
		Me._ImgCnt_0.Visible = False
		Me._ImgCnt_0.Enabled = True
		Me._ImgCnt_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._ImgCnt_0.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._ImgCnt_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._ImgCnt_0.Name = "_ImgCnt_0"
		Me.Controls.Add(HolderLne)
		Me.Controls.Add(Holder)
		Me.Controls.Add(_ImgCnt_1)
		Me.Controls.Add(_ImgCnt_3)
		Me.Controls.Add(_ImgCnt_2)
		Me.Controls.Add(_ImgCnt_0)
		Me.Holder.Controls.Add(HolderLbl)
		Me.Holder.Controls.Add(HolderBtn)
		Me.Holder.Controls.Add(HolderIcn)
		Me.ImgCnt.SetIndex(_ImgCnt_1, CType(1, Short))
		Me.ImgCnt.SetIndex(_ImgCnt_3, CType(3, Short))
		Me.ImgCnt.SetIndex(_ImgCnt_2, CType(2, Short))
		Me.ImgCnt.SetIndex(_ImgCnt_0, CType(0, Short))
		CType(Me.ImgCnt, System.ComponentModel.ISupportInitialize).EndInit()
	End Sub
#End Region 
	Private tImg As System.Drawing.Image = New System.Drawing.Bitmap(1, 1)
	
	Private lHldrStyle As Integer
	Private lHldrIcon As Integer
	Private sHldrTxt As String
	Private oHldrTxtClr As System.Drawing.Color
	Private bHldrLne As Boolean
	
	Private lPageState As Integer
	Private lPageHeight As Integer
	
	
	Public Enum eHldrStyle
		Normal = 0
		Simple = 1
		Text_Only = 2
	End Enum
	
	Public Enum eHldrIcon
		Default_Renamed = 0
		Planetary_1 = 1
		Planetary_2 = 2
	End Enum
	
	<System.Runtime.InteropServices.ProgId("HolderButtonClickEventArgs_NET.HolderButtonClickEventArgs")> Public NotInheritable Class HolderButtonClickEventArgs
		Inherits System.EventArgs
		Public Collapse As Boolean
		Public Sub New(ByVal Collapse As Boolean)
			MyBase.New()
			Me.Collapse = Collapse
		End Sub
	End Class
	Public Event HolderButtonClick(ByVal Sender As System.Object, ByVal e As HolderButtonClickEventArgs)
	<System.Runtime.InteropServices.ProgId("PageFlipEventArgs_NET.PageFlipEventArgs")> Public NotInheritable Class PageFlipEventArgs
		Inherits System.EventArgs
		Public Collapse As Boolean
		Public Sub New(ByVal Collapse As Boolean)
			MyBase.New()
			Me.Collapse = Collapse
		End Sub
	End Class
	Public Event PageFlip(ByVal Sender As System.Object, ByVal e As PageFlipEventArgs)
	
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' USERCONTROL
	'
	'
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'UPGRADE_WARNING: UserControl Event UserControl.InitProperties was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2050"'
	Private Sub UserControl_InitProperties()
		'UPGRADE_ISSUE: UserControl property PageHolder.Extender was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Extender.Name. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		sHldrTxt = Extender.Name
		oHldrTxtClr = System.Drawing.Color.White
		bHldrLne = True
		lPageHeight = VB6.PixelsToTwipsY(MyBase.Height)
		Call zRedrawHolder(lHldrStyle)
	End Sub
	Private Sub PageHolder_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		Holder.Width = MyBase.Width
		HolderLne.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Holder.Width))
		If lPageState = 0 Then
			lPageHeight = VB6.PixelsToTwipsY(MyBase.Height)
		Else
			'UPGRADE_ISSUE: UserControl property PageHolder.Extender was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Extender.Height. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			Extender.Height = VB6.PixelsToTwipsY(Holder.Height)
		End If
	End Sub
	Private Sub HolderBtn_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles HolderBtn.Click
		If lPageState = 0 Then
			lPageState = 1
			Call zCollapse()
		Else
			lPageState = 0
			Call zCollapse(False)
		End If
		RaiseEvent HolderButtonClick(Me, New HolderButtonClickEventArgs(lPageState))
	End Sub
	Private Sub HolderLbl_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles HolderLbl.Click
		Call HolderBtn_Click(HolderBtn, New System.EventArgs())
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
		lHldrStyle = PropBag.ReadProperty("HldrStyle", 0)
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		lHldrIcon = PropBag.ReadProperty("HldrIcon", 0)
		'UPGRADE_ISSUE: UserControl property PageHolder.Extender was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Extender.Name. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		sHldrTxt = PropBag.ReadProperty("HldrTxt", Extender.Name)
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		oHldrTxtClr = System.Drawing.ColorTranslator.FromOle(PropBag.ReadProperty("HldrTxtClr", System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)))
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		bHldrLne = PropBag.ReadProperty("HldrLne", True)
		
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		lPageState = PropBag.ReadProperty("PageState", 0)
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		lPageHeight = PropBag.ReadProperty("PageHeight", lPageHeight)
		
		Call zRedrawHolder(lHldrStyle)
	End Sub
	'UPGRADE_WARNING: PropertyBag object was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup6003"'
	'UPGRADE_WARNING: UserControl Event WriteProperties is not supported. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup6002"'
	Private Sub UserControl_WriteProperties(ByRef PropBag As PropertyBag)
		'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		PropBag.WriteProperty("HldrStyle", lHldrStyle, 0)
		'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		PropBag.WriteProperty("HldrIcon", lHldrIcon, 0)
		'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		PropBag.WriteProperty("HldrTxt", sHldrTxt)
		'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		PropBag.WriteProperty("HldrTxtClr", oHldrTxtClr)
		'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		PropBag.WriteProperty("HldrLne", bHldrLne)
		
		'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		PropBag.WriteProperty("PageState", lPageState, 0)
		'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		PropBag.WriteProperty("PageHeight", lPageHeight)
	End Sub
	
	
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' PROPERTY SECTION
	'
	'
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Public Property HolderStyle() As eHldrStyle
		Get
			HolderStyle = lHldrStyle
		End Get
		Set(ByVal Value As eHldrStyle)
			lHldrStyle = Value
			Call zRedrawHolder(Value)
			'UPGRADE_ISSUE: UserControl method PageHolder.PropertyChanged was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Call PropertyChanged()
		End Set
	End Property
	
	Public Property HolderIcon() As eHldrIcon
		Get
			HolderIcon = lHldrIcon
		End Get
		Set(ByVal Value As eHldrIcon)
			lHldrIcon = Value
			Call zRedrawHolder(lHldrStyle)
			'UPGRADE_ISSUE: UserControl method PageHolder.PropertyChanged was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Call PropertyChanged()
		End Set
	End Property
	
	Public Property HolderText() As String
		Get
			HolderText = sHldrTxt
		End Get
		Set(ByVal Value As String)
			sHldrTxt = Value
			Call zRedrawHolder(0, True)
			'UPGRADE_ISSUE: UserControl method PageHolder.PropertyChanged was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Call PropertyChanged()
		End Set
	End Property
	
	Public Property HolderTextColor() As System.Drawing.Color
		Get
			HolderTextColor = oHldrTxtClr
		End Get
		Set(ByVal Value As System.Drawing.Color)
			oHldrTxtClr = Value
			Call zRedrawHolder(0, True)
			'UPGRADE_ISSUE: UserControl method PageHolder.PropertyChanged was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Call PropertyChanged()
		End Set
	End Property
	
	Public Property HolderLine() As Boolean
		Get
			HolderLine = bHldrLne
		End Get
		Set(ByVal Value As Boolean)
			bHldrLne = Value
			Call zRedrawHolder(lHldrStyle)
			'UPGRADE_ISSUE: UserControl method PageHolder.PropertyChanged was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Call PropertyChanged()
		End Set
	End Property
	
	Public Property PageCollapse() As Boolean
		Get
			If lPageState = 0 Then
				PageCollapse = False
			Else
				PageCollapse = True
			End If
		End Get
		Set(ByVal Value As Boolean)
			If PageCollapse = True Then
				lPageState = 0
				Call zCollapse(False)
			Else
				lPageState = 1
				Call zCollapse()
			End If
			'UPGRADE_ISSUE: UserControl method PageHolder.PropertyChanged was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Call PropertyChanged()
		End Set
	End Property
	
	
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' PRIVATE METHOD & FUNCTION
	'
	'
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Private Sub zRedrawHolder(ByRef Style As Object, Optional ByRef TextOnly As Boolean = False)
		HolderLbl.Text = sHldrTxt
		HolderLbl.ForeColor = oHldrTxtClr
		If TextOnly = True Then Exit Sub
		
		HolderIcn.Image = ImgCnt(lHldrIcon + 1).Image
		HolderLne.Visible = bHldrLne
		Select Case Style
			Case 0
				HolderIcn.Left = VB6.TwipsToPixelsX(15)
				HolderBtn.Left = VB6.TwipsToPixelsX(255)
				HolderLbl.Left = VB6.TwipsToPixelsX(525)
				HolderIcn.Visible = True
				HolderBtn.Visible = True
				HolderLbl.Visible = True
			Case 1
				HolderIcn.Visible = False
				HolderBtn.Visible = True
				HolderLbl.Visible = True
				HolderBtn.Left = VB6.TwipsToPixelsX(15)
				HolderLbl.Left = VB6.TwipsToPixelsX(255)
			Case 2
				HolderIcn.Visible = False
				HolderBtn.Visible = False
				HolderLbl.Visible = True
				HolderLbl.Left = VB6.TwipsToPixelsX(15)
		End Select
	End Sub
	
	Private Sub zCollapse(Optional ByRef Collapse As Boolean = True)
		If Collapse = True Then
			'UPGRADE_ISSUE: UserControl property PageHolder.Extender was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Extender.Top. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Extender.Height. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			Extender.Top = (Extender.Top + Extender.Height) - VB6.PixelsToTwipsY(Holder.Height)
			'UPGRADE_ISSUE: UserControl property PageHolder.Extender was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Extender.Height. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			Extender.Height = VB6.PixelsToTwipsY(Holder.Height)
			tImg = HolderBtn
			HolderBtn.Image = ImgCnt(0).Image
		Else
			'UPGRADE_ISSUE: UserControl property PageHolder.Extender was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Extender.Top. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			Extender.Top = (Extender.Top + VB6.PixelsToTwipsY(Holder.Height)) - lPageHeight
			'UPGRADE_ISSUE: UserControl property PageHolder.Extender was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Extender.Height. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			Extender.Height = lPageHeight
			HolderBtn.Image = tImg
		End If
		RaiseEvent PageFlip(Me, New PageFlipEventArgs(Collapse))
	End Sub
End Class