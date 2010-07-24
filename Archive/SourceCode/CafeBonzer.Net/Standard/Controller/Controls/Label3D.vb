Option Strict Off
Option Explicit On
Friend Class Label3D
	Inherits System.Windows.Forms.UserControl
#Region "Windows Form Designer generated code "
	Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
		UserControl_Initialize()
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
	Friend WithEvents Label2 As System.Windows.Forms.Label
	Friend WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Label3D))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.ClientSize = New System.Drawing.Size(144, 97)
		MyBase.Location = New System.Drawing.Point(0, 0)
		MyBase.Name = "Label3D"
		Me.Label2.Text = "Label3D"
		Me.Label2.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Size = New System.Drawing.Size(177, 57)
		Me.Label2.Location = New System.Drawing.Point(34, 14)
		Me.Label2.TabIndex = 1
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label2.BackColor = System.Drawing.Color.Transparent
		Me.Label2.Enabled = True
		Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label2.UseMnemonic = True
		Me.Label2.Visible = True
		Me.Label2.AutoSize = False
		Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label2.Name = "Label2"
		Me.Label1.Text = "Label3D"
		Me.Label1.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.ForeColor = System.Drawing.SystemColors.highlightText
		Me.Label1.Size = New System.Drawing.Size(153, 25)
		Me.Label1.Location = New System.Drawing.Point(0, 0)
		Me.Label1.TabIndex = 0
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.BackColor = System.Drawing.Color.Transparent
		Me.Label1.Enabled = True
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
	End Sub
#End Region 
	Public Event EnabledChange()
	Public Event FontChange()
	Public Event BackColorChange()
	Public Event PhaseChange()
	Public Event CaptionChange()
	Public Event ForeColor1Change()
	Public Event ForeColor2Change()
	Public Event BorderStyleChange()
	Public Event AlignmentChange()
	
	Public Enum T_Phase
		TOPLEFT = 3
		TOPRIGHT = 2
		BOTTOMLEFT = 1
		BOTTOMRIGHT = 0
	End Enum
	
	Public Enum T_BorderStyle
		None = 0
		FixedSingle = 1
	End Enum
	
	Public Enum T_Align
		AlignLeft = 0
		AlignRight = 1
		AlignCenter = 2
	End Enum
	
	'Events declaration
	Shadows Event Click(ByVal Sender As System.Object, ByVal e As System.EventArgs) 'MappingInfo=Label1,Label1,-1,Click
	Event DblClick(ByVal Sender As System.Object, ByVal e As System.EventArgs) 'MappingInfo=Label1,Label1,-1,DblClick
	<System.Runtime.InteropServices.ProgId("MouseDownEventArgs_NET.MouseDownEventArgs")> Public NotInheritable Class MouseDownEventArgs
		Inherits System.EventArgs
		Public Button As Short
		Public Shift As Short
		Public x As Single
		Public y As Single
		Public Sub New(ByRef Button As Short, ByRef Shift As Short, ByRef x As Single, ByRef y As Single)
			MyBase.New()
			Me.Button = Button
			Me.Shift = Shift
			Me.x = x
			Me.y = y
		End Sub
	End Class
	Shadows Event MouseDown(ByVal Sender As System.Object, ByVal e As MouseDownEventArgs) 'MappingInfo=Label1,Label1,-1,MouseDown
	<System.Runtime.InteropServices.ProgId("MouseMoveEventArgs_NET.MouseMoveEventArgs")> Public NotInheritable Class MouseMoveEventArgs
		Inherits System.EventArgs
		Public Button As Short
		Public Shift As Short
		Public x As Single
		Public y As Single
		Public Sub New(ByRef Button As Short, ByRef Shift As Short, ByRef x As Single, ByRef y As Single)
			MyBase.New()
			Me.Button = Button
			Me.Shift = Shift
			Me.x = x
			Me.y = y
		End Sub
	End Class
	Shadows Event MouseMove(ByVal Sender As System.Object, ByVal e As MouseMoveEventArgs) 'MappingInfo=Label1,Label1,-1,MouseMove
	<System.Runtime.InteropServices.ProgId("MouseUpEventArgs_NET.MouseUpEventArgs")> Public NotInheritable Class MouseUpEventArgs
		Inherits System.EventArgs
		Public Button As Short
		Public Shift As Short
		Public x As Single
		Public y As Single
		Public Sub New(ByRef Button As Short, ByRef Shift As Short, ByRef x As Single, ByRef y As Single)
			MyBase.New()
			Me.Button = Button
			Me.Shift = Shift
			Me.x = x
			Me.y = y
		End Sub
	End Class
	Shadows Event MouseUp(ByVal Sender As System.Object, ByVal e As MouseUpEventArgs) 'MappingInfo=Label1,Label1,-1,MouseUp
	'default variabled definition
	Const m_def_Phase As Short = 0
	'veriables definition
	Dim m_Phase As Byte
	
	Private Sub UserControl_Initialize()
		Label3D_Resize(Me, New System.EventArgs())
	End Sub
	
	'UPGRADE_WARNING: UserControl Event UserControl.InitProperties was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2050"'
	Private Sub UserControl_InitProperties()
		m_Phase = m_def_Phase
		'UPGRADE_ISSUE: AmbientProperties property Ambient.Font was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		Label1.Font = Ambient.Font
		'UPGRADE_ISSUE: AmbientProperties property Ambient.Font was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		Label2.Font = Ambient.Font
	End Sub
	
	Private Sub Label3D_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		If VB6.PixelsToTwipsX(Width) < 100 Then Width = VB6.TwipsToPixelsX(100)
		If VB6.PixelsToTwipsY(Height) < 100 Then Height = VB6.TwipsToPixelsY(100)
		Label1.Width = Width
		Label1.Height = Height
		Label2.Width = Width
		Label2.Height = Height
		Select Case m_Phase
			Case 3
				Label1.Left = 0
				Label1.Top = 0
				Label2.Left = VB6.TwipsToPixelsX(15)
				Label2.Top = VB6.TwipsToPixelsY(15)
			Case 2
				Label1.Left = VB6.TwipsToPixelsX(15)
				Label1.Top = 0
				Label2.Left = 0
				Label2.Top = VB6.TwipsToPixelsY(15)
			Case 1
				Label1.Left = 0
				Label1.Top = VB6.TwipsToPixelsY(15)
				Label2.Left = VB6.TwipsToPixelsX(15)
				Label2.Top = 0
			Case 0
				Label1.Left = VB6.TwipsToPixelsX(15)
				Label1.Top = VB6.TwipsToPixelsY(15)
				Label2.Left = 0
				Label2.Top = 0
		End Select
	End Sub
	
	'MappingInfo=Label1,Label1,-1,Enabled
	
	Public Shadows Property Enabled() As Boolean
		Get
			Return Label1.Enabled
		End Get
		Set(ByVal Value As Boolean)
			Label1.Enabled = Value
			Label2.Enabled = Value
			RaiseEvent EnabledChange()
		End Set
	End Property
	
	'MappingInfo=Label1,Label1,-1,Font
	
	Public Overrides Property Font() As System.Drawing.Font
		Get
			Font = Label1.Font
		End Get
		Set(ByVal Value As System.Drawing.Font)
			Label1.Font = Value
			Label2.Font = Value
			RaiseEvent FontChange()
		End Set
	End Property
	
	'MappingInfo=Label1,Label1,-1,ForeColor
	
	Public Property ForeColor1() As System.Drawing.Color
		Get
			ForeColor1 = Label1.ForeColor
		End Get
		Set(ByVal Value As System.Drawing.Color)
			Label1.ForeColor = Value
			RaiseEvent ForeColor1Change()
		End Set
	End Property
	
	'MappingInfo=Label2,Label2,-1,ForeColor
	
	Public Property ForeColor2() As System.Drawing.Color
		Get
			ForeColor2 = Label2.ForeColor
		End Get
		Set(ByVal Value As System.Drawing.Color)
			Label2.ForeColor = Value
			RaiseEvent ForeColor2Change()
		End Set
	End Property
	
	'MappingInfo=Label1,Label1,-1,Caption
	
	Public Property Caption() As String
		Get
			Caption = Label1.Text
		End Get
		Set(ByVal Value As String)
			Label1.Text = Value
			Label2.Text = Value
			RaiseEvent CaptionChange()
		End Set
	End Property
	
	'MappingInfo=Label1,Label1,-1,Alignment
	
	Public Property Alignment() As T_Align
		Get
			Alignment = Label1.TextAlign
		End Get
		Set(ByVal Value As T_Align)
			Label1.TextAlign = Value
			Label2.TextAlign = Value
			RaiseEvent AlignmentChange()
		End Set
	End Property
	
	
	Public Property Phase() As T_Phase
		Get
			Phase = m_Phase
		End Get
		Set(ByVal Value As T_Phase)
			m_Phase = Value
			Label3D_Resize(Me, New System.EventArgs())
			RaiseEvent PhaseChange()
		End Set
	End Property
	
	'MappingInfo=UserControl,UserControl,-1,BackColor
	
	Public Overrides Property BackColor() As System.Drawing.Color
		Get
			Return MyBase.BackColor
		End Get
		Set(ByVal Value As System.Drawing.Color)
			MyBase.BackColor = Value
			RaiseEvent BackColorChange()
		End Set
	End Property
	
	'MappingInfo=Label2,Label2,-1,BorderStyle
	
	Public Property BorderStyle() As T_BorderStyle
		Get
			BorderStyle = Label2.BorderStyle
		End Get
		Set(ByVal Value As T_BorderStyle)
			Label2.BorderStyle = Value
			RaiseEvent BorderStyleChange()
		End Set
	End Property
	
	'MappingInfo=UserControl,UserControl,-1,Refresh
	Public Overrides Sub Refresh()
		MyBase.Refresh()
	End Sub
	
	Private Sub Label2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Label2.Click
		RaiseEvent Click(Me, Nothing)
	End Sub
	
	Private Sub Label2_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Label2.DoubleClick
		RaiseEvent DblClick(Me, Nothing)
	End Sub
	
	Private Sub Label2_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles Label2.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		RaiseEvent MouseDown(Me, New MouseDownEventArgs(Button, Shift, x, y))
	End Sub
	
	Private Sub Label2_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles Label2.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		RaiseEvent MouseMove(Me, New MouseMoveEventArgs(Button, Shift, x, y))
	End Sub
	
	Private Sub Label2_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles Label2.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		RaiseEvent MouseUp(Me, New MouseUpEventArgs(Button, Shift, x, y))
	End Sub
	
	'UPGRADE_WARNING: PropertyBag object was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup6003"'
	'UPGRADE_WARNING: UserControl Event ReadProperties is not supported. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup6002"'
	Private Sub UserControl_ReadProperties(ByRef PropBag As PropertyBag)
		
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Label1.Enabled = PropBag.ReadProperty("Enabled", True)
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Label2.Enabled = PropBag.ReadProperty("Enabled", True)
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Label1.ForeColor = System.Drawing.ColorTranslator.FromOle(PropBag.ReadProperty("ForeColor1", &H0s))
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Label2.ForeColor = System.Drawing.ColorTranslator.FromOle(PropBag.ReadProperty("ForeColor2", &HFFFFFF))
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Label1.Text = PropBag.ReadProperty("Caption", "Label3D")
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Label2.Text = PropBag.ReadProperty("Caption", "Label3D")
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Label1.TextAlign = PropBag.ReadProperty("Alignment", 0)
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Label2.TextAlign = PropBag.ReadProperty("Alignment", 0)
		'UPGRADE_ISSUE: AmbientProperties property Ambient.Font was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		Label1.Font = PropBag.ReadProperty("Font", Ambient.Font)
		'UPGRADE_ISSUE: AmbientProperties property Ambient.Font was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		Label2.Font = PropBag.ReadProperty("Font", Ambient.Font)
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Label2.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		m_Phase = PropBag.ReadProperty("Phase", m_def_Phase)
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		MyBase.BackColor = System.Drawing.ColorTranslator.FromOle(PropBag.ReadProperty("BackColor", &H8000000F))
		Label3D_Resize(Me, New System.EventArgs())
	End Sub
	
	'UPGRADE_WARNING: PropertyBag object was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup6003"'
	'UPGRADE_WARNING: UserControl Event WriteProperties is not supported. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup6002"'
	Private Sub UserControl_WriteProperties(ByRef PropBag As PropertyBag)
		
		'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		Call PropBag.WriteProperty("Enabled", Label1.Enabled, True)
		'UPGRADE_ISSUE: AmbientProperties property Ambient.Font was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		Call PropBag.WriteProperty("Font", Label1.Font, Ambient.Font)
		'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		Call PropBag.WriteProperty("ForeColor1", System.Drawing.ColorTranslator.ToOle(Label1.ForeColor), &H0s)
		'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		Call PropBag.WriteProperty("ForeColor2", System.Drawing.ColorTranslator.ToOle(Label2.ForeColor), &HFFFFFF)
		'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		Call PropBag.WriteProperty("Caption", Label1.Text, "Label3D")
		'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		Call PropBag.WriteProperty("Alignment", Label1.TextAlign, 0)
		'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		Call PropBag.WriteProperty("Phase", m_Phase, m_def_Phase)
		'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		Call PropBag.WriteProperty("BackColor", System.Drawing.ColorTranslator.ToOle(MyBase.BackColor), &H8000000F)
		'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		Call PropBag.WriteProperty("BorderStyle", Label2.BorderStyle, 0)
	End Sub
End Class