Option Strict Off
Option Explicit On
Friend Class Line3D
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
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Friend WithEvents bLight As System.Windows.Forms.Label
	Friend WithEvents bDark As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Line3D))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.bLight = New System.Windows.Forms.Label
		Me.bDark = New System.Windows.Forms.Label
		Me.ClientSize = New System.Drawing.Size(79, 204)
		MyBase.Location = New System.Drawing.Point(0, 0)
		MyBase.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		MyBase.Name = "Line3D"
		Me.bLight.BackColor = System.Drawing.Color.White
		Me.bLight.Visible = True
		Me.bLight.Location = New System.Drawing.Point(7, 2)
		Me.bLight.Width = 1
		Me.bLight.Height = 157
		Me.bLight.Name = "bLight"
		Me.bDark.BackColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me.bDark.Visible = True
		Me.bDark.Location = New System.Drawing.Point(6, 2)
		Me.bDark.Width = 1
		Me.bDark.Height = 159
		Me.bDark.Name = "bDark"
		Me.Controls.Add(bLight)
		Me.Controls.Add(bDark)
	End Sub
#End Region 
	
	Public Enum Align
		Vertical = 1
		Horizontal = 2
	End Enum
	
	'UPGRADE_WARNING: PropertyBag object was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup6003"'
	Public Pbag As New PropertyBag
	Private cHeight As Integer
	Private cWidth As Integer
	Private aHorizon As Boolean
	
	'UPGRADE_WARNING: PropertyBag object was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup6003"'
	'UPGRADE_WARNING: UserControl Event ReadProperties is not supported. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup6002"'
	Private Sub UserControl_ReadProperties(ByRef PropBag As PropertyBag)
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		aHorizon = PropBag.ReadProperty("horizon")
	End Sub
	
	Private Sub Line3D_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		If VB6.PixelsToTwipsX(MyBase.Width) > VB6.PixelsToTwipsY(MyBase.Height) Then aHorizon = True
		If aHorizon = False Then
			cWidth = 50
			cHeight = VB6.PixelsToTwipsY(MyBase.Height)
			'UPGRADE_ISSUE: Line property bDark.X1 is not supported at runtime. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2066"'
			bDark.X1 = 1
			'UPGRADE_ISSUE: Line property bDark.X2 is not supported at runtime. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2066"'
			bDark.X2 = 1
			'UPGRADE_ISSUE: Line property bDark.Y1 is not supported at runtime. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2066"'
			bDark.Y1 = 0
			'UPGRADE_ISSUE: Line property bDark.Y2 is not supported at runtime. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2066"'
			bDark.Y2 = cHeight
			'UPGRADE_ISSUE: Line property bLight.X1 is not supported at runtime. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2066"'
			bLight.X1 = 2
			'UPGRADE_ISSUE: Line property bLight.X2 is not supported at runtime. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2066"'
			bLight.X2 = 2
			'UPGRADE_ISSUE: Line property bLight.Y1 is not supported at runtime. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2066"'
			bLight.Y1 = 0
			'UPGRADE_ISSUE: Line property bLight.Y2 is not supported at runtime. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2066"'
			bLight.Y2 = cHeight
		Else
			cWidth = VB6.PixelsToTwipsX(MyBase.Width)
			cHeight = 50
			'UPGRADE_ISSUE: Line property bDark.X1 is not supported at runtime. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2066"'
			bDark.X1 = 0
			'UPGRADE_ISSUE: Line property bDark.X2 is not supported at runtime. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2066"'
			bDark.X2 = cWidth
			'UPGRADE_ISSUE: Line property bDark.Y1 is not supported at runtime. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2066"'
			bDark.Y1 = 1
			'UPGRADE_ISSUE: Line property bDark.Y2 is not supported at runtime. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2066"'
			bDark.Y2 = 1
			'UPGRADE_ISSUE: Line property bLight.X1 is not supported at runtime. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2066"'
			bLight.X1 = 0
			'UPGRADE_ISSUE: Line property bLight.X2 is not supported at runtime. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2066"'
			bLight.X2 = cWidth
			'UPGRADE_ISSUE: Line property bLight.Y1 is not supported at runtime. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2066"'
			bLight.Y1 = 2
			'UPGRADE_ISSUE: Line property bLight.Y2 is not supported at runtime. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2066"'
			bLight.Y2 = 2
		End If
		
		MyBase.Width = VB6.TwipsToPixelsX(cWidth)
		MyBase.Height = VB6.TwipsToPixelsY(cHeight)
	End Sub
	
	
	
	Public Property Alignment() As Align
		Get
			If aHorizon Then
				Alignment = Align.Horizontal
			Else
				Alignment = Align.Vertical
			End If
		End Get
		Set(ByVal Value As Align)
			Select Case Value
				Case 1
					If aHorizon = False Then Exit Property
					aHorizon = False
					MyBase.Height = VB6.TwipsToPixelsY(MyBase.Width)
				Case 2
					If aHorizon = True Then Exit Property
					aHorizon = True
					MyBase.Width = VB6.TwipsToPixelsX(MyBase.Height)
			End Select
			Call Line3D_Resize(Me, New System.EventArgs())
		End Set
	End Property
	
	'UPGRADE_WARNING: PropertyBag object was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup6003"'
	'UPGRADE_WARNING: UserControl Event WriteProperties is not supported. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup6002"'
	Private Sub UserControl_WriteProperties(ByRef PropBag As PropertyBag)
		'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		PropBag.WriteProperty("horizon", aHorizon)
	End Sub
End Class