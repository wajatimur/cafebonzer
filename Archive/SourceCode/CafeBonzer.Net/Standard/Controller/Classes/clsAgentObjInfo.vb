Option Strict Off
Option Explicit On
Friend Class clsAgInfo
	Private StackPrinter As New Collection
	Private Parent As clsAgent
	
	Private lMemLoad As Integer
	Private lMemPhyTotal As Integer
	Private lMemPhyAvail As Integer
	Private lMemVirTotal As Integer
	Private lMemVirAvail As Integer
	Private lMemPageTotal As Integer
	Private lMemPageAvail As Integer
	
	
	Public Sub Init(ByRef Agent As clsAgent)
		Parent = Agent
	End Sub
	
	
	Public ReadOnly Property Printers(ByVal Key As Object) As clsAgInfoPrinter
		Get
			'UPGRADE_WARNING: Couldn't resolve default property of object Key. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			If Key = "" Then Exit Property
			Printers = StackPrinter.Item(Key)
		End Get
	End Property
	Public ReadOnly Property PrintersCount() As Integer
		Get
			PrintersCount = StackPrinter.Count()
		End Get
	End Property
	
	
	Public ReadOnly Property MemLoad() As Integer
		Get
			MemLoad = lMemLoad
		End Get
	End Property
	Public ReadOnly Property MemPhyTotal() As Integer
		Get
			MemPhyTotal = lMemPhyTotal
		End Get
	End Property
	Public ReadOnly Property MemPhyAvail() As Integer
		Get
			MemPhyAvail = lMemPhyAvail
		End Get
	End Property
	Public ReadOnly Property MemVirTotal() As Integer
		Get
			MemVirTotal = lMemVirTotal
		End Get
	End Property
	Public ReadOnly Property MemVirAvail() As Integer
		Get
			MemVirAvail = lMemVirAvail
		End Get
	End Property
	Public ReadOnly Property MemPageTotal() As Integer
		Get
			MemPageTotal = lMemPageTotal
		End Get
	End Property
	Public ReadOnly Property MemPageAvail() As Integer
		Get
			MemPageAvail = lMemPageAvail
		End Get
	End Property
End Class