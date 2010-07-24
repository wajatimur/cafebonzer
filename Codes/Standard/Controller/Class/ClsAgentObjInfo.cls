VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAgInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private StackPrinter As New Collection
Private Parent As clsAgent

Private lMemLoad As Long
Private lMemPhyTotal As Long
Private lMemPhyAvail As Long
Private lMemVirTotal As Long
Private lMemVirAvail As Long
Private lMemPageTotal As Long
Private lMemPageAvail As Long


Public Sub Init(Agent As clsAgent)
    Set Parent = Agent
End Sub


Public Property Get Printers(Key) As clsAgInfoPrinter
    If Key = "" Then Exit Property
    Set Printers = StackPrinter(Key)
End Property
Public Property Get PrintersCount() As Long
    PrintersCount = StackPrinter.Count
End Property


Public Property Get MemLoad() As Long
    MemLoad = lMemLoad
End Property
Public Property Get MemPhyTotal() As Long
    MemPhyTotal = lMemPhyTotal
End Property
Public Property Get MemPhyAvail() As Long
    MemPhyAvail = lMemPhyAvail
End Property
Public Property Get MemVirTotal() As Long
    MemVirTotal = lMemVirTotal
End Property
Public Property Get MemVirAvail() As Long
    MemVirAvail = lMemVirAvail
End Property
Public Property Get MemPageTotal() As Long
    MemPageTotal = lMemPageTotal
End Property
Public Property Get MemPageAvail() As Long
    MemPageAvail = lMemPageAvail
End Property


