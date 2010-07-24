VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEnkrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Function Ekrip(Data, Pass)
    Dim d As String, p As String, d2 As String
    Dim rH As Integer, g As Integer
    Dim BitArray(128) As Integer
    
    
    d = StrReverse(Data)
    p = StrReverse(Pass)
    
    Randomize
    rH = 9 * Rnd
    
    For X = 1 To Len(d)
        g = Asc(Mid(d, 1, 1))
        g = g - rH
        d2 = d2 & Chr(g)
    Next X
End Function
