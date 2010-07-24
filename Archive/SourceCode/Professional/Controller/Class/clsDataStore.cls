VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private StackData As New Collection
Public Name As String


Public Sub Add(Data, Key)
    StackData.Add Data, "#" & Key
End Sub

Public Sub Remove(Key)
    StackData.Remove Key
End Sub

Public Sub Clear()
    Set StackData = Nothing
End Sub

Public Function Count() As Long
    Count = StackData.Count
End Function

Public Function Data(Key) As Variant
Attribute Data.VB_UserMemId = 0
    Data = StackData("#" & Key)
End Function


Private Sub Class_Terminate()
    Set StackData = Nothing
    Name = ""
End Sub
