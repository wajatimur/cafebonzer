VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsXdata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : clsXdata
'    Project    : CafeBonzer
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Private FieldIndex As Integer
Private ValueIndex As Integer
Private FieldArr() As ClsXdata
Private ValueArr() As String
Public Name As Variant


'=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~
' Field section
'=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~
Public Sub FieldCreate(Name As String)
    FieldIndex = FieldIndex + 1
    ReDim Preserve FieldArr(FieldIndex)
    Set FieldArr(FieldIndex) = New ClsXdata
    FieldArr(FieldIndex).Name = Name
End Sub

Public Function FieldName(FieldNumber As Integer) As String
    FieldName = FieldArr(FieldNumber).Name
End Function

Public Function FieldCount() As Integer
    FieldCount = FieldIndex
End Function


'=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~
' Data section
'=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~
Public Sub DataAdd(Field As String, Value As String, Optional Parent As Boolean = True)
    If Parent = True Then
        For X = 1 To FieldIndex
            If FieldArr(X).Name = Field Then
                FieldArr(X).DataAdd Field, Value, False
                ValueIndex = ValueIndex + 1
            End If
        Next
    Else
        ValueIndex = ValueIndex + 1
        ReDim Preserve ValueArr(ValueIndex)
        ValueArr(ValueIndex) = Value
        EnumField
    End If
End Sub

Public Function DataCount(Optional Parent As Boolean = True)
    If Parent = True Then
        DataCount = FieldArr(1).DataCount(False)
    Else
        DataCount = ValueIndex
    End If
End Function


'=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~
' Misc section
'=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~
Private Sub EnumField()
    Debug.Print "=================================="
    Debug.Print "Name : " & Name
    For X = 1 To ValueIndex
        Debug.Print X & " >> " & ValueArr(X)
    Next
End Sub



