VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Small Database Engine (DAO)
' Author : Azri Jamil
' Date : N/A
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

'Private local variable
Private CurFilter As String
Private CurDB As String
Private Db As Database
Private Rs As Recordset
Private RecordSetType As RecordsetTypeEnum

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
' Database Initialise
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
Property Let InitDb(DbName As String)
'On Error GoTo ErrInt
    Set Db = OpenDatabase(DbName)
    CurDB = DbName
    RecordSetType = dbOpenDynaset
' Error Control
Exit Property
ErrInt:
    Select Case Err.Number
    Case 3024
        MsgBox Err.Description, vbExclamation, "Nematix DataLib - InitDB(Property_Let)3024"
    Case Else
        MsgBox Err.Description, vbExclamation, "Nematix DataLib - InitDB(Property_Let)"
    End Select
End Property
Property Get InitDb() As String
'On Error GoTo ErrInt
    InitDb = CurDB

Exit Property
ErrInt:
    MsgBox Err.Description, vbExclamation, "Nematix DataLib - InitDb - InitDB(Property_Get)"
End Property

Property Let RsType(RsetType As RecordsetTypeEnum)
    RecordSetType = RsetType
End Property
Property Get RsType() As RecordsetTypeEnum
    RsType = RecordSetType
End Property


Public Sub DataSave(Recordset, DataField, DataValue, NewData As Boolean, UpdateOk As Boolean)
    On Error GoTo ErrInt
    If NewData = True Then Set Rs = Db.OpenRecordset(Recordset, dbOpenDynaset)
     
    If NewData = True Then Rs.AddNew
    Rs.Fields(DataField).Value = DataValue
    If UpdateOk = True Then Rs.Update: Rs.Close

Exit Sub
ErrInt:
    MsgBox Err.Description, vbExclamation, "Nematix DataLib - DataSave"
End Sub

Public Sub DataEdit(Recordset, DataField, IndexName, IndexValue, Value, NewEdit As Boolean, UpdateOk As Boolean)
    On Error GoTo ErrInt
    If NewEdit = True Then Set Rs = Db.OpenRecordset(Recordset, dbOpenTable)
    
    If NewEdit = True Then
        Rs.Index = IndexName
        Rs.Seek "=", IndexValue
        Rs.Edit
    End If
    Rs.Fields(DataField).Value = Value
    If UpdateOk = True Then Rs.Update

Exit Sub
ErrInt:
    MsgBox Err.Description, vbExclamation, "Nematix DataLib - DataEdit"
End Sub

Public Sub FilterAdd(FieldName, Value)
    If CurFilter = "" Then
        CurFilter = FieldName & " = '" & Value & "'"
    Else
        CurFilter = CurFilter & " AND " & FieldName & " = '" & Value & "'"
    End If
    Rs.Filter = CurFilter
End Sub

Public Sub FilterClear()
    CurFilter = ""
    Rs.Filter = CurFilter
End Sub

Public Function DataCount(Recordset)
    On Error GoTo ErrInt
    Set Rs = Db.OpenRecordset(Recordset, dbOpenDynaset)
    
    If Rs.BOF = True Then DataCount = 0: Exit Function
    Rs.MoveLast
    Rs.MoveFirst
    DataCount = Rs.RecordCount
    Rs.Close

Exit Function
ErrInt:
    MsgBox Err.Description, vbExclamation, "Nematix DataLib - DataCount"
End Function

Public Function DataGet(Recordset, DataField, Position)
    On Error GoTo ErrInt
    Set Rs = Db.OpenRecordset(Recordset, dbOpenDynaset)
    
    If Rs.BOF = True Then DataGet = "": Exit Function
    Rs.Move Position
    DataGet = Rs.Fields(DataField)
    Rs.Close
    
Exit Function
ErrInt:
    MsgBox Err.Description, vbExclamation, "Nematix DataLib - DataGet"
End Function

Public Function DataFind(Recordset, IndexName, DataField, Value)
    On Error GoTo ErrInt
    Set Rs = Db.OpenRecordset(Recordset, dbOpenTable)
    
    Rs.Index = IndexName
    Rs.Seek "=", Value
    If Rs.NoMatch = False Then
        DataFind = Rs.Fields(DataField).Value
    Else
        DataFind = -1
    End If
    Rs.Close

Exit Function
ErrInt:
    MsgBox Err.Description, vbExclamation, "Nematix DataLib - DataFind"
End Function


Public Function DataSeek(Recordset, IndexName, Value) As Boolean
On Error GoTo ErrInt
    Set Rs = Db.OpenRecordset(Recordset, dbOpenTable)
    
    Rs.Index = IndexName
    Rs.Seek "=", Value
    If Rs.NoMatch = False Then DataSeek = True Else DataSeek = False

Exit Function
ErrInt:
    MsgBox Err.Description, vbExclamation, "Nematix DataLib - DataSeek"
End Function

Public Function DataRemove(Recordset, IndexName, Value) As Boolean
On Error GoTo ErrInt
    Set Rs = Db.OpenRecordset(Recordset, dbOpenTable)
    
    Rs.Index = IndexName
    Rs.Seek "=", Value
    If Rs.NoMatch = False Then
        Rs.Delete
        DataRemove = True
    End If
    Rs.Close
Exit Function
ErrInt:
    MsgBox Err.Description, vbExclamation, "Nematix DataLib - DataRemove"
End Function


'=--==--==--==--==--==--==--==--==--==--==--==--==--==--==--=
' Internal Setting System
'=--==--==--==--==--==--==--==--==--==--==--==--==--==--==--=
' - Save Setting
Public Sub DbSaveSetting(SetName, SetValue)
    Set Rs = Db.OpenRecordset(":setting", dbOpenDynaset)
    
    With Rs
        .FindFirst "setting = '" & SetName & "'"
        If .NoMatch = False Then
            .Edit
            !Value = SetValue
            .Update
        Else
            .AddNew
            !Setting = SetName
            !Value = SetValue
        End If
    End With
End Sub
' - Get Setting
Public Function DbGetSetting(SetName) As String
    Set Rs = Db.OpenRecordset(":setting", dbOpenDynaset)
    
    With Rs
        .FindFirst "setting = '" & SetName & "'"
        If .NoMatch = False Then
            DbGetSetting = !Value & ""
        Else
            DbGetSetting = ""
        End If
    End With
End Function
