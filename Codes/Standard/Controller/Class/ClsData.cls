VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : clsData
'    Project    : CafeBonzer
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Small Database Engine (DAO)
' Author : Azri Jamil
' Date : N/A
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

'Private local variable
Private CurFilter As String
Private CurDB As String
Private Db As Database
Private CRset As Recordset
Private RecordSetType As RecordsetTypeEnum

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
' Database Initialise
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
Property Let InitDb(DbName As String)
'On Error GoTo ErrInt
    Set Db = OpenDatabase(DbName, False, False, ";pwd=nsb2003")
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
    If NewData = True Then Set CRset = Db.OpenRecordset(Recordset, dbOpenDynaset)
     
    If NewData = True Then CRset.AddNew
    CRset.Fields(DataField).Value = DataValue
    If UpdateOk = True Then CRset.Update: CRset.Close

Exit Sub
ErrInt:
    MsgBox Err.Description, vbExclamation, "Nematix DataLib - DataSave"
End Sub

Public Sub DataEdit(Recordset, DataField, IndexName, IndexValue, Value, NewEdit As Boolean, UpdateOk As Boolean)
    On Error GoTo ErrInt
    If NewEdit = True Then Set CRset = Db.OpenRecordset(Recordset, dbOpenTable)
    
    If NewEdit = True Then
        CRset.Index = IndexName
        CRset.Seek "=", IndexValue
        CRset.Edit
    End If
    CRset.Fields(DataField).Value = Value
    If UpdateOk = True Then CRset.Update

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
    CRset.Filter = CurFilter
End Sub

Public Sub FilterClear()
    CurFilter = ""
    CRset.Filter = CurFilter
End Sub

Public Function DataCount(Recordset)
    On Error GoTo ErrInt
    Set CRset = Db.OpenRecordset(Recordset, dbOpenDynaset)
    
    If CRset.BOF = True Then DataCount = 0: Exit Function
    CRset.MoveLast
    CRset.MoveFirst
    DataCount = CRset.RecordCount
    CRset.Close

Exit Function
ErrInt:
    MsgBox Err.Description, vbExclamation, "Nematix DataLib - DataCount"
End Function

Public Function DataGet(Recordset, DataField, Position)
    On Error GoTo ErrInt
    Set CRset = Db.OpenRecordset(Recordset, dbOpenDynaset)
    
    If CRset.BOF = True Then DataGet = "": Exit Function
    CRset.Move Position
    DataGet = CRset.Fields(DataField)
    CRset.Close
    
Exit Function
ErrInt:
    MsgBox Err.Description, vbExclamation, "Nematix DataLib - DataGet"
End Function

Public Function DataFind(Recordset, IndexName, DataField, Value)
    On Error GoTo ErrInt
    Set CRset = Db.OpenRecordset(Recordset, dbOpenTable)
    
    CRset.Index = IndexName
    CRset.Seek "=", Value
    If CRset.NoMatch = False Then
        DataFind = CRset.Fields(DataField).Value
    Else
        DataFind = -1
    End If
    CRset.Close

Exit Function
ErrInt:
    MsgBox Err.Description, vbExclamation, "Nematix DataLib - DataFind"
End Function


Public Function DataSeek(Recordset, IndexName, Value) As Boolean
On Error GoTo ErrInt
    Set CRset = Db.OpenRecordset(Recordset, dbOpenTable)
    
    CRset.Index = IndexName
    CRset.Seek "=", Value
    If CRset.NoMatch = False Then DataSeek = True Else DataSeek = False

Exit Function
ErrInt:
    MsgBox Err.Description, vbExclamation, "Nematix DataLib - DataSeek"
End Function

Public Function DataRemove(Recordset, IndexName, Value) As Boolean
On Error GoTo ErrInt
    Set CRset = Db.OpenRecordset(Recordset, dbOpenTable)
    
    CRset.Index = IndexName
    CRset.Seek "=", Value
    If CRset.NoMatch = False Then
        CRset.Delete
        DataRemove = True
    End If
    CRset.Close
Exit Function
ErrInt:
    MsgBox Err.Description, vbExclamation, "Nematix DataLib - DataRemove"
End Function


'=--==--==--==--==--==--==--==--==--==--==--==--==--==--==--=
' Internal Setting System
'=--==--==--==--==--==--==--==--==--==--==--==--==--==--==--=
' - Save Setting
Public Sub DbSaveSetting(SetName, SetValue)
    Set CRset = Db.OpenRecordset(":setting", dbOpenDynaset)
    
    With CRset
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
    Set CRset = Db.OpenRecordset(":setting", dbOpenDynaset)
    
    With CRset
        .FindFirst "setting = '" & SetName & "'"
        If .NoMatch = False Then
            DbGetSetting = !Value & ""
        Else
            DbGetSetting = ""
        End If
    End With
End Function
