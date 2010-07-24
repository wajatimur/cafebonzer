Attribute VB_Name = "MdlData"
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Setting Save | Database
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Sub SetSaveDb(Setting As String, Value As Variant)
    Dim Db As Database, Rs As Recordset
    Set Db = OpenDatabase(App.Path & "\data\sdata.mdb", False, False, ";pwd=nsb2003")
    Set Rs = Db.OpenRecordset(":setting", dbOpenTable)
    
    With Rs
        .Index = "setting"
        .Seek "=", Setting
        If .NoMatch = True Then
            .AddNew
            !Setting = Setting
            !Value = Value
            .Update
        Else
            .Edit
            !Value = Value
            .Update
        End If
    End With
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Setting Get | Database
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function SetGetDb(Setting As String, Optional Default As Variant = "") As Variant
    Dim Db As Database, Rs As Recordset
    Set Db = OpenDatabase(App.Path & "\data\sdata.mdb", False, False, ";pwd=nsb2003")
    Set Rs = Db.OpenRecordset(":setting", dbOpenSnapshot)
    
    SetGetDb = Default
    With Rs
        .FindFirst "setting = '" & Setting & "'"
        If .NoMatch = False Then SetGetDb = !Value
    End With
End Function
