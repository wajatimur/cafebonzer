Option Strict Off
Option Explicit On
Friend Class clsData
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Small Database Engine (DAO)
	' Author : Azri Jamil
	' Date : N/A
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	
	'Private local variable
	Private CurFilter As String
	Private CurDB As String
	Private Db As DAO.Database
	Private Rs As DAO.Recordset
	Private RecordSetType As DAO.RecordsetTypeEnum
	
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
	' Database Initialise
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
	Property InitDb() As String
		Get
			'On Error GoTo ErrInt
			InitDb = CurDB
			
			Exit Property
ErrInt: 
			MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Nematix DataLib - InitDb - InitDB(Property_Get)")
		End Get
		Set(ByVal Value As String)
			'On Error GoTo ErrInt
			Db = DAODBEngine_definst.OpenDatabase(Value, False, False, ";pwd=nsb2003")
			CurDB = Value
			RecordSetType = DAO.RecordsetTypeEnum.dbOpenDynaset
			' Error Control
			Exit Property
ErrInt: 
			Select Case Err.Number
				Case 3024
					MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Nematix DataLib - InitDB(Property_Let)3024")
				Case Else
					MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Nematix DataLib - InitDB(Property_Let)")
			End Select
		End Set
	End Property
	
	Property RsType() As DAO.RecordsetTypeEnum
		Get
			RsType = RecordSetType
		End Get
		Set(ByVal Value As DAO.RecordsetTypeEnum)
			RecordSetType = Value
		End Set
	End Property
	
	
	Public Sub DataSave(ByRef Recordset As Object, ByRef DataField As Object, ByRef DataValue As Object, ByRef NewData As Boolean, ByRef UpdateOk As Boolean)
		On Error GoTo ErrInt
		'UPGRADE_WARNING: Couldn't resolve default property of object Recordset. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If NewData = True Then Rs = Db.OpenRecordset(Recordset, DAO.RecordsetTypeEnum.dbOpenDynaset)
		
		If NewData = True Then Rs.AddNew()
		'UPGRADE_WARNING: Couldn't resolve default property of object DataValue. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Rs.Fields(DataField).Value = DataValue
		If UpdateOk = True Then Rs.Update() : Rs.Close()
		
		Exit Sub
ErrInt: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Nematix DataLib - DataSave")
	End Sub
	
	Public Sub DataEdit(ByRef Recordset As Object, ByRef DataField As Object, ByRef IndexName As Object, ByRef IndexValue As Object, ByRef Value As Object, ByRef NewEdit As Boolean, ByRef UpdateOk As Boolean)
		On Error GoTo ErrInt
		'UPGRADE_WARNING: Couldn't resolve default property of object Recordset. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If NewEdit = True Then Rs = Db.OpenRecordset(Recordset, DAO.RecordsetTypeEnum.dbOpenTable)
		
		If NewEdit = True Then
			'UPGRADE_WARNING: Couldn't resolve default property of object IndexName. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			Rs.Index = IndexName
			Rs.Seek("=", IndexValue)
			Rs.Edit()
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object Value. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Rs.Fields(DataField).Value = Value
		If UpdateOk = True Then Rs.Update()
		
		Exit Sub
ErrInt: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Nematix DataLib - DataEdit")
	End Sub
	
	Public Sub FilterAdd(ByRef FieldName As Object, ByRef Value As Object)
		If CurFilter = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Value. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Couldn't resolve default property of object FieldName. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			CurFilter = FieldName & " = '" & Value & "'"
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object Value. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Couldn't resolve default property of object FieldName. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			CurFilter = CurFilter & " AND " & FieldName & " = '" & Value & "'"
		End If
		Rs.Filter = CurFilter
	End Sub
	
	Public Sub FilterClear()
		CurFilter = ""
		Rs.Filter = CurFilter
	End Sub
	
	Public Function DataCount(ByRef Recordset As Object) As Object
		On Error GoTo ErrInt
		'UPGRADE_WARNING: Couldn't resolve default property of object Recordset. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Rs = Db.OpenRecordset(Recordset, DAO.RecordsetTypeEnum.dbOpenDynaset)
		
		If Rs.BOF = True Then
			'UPGRADE_WARNING: Couldn't resolve default property of object DataCount. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			DataCount = 0 : Exit Function
		End If
		Rs.MoveLast()
		Rs.MoveFirst()
		'UPGRADE_WARNING: Couldn't resolve default property of object DataCount. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		DataCount = Rs.RecordCount
		Rs.Close()
		
		Exit Function
ErrInt: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Nematix DataLib - DataCount")
	End Function
	
	Public Function DataGet(ByRef Recordset As Object, ByRef DataField As Object, ByRef Position As Object) As Object
		On Error GoTo ErrInt
		'UPGRADE_WARNING: Couldn't resolve default property of object Recordset. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Rs = Db.OpenRecordset(Recordset, DAO.RecordsetTypeEnum.dbOpenDynaset)
		
		If Rs.BOF = True Then
			'UPGRADE_WARNING: Couldn't resolve default property of object DataGet. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			DataGet = "" : Exit Function
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object Position. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Rs.Move(Position)
		'UPGRADE_WARNING: Couldn't resolve default property of object DataGet. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		DataGet = Rs.Fields(DataField).Value
		Rs.Close()
		
		Exit Function
ErrInt: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Nematix DataLib - DataGet")
	End Function
	
	Public Function DataFind(ByRef Recordset As Object, ByRef IndexName As Object, ByRef DataField As Object, ByRef Value As Object) As Object
		On Error GoTo ErrInt
		'UPGRADE_WARNING: Couldn't resolve default property of object Recordset. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Rs = Db.OpenRecordset(Recordset, DAO.RecordsetTypeEnum.dbOpenTable)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object IndexName. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Rs.Index = IndexName
		Rs.Seek("=", Value)
		If Rs.NoMatch = False Then
			'UPGRADE_WARNING: Couldn't resolve default property of object DataFind. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			DataFind = Rs.Fields(DataField).Value
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object DataFind. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			DataFind = -1
		End If
		Rs.Close()
		
		Exit Function
ErrInt: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Nematix DataLib - DataFind")
	End Function
	
	
	Public Function DataSeek(ByRef Recordset As Object, ByRef IndexName As Object, ByRef Value As Object) As Boolean
		On Error GoTo ErrInt
		'UPGRADE_WARNING: Couldn't resolve default property of object Recordset. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Rs = Db.OpenRecordset(Recordset, DAO.RecordsetTypeEnum.dbOpenTable)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object IndexName. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Rs.Index = IndexName
		Rs.Seek("=", Value)
		If Rs.NoMatch = False Then DataSeek = True Else DataSeek = False
		
		Exit Function
ErrInt: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Nematix DataLib - DataSeek")
	End Function
	
	Public Function DataRemove(ByRef Recordset As Object, ByRef IndexName As Object, ByRef Value As Object) As Boolean
		On Error GoTo ErrInt
		'UPGRADE_WARNING: Couldn't resolve default property of object Recordset. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Rs = Db.OpenRecordset(Recordset, DAO.RecordsetTypeEnum.dbOpenTable)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object IndexName. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Rs.Index = IndexName
		Rs.Seek("=", Value)
		If Rs.NoMatch = False Then
			Rs.Delete()
			DataRemove = True
		End If
		Rs.Close()
		Exit Function
ErrInt: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Nematix DataLib - DataRemove")
	End Function
	
	
	'=--==--==--==--==--==--==--==--==--==--==--==--==--==--==--=
	' Internal Setting System
	'=--==--==--==--==--==--==--==--==--==--==--==--==--==--==--=
	' - Save Setting
	Public Sub DbSaveSetting(ByRef SetName As Object, ByRef SetValue As Object)
		Rs = Db.OpenRecordset(":setting", DAO.RecordsetTypeEnum.dbOpenDynaset)
		
		With Rs
			'UPGRADE_WARNING: Couldn't resolve default property of object SetName. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			.FindFirst("setting = '" & SetName & "'")
			If .NoMatch = False Then
				.Edit()
				'UPGRADE_WARNING: Couldn't resolve default property of object SetValue. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				.Fields("Value").Value = SetValue
				.Update()
			Else
				.AddNew()
				'UPGRADE_WARNING: Couldn't resolve default property of object SetName. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				.Fields("Setting").Value = SetName
				'UPGRADE_WARNING: Couldn't resolve default property of object SetValue. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				.Fields("Value").Value = SetValue
			End If
		End With
	End Sub
	' - Get Setting
	Public Function DbGetSetting(ByRef SetName As Object) As String
		Rs = Db.OpenRecordset(":setting", DAO.RecordsetTypeEnum.dbOpenDynaset)
		
		With Rs
			'UPGRADE_WARNING: Couldn't resolve default property of object SetName. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			.FindFirst("setting = '" & SetName & "'")
			If .NoMatch = False Then
				DbGetSetting = .Fields("Value").Value & ""
			Else
				DbGetSetting = ""
			End If
		End With
	End Function
End Class