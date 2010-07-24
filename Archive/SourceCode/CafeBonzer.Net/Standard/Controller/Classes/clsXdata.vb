Option Strict Off
Option Explicit On
Friend Class clsXdata
	Private FieldIndex As Short
	Private ValueIndex As Short
	Private FieldArr() As clsXdata
	Private ValueArr() As String
	Public Name As Object
	
	
	'=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~
	' Field section
	'=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~
	Public Sub FieldCreate(ByRef Name As String)
		FieldIndex = FieldIndex + 1
		ReDim Preserve FieldArr(FieldIndex)
		FieldArr(FieldIndex) = New clsXdata
		'UPGRADE_WARNING: Couldn't resolve default property of object FieldArr().Name. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		FieldArr(FieldIndex).Name = Name
	End Sub
	
	Public Function FieldName(ByRef FieldNumber As Short) As String
		'UPGRADE_WARNING: Couldn't resolve default property of object FieldArr().Name. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		FieldName = FieldArr(FieldNumber).Name
	End Function
	
	Public Function FieldCount() As Short
		FieldCount = FieldIndex
	End Function
	
	
	'=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~
	' Data section
	'=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~
	Public Sub DataAdd(ByRef Field As String, ByRef Value As String, Optional ByRef Parent As Boolean = True)
		Dim X As Object
		If Parent = True Then
			For X = 1 To FieldIndex
				'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				'UPGRADE_WARNING: Couldn't resolve default property of object FieldArr(X).Name. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				If FieldArr(X).Name = Field Then
					'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					FieldArr(X).DataAdd(Field, Value, False)
					ValueIndex = ValueIndex + 1
				End If
			Next X
		Else
			ValueIndex = ValueIndex + 1
			ReDim Preserve ValueArr(ValueIndex)
			ValueArr(ValueIndex) = Value
			EnumField()
		End If
	End Sub
	
	Public Function DataCount(Optional ByRef Parent As Boolean = True) As Object
		If Parent = True Then
			'UPGRADE_WARNING: Couldn't resolve default property of object FieldArr().DataCount(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Couldn't resolve default property of object DataCount. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			DataCount = FieldArr(1).DataCount(False)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object DataCount. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			DataCount = ValueIndex
		End If
	End Function
	
	
	'=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~
	' Misc section
	'=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~
	Private Sub EnumField()
		Dim X As Object
		System.Diagnostics.Debug.WriteLine("==================================")
		'UPGRADE_WARNING: Couldn't resolve default property of object Name. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		System.Diagnostics.Debug.WriteLine("Name : " & Name)
		For X = 1 To ValueIndex
			'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			System.Diagnostics.Debug.WriteLine(X & " >> " & ValueArr(X))
		Next X
	End Sub
End Class