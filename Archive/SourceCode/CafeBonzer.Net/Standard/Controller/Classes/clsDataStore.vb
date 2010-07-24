Option Strict Off
Option Explicit On
Friend Class clsDataStore
	Private StackData As New Collection
	Public Name As String
	
	
	Public Sub Add(ByRef Data As Object, ByRef Key As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object Key. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		StackData.Add(Data, "#" & Key)
	End Sub
	
	Public Sub Remove(ByRef Key As Object)
		StackData.Remove(Key)
	End Sub
	
	Public Sub Clear()
		'UPGRADE_NOTE: Object StackData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
		StackData = Nothing
	End Sub
	
	Public Function Count() As Integer
		Count = StackData.Count()
	End Function
	
	Public Function Data(ByRef Key As Object) As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object Key. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object StackData(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		DataControl = StackData.Item("#" & Key)
	End Function
	
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object StackData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
		StackData = Nothing
		Name = ""
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class