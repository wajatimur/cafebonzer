Option Strict Off
Option Explicit On
Friend Class clsAgInfoPrinter
	Private StackJob As New Collection
	
	Private sDeviceName As String
	Private sPort As String
	Private sDriverName As String
	Private sPaperSize As String
	Private sOrientation As String
	
	
	Public Sub Init(ByRef sCmd As String, ByRef Index As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object Index. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object SubVal(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		sDeviceName = SubVal(sCmd, Index & "name")
		'UPGRADE_WARNING: Couldn't resolve default property of object Index. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object SubVal(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		sPort = SubVal(sCmd, Index & "port")
		'UPGRADE_WARNING: Couldn't resolve default property of object Index. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object SubVal(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		sDriverName = SubVal(sCmd, Index & "drivername")
		'UPGRADE_WARNING: Couldn't resolve default property of object Index. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object SubVal(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		sPaperSize = SubVal(sCmd, Index & "papersize")
		'UPGRADE_WARNING: Couldn't resolve default property of object Index. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object SubVal(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		sOrientation = SubVal(sCmd, Index & "orientation")
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
	Private Sub Class_Terminate_Renamed()
		sDeviceName = ""
		sPort = ""
		sDriverName = ""
		sPaperSize = ""
		sOrientation = ""
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	
	Public ReadOnly Property DeviceName() As String
		Get
			DeviceName = sDeviceName
		End Get
	End Property
	Public ReadOnly Property Port() As String
		Get
			Port = sPort
		End Get
	End Property
	Public ReadOnly Property DriverName() As String
		Get
			DriverName = sDriverName
		End Get
	End Property
	Public ReadOnly Property PaperSize() As String
		Get
			PaperSize = sPaperSize
		End Get
	End Property
	Public ReadOnly Property Orientation() As String
		Get
			Orientation = sOrientation
		End Get
	End Property
	
	
	Public ReadOnly Property Jobs(ByVal JobId As Object) As clsAgInfoPrinterJob
		Get
			Jobs = StackJob.Item(JobId)
		End Get
	End Property
	
	Public ReadOnly Property JobsCount() As Integer
		Get
			JobsCount = StackJob.Count()
		End Get
	End Property
	
	Public Sub JobsAdd(ByRef sCmd As String)
		Dim c_tJob As New clsAgInfoPrinterJob
		Dim s_JobId As String
		
		'UPGRADE_WARNING: Couldn't resolve default property of object SubVal(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		s_JobId = Trim(SubVal(sCmd, "jobid"))
		For	Each c_tJob In StackJob
			If c_tJob.JobId = s_JobId Then
				c_tJob.Parse(sCmd)
				Exit Sub
			End If
		Next c_tJob
		
		c_tJob.Parse(sCmd)
		StackJob.Add(c_tJob, c_tJob.JobId)
		'UPGRADE_NOTE: Object c_tJob may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
		c_tJob = Nothing
	End Sub
End Class