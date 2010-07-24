Option Strict Off
Option Explicit On
Friend Class clsAgInfoPrinterJob
	Private s_PrinterName As String
	Private s_JobId As String
	Private s_Status As String
	Private s_Document As String
	Private s_PagePrinted As Integer
	Private s_TotalPages As Integer
	
	Public ReadOnly Property PrinterName() As String
		Get
			PrinterName = s_PrinterName
		End Get
	End Property
	Public ReadOnly Property JobId() As String
		Get
			JobId = s_JobId
		End Get
	End Property
	Public ReadOnly Property Status() As String
		Get
			Status = s_Status
		End Get
	End Property
	Public ReadOnly Property Document() As String
		Get
			Document = s_Document
		End Get
	End Property
	Public ReadOnly Property PagePrinted() As Integer
		Get
			PagePrinted = s_PagePrinted
		End Get
	End Property
	Public ReadOnly Property TotalPages() As Integer
		Get
			TotalPages = s_TotalPages
		End Get
	End Property
	
	Public Sub Parse(ByRef sCmd As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object SubVal(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		s_PrinterName = Trim(SubVal(sCmd, "printername"))
		'UPGRADE_WARNING: Couldn't resolve default property of object SubVal(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		s_JobId = Trim(SubVal(sCmd, "jobid"))
		'UPGRADE_WARNING: Couldn't resolve default property of object SubVal(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		s_Document = Trim(SubVal(sCmd, "document"))
		'UPGRADE_WARNING: Couldn't resolve default property of object SubVal(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		s_TotalPages = CInt(SubVal(sCmd, "totalpages", 0))
		'UPGRADE_WARNING: Couldn't resolve default property of object SubVal(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		s_PagePrinted = CInt(SubVal(sCmd, "pageprinted", 0))
		'UPGRADE_WARNING: Couldn't resolve default property of object SubVal(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		s_Status = Trim(SubVal(sCmd, "status"))
	End Sub
End Class