VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAgInfoPrinterJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private s_PrinterName As String
Private s_JobId As String
Private s_Status As String
Private s_Document As String
Private s_PagePrinted As Long
Private s_TotalPages As Long

Public Property Get PrinterName() As String
    PrinterName = s_PrinterName
End Property
Public Property Get JobId() As String
    JobId = s_JobId
End Property
Public Property Get Status() As String
    Status = s_Status
End Property
Public Property Get Document() As String
    Document = s_Document
End Property
Public Property Get PagePrinted() As Long
    PagePrinted = s_PagePrinted
End Property
Public Property Get TotalPages() As Long
    TotalPages = s_TotalPages
End Property

Public Sub Parse(sCmd As String)
    s_PrinterName = Trim(SubVal(sCmd, "printername"))
    s_JobId = Trim(SubVal(sCmd, "jobid"))
    s_Document = Trim(SubVal(sCmd, "document"))
    s_TotalPages = CLng(SubVal(sCmd, "totalpages", 0))
    s_PagePrinted = CLng(SubVal(sCmd, "pageprinted", 0))
    s_Status = Trim(SubVal(sCmd, "status"))
End Sub
