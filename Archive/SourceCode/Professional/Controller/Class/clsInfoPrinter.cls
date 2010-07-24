VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAgInfoPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private StackJob As New Collection

Private sDeviceName As String
Private sPort As String
Private sDriverName As String
Private sPaperSize As String
Private sOrientation As String


Public Sub Init(sCmd As String, Index)
    sDeviceName = SubVal(sCmd, Index & "name")
    sPort = SubVal(sCmd, Index & "port")
    sDriverName = SubVal(sCmd, Index & "drivername")
    sPaperSize = SubVal(sCmd, Index & "papersize")
    sOrientation = SubVal(sCmd, Index & "orientation")
End Sub

Private Sub Class_Terminate()
    sDeviceName = ""
    sPort = ""
    sDriverName = ""
    sPaperSize = ""
    sOrientation = ""
End Sub


Public Property Get DeviceName() As String
    DeviceName = sDeviceName
End Property
Public Property Get Port() As String
    Port = sPort
End Property
Public Property Get DriverName() As String
    DriverName = sDriverName
End Property
Public Property Get PaperSize() As String
    PaperSize = sPaperSize
End Property
Public Property Get Orientation() As String
    Orientation = sOrientation
End Property


Public Property Get Jobs(JobId) As clsAgInfoPrinterJob
    Set Jobs = StackJob(JobId)
End Property

Public Sub JobsAdd(sCmd As String)
    Dim c_tJob As New clsAgInfoPrinterJob, s_JobId As String
    
    s_JobId = Trim(SubVal(sCmd, "jobid"))
    For Each c_tJob In StackJob
        If c_tJob.JobId = s_JobId Then
            c_tJob.Parse sCmd
            Exit Sub
        End If
    Next
    
    c_tJob.Parse sCmd
    StackJob.Add c_tJob, c_tJob.JobId
    Set c_tJob = Nothing
End Sub

Public Property Get JobsCount() As Long
    JobsCount = StackJob.Count
End Property
