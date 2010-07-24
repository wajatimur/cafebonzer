Attribute VB_Name = "CasGenApplication"
Public sVsInputRet As String

Public BlnSecPassOnly As Boolean
Public StrSecPassUser As String
Public StrSecPassPassword As String

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Error handler
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub ErrHand(ErrObj As ErrObject, ProcName As String)
On Error GoTo ErrInt
    Dim IntErrNum As Integer, StrErrDesc As String, StrErrSource As String
    Dim ErrDesc As String
    
    IntErrNum = ErrObj.Number
    StrErrSource = ErrObj.Source
    StrErrDesc = ErrObj.Description
    
    MsgBox IntErrNum & " / " & StrErrSource & vbNewLine & StrErrDesc, vbExclamation, "CgAgXLib | " & ProcName
Exit Sub
ErrInt:
    MsgBox Err.Number & " / " & Err.Source & vbNewLine & Err.Description, vbExclamation, "CbAgXLib | Error Handler"
End Sub


Function SubBuild(DataName As String, DataValue As String) As String
On Error GoTo ErrInt
    SubBuild = "{" & DataName & "|'" & DataValue & "'}"
Exit Function

ErrInt:
    ErrHand Err, "Module Command | SubBuild"
End Function


'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Is Windows In Task] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Function IsTaskWindow(hWnd As Long) As Boolean
    Dim lngStyle As Long, IsTask As Long
    IsTask = WS_VISIBLE Or WS_BORDER
    lngStyle = GetWindowLong(hWnd, GWL_STYLE)
    If (lngStyle And IsTask) = IsTask Then IsTaskWindow = True
End Function

Public Function GetTitle(hWnd As Long) As String
    Dim sBuffer As String * 64
    GetWindowText hWnd, sBuffer, 64
    GetTitle = Left$(sBuffer, InStr(1, sBuffer, Chr(0)) - 1)
End Function

Public Function GetClass(hWnd As Long) As String
    Dim sBuffer As String * 64
    GetClassName hWnd, sBuffer, 64
    GetClass = Left$(sBuffer, InStr(1, sBuffer, Chr(0)) - 1)
End Function

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Find Taskbar Handle] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Function FindShellTaskBar() As Long
    Dim hWnd As Long
    On Error Resume Next
    hWnd = FindWindowEx(0&, 0&, "Shell_TrayWnd", vbNullString)
    If hWnd <> 0 Then
      FindShellTaskBar = hWnd
    End If
End Function


Function PrinterJobsCount(PrinterName As String)
On Error GoTo ErrInt
    Dim hPrinter As Long, l_FirstJob As Long, l_EnumJob As Long, l_Level As Long
    Dim l_JobsNeed As Long, l_JobsCount As Long, byteJobsBuffer() As Byte
    Dim lngResult As Long
    
    lngResult = OpenPrinter(PrinterName, hPrinter, ByVal vbNullString)
    l_FirstJob = 0
    l_EnumJob = 99
    l_Level = 1
    
    lngResult = EnumJobs(hPrinter, l_FirstJob, l_EnumJob, l_Level, ByVal vbNullString, 0, l_JobsNeed, l_JobsCount)
    
    If l_JobsNeed > 0 Then
        ReDim byteJobsBuffer(l_JobsNeed - 1)
        lngResult = EnumJobs(hPrinter, l_FirstJob, l_EnumJob, l_Level, byteJobsBuffer(0), l_JobsNeed, l_JobsNeed, l_JobsCount)
        PrinterJobsCount = l_JobsCount
    End If
    ClosePrinter hPrinter
Exit Function

ErrInt:
    ErrHand Err, "Module System | PrinterJobsCount"
End Function


Function PrinterGetStatus(RetStatus As Long) As String
On Error GoTo ErrInt
    Dim TmpStatusFlag As Long, TmpStatusStr As String, RetStatusStr As String
    For s = 1 To 8
        TmpStatusFlag = Choose(s, JOB_STATUS_DELETING, JOB_STATUS_ERROR, JOB_STATUS_OFFLINE, _
                                        JOB_STATUS_PAPEROUT, JOB_STATUS_PAUSED, JOB_STATUS_PRINTED, _
                                        JOB_STATUS_PRINTING, JOB_STATUS_SPOOLING)
        TmpStatusStr = Choose(s, "Deleting", "Error", "Offline", "Out of paper", "Paused", "Printed", "Printing", "Spooling")
        If RetStatus And TmpStatusFlag Then
            If Trim$(RetStatusStr) <> "" Then
                RetStatusStr = RetStatusStr & " - " & TmpStatusStr
            Else
                RetStatusStr = TmpStatusStr
            End If
        End If
    Next s
    PrinterGetStatus = RetStatusStr
Exit Function

ErrInt:
    ErrHand Err, "Module System | PrinterGetStatus"
End Function


Public Function ConvMem2Str(MemLong As Long) As String
    Dim byteBuffer(64) As Byte, lRet As Long
    lRet = lstrcpy(byteBuffer(0), ByVal MemLong)
    ConvMem2Str = StrConv(byteBuffer(), vbUnicode)
    ConvMem2Str = Left$(ConvMem2Str, InStr(ConvMem2Str, vbNullChar) - 1)
End Function

Public Function RemoveNull(StringWithNull As String) As String
    RemoveNull = Left$(StringWithNull, InStr(StringWithNull, vbNullChar) - 1)
End Function
