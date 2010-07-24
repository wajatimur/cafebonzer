Attribute VB_Name = "mdlMonitoring"
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Monitor Printer] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub MonPrinter()
    Dim Prn As Printer, lRet As Long, sCmd As String
    Dim byteJobsBuffer() As Byte, byteBuffer(64) As Byte, ut_JobInfo() As JOB_INFO_1
    Dim hPrinter As Long, l_FirstJob As Long, l_EnumJob As Long, l_Level As Long
    Dim l_JobsNeed As Long, l_JobsCount As Long
    
    
    For Each Prn In Printers
        l_JobsCount = PrinterJobsCount(Prn.DeviceName)
        If l_JobsCount > 0 Then
            l_FirstJob = 0
            l_EnumJob = 99
            l_Level = 1
            
            lRet = OpenPrinter(Prn.DeviceName, hPrinter, ByVal vbNullString)
            lRet = EnumJobs(hPrinter, l_FirstJob, l_EnumJob, l_Level, ByVal vbNullString, 0, l_JobsNeed, l_JobsCount)
            
            ReDim byteJobsBuffer(l_JobsNeed - 1)
            ReDim ut_JobInfo(l_JobsNeed - 1)
            lRet = EnumJobs(hPrinter, l_FirstJob, l_EnumJob, l_Level, byteJobsBuffer(0), l_JobsNeed, l_JobsNeed, l_JobsCount)
            
            If l_JobsCount > 0 Then
                MoveMemory ut_JobInfo(0), byteJobsBuffer(0), Len(ut_JobInfo(0)) * l_JobsCount
                For c = 0 To l_JobsCount - 1
                    With ut_JobInfo(c)
                        sCmd = SubBuild("jobid", CStr(.JobId))
                        sCmd = sCmd & SubBuild("printername", ConvMem2Str(.pPrinterName))
                        sCmd = sCmd & SubBuild("machinename", ConvMem2Str(.pMachineName))
                        sCmd = sCmd & SubBuild("username", ConvMem2Str(.pUserName))
                        sCmd = sCmd & SubBuild("document", ConvMem2Str(.pDocument))
                        sCmd = sCmd & SubBuild("datatype", ConvMem2Str(.pDatatype))
                        sCmd = sCmd & SubBuild("status", PrinterGetStatus(.Status))
                        sCmd = sCmd & SubBuild("priority", CStr(.Priority))
                        sCmd = sCmd & SubBuild("position", CStr(.Position))
                        sCmd = sCmd & SubBuild("totalpages", CStr(.TotalPages))
                        sCmd = sCmd & SubBuild("pagesprinted", CStr(.PagesPrinted))
                        NetSend Index, "/info.printerjob:" & sCmd
                    End With
                Next c
            End If
            lRet = ClosePrinter(hPrinter)
        End If
    Next
End Sub

Public Sub MonResource()
    Dim tMemStat As MEMORYSTATUS, sCmd As String
    
    Call GlobalMemoryStatus(tMemStat)

    With tMemStat
        sCmd = SubBuild("memload", CStr(.dwMemoryLoad))
        sCmd = sCmd + SubBuild("memtotal", CStr(.dwTotalPhys)) + SubBuild("memavail", CStr(.dwAvailPhys))
        sCmd = sCmd + SubBuild("virtotal", CStr(.dwTotalVirtual)) + SubBuild("viravail", CStr(.dwAvailVirtual))
        sCmd = sCmd + SubBuild("pagetotal", CStr(.dwTotalPageFile)) + SubBuild("pageavail", CStr(.dwAvailPageFile))
    End With
    NetSend Index, "/info.resource:" & sCmd
End Sub

Public Sub MonApp()
    Dim hwnd As Long, sCmd As String, hProcess As PROCESSENTRY32
    Dim hProcessFnd As Long, hSnapShot As Long
    Dim sTitle As String, sClass As String, lCnt As Long
    
    hwnd = GetWindow(FindShellTaskBar, GW_HWNDFIRST)
    Do While hwnd
        DoEvents
        If IsTaskWindow(hwnd) = True Then
            If IsWindowVisible(hwnd) <> 0 Then
                sTitle = GetTitle(hwnd)
                sClass = GetClass(hwnd)
                If sClass <> "Shell_TrayWnd" Then
                    lCnt = lCnt + 1
                    sCmd = sCmd & SubBuild(lCnt & "hwnd", CStr(hwnd)) & SubBuild(lCnt & "name", sTitle)
                End If
            End If
        End If
        hwnd = GetWindow(hwnd, GW_HWNDNEXT)
    Loop
    sCmd = sCmd & SubBuild("windowtotal", CStr(lCnt))
    
    lCnt = 0
    hProcess.dwSize = Len(hProcess)
    hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    hProcessFnd = ProcessFirst(hSnapShot, hProcess)
    Do While hProcessFnd
        DoEvents
        lCnt = lCnt + 1
        sCmd = sCmd & SubBuild(lCnt & "pid", CStr(hProcess.th32ProcessID))
        sCmd = sCmd & SubBuild(lCnt & "exename", RemoveNull(hProcess.szexeFile))
        hProcessFnd = ProcessNext(hSnapShot, hProcess)
    Loop
    sCmd = sCmd & SubBuild("processtotal", CStr(lCnt))
    NetSend Index, "/info.app:" & sCmd
End Sub
