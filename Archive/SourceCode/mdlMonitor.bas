Attribute VB_Name = "mdlMonitoring"
Private sJobsLast As String
Private sResLast As String
Private sAppLast As String

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Monitor Printer] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub MonPrinter()
On Error GoTo ErrInt
    Dim Prtr As Printer, bJobsBffr() As Byte, bBffr(64) As Byte, uJobInfo() As JOB_INFO_1
    Dim lRet As Long, sCmd As String, sMachine As String
    Dim hPrinter As Long, lJobFirst As Long, lJobEnum As Long, lLevel As Long
    Dim lJobsNeed As Long, lJobsCount As Long
    
    For Each Prtr In Printers
        lJobsCount = PrinterJobsCount(Prtr.DeviceName)
        If lJobsCount > 0 Then
            lLevel = 1: lJobFirst = 0: lJobEnum = 99
            
            lRet = OpenPrinter(Prtr.DeviceName, hPrinter, ByVal vbNullString)
            lRet = EnumJobs(hPrinter, lJobFirst, lJobEnum, lLevel, ByVal vbNullString, 0, lJobsNeed, lJobsCount)
            
            ReDim bJobsBffr(lJobsNeed - 1)
            ReDim uJobInfo(lJobsNeed - 1)
            lRet = EnumJobs(hPrinter, lJobFirst, lJobEnum, lLevel, bJobsBffr(0), lJobsNeed, lJobsNeed, lJobsCount)
            
            If lJobsCount > 0 Then
                CopyMemory uJobInfo(0), bJobsBffr(0), Len(uJobInfo(0)) * lJobsCount
                For c = 0 To lJobsCount - 1
                    With uJobInfo(c)
                        sMachine = Mid(ConvMem2Str(.pMachineName), 3)
                        sCmd = SubBuild("jobid", CStr(.JobId))
                        sCmd = sCmd & SubBuild("machinename", sMachine)
                        sCmd = sCmd & SubBuild("printername", ConvMem2Str(.pPrinterName))
                        sCmd = sCmd & SubBuild("username", ConvMem2Str(.pUserName))
                        sCmd = sCmd & SubBuild("document", ConvMem2Str(.pDocument))
                        sCmd = sCmd & SubBuild("datatype", ConvMem2Str(.pDatatype))
                        sCmd = sCmd & SubBuild("status", PrinterGetStatus(.Status))
                        sCmd = sCmd & SubBuild("priority", CStr(.Priority))
                        sCmd = sCmd & SubBuild("position", CStr(.Position))
                        sCmd = sCmd & SubBuild("totalpages", CStr(.TotalPages))
                        sCmd = sCmd & SubBuild("pagesprinted", CStr(.PagesPrinted))
                        
                        If LCase(MyName) = LCase(sMachine) Then
                            If sCmd <> sJobsLast Then
                                NetSend "/info.printerjob:" & sCmd
                                sJobsLast = sCmd
                            End If
                        End If
                    End With
                Next c
            End If
            lRet = ClosePrinter(hPrinter)
            sCmd = ""
            Erase bJobsBffr
            Erase bBffr
        End If
    Next
Exit Sub

ErrInt:
    ErrHand Err, "Module Monitoring | MonPrinter"
End Sub

Public Sub MonResource()
On Error GoTo ErrInt
    Dim tMemStat As MEMORYSTATUS, sCmd As String
    
    Call GlobalMemoryStatus(tMemStat)

    With tMemStat
        sCmd = SubBuild("memload", CStr(.dwMemoryLoad))
        sCmd = sCmd + SubBuild("memtotal", CStr(.dwTotalPhys)) + SubBuild("memavail", CStr(.dwAvailPhys))
        sCmd = sCmd + SubBuild("virtotal", CStr(.dwTotalVirtual)) + SubBuild("viravail", CStr(.dwAvailVirtual))
        sCmd = sCmd + SubBuild("pagetotal", CStr(.dwTotalPageFile)) + SubBuild("pageavail", CStr(.dwAvailPageFile))
    End With
    If sCmd <> sResLast Then
        NetSend "/info.resource:" & sCmd
        sResLast = sCmd
    End If
Exit Sub

ErrInt:
    ErrHand Err, "Module Monitoring | MonResource"
End Sub

Public Sub MonApp()
On Error GoTo ErrInt
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
    If sCmd <> sAppLast Then
        NetSend "/info.app:" & sCmd
        sAppLast = sCmd
    End If
Exit Sub

ErrInt:
    ErrHand Err, "Module Monitor | MonApp"
End Sub
