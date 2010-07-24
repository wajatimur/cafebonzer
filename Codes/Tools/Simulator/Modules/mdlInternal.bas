Attribute VB_Name = "mdlSystem"


'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Change Priority] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub SetPriority(Optional pNormal As Boolean = False)
    Dim pId As Long
    Dim hProcess As Long
    
    pId = GetCurrentProcessId
    hProcess = OpenProcess(PROCESS_DUP_HANDLE, True, pId)
    
    If pNormal = False Then
        SetPriorityClass hProcess, REALTIME_PRIORITY_CLASS
    Else
        SetPriorityClass hProcess, IDLE_PRIORITY_CLASS
    End If
    Call CloseHandle(hProcess)
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Disabling CTL+ALT+DEL] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub DisableCtlAltDel(Opt As Boolean)
    Dim lRet As Long
    lRet = SystemParametersInfo(SPI_SCREENSAVERRUNNING, Opt, vbNull, 0)
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Get Machine Name] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Function MyName()
    Dim StrNama As String
    StrNama = String(255, Chr(0))
    GetComputerName StrNama, 255
    StrNama = Left(StrNama, InStr(1, StrNama, Chr(0)) - 1)
    MyName = StrNama
End Function

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Set Machine Name] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub MyNameSet(NetName As String)
    If NetName = "" Then Exit Sub
    SetComputerName NetName
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Get MAC Address] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Function GetMACAddress() As String
On Error GoTo ErrInt
   Dim s_Mac As String, lASTAT As Long
   Dim NCB As NET_CONTROL_BLOCK
   Dim AST As ASTAT

  'The IBM NetBIOS 3.0 specifications defines four basic
  'NetBIOS environments under the NCBRESET command. Win32
  'follows the OS/2 Dynamic Link Routine (DLR) environment.
  'This means that the first NCB issued by an application
  'must be a NCBRESET, with the exception of NCBENUM.
  'The Windows NT implementation differs from the IBM
  'NetBIOS 3.0 specifications in the NCB_CALLNAME field.
   NCB.ncb_command = NCBRESET
   Call Netbios(NCB)
   
  'To get the Media Access Control (MAC) address for an
  'ethernet adapter programmatically, use the Netbios()
  'NCBASTAT command and provide a "*" as the name in the
  'NCB.ncb_CallName field (in a 16-chr string).
   NCB.ncb_callname = "*               "
   NCB.ncb_command = NCBASTAT
   
  'For machines with multiple network adapters you need to
  'enumerate the LANA numbers and perform the NCBASTAT
  'command on each. Even when you have a single network
  'adapter, it is a good idea to enumerate valid LANA numbers
  'first and perform the NCBASTAT on one of the valid LANA
  'numbers. It is considered bad programming to hardcode the
  'LANA number to 0 (see the comments section below).
   NCB.ncb_lana_num = 0
   NCB.ncb_length = Len(AST)
   
   lASTAT = HeapAlloc(GetProcessHeap(), HEAP_GENERATE_EXCEPTIONS Or HEAP_ZERO_MEMORY, NCB.ncb_length)
            
   If lASTAT = 0 Then
      GoTo ErrInt
      Exit Function
   End If
   
   NCB.ncb_buffer = lASTAT
   Call Netbios(NCB)
   CopyMemory AST, NCB.ncb_buffer, Len(AST)
    
   s_Mac = Format$(Hex(AST.adapt.adapter_address(0)), "00") & _
         Format$(Hex(AST.adapt.adapter_address(1)), "00") & _
         Format$(Hex(AST.adapt.adapter_address(2)), "00") & _
         Format$(Hex(AST.adapt.adapter_address(3)), "00") & _
         Format$(Hex(AST.adapt.adapter_address(4)), "00") & _
         Format$(Hex(AST.adapt.adapter_address(5)), "00")

   HeapFree GetProcessHeap(), 0, lASTAT
   GetMACAddress = s_Mac
Exit Function

ErrInt:
    ErrHand Err, "GetMACAddress"
End Function

Function PrinterJobsNeed(PrinterName As String)
    Dim hPrinter As Long, l_FirstJob As Long, l_EnumJob As Long, l_Level As Long
    Dim l_JobsNeed As Long, l_JobsCount As Long
    Dim lngResult As Long
    
    lngResult = OpenPrinter(PrinterName, hPrinter, ByVal vbNullString)
    l_FirstJob = 0
    l_EnumJob = 99
    l_Level = 1
    
    lngResult = EnumJobs(hPrinter, l_FirstJob, l_EnumJob, l_Level, ByVal vbNullString, 0, l_JobsNeed, l_JobsCount)
    PrinterJobsNeed = l_JobsNeed
End Function


Function PrinterJobsCount(PrinterName As String)
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
End Function


Function PrinterGetStatus(RetStatus As Long) As String
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
End Function


Function GetPrinters() As String
    Dim sRet As String, l_Cnt As Long, Prn As Printer
    ' Format
    '  {total|'1'}{1name|'Bubblejet 250'}{1default|'True'}{1port|'lpt1'}{1drivername|'bbjet'}... and so on
    
    sRet = SubBuild("total", Printers.Count)
    For Each Prn In Printers
        l_Cnt = l_Cnt + 1
        
        sRet = sRet & SubBuild(l_Cnt & "name", Prn.DeviceName)
        sRet = sRet & SubBuild(l_Cnt & "default", Prn.TrackDefault)
        sRet = sRet & SubBuild(l_Cnt & "port", Prn.Port)
        sRet = sRet & SubBuild(l_Cnt & "drivername", Prn.DriverName)
        sRet = sRet & SubBuild(l_Cnt & "papersize", Prn.PaperSize)
        sRet = sRet & SubBuild(l_Cnt & "orientation", Prn.Orientation)
    Next
    GetPrinters = sRet
End Function
