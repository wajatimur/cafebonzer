VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsPipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const FILE_FLAG_WRITE_THROUGH = &H80000000
Private Const FILE_FLAG_OVERLAPPED = &H40000000
Private Const PIPE_ACCESS_DUPLEX = &H3
Private Const PIPE_TYPE_MESSAGE = &H4
Private Const PIPE_READMODE_MESSAGE = &H2
Private Const PIPE_WAIT = &H0
Private Const PIPE_NOWAIT = &H1
Private Const PIPE_UNLIMITED_INSTANCES = 255
Private Const NMPWAIT_USE_DEFAULT_WAIT = &H0
Private Const INVALID_HANDLE_VALUE = -1
Private Const ERROR_IO_PENDING = 997 '  dderror
Private Const ERROR_PIPE_CONNECTED = 535&

Private Const SECURITY_DESCRIPTOR_MIN_LENGTH = (20)
Private Const SECURITY_DESCRIPTOR_REVISION = (1)
Private Const GMEM_FIXED = &H0
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
   
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Declare Function WaitNamedPipe Lib "kernel32" Alias "WaitNamedPipeA" (ByVal lpNamedPipeName As String, ByVal nTimeOut As Long) As Long
Private Declare Function CreateNamedPipe Lib "kernel32" Alias "CreateNamedPipeA" (ByVal lpName As String, ByVal dwOpenMode As Long, ByVal dwPipeMode As Long, ByVal nMaxInstances As Long, ByVal nOutBufferSize As Long, ByVal nInBufferSize As Long, ByVal nDefaultTimeOut As Long, lpSecurityAttributes As Any) As Long
Private Declare Function ConnectNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long, lpOverlapped As Any) As Long
Private Declare Function CallNamedPipe Lib "kernel32" Alias "CallNamedPipeA" (ByVal lpNamedPipeName As String, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesRead As Long, ByVal nTimeOut As Long) As Long
Private Declare Function DisconnectNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function InitializeSecurityDescriptor Lib "advapi32.dll" (ByVal pSecurityDescriptor As Long, ByVal dwRevision As Long) As Long
Private Declare Function SetSecurityDescriptorDacl Lib "advapi32.dll" (ByVal pSecurityDescriptor As Long, ByVal bDaclPresent As Long, ByVal pDacl As Long, ByVal bDaclDefaulted As Long) As Long

Private Const StrShutDownMsg = "shutdown pipe "
Private Const StrPipeFormat = "\\.\pipe\"
Private Const LngBufSize = 20000
    
Public Event ReceiveData(Data As String)

Private DtpSa As SECURITY_ATTRIBUTES
Private BytBuffer(LngBufSize) As Byte
Private LngSecDes As Long
Private LngServerHandle As Long
Private StrServerName As String
Private StrHostName As String
Private BlnTerminate As Boolean
Private LngTimerId As Long

Public Sub ServerCreatePipe(PipeName As String, StartServer As Boolean)
    BlnTerminate = False
    StrServerName = StrPipeFormat & PipeName
    LngServerHandle = INVALID_HANDLE_VALUE
    If StartServer Then Call ServerStart
End Sub

Public Sub ServerStart()
    Dim LngPipeMode As Long, LngOpenMode As Long, LngCount As Long
    Dim LngBSize As Long, LngBytRead As Long
    Dim LngIdxA As Long, LngRet As Long
    
    For LngIdxA = 0 To LngBufSize - 1
        BytBuffer(LngIdxA) = LngIdxA Mod 256
    Next
    
    LngSecDes = GlobalAlloc(GPTR, SECURITY_DESCRIPTOR_MIN_LENGTH)
    LngRet = InitializeSecurityDescriptor(LngSecDes, SECURITY_DESCRIPTOR_REVISION)
    LngRet = SetSecurityDescriptorDacl(LngSecDes, -1, 0, 0)
    DtpSa.nLength = LenB(DtpSa)
    DtpSa.lpSecurityDescriptor = LngSecDes
    DtpSa.bInheritHandle = True
    
    LngOpenMode = PIPE_ACCESS_DUPLEX Or FILE_FLAG_WRITE_THROUGH
    LngPipeMode = PIPE_TYPE_MESSAGE Or PIPE_READMODE_MESSAGE Or PIPE_WAIT
    LngServerHandle = CreateNamedPipe(StrServerName, LngOpenMode, LngPipeMode, 10, 10000, 2000, 10000, DtpSa)
    
    Do
        LngRet = ConnectNamedPipe(LngServerHandle, ByVal 0)
        LngBytRead = 4
        
        LngRet = ReadFile(LngServerHandle, LngCount, LenB(LngCount), LngBytRead, ByVal 0)
        If LngCount <> 0 Then
            If LngCount > LngBufSize Then LngCount = LngBufSize
            LngRet = WriteFile(LngServerHandle, BytBuffer(0), LngCount, LngBytRead, ByVal 0)
            LngRet = FlushFileBuffers(LngServerHandle)
        End If
        LngRet = DisconnectNamedPipe(LngServerHandle)
    Loop Until LngCount = 0
    
    CloseHandle LngServerHandle
End Sub

Public Sub ServerRead()
    Dim DtpInMsg As UdtPipeMessage, DtpOutMsg As UdtPipeMessage
    Dim LngRet As Long, LngEventRet As Long, DtpOvl As OVERLAPPED
       
    If LngPipeServerHandle = INVALID_HANDLE_VALUE Then
        Sleep 250
    Else
        LngRet = ConnectNamedPipe(LngPipeServerHandle, DtpOvl)
        If LngRet Then
            DtpInMsg.Size = Len(DtpInMsg)
            ReadFile LngPipeServerHandle, DtpInMsg, DtpInMsg.Size, DtpInMsg.Size, ByVal 0&
            If DtpInMsg.Kind = 0 And DtpInMsg.Data = (StrShutDownMsg & StrPipeServerName) Then
                DtpOutMsg.Kind = 0
                DtpOutMsg.Count = 3
                DtpOutMsg.Data = "OK" & vbNullString
            ElseIf Not (DtpInMsg.Data = "") Then
                DtpOutMsg = DtpInMsg
                RaiseEvent ReceiveData(DtpInMsg.Data)
                CalcSize DtpOutMsg
                WriteFile LngPipeServerHandle, DtpOutMsg, DtpOutMsg.Size, 0, ByVal 0&
            End If
        End If
        DisconnectNamedPipe LngPipeServerHandle
    End If
End Sub

Public Sub ServerStop()
    Dim DtpInMsg As UdtPipeMessage, DtpOutMsg As UdtPipeMessage
    Dim StrShutMsg As String
    
    If LngPipeServerHandle <> INVALID_HANDLE_VALUE Then
        DtpOutMsg.Size = Len(DtpOutMsg)
        StrShutMsg = StrShutDownMsg & StrPipeServerName
        With DtpInMsg
            .Kind = 0
            .Count = Len(StrShutMsg) + 1
            .Data = StrShutMsg
        End With
        CalcSize DtpInMsg
        CallNamedPipe StrPipeServerName, DtpInMsg, DtpInMsg.Size, DtpOutMsg, DtpOutMsg.Size, 0, 100
        CloseHandle LngPipeServerHandle
        LngPipeServerHandle = INVALID_HANDLE_VALUE
        BlnTerminate = True
    End If
End Sub

Public Sub ClientSend(PipeHost As String, Data As String)
    '
End Sub

Private Sub Class_Terminate()
    KillTimer LngTimerId, 0&
    Call ServerStop
End Sub
