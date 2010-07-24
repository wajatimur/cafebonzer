VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsDesktop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessId As Long
   dwThreadId As Long
End Type

Private Type STARTUPINFOW
   cbSize As Long
   lpReserved As Long
   lpDesktop As Long
   lpTitle As Long
   dwX As Long
   dwY As Long
   dwXSize As Long
   dwYSize As Long
   dwXCountChars As Long
   dwYCountChars As Long
   dwFillAttribute As Long
   dwFlags As Long
   wShowWindow As Integer
   cbReserved2 As Integer
   lpReserved2 As Long
   hStdInput As Long
   hStdOutput As Long
   hStdError As Long
End Type

Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessW" ( _
      ByVal lpApplicationName As Long, _
      ByVal lpCommandLine As Long, _
      lpProcessAttributes As Any, _
      lpThreadAttributes As Any, _
      ByVal bInheritHandles As Long, _
      ByVal dwCreationFlags As Long, _
      lpEnvironment As Any, _
      ByVal lpCurrentDirectory As Long, _
      lpStartupInfo As STARTUPINFOW, _
      lpProcessInformation As PROCESS_INFORMATION _
   ) As Long
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetThreadDesktop Lib "user32" (ByVal dwThread As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function OpenInputDesktop Lib "user32" ( _
      ByVal dwFlags As Long, _
      ByVal fInherit As Boolean, _
      ByVal dwDesiredAccess As Long _
   ) As Long
Private Declare Function CreateDesktop Lib "user32" Alias "CreateDesktopW" ( _
      ByVal lpszDesktop As Long, _
      ByVal lpszDevice As Long, _
      pDevmode As Any, _
      ByVal dwFlags As Long, _
      ByVal dwDesiredAccess As Long, _
      lpsa As Any _
   ) As Long
Private Declare Function OpenDesktop Lib "user32" Alias "CreateDesktopW" ( _
    ByVal lpszDesktop As Long, _
    ByVal dwFlags As Long, _
    ByVal fInherit As Boolean, _
    ByVal dwDesiredAccess As Long) As Long
    
Private Declare Function SetThreadDesktop Lib "user32" (ByVal hDesktop As Long) As Long
Private Declare Function SwitchDesktop Lib "user32" (ByVal hDesktop As Long) As Long
Private Declare Function CloseDesktop Lib "user32" (ByVal hDesktop As Long) As Long
Private Const GENERIC_ALL = &H10000000
Private Const DESKTOP_SWITCHDESKTOP = &H100&
Private Const STILL_ACTIVE = &H103
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Const INFINITE As Long = &HFFFFFFFF       '  Infinite timeout

Private StrDesktopName As String
Private HDesktopNew As Long
Private HDesktopOldThread As Long
Private HDesktopOldInput As Long
Private LngProcessId As Long


Public Sub Create(ByVal sDesktopName As String)
    HDesktopOldThread = GetThreadDesktop(GetCurrentThreadId())
    HDesktopOldInput = OpenInputDesktop(0, False, DESKTOP_SWITCHDESKTOP)
    HDesktopNew = CreateDesktop(StrPtr(sDesktopName), ByVal 0&, ByVal 0&, 0, GENERIC_ALL, ByVal 0&)
    StrDesktopName = sDesktopName
End Sub


Public Sub Switch(Optional DefaultDesktop As Boolean = False)
    Dim lR As Long, LngHdesk As Long
    
    If DefaultDesktop = False Then
        If Not (HDesktopNew = 0) Then
            lR = SetThreadDesktop(HDesktopNew)
            lR = SwitchDesktop(HDesktopNew)
        End If
    Else
        LngHdesk = OpenDesktop(StrPtr("Default"), 0, False, GENERIC_ALL)
        SetThreadDesktop LngHdesk
        SwitchDesktop LngHdesk
    End If
End Sub


Public Sub StartProcess(ByVal sPath As String)
    Dim tSi As STARTUPINFOW
    Dim tPi As PROCESS_INFORMATION

    ' Must set the desktop to run on in the
    ' STARTUPINFO structure:
    tSi.cbSize = Len(tSi)
    tSi.lpDesktop = StrPtr(StrDesktopName)
    
    LngProcessId = CreateProcess(StrPtr(sPath), ByVal 0&, ByVal 0&, ByVal 0&, True, 0, ByVal 0&, ByVal 0&, tSi, tPi)
    LngProcessId = tPi.hProcess
    
    If (LngProcessId = 0) Then
        ClearUp
    Else
        ' Done. Not sure if we need to close these handles, but it doesn't cause a problem
        CloseHandle tPi.hProcess
        CloseHandle tPi.hThread
    End If
End Sub


Public Sub ClearUp()
    If Not (HDesktopOldInput = 0) Then
        SwitchDesktop HDesktopOldInput
        HDesktopOldInput = 0
    End If
    If Not (HDesktopOldThread = 0) Then
        SetThreadDesktop HDesktopOldThread
        HDesktopOldThread = 0
    End If
    If Not (HDesktopNew = 0) Then
        CloseDesktop HDesktopNew
        HDesktopNew = 0
    End If
End Sub


Private Sub Class_Terminate()
   ClearUp
End Sub
