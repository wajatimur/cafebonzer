Attribute VB_Name = "CasGenAPi"
'///////////////////////////////////////////////////////////////////////////////////////////////////////
'/ SYSTEM
'///////////////////////////////////////////////////////////////////////////////////////////////////////
'/  Memory Function
    Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
    Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
    Public Declare Function GetProcessHeap Lib "kernel32" () As Long
    Public Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
    Public Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
    Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (lpString1 As Any, lpString2 As Any) As Long
'/  Process Function
    Public Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
    Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
    Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
    Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long
    Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
    Public Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
    Public Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
    Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
    Public Declare Function SetPriorityClass& Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long)
'/  Information
    Public Declare Function SetComputerName Lib "kernel32" Alias "SetComputerNameA" (ByVal lpComputerName As String) As Long
    Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
    Public Declare Function SetSystemPowerState Lib "kernel32" (ByVal fSuspend As Long, ByVal fForce As Long) As Long
    Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
'/  Hook
    Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpFn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
    Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
    Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

'///////////////////////////////////////////////////////////////////////////////////////////////////////
'/  WINDOWS
'///////////////////////////////////////////////////////////////////////////////////////////////////////
    Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
    Public Declare Function MapWindowPoints Lib "user32" (ByVal hwndFrom As Long, ByVal hWndTo As Long, lpPoints As Any, ByVal cPoints As Long) As Long
    Public Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal flags As Long) As Boolean
    Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Public Declare Function GetActiveWindow Lib "user32" () As Long
    Public Declare Function GetDesktopWindow Lib "user32" () As Long
    Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
    Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
    Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
    Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
    Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Boolean
    Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal Parenthwnd As Long, ByVal Firsthwnd As Long, ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
'/  Geometric Function
    Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As Any) As Boolean
    Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'/  Shell Function
    Public Declare Sub SHAddToRecentDocs Lib "shell32" (ByVal uFlags As Long, ByVal pv As String)
    Public Declare Function SHEmptyRecycleBin Lib "shell32" Alias "SHEmptyRecycleBinA" (ByVal hWnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
    Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATAA) As Boolean
    Public Declare Function Shell_NotifyIconW Lib "shell32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATAW) As Boolean
    Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
    Public Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
    Public Declare Function BlockInput Lib "user32" (ByVal fBlock As Long) As Long
'/  System Message
    Public Declare Function SHAppBarMessage Lib "shell32" (ByVal dwMessage As SHAppBar_Messages, pData As AppBarData) As Long
    Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any) As Long
    Public Declare Function GetLastError Lib "kernel32" () As Long
    Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long


'///////////////////////////////////////////////////////////////////////////////////////////////////////
'/ GRAPHIC
'///////////////////////////////////////////////////////////////////////////////////////////////////////
    Public Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long


'///////////////////////////////////////////////////////////////////////////////////////////////////////
'/ MISCELANEOUS
'///////////////////////////////////////////////////////////////////////////////////////////////////////
'/  File I/O
    Public Declare Function SHFileOperation Lib "shell32" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
'/  Keyboard
    Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
'/  Common Dialogs
    Public Declare Function EnumFonts Lib "gdi32" Alias "EnumFontsA" (ByVal hdc As Long, ByVal lpsz As String, ByVal lpFontEnumProc As Long, ByVal lParam As Long) As Long
    Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
'/  Printer
    Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As Any) As Long
    Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
    Public Declare Function EnumJobs Lib "winspool.drv" Alias "EnumJobsA" (ByVal hPrinter As Long, ByVal FirstJob As Long, ByVal NoJobs As Long, ByVal Level As Long, pJob As Any, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
'/  Network
    Public Declare Function Netbios Lib "netapi32.dll" (pncb As NET_CONTROL_BLOCK) As Byte
'/  Permission
    Public Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
    Public Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
    Public Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long






Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1
Const SWP_WNDFLAGS = SWP_NOMOVE Or SWP_NOSIZE
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2


'// Constants for ShowWindow()
Const SW_HIDE = 0
Const SW_NORMAL = 1
Const SW_SHOWMINIMIZED = 2
Const SW_SHOWMAXIMIZED = 3
Const SW_SHOWNOACTIVATE = 4
Const SW_SHOW = 5
Const SW_MINIMIZE = 6
Const SW_SHOWMINNOACTIVE = 7
Const SW_SHOWNA = 8
Const SW_RESTORE = 9
Const SW_SHOWDEFAULT = 10
Const NORMAL_PRIORITY_CLASS = &H20
Const IDLE_PRIORITY_CLASS = &H40
Const HIGH_PRIORITY_CLASS = &H80
Const REALTIME_PRIORITY_CLASS = &H100
Const PROCESS_DUP_HANDLE = &H40
Const PROCESS_ALL_ACCESS = 0
Const TH32CS_SNAPPROCESS As Long = 2&
Const MAX_PATH& = 260

Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Const LANG_NEUTRAL = &H0
Const SUBLANG_DEFAULT = &H1

Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4

Const NIM_ADD = &H0
Const NIM_MODIFY = &H1
Const NIM_DELETE = &H2

Const NIF_MESSAGE = &H1
Const NIF_ICON = &H2
Const NIF_TIP = &H4

Const RDW_INVALIDATE = &H1
Const RDW_ALLCHILDREN = &H80
Const RDW_UPDATENOW = &H100

Const SWP_NOZORDER = &H4
Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Const SWP_HIDEWINDOW = &H80
Const SWP_SHOWWINDOW = &H40

Const SHERB_NOCONFIRMATION = &H1
Const SHERB_NOPROGRESSUI = &H2
Const SHERB_NOSOUND = &H4

Const WM_COMMAND = &H111
Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_RBUTTONDBLCLK = &H206
Const WM_RBUTTONDOWN = &H204
Const WM_RBUTTONUP = &H205
Const WM_SETREDRAW = &HB
Const WM_USER As Long = &H400
Const WM_MYHOOK As Long = WM_USER + 1

Const MIN_ALL = 419
Const MIN_ALL_UNDO = 416

Const GWL_STYLE = (-16)
Const GWL_EXSTYLE = (-20)

Const GW_HWNDFIRST = 0
Const GW_HWNDLAST = 1
Const GW_HWNDNEXT = 2
Const GW_HWNDPREV = 3
Const GW_MAX = 5
Const GW_OWNER = 4

Const WS_BORDER = &H800000
Const WS_EX_STATICEDGE = &H20000

Const SPI_SCREENSAVERRUNNING = 97
Const SPI_SETDESKWALLPAPER = 20
Const SPIF_SENDWININICHANGE = &H2
Const SPIF_UPDATEINIFILE = &H1

Const SM_CYCAPTION = 4
Const SM_CXSCREEN = 0
Const SM_CYSCREEN = 1

Const FO_COPY = &H2
Const FO_DELETE = &H3
Const FO_MOVE = &H1
Const FO_RENAME = &H4

Const FOF_ALLOWUNDO = &H40
Const FOF_CONFIRMMOUSE = &H2
Const FOF_FILESONLY = &H80
Const FOF_MULTIDESTFILES = &H1
Const FOF_NO_CONNECTED_ELEMENTS = &H2000
Const FOF_NOCONFIRMATION = &H10
Const FOF_NOCONFIRMMKDIR = &H200
Const FOF_NOCOPYSECURITYATTRIBS = &H800
Const FOF_NOERRORUI = &H400
Const FOF_NORECURSION = &H1000
Const FOF_RENAMEONCOLLISION = &H8
Const FOF_SILENT = &H4
Const FOF_SIMPLEPROGRESS = &H100
Const FOF_WANTMAPPINGHANDLE = &H20
Const FOF_WANTNUKEWARNING = &H4000

Const CF_PRINTERFONTS = &H2
Const CF_SCREENFONTS = &H1
Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Const CF_EFFECTS = &H100&
Const CF_FORCEFONTEXIST = &H10000
Const CF_INITTOLOGFONTSTRUCT = &H40&
Const CF_LIMITSIZE = &H2000&
Const REGULAR_FONTTYPE = &H400

Const FW_NORMAL = 400
Const DEFAULT_CHARSET = 1
Const OUT_DEFAULT_PRECIS = 0
Const CLIP_DEFAULT_PRECIS = 0
Const DEFAULT_QUALITY = 0
Const DEFAULT_PITCH = 0
Const FF_ROMAN = 16
Const LF_FACESIZE = 32

Const GMEM_MOVEABLE = &H2
Const GMEM_ZEROINIT = &H40
Const HEAP_ZERO_MEMORY As Long = &H8
Const HEAP_GENERATE_EXCEPTIONS As Long = &H4

Const CCHDEVICENAME = 32
Const CCHFORMNAME = 32

Const JOB_STATUS_PAUSED = &H1
Const JOB_STATUS_ERROR = &H2
Const JOB_STATUS_DELETING = &H4
Const JOB_STATUS_SPOOLING = &H8
Const JOB_STATUS_PRINTING = &H10
Const JOB_STATUS_OFFLINE = &H20
Const JOB_STATUS_PAPEROUT = &H40
Const JOB_STATUS_PRINTED = &H80
Const JOB_STATUS_DELETED = &H100
Const JOB_STATUS_BLOCKED_DEVQ = &H200
Const JOB_STATUS_USER_INTERVENTION = &H400     ' Windows 95 Only

Const NO_PRIORITY = 0
Const MAX_PRIORITY = 99
Const MIN_PRIORITY = 1
Const DEF_PRIORITY = 1

Const NCBASTAT As Long = &H33
Const NCBNAMSZ As Long = 16
Const NCBRESET As Long = &H32

Const HC_ACTION = 0
Const WM_KEYDOWN = &H100
Const WM_KEYUP = &H101
Const WM_SYSKEYDOWN = &H104
Const WM_SYSKEYUP = &H105
Const VK_TAB = &H9
Const VK_CONTROL = &H11
Const VK_ESCAPE = &H1B
Const VK_LWIN = &H5B
Const VK_RWIN = &H5C

Const WH_KEYBOARD_LL = 13
Const LLKHF_ALTDOWN = &H20

Public Type KBDLLHOOKSTRUCT
    vkCode As Long
    scanCode As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128 ' Maintenance string for PSS usage
End Type

Private Type LUID
   UsedPart As Long
   IgnoredForNowHigh32BitPart As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    TheLuid As LUID
    Attributes As Long
End Type

Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * MAX_PATH
End Type

Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Type NOTIFYICONDATAA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Type NOTIFYICONDATAW
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip(0 To 127) As Byte
End Type

Public Enum SHAppBar_Messages
    ABM_NEW = &H0
    ABM_REMOVE = &H1
    ABM_QUERYPOS = &H2
    ABM_SETPOS = &H3
    ABM_GETSTATE = &H4
    ABM_GETTASKBARPOS = &H5
    ABM_ACTIVATE = &H6
    ABM_GETAUTOHIDEBAR = &H7
    ABM_SETAUTOHIDEBAR = &H8
    ABM_WINDOWPOSCHANGED = &H9
End Enum

Public Enum SHAppBar_Notifications
    ABN_STATECHANGE = &H0
    ABN_POSCHANGED = &H1
    ABN_FULLSCREENAPP = &H2
    ABN_WINDOWARRANGE = &H3
End Enum

Public Enum SHAppBar_States
    ABS_AUTOHIDE = &H1
    ABS_ALWAYSONTOP = &H2
End Enum

Public Enum SHAppBar_Edges
    ABE_LEFT = 0
    ABE_TOP = 1
    ABE_RIGHT = 2
    ABE_BOTTOM = 3
End Enum

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type AppBarData
    cbSize As Long
    hWnd As Long
    uCallbackMessage As Long
    uEdge As SHAppBar_Edges
    rc As RECT
    lParam As Long
End Type

Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type

Public Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName(LF_FACESIZE) As Byte
End Type


Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type


Public Type NET_CONTROL_BLOCK  'NCB
   ncb_command    As Byte
   ncb_retcode    As Byte
   ncb_lsn        As Byte
   ncb_num        As Byte
   ncb_buffer     As Long
   ncb_length     As Integer
   ncb_callname   As String * NCBNAMSZ
   ncb_name       As String * NCBNAMSZ
   ncb_rto        As Byte
   ncb_sto        As Byte
   ncb_post       As Long
   ncb_lana_num   As Byte
   ncb_cmd_cplt   As Byte
   ncb_reserve(9) As Byte ' Reserved, must be 0
   ncb_event      As Long
End Type

Public Type ADAPTER_STATUS
   adapter_address(5) As Byte
   rev_major         As Byte
   reserved0         As Byte
   adapter_type      As Byte
   rev_minor         As Byte
   duration          As Integer
   frmr_recv         As Integer
   frmr_xmit         As Integer
   iframe_recv_err   As Integer
   xmit_aborts       As Integer
   xmit_success      As Long
   recv_success      As Long
   iframe_xmit_err   As Integer
   recv_buff_unavail As Integer
   t1_timeouts       As Integer
   ti_timeouts       As Integer
   Reserved1         As Long
   free_ncbs         As Integer
   max_cfg_ncbs      As Integer
   max_ncbs          As Integer
   xmit_buf_unavail  As Integer
   max_dgram_size    As Integer
   pending_sess      As Integer
   max_cfg_sess      As Integer
   max_sess          As Integer
   max_sess_pkt_size As Integer
   name_count        As Integer
End Type
   
Public Type NAME_BUFFER
   name        As String * NCBNAMSZ
   name_num    As Integer
   name_flags  As Integer
End Type

Public Type ASTAT
   adapt          As ADAPTER_STATUS
   NameBuff(30)   As NAME_BUFFER
End Type
