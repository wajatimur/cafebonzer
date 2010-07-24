VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GCbApi"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'///////////////////////////////////////////////////////////////////////////////////////////////////////
'/ SYSTEM
'///////////////////////////////////////////////////////////////////////////////////////////////////////
'/  Memory Function
       Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
      Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
      Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
      Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
      Declare Function GetProcessHeap Lib "kernel32" () As Long
      Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
      Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
      Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
      Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
      Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (lpString1 As Any, lpString2 As Any) As Long

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
