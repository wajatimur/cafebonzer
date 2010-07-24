Attribute VB_Name = "MdlApiDeclare"
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : MdlApiDeclare
'    Project    : CafeBonzer
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Public Declare Function DeleteFile Lib "Kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function GetDateFormat Lib "Kernel32" Alias "GetDateFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpDate As SYSTEMTIME, ByVal lpFormat As String, ByVal lpDateStr As String, ByVal cchDate As Long) As Long


'Retrieves information about the bounding rectangle
'for a subitem in a list view control.
Public Const LVM_FIRST = &H1000
Public Const LVM_GETSUBITEMRECT = (LVM_FIRST + 56)

'Returns the bounding rectangle of the entire item,
'including the icon and label.
Public Const LVIR_LABEL = 2
Public Const CB_SETDROPPEDWIDTH = &H160

'The HDN_ENDTRACK Notifies a header control's parent window
'that the user has finished dragging a divider.
Public Const HDN_FIRST      As Long = (0 - 300)
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const SW_NORMAL = 1

Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const SWP_WNDFLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HTCAPTION = 2

Public Const GWL_STYLE = (-16)
Public Const ES_NUMBER = &H2000&


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

Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type
