Option Strict Off
Option Explicit On
Module mApiDeclare
	Public Declare Function SetUnhandledExceptionFilter Lib "kernel32" (ByVal lpTopLevelExceptionFilter As Integer) As Integer
	Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)
	Public Declare Function ShellExecute Lib "shell32.dll"  Alias "ShellExecuteA"(ByVal hwnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer
	
	Public Declare Sub ReleaseCapture Lib "user32" ()
	Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Integer, ByVal hWndNewParent As Integer) As Integer
	Public Declare Function GetWindowLong Lib "user32"  Alias "GetWindowLongA"(ByVal hwnd As Integer, ByVal nIndex As Integer) As Integer
	Public Declare Function SetWindowLong Lib "user32"  Alias "SetWindowLongA"(ByVal hwnd As Integer, ByVal nIndex As Integer, ByVal dwNewLong As Integer) As Integer
	Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Integer, ByVal hWndInsertAfter As Integer, ByVal x As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1016"'
	Public Declare Function SendMessage Lib "user32"  Alias "SendMessageA"(ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As Any) As Integer
	
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1050"'
	Public Declare Function DrawText Lib "user32"  Alias "DrawTextA"(ByVal hdc As Integer, ByVal lpStr As String, ByVal nCount As Integer, ByRef lpRect As RECT, ByVal wFormat As Integer) As Object
	
	Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Integer, ByVal nIDEvent As Integer, ByVal uElapse As Integer, ByVal lpTimerFunc As Integer) As Integer
	Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Integer, ByVal nIDEvent As Integer) As Integer
	
	'UPGRADE_WARNING: Structure SYSTEMTIME may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1050"'
	Public Declare Function GetDateFormat Lib "kernel32"  Alias "GetDateFormatA"(ByVal Locale As Integer, ByVal dwFlags As Integer, ByRef lpDate As SYSTEMTIME, ByVal lpFormat As String, ByVal lpDateStr As String, ByVal cchDate As Integer) As Integer
	
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1016"'
	Public Declare Function GetPrivateProfileString Lib "kernel32"  Alias "GetPrivateProfileStringA"(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1016"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1016"'
	Public Declare Function WritePrivateProfileString Lib "kernel32"  Alias "WritePrivateProfileStringA"(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Integer
	
	'Retrieves information about the bounding rectangle
	'for a subitem in a list view control.
	Public Const LVM_FIRST As Short = &H1000s
	Public Const LVM_GETSUBITEMRECT As Integer = (LVM_FIRST + 56)
	'Returns the bounding rectangle of the entire item,
	'including the icon and label.
	Public Const LVIR_LABEL As Short = 2
	Public Const CB_SETDROPPEDWIDTH As Short = &H160s
	
	'The HDN_ENDTRACK Notifies a header control's parent window
	'that the user has finished dragging a divider.
	Public Const HDN_FIRST As Integer = (0 - 300)
	Public Const HDN_ENDTRACK As Integer = (HDN_FIRST - 1)
	
	Public Const WM_NOTIFY As Integer = &H4Es
	Public Const WM_HSCROLL As Integer = &H114s
	Public Const WM_VSCROLL As Integer = &H115s
	Public Const WM_KEYDOWN As Integer = &H100s
	Public Const WM_NCLBUTTONDOWN As Short = &HA1s
	
	Public Const SW_HIDE As Short = 0
	Public Const SW_NORMAL As Short = 1
	Public Const SW_SHOWMINIMIZED As Short = 2
	Public Const SW_SHOWMAXIMIZED As Short = 3
	Public Const SW_MAXIMIZE As Short = 3
	Public Const SW_SHOWNOACTIVATE As Short = 4
	Public Const SW_SHOW As Short = 5
	Public Const SW_MINIMIZE As Short = 6
	Public Const SW_SHOWMINNOACTIVE As Short = 7
	Public Const SW_SHOWNA As Short = 8
	Public Const SW_RESTORE As Short = 9
	Public Const SW_SHOWDEFAULT As Short = 10
	
	Public Const SWP_NOMOVE As Short = 2
	Public Const SWP_NOSIZE As Short = 1
	Public Const SWP_WNDFLAGS As Boolean = SWP_NOMOVE Or SWP_NOSIZE
	Public Const HWND_TOPMOST As Short = -1
	Public Const HWND_NOTOPMOST As Short = -2
	Public Const HTCAPTION As Short = 2
	
	Public Const GWL_STYLE As Short = (-16)
	Public Const ES_NUMBER As Integer = &H2000
	
	Public Const DT_CENTER As Short = &H1s
	Public Const DC_GRADIENT As Short = &H20s 'Only Win98/2000 !!
	
	Public Structure SYSTEMTIME
		Dim wYear As Short
		Dim wMonth As Short
		Dim wDayOfWeek As Short
		Dim wDay As Short
		Dim wHour As Short
		Dim wMinute As Short
		Dim wSecond As Short
		Dim wMilliseconds As Short
	End Structure
	
	Structure RECT
		'UPGRADE_NOTE: Left was upgraded to Left_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
		Dim Left_Renamed As Integer
		Dim Top As Integer
		'UPGRADE_NOTE: Right was upgraded to Right_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
		Dim Right_Renamed As Integer
		Dim Bottom As Integer
	End Structure
End Module