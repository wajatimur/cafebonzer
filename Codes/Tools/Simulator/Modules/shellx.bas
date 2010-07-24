Attribute VB_Name = "mdlShellx"
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd As Long, ByVal hWndChild As Long, ByVal lpszClassName As String, ByVal lpszWindow As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const SWP_WNDFLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2


'// Constants for ShowWindow()
Private Const SW_HIDE = 0
Private Const SW_NORMAL = 1
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_SHOWNOACTIVATE = 4
Private Const SW_SHOW = 5
Private Const SW_MINIMIZE = 6
Private Const SW_SHOWMINNOACTIVE = 7
Private Const SW_SHOWNA = 8
Private Const SW_RESTORE = 9
Private Const SW_SHOWDEFAULT = 10

'// Names of the shell windows we'll be looking for(Windows Class)
Private Const g_cstrShellViewWnd As String = "Progman"
Private Const g_cstrShellTaskBarWnd As String = "Shell_TrayWnd"


'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Put Windows On Top] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Sub PutOnTop(hWnd As Long)
    Dim i As Long
    i = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_WNDFLAGS)
End Sub

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
    hWnd = FindWindowEx(0&, 0&, g_cstrShellTaskBarWnd, vbNullString)
    If hWnd <> 0 Then
      FindShellTaskBar = hWnd
    End If
End Function


'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Find Desktop Handle] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
'   Locates the shell_defview window
Public Function FindShellWindow() As Long
    Dim hWnd As Long
    On Error Resume Next
    hWnd = FindWindowEx(0&, 0&, g_cstrShellViewWnd, vbNullString)
    If hWnd <> 0 Then
      FindShellWindow = hWnd
    End If
End Function


'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Hide Window] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
'   Toggles a window's visibility
Public Sub HideShowWindow(ByVal hWnd As Long, Optional ByVal Hide As Boolean = False)
    Dim lngShowCmd As Long
    On Error Resume Next
    If Hide = True Then
       lngShowCmd = SW_HIDE
    Else
       lngShowCmd = SW_SHOW
    End If
    Call ShowWindow(hWnd, lngShowCmd)
End Sub


'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Hide All Windows] - still buggy ! dokleh nok show semula
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Sub HideAllWindow(Hide As Boolean)
    Dim hWndCur As Long, buf As String * 255, buf2 As String
    
    hWndCur = GetWindow(FrmKey.hWnd, GW_HWNDFIRST)

    Do While hWndCur
        If IsTaskWindow(hWndCur) = True And hWndCur <> FrmKey.hWnd Then
                If Hide = True Then
                    ShowWindow hWndCur, SW_HIDE
                Else
                    ShowWindow hWndCur, SW_SHOW Or SW_SHOWNORMAL
                End If
        End If
    
        hWndCur = GetWindow(hWndCur, GW_HWNDNEXT)
    Loop
End Sub


'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Minimize All Windows] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Sub MinAllWindow(Minimize As Boolean)
    hWnd = FindWindow("Shell_TrayWnd", vbNullString)
    If Minimize = True Then
        PostMessage hWnd, WM_COMMAND, MIN_ALL, 0&
    Else
        PostMessage hWnd, WM_COMMAND, MIN_ALL_UNDO, 0&
    End If
End Sub


'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Hide Active Windows] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Sub HideActiveWin(Hide As Boolean)
    hWnd = GetActiveWindow
    If Hide Then
        SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_HIDEWINDOW
    Else
        SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
    End If
End Sub


'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Lock Updates For Active Windows] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Sub LockUpdates(LockUpd As Boolean)
    hWnd = GetActiveWindow
    If LockUpd = True Then
        LockWindowUpdate hWnd
    Else
        LockWindowUpdate vbNull
    End If
End Sub


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


'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Trap Mouse In Form] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub FormTrap(CurForm As Form, Trap As Boolean)
    Dim x As Long, y As Long, erg As Long
    Dim NewRect As RECT
    Dim DeskRect As RECT
    GetWindowRect GetDesktopWindow, DeskRect
    
    x& = Screen.TwipsPerPixelX
    y& = Screen.TwipsPerPixelY
        
    If Trap = True Then
        With NewRect
            .Left = CurForm.Left / x& '- 8
            .Top = CurForm.Top / y& ' - 8
            .Right = .Left + (CurForm.Width / x&) '- 14
            .Bottom = .Top + (CurForm.Height / y&) '- 15
        End With
    Else
        With NewRect
            .Left = 0&
            .Top = 0&
            .Right = DeskRect.Right
            .Bottom = DeskRect.Bottom
        End With
    End If
    
    erg& = ClipCursor(NewRect)
End Sub


Public Function GetWallPaper() As String
    GetWallPaper = GetString(HKEY_CURRENT_USER, "Control Panel\Desktop", "wallpaper")
End Function
