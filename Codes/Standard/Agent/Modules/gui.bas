Attribute VB_Name = "MdlInterface"
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : MdlInterface
'    Project    : CafeBonzerAG
'
'    Description: [type_description_here]
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Enable Group] - Enable\Disable Group Base On GrpName Tag
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub EnableGroup(GrpName As String, Optional Enable As Boolean = False)
On Error Resume Next
    For Each Control In FrmMain.Controls
        If Control.Tag = GrpName Then Control.Enabled = Enable
    Next
End Sub



'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Enum Font Procedure] - Enum All Font Callback Procedure
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Function EnumFontProc(ByVal lplf As Long, ByVal lptm As Long, ByVal dwType As Long, ByVal lpData As Long) As Long
    Dim lfRet As LOGFONT, s_FntName As String
    
    CopyMemory lfRet, ByVal lplf, LenB(lfRet)
    s_FntName = StrConv(lfRet.lfFaceName, vbUnicode)
    s_FntName = Trim$(s_FntName)
    
    FrmPickFont.FntList.AddItem s_FntName
    EnumFontProc = 1
End Function

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [DrawBorder] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub DrawBorder(hWnd As Long)
    Dim Stle As Long
  
  '{ Buang 'Border' asal }'
    Stle = GetWindowLong(hWnd, GWL_STYLE)
    Stle = Stle And Not WS_BORDER
    SetWindowLong hWnd, GWL_STYLE, Stle
    
  '{ Set 'Style' baru }'
    Stle = GetWindowLong(hWnd, GWL_EXSTYLE)
    Stle = Stle Or WS_EX_STATICEDGE
    SetWindowLong hWnd, GWL_EXSTYLE, Stle

End Sub


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
    GetTitle = Left$(sBuffer, InStr(1, sBuffer, Chr$(0)) - 1)
End Function

Public Function GetClass(hWnd As Long) As String
    Dim sBuffer As String * 64
    GetClassName hWnd, sBuffer, 64
    GetClass = Left$(sBuffer, InStr(1, sBuffer, Chr$(0)) - 1)
End Function


'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Find Taskbar Handle] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Function FindShellTaskBar() As Long
    Dim hWnd As Long
    On Error Resume Next
    hWnd = FindWindowEx(0&, 0&, StrShellTaskBarClass, vbNullString)
    If hWnd <> 0 Then
      FindShellTaskBar = hWnd
    End If
End Function


'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Find Desktop Handle] - Locates the shell_defview window
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Function FindShellWindow() As Long
    Dim hWnd As Long
    On Error Resume Next
    hWnd = FindWindowEx(0&, 0&, StrShellViewClass, vbNullString)
    If hWnd <> 0 Then
      FindShellWindow = hWnd
    End If
End Function


'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Hide Window] - Toggles a window's visibility
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub HideDesktop(Optional Hide As Boolean = False)
    Dim lngShowCmd As Long
    On Error Resume Next
    If Hide = True Then
       lngShowCmd = SW_HIDE
    Else
       lngShowCmd = SW_SHOW
    End If
    Call ShowWindow(FindShellWindow, lngShowCmd)
    Call ShowWindow(FindShellTaskBar, lngShowCmd)
End Sub


'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Hide Window] - Toggles a window's visibility
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
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
    Dim hWndCur As Long
    
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
    Dim DtpDeskRect As RECT
    GetWindowRect GetDesktopWindow, DtpDeskRect
    
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
            .Right = DtpDeskRect.Right
            .Bottom = DtpDeskRect.Bottom
        End With
    End If
    
    erg& = ClipCursor(NewRect)
End Sub


Public Function GetWallPaper() As String
    GetWallPaper = GetString(HKEY_CURRENT_USER, "Control Panel\Desktop", "wallpaper")
End Function

