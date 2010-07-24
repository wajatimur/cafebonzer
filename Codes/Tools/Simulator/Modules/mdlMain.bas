Attribute VB_Name = "mdlTicker"
Option Explicit

'Public variables
Public hDesktopWnd As Long
Public hTaskBarWnd As Long        ' Taskbar window handle
Public hTrayWnd As Long           ' Tray window handle
Public hClockWnd As Long          ' Clock window handle
Public LastEdge As SHAppBar_Edges ' Last checked edge where the taskbar was
Public LastWidth As Long          ' Last checked Tray width
Public LastHeight As Long         ' Last checked Tray height

'Private variables
Dim TickerWidth As Long, TickerHeight As Long
Public IconCount As Long
Public IsHidden As Boolean

'Private const
Private Const MB01 = "Please move the taskbar to the top or bottom edge of screen, or chage its size to one row."
Private Const TT01 = "Please resize your taskbar"

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' IconsAdd: Add the icons to the system tray
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub IconsAdd()
    Dim NID As NOTIFYICONDATAA, i As Long
    With NID
        .cbSize = LenB(NID)
        .hwnd = FrmTicker.hwnd
        .uCallbackMessage = WM_MOUSEMOVE
        .uFlags = NIF_MESSAGE
    End With
    
    For i = 1 To IconCount
        NID.uID = i
        Shell_NotifyIcon NIM_ADD, NID
    Next
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' FindTaskBar: Finds taskbar, tray and clock windows
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub FindTaskBar()
    ' Find TaskBar handle
    hTaskBarWnd = FindWindow("Shell_TrayWnd", vbNullString)
    
    If hTaskBarWnd Then
        ' Find Tray handle (anak kepada taskbar)
        hTrayWnd = FindWindowEx(hTaskBarWnd, 0, "TrayNotifyWnd", vbNullString)
        If hTrayWnd Then
            ' Find Clock handle (anak kepada tray)
            hClockWnd = FindWindowEx(hTrayWnd, 0, "TrayClockWClass", vbNullString)
            If hClockWnd = 0 Then Err.Raise vbObjectError + 2
        Else
            Err.Raise vbObjectError + 1
        End If
    Else
        Err.Raise vbObjectError
    End If
    
    ' Find Desktop handle
    hDesktopWnd = GetDesktopWindow
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Menyembunyikan Ticker
' note : menyembunyikan ticker dengan move, dan visible prop. tidak boleh
'        digunakan kerana ia akan mmberhentikan vbModal loop. <-- ?? apa aku kata nih(not valid)
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub TickerHide(Optional AddIcon As Boolean = True)
    Dim NID As NOTIFYICONDATAA
    ' Check if already hidden
    If IsHidden Then Exit Sub
    IsHidden = True
    ' Remove the icons
    IconsRemove
    ' Move ticker outside
    SetWindowPos FrmTicker.hwnd, 0, -2000, -2000, 0, 0, SWP_NOSIZE Or SWP_NOZORDER
    
    If AddIcon = True Then
        ' Add a "standard" icon
        With NID
            .cbSize = LenB(NID)
            .hwnd = FrmTray.hwnd
            .uID = -100
            .hIcon = FrmTicker.Icon.Handle
            .szTip = TT01
            .uCallbackMessage = WM_MOUSEMOVE
            .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        End With
        Shell_NotifyIcon NIM_ADD, NID
    End If
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' TickerFly: Makes the ticker hover and ontop
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub TickerFly()
    ' Position of TaskBar...(untuk menentukan kedudukan frmticker)
    ' 1 = top
    ' 2 = bottom
    ' 3 = left
    ' 4 = right
    
    Dim DeskRect As RECT, Edge As SHAppBar_Edges, TaskPos As Long
    Dim tHeight As Long, tWidth As Long, tTop As Long, tRight As Long, Stle As Long
    
    IsHidden = False
    SendMessage hDesktopWnd, WM_SETREDRAW, 0, ByVal 0&
    
    'The Edge
    Edge = GetTaskBarEdge()
    Select Case Edge
    Case ABE_TOP: TaskPos = 1
    Case ABE_BOTTOM: TaskPos = 2
    Case ABE_LEFT: TaskPos = 3
    Case ABE_RIGHT: TaskPos = 4
    End Select
    
    'Ticker size other metric size
    GetClientRect GetDesktopWindow, DeskRect
    
    tHeight = GetSystemMetrics(SM_CYCAPTION)
    tWidth = tHeight * IconCount
    tTop = DeskRect.Bottom - (tHeight * 4)
    tRight = DeskRect.Right - (tWidth + 2)
    
    If TaskPos = 3 Or TaskPos = 1 Then tTop = DeskRect.Bottom - tHeight
    If TaskPos = 4 Then tTop = DeskRect.Bottom - tHeight: tRight = 0
    
    'Setting the parents and moving..
    SetParent FrmTicker.hwnd, hDesktopWnd
    SetWindowPos FrmTicker.hwnd, HWND_TOPMOST, tRight, tTop, tWidth, tHeight, 0&
    
    'Redraw !
    SendMessage hDesktopWnd, WM_SETREDRAW, 1, ByVal 0&
    RedrawWindow hDesktopWnd, ByVal 0&, 0&, RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_UPDATENOW
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' TickerShow: Moves the form to the correct position
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub TickerShow()
    Dim ClkR As RECT, TryR As RECT
    Dim NID As NOTIFYICONDATAA

    'Remove the "standard" icon
    With NID
        .cbSize = LenB(NID)
        .hwnd = FrmTicker.hwnd
        .uID = -100
    End With
    
    Shell_NotifyIcon NIM_DELETE, NID
    IsHidden = False

    'jika clock visible
    If IsWindowVisible(hClockWnd) Then
        'Get clock rect
        GetWindowRect hClockWnd, ClkR
        'Calculate clock width
        ClkR.Right = ClkR.Right - ClkR.Left
    End If

    ' If the taskbar has autohide enabled and it's hidden the icons are added when the taskbar
    ' is shown. So we can't use the tray size.
    ' Get tray client rect
    GetClientRect hTrayWnd, TryR
    
    ' Makesure the parent is tray
    SetParent FrmTicker.hwnd, hTrayWnd
    
    If IsAutohide() Then
        ' Move the ticker.
        SetWindowPos FrmTicker.hwnd, 0, TryR.Right - ClkR.Right, 1, 0, 0, SWP_NOSIZE Or SWP_NOZORDER
    Else
        ' Move the ticker.
        SetWindowPos FrmTicker.hwnd, 0, TryR.Right - ClkR.Right - TickerWidth + 1, 1, 0, 0, SWP_NOSIZE Or SWP_NOZORDER
    End If
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' IsAutohide: returns if the taskbar has autohide enabled
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Function IsAutohide() As Boolean
    Dim ABD As AppBarData
    With ABD
        .cbSize = LenB(ABD)
        .hwnd = hTaskBarWnd
    End With
    IsAutohide = SHAppBarMessage(ABM_GETSTATE, ABD) And ABS_AUTOHIDE
End Function

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' TickerResize: Resize the form to the correct size.
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub TickerResize()
    Dim Before As RECT, After As RECT, Stle As Long
    
    ' Stop the updating of taskbar
    SendMessage hTaskBarWnd, WM_SETREDRAW, 0, ByVal 0&
    
    ' Remove and add the icons so we are sure that the ticker will not be over
    ' other icons
    IconsRemove
    IconsAdd

    ' The size of icons is the same as title bar system menu icon
    TickerHeight = GetSystemMetrics(SM_CYCAPTION) - 2
    TickerWidth = TickerHeight * IconCount
    
    ' Change ticker with and height
    SetWindowPos FrmTicker.hwnd, 0, 0, 0, TickerWidth, TickerHeight, SWP_NOMOVE Or SWP_NOZORDER
    
    ' Redraw task bar
    SendMessage hTaskBarWnd, WM_SETREDRAW, 1, ByVal 0&
    RedrawWindow hTaskBarWnd, ByVal 0&, 0&, RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_UPDATENOW
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' IconsRemove: removes the icons from the system tray
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub IconsRemove()
    Dim NID As NOTIFYICONDATAA, i As Long
    With NID
        .cbSize = LenB(NID)
        .hwnd = FrmTicker.hwnd
    End With
    For i = 1 To IconCount
        NID.uID = i
        Shell_NotifyIcon NIM_DELETE, NID
    Next
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' TrayStart: alternate tray start
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub TrayStart()
    Dim NID As NOTIFYICONDATAA
    
    'the ticker is hidden (variable is global)
    IsHidden = True
    
    With NID
        .cbSize = LenB(NID)
        .hwnd = FrmTray.hwnd
        .uID = 1
        .hIcon = FrmTray.Icon.Handle
        .uCallbackMessage = WM_MOUSEMOVE
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    End With
    
    Shell_NotifyIcon NIM_ADD, NID
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [TrayRemove] - Buang TrayIcon
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub TrayRemove()
    Dim NID As NOTIFYICONDATAA
    
    'ticker now is visible, global juga
    IsHidden = False
    
    With NID
        .cbSize = LenB(NID)
        .uID = 1
        .hwnd = FrmTray.hwnd
    End With
    
    Shell_NotifyIcon NIM_DELETE, NID
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [TickerStart] - Main - where the ticker start
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub TickerStart()
    Dim Edge As SHAppBar_Edges, sTmpRet As String
    Dim sTickFontName As String, sTickFontSize As String
    ' Create a new CTaskbar object
    FindTaskBar
    
    ' Set how many icons we will add to system tray.
    sTmpRet = SetGet("ticker.size")
    If sTmpRet = "" Then
        If IconCount = 0 Then IconCount = 7
    Else
        IconCount = sTmpRet
    End If

    ' Initialize public variables
    LastEdge = GetTaskBarEdge()
    GetTraySize LastWidth, LastHeight
    ' Load ticker form
    Load FrmTicker
    
    ' Load ticker setting
    FrmTicker.picTicker.FontName = SetGet("ticker.fontface", "Verdana")
    FrmTicker.picTicker.FontSize = SetGet("ticker.fontsize", 8)
    FrmTicker.picTicker.ForeColor = SetGet("ticker.fontcolor", vbBlack)
    FrmTicker.picTicker.BackColor = SetGet("ticker.backcolor", &HE0E0E0)
    
    ' Change parent of ticker form to system tray
    SetParent FrmTicker.hwnd, hTrayWnd
    
    ' Show the ticker only if the taskbar is at the top or bottom edge of the screen, and has only
    ' one row
    Edge = GetTaskBarEdge()
    If Edge <> ABE_LEFT And Edge <> ABE_RIGHT And TrayIconRows() = 1 Then
        TickerResize
        TickerShow
    Else
        TickerHide False
        TickerFly
    End If
    
    FrmTicker.Show
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [TickerStop] - Time to stop and getout of the here
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub TickerStop()
    ' Ticker Termination Process Restore ticker parent to desktop
    SetParent FrmTicker.hwnd, 0&
    ' Stop painting the task bar
    SendMessage hTaskBarWnd, WM_SETREDRAW, 0, ByVal 0&
    ' Remove the icons
    IconsRemove
    ' Start painting in task bar and force a repaint
    SendMessage hTaskBarWnd, WM_SETREDRAW, 1, ByVal 0&
    RedrawWindow hTaskBarWnd, ByVal 0&, 0&, RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_UPDATENOW
 
    Unload FrmTicker
End Sub


'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [TrayIconRows] - return how many rows has the tray
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Function TrayIconRows() As Long
    Dim CR As RECT, ClkR As RECT
    Dim IcnSize As Long

    GetClientRect hTrayWnd, CR
    
    ' If the clock is visible get its size
    If IsWindowVisible(hClockWnd) Then
        ' Get clock rect
        GetWindowRect hClockWnd, ClkR
        ' Map clock rect to tray coordinates
        MapWindowPoints 0&, hTrayWnd, ClkR, 2
        ' Ignore Clock size if it isn't at the top
        If ClkR.Top <> 0 Then
            ClkR.Top = 0
            ClkR.Bottom = 0
        End If
    End If
    
    ' Get the icon height.
    IcnSize = GetSystemMetrics(SM_CYCAPTION) - 3
    ' Calculate rows
    TrayIconRows = (CR.Bottom - (ClkR.Bottom - ClkR.Top)) \ IcnSize
End Function

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [GetTraySize] - returns the tray size
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub GetTraySize(ByRef Width As Long, ByRef Height As Long)
Dim CR As RECT
    GetClientRect hTrayWnd, CR

    Width = CR.Right
    Height = CR.Bottom
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [GetTaskBarEdge] - return the edge where the taskbar is
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Function GetTaskBarEdge() As SHAppBar_Edges
    Dim ABD As AppBarData
    With ABD
        .cbSize = LenB(ABD)
        .hwnd = hTaskBarWnd
    End With
    SHAppBarMessage ABM_GETTASKBARPOS, ABD
    
    GetTaskBarEdge = ABD.uEdge
End Function
