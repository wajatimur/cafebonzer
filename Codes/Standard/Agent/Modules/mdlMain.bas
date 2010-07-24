Attribute VB_Name = "MdlTicker"
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : MdlTicker
'    Project    : CafeBonzerAG
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Option Explicit
Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Const OPAQUE = 2
Private Const TRANSPARENT = 1
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Const COLOR_BTNFACE = 15
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Enum EDrawTextFormat
   DT_BOTTOM = &H8
   DT_CALCRECT = &H400
   DT_CENTER = &H1
   DT_EXPANDTABS = &H40
   DT_EXTERNALLEADING = &H200
   DT_INTERNAL = &H1000
   DT_LEFT = &H0
   DT_NOCLIP = &H100
   DT_NOPREFIX = &H800
   DT_RIGHT = &H2
   DT_SINGLELINE = &H20
   DT_TABSTOP = &H80
   DT_TOP = &H0
   DT_VCENTER = &H4
   DT_WORDBREAK = &H10
   DT_EDITCONTROL = &H2000&
   DT_PATH_ELLIPSIS = &H4000&
   DT_END_ELLIPSIS = &H8000&
   DT_MODIFYSTRING = &H10000
   DT_RTLREADING = &H20000
   DT_WORD_ELLIPSIS = &H40000
End Enum


'Public variables
Public StrTickerText As String
Public LngIconCount As Long

'Private variables
Private HwndDesktop As Long
Private HwndTaskBar As Long             ' Taskbar window handle
Private HwndTray As Long                ' Tray window handle
Private HwndClock As Long               ' Clock window handle
Private TickerWidth As Long
Private TickerHeight As Long
Private LngLastWidth As Long            ' Last checked Tray width
Private LngLastHeight As Long           ' Last checked Tray height

Private CMemDc As New ClsMemDC
Private CStdFont As IFont
Private DtpAreaRect As RECT
Private DtpStrRect As RECT
Private DtpLastEdge As SHAppBar_Edges


'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [TickerStart] - Main - where the ticker start
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub TickerStart(TickerForm As Form)
    Dim DtpEdge As SHAppBar_Edges, CFont As New StdFont
    
   '[ Load Configuration ]'
    If StrTickerText = "" Then StrTickerText = StrTickMsgWelcome
    LngIconCount = SettingGet("TickGuiSize", 5)
    FrmTicker.BackColor = SettingGet("TickGuiBackColor", &H8000000F)
    FrmTicker.ForeColor = SettingGet("TickGuiForeColor", &H80000012)
    CFont.name = SettingGet("TickGuiFont", "Verdana")
    CFont.Size = SettingGet("TickGuiFontSize", 7)
    Set CStdFont = CFont
    
   '[ General Metrics ]'
    Call FindTaskBar
    Call GetTraySize(LngLastWidth, LngLastHeight)
    TickerHeight = GetSystemMetrics(SM_CYCAPTION) - 2
    TickerWidth = TickerHeight * LngIconCount
    CMemDc.Height = TickerHeight
    CMemDc.Width = TickerWidth
    
   '[ Check Edge and Draw Ticker ]'
    DtpLastEdge = GetTaskBarEdge
    DtpEdge = GetTaskBarEdge
    If DtpEdge <> ABE_LEFT And DtpEdge <> ABE_RIGHT And GetTrayIconRow = 1 Then
        Call TickerNormal
    Else
        Call TickerHover
    End If
    
   '[ Text Drawing Metrics ]'
    GetClientRect FrmTicker.hWnd, DtpAreaRect
    DtpStrRect = DtpAreaRect
    DtpStrRect.Right = DtpStrRect.Left + (FrmTicker.TextWidth(StrTickerText) \ Screen.TwipsPerPixelX)
    
   '[ Activate Ticker ]'
    FrmTicker.Show
    FrmTicker.TmrTicker.Enabled = True
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [TickerStop] - Time to stop and getout of the here
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub TickerStop()
    SetParent FrmTicker.hWnd, 0&
    SendMessage HwndTaskBar, WM_SETREDRAW, 0, ByVal 0&
    TickerIconRemove
    SendMessage HwndTaskBar, WM_SETREDRAW, 1, ByVal 0&
    RedrawWindow HwndTaskBar, ByVal 0&, 0&, RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_UPDATENOW
 
    Unload FrmTicker
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [TickerNormal] - Moves the form to the correct position
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub TickerNormal()
    Dim DtpClockRect As RECT, DtpTrayRect As RECT
    
    SendMessage HwndTaskBar, WM_SETREDRAW, 0, ByVal 0&
    Call TickerIconRemove
    Call TickerIconAdd
    SendMessage HwndTaskBar, WM_SETREDRAW, 1, ByVal 0&
    RedrawWindow HwndTaskBar, ByVal 0&, 0&, RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_UPDATENOW

    SetParent FrmTicker.hWnd, HwndTray
    If IsAutohide() Then
        SetWindowPos FrmTicker.hWnd, 0, DtpTrayRect.Left + 1, DtpTrayRect.Top + 1, TickerWidth, TickerHeight, SWP_NOZORDER
    Else
        SetWindowPos FrmTicker.hWnd, 0, 1, 1, TickerWidth, TickerHeight, SWP_NOZORDER
    End If
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [TickerHover] - Makes the ticker hover and ontop
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub TickerHover()
    Dim DtpDeskRect As RECT, DtpEdge As SHAppBar_Edges
    Dim LngTop As Long, LngLeft As Long
    
    SendMessage HwndDesktop, WM_SETREDRAW, 0, ByVal 0&
    Call TickerIconRemove
    GetClientRect GetDesktopWindow, DtpDeskRect
    
    DtpEdge = GetTaskBarEdge()
    If DtpEdge = ABE_LEFT Then
        LngTop = DtpDeskRect.Bottom - TickerHeight
        LngLeft = DtpDeskRect.Right - TickerWidth
    ElseIf DtpEdge = ABE_RIGHT Then
        LngTop = DtpDeskRect.Bottom - TickerHeight
        LngLeft = 0
    End If

    SetParent FrmTicker.hWnd, HwndDesktop
    SetWindowPos FrmTicker.hWnd, HWND_TOPMOST, LngLeft, LngTop, TickerWidth, TickerHeight, 0&

    SendMessage HwndDesktop, WM_SETREDRAW, 1, ByVal 0&
    RedrawWindow HwndDesktop, ByVal 0&, 0&, RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_UPDATENOW
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [TickerDrawText] - Draw Text
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub TickerDrawText()
    Dim LngHbrush As Long, LngStrWidth As Long
    Dim hDcDraw As Long, hDcOutPut As Long
    
    LngStrWidth = DtpStrRect.Right - DtpStrRect.Left
    If DtpStrRect.Right < 0 Then
        DtpStrRect.Left = DtpAreaRect.Right
        DtpStrRect.Right = DtpStrRect.Left + LngStrWidth
    Else
        DtpStrRect.Left = DtpStrRect.Left - 1
        DtpStrRect.Right = DtpStrRect.Left + LngStrWidth
    End If
    
    hDcOutPut = FrmTicker.hdc
    hDcDraw = CMemDc.hdc
    SelectObject hDcDraw, CStdFont.hFont
    SetTextColor hDcDraw, FrmTicker.ForeColor
    SetBkMode hDcDraw, TRANSPARENT
    
    LngHbrush = CreateSolidBrush(GetBkColor(FrmTicker.hdc))
    FillRect hDcDraw, DtpAreaRect, LngHbrush
    DeleteObject LngHbrush
    
    DrawText hDcDraw, StrTickerText, -1, DtpStrRect, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
    CMemDc.Draw hDcOutPut, , , , , DtpAreaRect.Left, DtpAreaRect.Top
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [TickerCheck] - Check Taskbar Position
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub TickerCheck()
    Dim LngWidth As Long, LngHeight As Long, DtpCurEdge As SHAppBar_Edges

    DtpCurEdge = GetTaskBarEdge()
    Select Case DtpCurEdge
    Case ABE_BOTTOM, ABE_TOP
        If DtpLastEdge <> DtpCurEdge Then
            Call TickerNormal
        Else
        '[ Check Taskbar Size ]'
            GetTraySize LngWidth, LngHeight
            If GetTrayIconRow() = 1 Then
                If LngHeight < LngLastHeight Or LngWidth < LngLastWidth Then
                    Call TickerNormal
                End If
            Else
                If LngHeight < LngLastHeight Or LngWidth < LngLastWidth Then
                    Call TickerHover
                End If
            End If
        End If
    Case ABE_LEFT, ABE_RIGHT
        If DtpLastEdge <> DtpCurEdge Then
            Call TickerHover
        End If
    End Select
    
    DtpLastEdge = DtpCurEdge
    LngLastHeight = LngHeight
    LngLastWidth = LngWidth
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [TickerIconAdd] - Add the icons to the system tray
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub TickerIconAdd()
    Dim DtpNID As NOTIFYICONDATAA, LngIdxA As Long
    With DtpNID
        .cbSize = LenB(DtpNID)
        .hWnd = FrmTicker.hWnd
        .uCallbackMessage = WM_MOUSEMOVE
        .uFlags = NIF_MESSAGE
    End With
    
    For LngIdxA = 1 To LngIconCount
        DtpNID.uID = LngIdxA
        Shell_NotifyIcon NIM_ADD, DtpNID
    Next
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [TickerIconRemove] - Removes the icons from the system tray
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub TickerIconRemove()
    Dim DtpNID As NOTIFYICONDATAA, i As Long
    With DtpNID
        .cbSize = LenB(DtpNID)
        .hWnd = FrmTicker.hWnd
    End With
    For i = 1 To LngIconCount
        DtpNID.uID = i
        Shell_NotifyIcon NIM_DELETE, DtpNID
    Next
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [TrayStart] - alternate tray start
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub TrayStart()
    Dim DtpNID As NOTIFYICONDATAA
    
    With DtpNID
        .cbSize = LenB(DtpNID)
        .hWnd = FrmTray.hWnd
        .uID = 1
        .hIcon = FrmHost.IconStatOff.Picture.Handle
        .uCallbackMessage = WM_MOUSEMOVE
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    End With
    
    Shell_NotifyIcon NIM_ADD, DtpNID
    LngStatusTray = 1
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [TrayEdit] - Buang TrayIcon
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub TrayEdit(IconHandle As Long)
    Dim DtpNID As NOTIFYICONDATAA
    
    With DtpNID
        .cbSize = LenB(DtpNID)
        .hWnd = FrmTray.hWnd
        .uID = 1
        .hIcon = IconHandle
        .uCallbackMessage = WM_MOUSEMOVE
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    End With
    
    Shell_NotifyIcon NIM_MODIFY, DtpNID
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [TrayRemove] - Buang TrayIcon
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub TrayRemove()
    Dim DtpNID As NOTIFYICONDATAA
    
    With DtpNID
        .cbSize = LenB(DtpNID)
        .uID = 1
        .hWnd = FrmTray.hWnd
    End With
    
    Shell_NotifyIcon NIM_DELETE, DtpNID
    LngStatusTray = 0
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [GetTrayIconRow] - return how many rows has the tray
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Function GetTrayIconRow() As Long
    Dim DtpClientRect As RECT, DtpClockRect As RECT
    Dim LngIconHeight As Long

    GetClientRect HwndTray, DtpClientRect
    
    ' If the clock is visible get its size
    If IsWindowVisible(HwndClock) Then
        GetWindowRect HwndClock, DtpClockRect
        ' Map clock rect to tray coordinates
        MapWindowPoints 0&, HwndTray, DtpClockRect, 2
        ' Ignore Clock size if it isn't at the top
        If DtpClockRect.Top <> 0 Then
            DtpClockRect.Top = 0
            DtpClockRect.Bottom = 0
        End If
    End If
    
    LngIconHeight = GetSystemMetrics(SM_CYCAPTION) - 3
    GetTrayIconRow = (DtpClientRect.Bottom - (DtpClockRect.Bottom - DtpClockRect.Top)) \ LngIconHeight
End Function

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [GetTraySize] - returns the tray size
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub GetTraySize(ByRef Width As Long, ByRef Height As Long)
    Dim DtpClientRect As RECT
    GetClientRect HwndTray, DtpClientRect

    Width = DtpClientRect.Right
    Height = DtpClientRect.Bottom
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [GetTaskBarEdge] - return the DtpEdge where the taskbar is
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Function GetTaskBarEdge() As SHAppBar_Edges
On Error GoTo ErrInt
    Dim ABD As AppBarData
    With ABD
        .cbSize = LenB(ABD)
        .hWnd = HwndTaskBar
    End With
    SHAppBarMessage ABM_GETTASKBARPOS, ABD
    
    GetTaskBarEdge = ABD.uEdge
Exit Function

ErrInt:
    AppErrorLog Err, "Module Ticker | GetTaskBarEdge"
End Function

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [FindTaskBar] - Finds taskbar, tray and clock windows
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Private Sub FindTaskBar()
   '[ Find TaskBar handle ]'
    HwndTaskBar = FindWindow("Shell_TrayWnd", vbNullString)
    
    If HwndTaskBar Then
       '[ Find Tray handle ]'
        HwndTray = FindWindowEx(HwndTaskBar, 0, "TrayNotifyWnd", vbNullString)
        If HwndTray Then
           '[ Find Clock handle ]'
            HwndClock = FindWindowEx(HwndTray, 0, "TrayClockWClass", vbNullString)
        End If
    End If
    
    ' Find Desktop handle
    HwndDesktop = GetDesktopWindow
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [IsAutohide] - Returns if the taskbar has autohide enabled
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Private Function IsAutohide() As Boolean
    Dim ABD As AppBarData
    With ABD
        .cbSize = LenB(ABD)
        .hWnd = HwndTaskBar
    End With
    IsAutohide = SHAppBarMessage(ABM_GETSTATE, ABD) And ABS_AUTOHIDE
End Function
