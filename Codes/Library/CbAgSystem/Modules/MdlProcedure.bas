Attribute VB_Name = "CasGenProc"
Public Function ProcLowLevelKeyboard1(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Dim BlnKeyStroke As Boolean, UtKeyHook As KBDLLHOOKSTRUCT
   Dim BlnKeyAltTab As Boolean, BlnKeyAltEsc As Boolean, BlnKeyCtlEsc As Boolean, BlnKeyWin As Boolean
   
   BlnKeyAltTab = ((UtKeyHook.vkCode = VK_TAB) And ((UtKeyHook.flags And LLKHF_ALTDOWN) <> 0))
   BlnKeyAltEsc = ((UtKeyHook.vkCode = VK_ESCAPE) And ((UtKeyHook.flags And LLKHF_ALTDOWN) <> 0))
   BlnKeyCtlEsc = ((UtKeyHook.vkCode = VK_ESCAPE) And ((GetKeyState(VK_CONTROL) And &H8000) <> 0))
   BlnKeyWin = (UtKeyHook.vkCode = VK_LWIN) Or (UtKeyHook.vkCode = VK_RWIN)
   
   If (nCode = HC_ACTION) Then
      If wParam = WM_KEYDOWN Or wParam = WM_SYSKEYDOWN Or wParam = WM_KEYUP Or wParam = WM_SYSKEYUP Then
         CopyMemory UtKeyHook, ByVal lParam, Len(UtKeyHook)
         BlnKeyStroke = BlnKeyAltTab Or BlnKeyAltEsc Or BlnKeyCtlEsc Or BlnKeyWin
        End If
    End If
    
    If BlnKeyStroke Then
        LowLevelKeyboardProc = -1
    Else
        LowLevelKeyboardProc = CallNextHookEx(0, nCode, wParam, ByVal lParam)
    End If
End Function

