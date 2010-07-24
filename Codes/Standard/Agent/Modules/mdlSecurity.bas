Attribute VB_Name = "MdlSecurity"
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : MdlSecurity
'    Project    : CafeBonzerAG
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Public Function SecCheckPassword(Password As String) As Long
    ' 1 = Granted
    Dim StrPassword As String
    StrPassword = Trim$(SettingGet("GenAdminPass"))
    
    If StrPassword = Trim$(Password) Then SecCheckPassword = 1
End Function


Public Sub SecAdjustToken()
    Const TOKEN_ADJUST_PRIVILEGES = &H20
    Const TOKEN_QUERY = &H8
    Const SE_PRIVILEGE_ENABLED = &H2
    Dim hdlProcessHandle As Long
    Dim hdlTokenHandle As Long
    Dim tmpLuid As LUID
    Dim tkp As TOKEN_PRIVILEGES
    Dim tkpNewButIgnored As TOKEN_PRIVILEGES
    Dim lBufferNeeded As Long
    
    hdlProcessHandle = GetCurrentProcess()
    OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY), hdlTokenHandle
    
    ' Get the LUID for shutdown privilege.
    LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
    
    tkp.PrivilegeCount = 1    ' One privilege to set
    tkp.TheLuid = tmpLuid
    tkp.Attributes = SE_PRIVILEGE_ENABLED
    
    ' Enable the shutdown privilege in the access token of this process.
    AdjustTokenPrivileges hdlTokenHandle, False, tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
End Sub
     

Public Sub DeskWallProtect()
On Error GoTo ErrInt
    Dim s_CurWpaper As String, s_BackWpaper As String
    Dim l_Flag As Long, lRet As Long
    s_BackWpaper = "c:\windows\winwall.dat"
    
    If SettingGet("SysSecWallpaper", 1) = 1 Then
        If FileCheckExist(s_BackWpaper) = False Then
            s_CurWpaper = GetWallPaper
            l_Flag = SettingGet("persist.wpaperf", 0)
            
            If s_CurWpaper = "" Then
                If l_Flag = 0 Then
                    SettingSave "persist.wpaperf", 2
                ElseIf l_Flag = 1 Then
                    s_BackWpaper = ""
                End If
            Else
                If l_Flag = 2 Then
                    s_BackWpaper = ""
                ElseIf l_Flag = 0 Then
                    FileCopy s_CurWpaper, s_BackWpaper
                    SettingSave "persist.wpaperf", 1
                End If
            End If
        End If
        lRet = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, s_BackWpaper, 0)
    Else
        If FileCheckExist(s_BackWpaper) = True Then
            SettingSave "persist.wpaperf", 0
            Kill s_BackWpaper
        End If
    End If
Exit Sub

ErrInt:
    AppErrorLog Err, "DeskWallProtect"
End Sub


Public Sub DeskIconProtect()
    If SettingGet("SysSecDesktop", 0) = 1 Then
        PathFolderDisable "c:\windows\desktop"
    Else
        PathFolderEnable
    End If
End Sub
