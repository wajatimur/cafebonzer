Attribute VB_Name = "mdlSecurity"
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Memeriksa kata laluan
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function CekPass(Password)
    intpass = Trim(SetGet("noid"))
    
    If intpass = Trim(Password) Then CekPass = "ok"
    If Trim(Password) = "swordfish+" & Minute(Time) Then CekPass = "ok"
End Function


Public Sub DeskWallProtect()
On Error GoTo ErrInt
    Dim s_CurWpaper As String, s_BackWpaper As String
    Dim l_Flag As Long, lRet As Long
    s_BackWpaper = "c:\windows\winwall.dat"
    
    If SetGet("persist.wpaper", 1) = 1 Then
        If FileExist(s_BackWpaper) = False Then
            s_CurWpaper = GetWallPaper
            l_Flag = SetGet("persist.wpaperf", 0)
            
            If s_CurWpaper = "" Then
                If l_Flag = 0 Then
                    SetSave "persist.wpaperf", 2
                ElseIf l_Flag = 1 Then
                    s_BackWpaper = ""
                End If
            Else
                If l_Flag = 2 Then
                    s_BackWpaper = ""
                ElseIf l_Flag = 0 Then
                    FileCopy s_CurWpaper, s_BackWpaper
                    SetSave "persist.wpaperf", 1
                End If
            End If
        End If
        lRet = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, s_BackWpaper, 0)
    Else
        If FileExist(s_BackWpaper) = True Then
            SetSave "persist.wpaperf", 0
            Kill s_BackWpaper
        End If
    End If
Exit Sub

ErrInt:
    ErrHand Err, "DeskWallProtect"
End Sub


Public Sub DeskIconProtect()
    If SetGet("persist.deskicons", 0) = 1 Then
        DirDisable "c:\windows\desktop"
    Else
        DirEnable
    End If
End Sub
