Attribute VB_Name = "mdlIO"
Private fcolProtect As New Collection
Private s_CurProtectDir As String

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Cek kewujudan File
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function FileExist(ByVal PathName As String) As Boolean
    FileExist = IIf(Dir$(PathName) = "", False, True)
End Function

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [ClearRbin] - Empty Recycle Bin
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub ClearRbin()
    SHEmptyRecycleBin FrmHost.hwnd, vbNullString, SHERB_NOCONFIRMATION + SHERB_NOSOUND
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Clear History] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub ClearHistory()
    Dim uRl As New UrlHistory
    uRl.ClearHistory
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Clear Recent Docs] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub ClearRecentDocs()
    SHAddToRecentDocs 2, vbNullString
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Delete Tree] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub DelTree(PathStr)
    Dim PathRet As String
    
    If Right(PathStr, 1) <> "\" Then PathStr = PathStr & "\"
    PathRet = Dir(PathStr, vbDirectory)
    
    Do Until PathRet = ""
        DoEvents
        If GetAttr(PathStr & PathRet) = vbDirectory Then
            If PathRet <> "." And PathRet <> ".." Then
                DelTree PathStr & PathRet
                RmDir PathStr & PathRet
                PathRet = Dir(PathStr, vbDirectory)
            End If
        End If
        PathRet = Dir
    Loop
    FileWipe PathStr
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Wipe File] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub FileWipe(PathStr)
    Dim shF As SHFILEOPSTRUCT, PathRet As String
    
    If Right(PathStr, 1) <> "\" Then PathStr = PathStr & "\"
    shF.hwnd = FrmHost.hwnd
    shF.wFunc = FO_DELETE
    shF.fFlags = FOF_SILENT + FOF_NOCONFIRMATION + FOF_NOERRORUI
    PathRet = Dir(PathStr, vbNormal + vbArchive + vbHidden + vbSystem)
    
    Do Until PathRet = ""
        DoEvents
        shF.pFrom = PathStr & PathRet & Chr$(0) & Chr$(0)
        SHFileOperation shF
        PathRet = Dir
    Loop
End Sub


Public Function DlgFileOpen(sTitle As String, sInitialDir As String, hwndOwner As Long, sFilter As String, Optional lFlags As Long = 0, Optional bTitleOnly As Boolean = False)
    Dim tOfn As OPENFILENAME
    
    tOfn.lStructSize = Len(tOfn)
    tOfn.hInstance = App.hInstance
    tOfn.hwndOwner = hwndOwner
    tOfn.flags = lFlags
    tOfn.lpstrTitle = sTitle
    tOfn.lpstrInitialDir = sInitialDir
    tOfn.lpstrFilter = sFilter
    
    tOfn.lpstrFile = Space$(256)
    tOfn.nMaxFile = 256
    
    tOfn.lpstrFileTitle = Space$(256)
    tOfn.nMaxFileTitle = 256
    
    If GetOpenFileName(tOfn) Then
        If bTitleOnly = True Then
            DlgFileOpen = Trim$(tOfn.lpstrFileTitle)
            
        Else
            DlgFileOpen = Trim$(tOfn.lpstrFile)
        End If
       'remove null terminated, feel ok now haha
        DlgFileOpen = Left(DlgFileOpen, Len(DlgFileOpen) - 1)
    End If
End Function


'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Folder Guard Section] - File & Folder Security
'
'
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub DirDisable(sPath As String)
On Error Resume Next
    Dim RetPath As String, lFileNum As Long
    
    If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
    If s_CurProtectDir = sPath Then Exit Sub
    s_CurProtectDir = sPath
    RetPath = Dir(sPath, vbNormal + vbArchive + vbHidden + vbSystem)
    
    Do Until RetPath = ""
        DoEvents
        lFileNum = FreeFile
        fcolProtect.Add lFileNum
        Open sPath & RetPath For Random Lock Write As #lFileNum
        RetPath = Dir
    Loop
End Sub

Public Sub DirEnable()
    Dim fFile As Variant
    s_CurProtectDir = ""
    For Each fFile In fcolProtect
        Close #fFile
    Next
End Sub


