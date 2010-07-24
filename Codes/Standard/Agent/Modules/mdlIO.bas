Attribute VB_Name = "MdlInputOutput"
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : MdlInputOutput
'    Project    : CafeBonzerAG
'
'    Description: [type_description_here]
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Private ColFolderProtect As New Collection
Private StrFolderProtect As String

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Cek kewujudan File
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Function FileCheckExist(ByVal PathName As String) As Boolean
    FileCheckExist = IIf(Dir$(PathName) = "", False, True)
End Function

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Wipe File] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub FileWipe(PathStr As String)
    Dim DtpSfopstruct As SHFILEOPSTRUCT, PathRet As String
    
    If Right$(PathStr, 1) <> "\" Then PathStr = PathStr & "\"
    DtpSfopstruct.hWnd = FrmHost.hWnd
    DtpSfopstruct.wFunc = FO_DELETE
    DtpSfopstruct.fFlags = FOF_SILENT + FOF_NOCONFIRMATION + FOF_NOERRORUI
    PathRet = Dir$(PathStr, vbNormal + vbArchive + vbHidden + vbSystem)
    
    Do Until PathRet = ""
        DoEvents
        DtpSfopstruct.pFrom = PathStr & PathRet & Chr$(0) & Chr$(0)
        SHFileOperation DtpSfopstruct
        PathRet = Dir
    Loop
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Delete Tree] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub PathDelTree(PathStr As String)
    Dim PathRet As String
    
    If Right$(PathStr, 1) <> "\" Then PathStr = PathStr & "\"
    PathRet = Dir$(PathStr, vbDirectory)
    
    Do Until PathRet = ""
        DoEvents
        If GetAttr(PathStr & PathRet) = vbDirectory Then
            If PathRet <> "." And PathRet <> ".." Then
                PathDelTree PathStr & PathRet
                RmDir PathStr & PathRet
                PathRet = Dir$(PathStr, vbDirectory)
            End If
        End If
        PathRet = Dir
    Loop
    FileWipe PathStr
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Folder Guard Section] - File & Folder Security
'
'
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub PathFolderDisable(sPath As String)
On Error Resume Next
    Dim RetPath As String, lFileNum As Long
    
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    If StrFolderProtect = sPath Then Exit Sub
    StrFolderProtect = sPath
    RetPath = Dir$(sPath, vbNormal + vbArchive + vbHidden + vbSystem)
    
    Do Until RetPath = ""
        DoEvents
        lFileNum = FreeFile
        ColFolderProtect.Add lFileNum
        Open sPath & RetPath For Random Lock Write As #lFileNum
        RetPath = Dir
    Loop
End Sub

Public Sub PathFolderEnable()
    Dim fFile As Variant
    StrFolderProtect = ""
    For Each fFile In ColFolderProtect
        Close #fFile
    Next
End Sub

Public Function DlgFileOpen(sTitle As String, sInitialDir As String, hwndOwner As Long, sFilter As String, Optional lFlags As Long = 0, Optional BlnTitleOnly As Boolean = False) As String
    Dim DtpOpenFile As OPENFILENAME
    
    DtpOpenFile.lStructSize = Len(DtpOpenFile)
    DtpOpenFile.hInstance = App.hInstance
    DtpOpenFile.hwndOwner = hwndOwner
    DtpOpenFile.flags = lFlags
    DtpOpenFile.lpstrTitle = sTitle
    DtpOpenFile.lpstrInitialDir = sInitialDir
    DtpOpenFile.lpstrFilter = sFilter
    
    DtpOpenFile.lpstrFile = Space$(256)
    DtpOpenFile.nMaxFile = 256
    
    DtpOpenFile.lpstrFileTitle = Space$(256)
    DtpOpenFile.nMaxFileTitle = 256
    
    If GetOpenFileName(DtpOpenFile) Then
        If BlnTitleOnly = True Then
            DlgFileOpen = Trim$(DtpOpenFile.lpstrFileTitle)
            
        Else
            DlgFileOpen = Trim$(DtpOpenFile.lpstrFile)
        End If
       '{ Remove null terminated, feel ok now haha. }'
        DlgFileOpen = Left$(DlgFileOpen, Len(DlgFileOpen) - 1)
    End If
End Function
