Attribute VB_Name = "mdlCommonDialog"
'KPD-Team 1999
'URL: http://users.turboline.be/btl10148/
'E-Mail: KPD_Team@Hotmail.com

Public Const LF_FACESIZE = 32
Public Const MAX_PATH = 260

'ShowOpen/ShowSave flags:
Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_EXPLORER = &H80000
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_LONGNAMES = &H200000
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_NOLONGNAMES = &H40000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_READONLY = &H1
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHAREWARN = 0
Public Const OFN_SHOWHELP = &H10
Public Const OFS_MAXPATHNAME = 128

'BrowseForFolder flags:
Public Const BIF_RETURNONLYFSDIRS = &H1       ' For finding a folder to start document searching
Public Const BIF_DONTGOBELOWDOMAIN = &H2      ' For starting the Find Computer
Public Const BIF_STATUSTEXT = &H4
Public Const BIF_RETURNFSANCESTORS = &H8
Public Const BIF_BROWSEFORCOMPUTER = &H1000   ' Browsing for Computers.
Public Const BIF_BROWSEFORPRINTER = &H2000    ' Browsing for Printers
Public Const BIF_BROWSEINCLUDEFILES = &H4000  ' Browsing for Everything

'Error constants
Public Const CDERR_DIALOGFAILURE = &HFFFF
Public Const CDERR_FINDRESFAILURE = &H6
Public Const CDERR_GENERALCODES = &H0
Public Const CDERR_INITIALIZATION = &H2
Public Const CDERR_LOADRESFAILURE = &H7
Public Const CDERR_LOADSTRFAILURE = &H5
Public Const CDERR_LOCKRESFAILURE = &H8
Public Const CDERR_MEMALLOCFAILURE = &H9
Public Const CDERR_MEMLOCKFAILURE = &HA
Public Const CDERR_NOHINSTANCE = &H4
Public Const CDERR_NOHOOK = &HB
Public Const CDERR_REGISTERMSGFAIL = &HC
Public Const CDERR_NOTEMPLATE = &H3
Public Const CDERR_STRUCTSIZE = &H1

'ShowHelp Enum
Enum enumHelpState
     SW_HIDE = 0
     SW_NORMAL = 1
     SW_MAXIMIZE = 3
     SW_MINIMIZE = 6
     SW_SHOWDEFAULT = 10
End Enum

Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Type BROWSEINFO
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BROWSEINFO) As Long
Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Declare Function GetFileTitleAPI Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer
Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Global OFName As OPENFILENAME
Global BInfo As BROWSEINFO
Dim CustomColors() As Byte

'Use the vbNullChar character to seperate extensions in the filter
'   e.g.  sFilter = "Text Files (*.txt)" + vbNullChar + "*.txt" + "All Files (*.*)" + vbNullChar + "*.*"
'Use the OR operator for multiple flags
'   e.g. nFlags = OFN_EXPLORER Or OFN_FILEMUSTEXIST
Public Function ShowOpen(hWndOwner As Long, sFilter As String, sTitle As String, Optional nFlags As Long = OFN_EXPLORER) As String
    OFName.lStructSize = Len(OFName)
    OFName.hWndOwner = hWndOwner
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = sFilter
    OFName.lpstrFile = String(254, vbNullChar)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = String(254, vbNullChar)
    OFName.nMaxFileTitle = 255
    OFName.lpstrTitle = sTitle
    OFName.flags = nFlags

    If GetOpenFileName(OFName) Then
        ShowOpen = StripTerminator(OFName.lpstrFile)
    Else
        ShowOpen = ""
    End If
End Function
'Use the vbNullChar character to seperate extensions in the filter
'   e.g.  sFilter = "Text Files (*.txt)" + vbNullChar + "*.txt" + "All Files (*.*)" + vbNullChar + "*.*"
'Use the OR operator for multiple flags
'   e.g. nFlags = OFN_EXPLORER Or OFN_FILEMUSTEXIST
Public Function ShowSave(hWndOwner As Long, sFilter As String, sTitle As String, Optional nFlags As Long = OFN_EXPLORER) As String
    OFName.lStructSize = Len(OFName)
    OFName.hWndOwner = hWndOwner
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = sFilter
    OFName.lpstrFile = String(254, vbNullChar)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = String(254, vbNullChar)
    OFName.nMaxFileTitle = 255
    OFName.lpstrTitle = sTitle
    OFName.flags = nFlags

    If GetSaveFileName(OFName) Then
        ShowSave = StripTerminator(OFName.lpstrFile)
    Else
        ShowSave = ""
    End If
End Function


'Use the OR operator for multiple flags
'   e.g. nFlags = BIF_BROWSEFORCOMPUTER Or BIF_BROWSEFORPRINTER
Public Function BrowseForFolder(hWndOwner As Long, sTitle As String) As String
    Dim iNull As Integer, lpIDList As Long, lResult As Long

    With BInfo
        .hWndOwner = hWndOwner
        .lpszTitle = lstrcat(sTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    lpIDList = SHBrowseForFolder(BInfo)
    If lpIDList Then
        BrowseForFolder = String$(MAX_PATH, 0)
        SHGetPathFromIDList lpIDList, BrowseForFolder
        CoTaskMemFree lpIDList
        BrowseForFolder = StripTerminator(BrowseForFolder)
    End If
End Function
Public Function GetFileTitle(sFile As String) As String
    GetFileTitle = String(255, vbNullChar)
    GetFileTitleAPI sFile, GetFileTitle, 255
    GetFileTitle = StripTerminator(GetFileTitle)
End Function
Public Function GetErrorString() As String
    Select Case CommDlgExtendedError
        Case CDERR_DIALOGFAILURE
            GetErrorString = "The dialog box could not be created."
        Case CDERR_FINDRESFAILURE
            GetErrorString = "The common dialog box function failed to find a specified resource."
        Case CDERR_INITIALIZATION
            GetErrorString = "The common dialog box function failed during initialization."
        Case CDERR_LOADRESFAILURE
            GetErrorString = "The common dialog box function failed to load a specified resource."
        Case CDERR_LOADSTRFAILURE
            GetErrorString = "The common dialog box function failed to load a specified string."
        Case CDERR_LOCKRESFAILURE
            GetErrorString = "The common dialog box function failed to lock a specified resource."
        Case CDERR_MEMALLOCFAILURE
            GetErrorString = "The common dialog box function was unable to allocate memory for internal structures."
        Case CDERR_MEMLOCKFAILURE
            GetErrorString = "The common dialog box function was unable to lock the memory associated with a handle."
        Case CDERR_NOHINSTANCE
            GetErrorString = "The ENABLETEMPLATE flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but you failed to provide a corresponding instance handle."
        Case CDERR_NOHOOK
            GetErrorString = "The ENABLEHOOK flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but you failed to provide a pointer to a corresponding hook procedure."
        Case CDERR_REGISTERMSGFAIL
            GetErrorString = "The RegisterWindowMessage function returned an error code when it was called by the common dialog box function."
        Case CDERR_NOTEMPLATE
            GetErrorString = "The ENABLETEMPLATE flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but you failed to provide a corresponding template."
        Case CDERR_STRUCTSIZE
            GetErrorString = "The lStructSize member of the initialization structure for the corresponding common dialog box is invalid."
        Case Else
            GetErrorString = "Undefined error ..."
    End Select
End Function

Private Function StripTerminator(sInput As String) As String
    Dim ZeroPos As Integer
    ZeroPos = InStr(1, sInput, vbNullChar)
    If ZeroPos > 0 Then
        StripTerminator = Left$(sInput, ZeroPos - 1)
    Else
        StripTerminator = sInput
    End If
End Function
