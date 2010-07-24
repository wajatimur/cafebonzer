Attribute VB_Name = "mIoControl"
'==================================================================
' Aplication codename : CafeBonzer
' Programmer          : Azri Jamil a.k.a wajatimur
' Module Name         : Fail
' Description         :
'==================================================================
Public Enum EnuModule
    [Help] = -1
    [CafeReport] = 0
    [CafeSnmMgr] = 1
End Enum

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public EnumArray() As String

Public Sub LoadModule(Module As EnuModule)
    Dim lngRet As Long
    
    Select Case Module
    '[ Load Help Module ]'
    Case -1
        If Len(Dir(App.Path & "\help.htm", vbNormal)) = 0 Then
            Call ShellExecute(FrmMain.hwnd, "open", "http://www.nematix.net", vbNullString, vbNullString, SW_NORMAL)
            Exit Sub
        End If
        Call ShellExecute(FrmMain.hwnd, "open", App.Path & "\help.htm", vbNullString, vbNullString, SW_NORMAL)
    '[ Load Report Module ]'
    Case 0
        lngRet = ShellExecute(FrmMain.hwnd, "open", App.Path & "\CafeReport.exe", "pc-usage", vbNullString, SW_NORMAL)
    '[ Load SM Manager Module ]'
    Case 1
        lngRet = ShellExecute(FrmMain.hwnd, "open", App.Path & "\CafeSmMgr.exe", vbNullString, vbNullString, SW_NORMAL)
    End Select
    
    If lngRet <= 32 Then MsgBox MB(20), vbCritical, CbMsgWarn
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Cek kewujudan File
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function FileExist(ByVal PathName As String) As Boolean
    FileExist = IIf(Dir$(PathName) = "", False, True)
End Function

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Simpan error
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub ErrLog(errType As ErrObject, procName As String, Optional DisplayMsg As Boolean = True)
    Dim ErrDesc As String
    Dim i_errNum As Integer, s_errDesc As String, s_errSource As String
        
    i_errNum = errType.Number
    s_errSource = errType.Source
    s_errDesc = errType.Description
    
    MsgBox i_errNum & " / " & s_errSource & vbNewLine & s_errDesc, vbExclamation, procName
    ErrDesc = Now & " - " & s_errDesc & " - " & s_errSource & " - " & i_errNum
    
    Open "ErrLog.txt" For Append As #1
    Write #1, ErrDesc
    Close #1
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Simpan string ke File *.ini
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Sub INIsimpan(NamaFail, Bahagian, Kunci, Nilai, Optional NoEncrypt As Boolean = False)
        NoEncrypt = True
    Bahagian = CStr(Bahagian)
    Kunci = CStr(Kunci)
    If NoEncrypt = False Then Nilai = Crypt(CStr(Nilai), EcKey2)
    NamaFail = CStr(NamaFail)
    
    WritePrivateProfileString Bahagian, CStr(Kunci), CStr(Nilai), NamaFail
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Ambil string dari file *.ini
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function INIambil(NamaFail, Bahagian, Kunci, Optional NoEncrypt As Boolean = False) As String
    Dim retval As String * 255
        NoEncrypt = True
    Bahagian = CStr(Bahagian)
    Kunci = CStr(Kunci)
    NamaFail = CStr(NamaFail)
    
    GetPrivateProfileString Bahagian, CStr(Kunci), "", retval, Len(retval), NamaFail
    ostr = retval & Chr(0)
    INIambil = Left(ostr, InStr(1, ostr, Chr$(0)) - 1)
    If NoEncrypt = False Then INIambil = Crypt(INIambil, EcKey2)
End Function

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Enumerate Section dan Key dalam *.ini file
'  jika enumkey dinyatakan.. enumerator akan
'  mengambil nilai bagi setiap key.. jika tidak
'  ia akan enumerate bahagian dan bukannyer kunci
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function INIenumSection(NamaFail, Optional EnumKey As String = "") As String
    Dim Buff As String
    Dim idx As Integer
    ReDim EnumArray(10)
    
    If Len(Dir(NamaFail)) = 0 Then Open NamaFail For Output As #1: Close #1
    
    Open NamaFail For Input As #1
    
        If EnumKey = "" Then
            Do Until EOF(1)
                Line Input #1, Buff
                If Left(Buff, 1) = "[" Then
                    If idx = UBound(EnumArray) Then ReDim Preserve EnumArray(UBound(EnumArray) + 10)
                    EnumArray(idx) = StrReverse(Mid(Buff, 2))
                    EnumArray(idx) = StrReverse(Mid(EnumArray(idx), 2))
                    idx = idx + 1
                End If
            Loop
        Else
            Do Until EOF(1)
                Line Input #1, Buff
                If LCase(Left(Buff, Len(EnumKey))) = LCase(EnumKey) Then
                    If idx = UBound(EnumArray) Then ReDim Preserve EnumArray(UBound(EnumArray) + 10)
                    'EnumArray(idx) = Crypt(Mid(Buff, InStr(1, Buff, "=") + 1), EcKey2)
                    EnumArray(idx) = Mid(Buff, InStr(1, Buff, "=") + 1)
                    idx = idx + 1
                End If
            Loop
        End If
        
        Close #1
        ReDim Preserve EnumArray(idx)
End Function
