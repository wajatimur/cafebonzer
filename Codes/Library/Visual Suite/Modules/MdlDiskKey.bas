Attribute VB_Name = "MdlDiskKey"
Private Declare Function GetVolumeInformation Lib "Kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Const StrKeyFile = ":\Boot"
Private Const StrSeperator = "/:"


'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Membina DiskKey
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Function CreateDiskKey(Name As String, Key As String, CbDrvStr As String) As Boolean
On Error GoTo ErrInt
    Dim Buffer As String, LngSerial As Long
    
    LngSerial = GetSerial(CbDrvStr)
    Buffer = Crypt(LngSerial & StrSeperator & Name & StrSeperator & Key)
    
    Open CbDrvStr & StrKeyFile For Output As #1
    Print #1, Buffer
    Close #1
    SetAttr CbDrvStr & StrKeyFile, vbHidden
    CreateDiskKey = True
Exit Function

ErrInt:
    CreateDiskKey = False
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Mengesahkan DiskKey
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Function ValidateDisk(CbDrvStr As String) As Boolean
On Error GoTo ErrInt
    Dim LngSerial As Long, StrSerial As String
    Dim IntPos As Integer
    
    ValidateDisk = False
    If Len(Dir(CbDrvStr & StrKeyFile, vbHidden)) = 0 Then Exit Function
    
    '{ Get disk serial }'
    LngSerial = GetSerial(CbDrvStr)
    
    '{ Get infile disk serial }'
    Open CbDrvStr & StrKeyFile For Input As #1
    Line Input #1, StrSerial
    Close #1
    
    StrSerial = Crypt(StrSerial)
    IntPos = InStr(1, StrSerial, StrSeperator)
    StrSerial = left(StrSerial, IntPos - 1)

    If LngSerial = StrSerial Then ValidateDisk = True
Exit Function

ErrInt:
    ValidateDisk = False
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Mendapatkan Nama Pendaftaran Dari DiskKey
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Function GetName(CbDrvStr) As String
    Dim Buffer As String, TmpName As String
    Dim LngPosA As Long, LngPosB As Long
    
    Open CbDrvStr & StrKeyFile For Input As #1
    Line Input #1, Buffer
    Buffer = Crypt(Buffer)
    LngPosA = InStr(1, Buffer, StrSeperator)
    LngPosB = InStr(LngPosA + 2, Buffer, StrSeperator)
    TmpName = Mid(Buffer, LngPosA + 2, LngPosB - (LngPosA + 2))
    GetName = TmpName
    Close #1
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Mendapatkan Key Pendaftaran dari DiskKey
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Function GetKey(CbDrvStr) As String
    Dim Buffer As String, TmpKey As String
    Dim LngPosA As Long
    
    Open CbDrvStr & StrKeyFile For Input As #1
    Line Input #1, Buffer
    Buffer = Crypt(Buffer)
    LngPosA = InStrRev(Buffer, StrSeperator, -1)
    TmpKey = right(Buffer, Len(Buffer) - LngPosA - 1)
    GetKey = TmpKey
    Close #1
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Internal Function - not for outdoor use.
' Please caution
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Function GetSerial(CbDrvStr As String) As Long
    Dim SerialNum As Long
    Dim Res As Long
    Dim Temp1 As String
    Dim Temp2 As String
    Temp1 = String$(255, Chr$(0))
    Temp2 = String$(255, Chr$(0))
    Res = GetVolumeInformation(CbDrvStr & ":\", Temp1, Len(Temp1), SerialNum, 0, 0, Temp2, Len(Temp2))
    GetSerial = SerialNum
End Function

Private Function Crypt(Text As String)
    Dim Tmp As String, Itg As Integer
    For Itg = 1 To Len(Text)
    Tmp$ = Tmp$ + Chr$(Asc(Mid(Text, Itg, 1)) Xor 4)
    Next Itg
    Crypt = Tmp
End Function
