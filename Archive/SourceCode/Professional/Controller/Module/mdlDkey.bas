Attribute VB_Name = "mDkey"
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Membina DiskKey
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Function CreateDiskKey(Name As String, Key As String, CbDrvStr As String) As Boolean
On Error GoTo ErrInt
    Dim Buffer As String
    Dim Sn As Long
    Sn = GetSerial(CbDrvStr)
    
    Buffer = Crypt(Sn & "/:" & Name & "/:" & Key)
    
    Open CbDrvStr & "\nmtdkey.k" For Output As #1
    Print #1, Buffer
    Close #1
    SetAttr CbDrvStr & "\nmtdkey.k", vbHidden
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
    Dim Sn As Long, SnF As String
    Dim sKeyFile As String
    
    ValidateDisk = False
    sKeyFile = "\nmtdkey.k"
        
    If Len(Dir(CbDrvStr & sKeyFile, vbHidden)) = 0 Then Exit Function
    
    'get the disk volume
    Sn = GetSerial(CbDrvStr)
    'get the internal disk volume
    Open CbDrvStr & sKeyFile For Input As #1
    Line Input #1, SnF
    Close #1
    
    SnF = Crypt(SnF)
    X = InStr(1, SnF, "/:")
    SnF = Left(SnF, X - 1)

    If Sn = SnF Then ValidateDisk = True
Exit Function

ErrInt:
    ValidateDisk = False
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Mendapatkan Nama Pendaftaran Dari DiskKey
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Function GetName(CbDrvStr) As String
    Dim Buffer As String
    Dim TmpName As String
    
    Open CbDrvStr & "\nmtdkey.k" For Input As #1
    Line Input #1, Buffer
    Buffer = Crypt(Buffer)
    X = InStr(1, Buffer, "/:")
    Y = InStr(X + 2, Buffer, "/:")
    TmpName = Mid(Buffer, X + 2, Y - (X + 2))
    GetName = TmpName
    Close #1
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Mendapatkan Key Pendaftaran dari DiskKey
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Function GetKey(CbDrvStr) As String
    Dim Buffer As String
    Dim TmpKey As String
    
    Open CbDrvStr & "\nmtdkey.k" For Input As #1
    Line Input #1, Buffer
    Buffer = Crypt(Buffer)
    X = InStrRev(Buffer, "/:", -1)
    TmpKey = Right(Buffer, Len(Buffer) - X - 1)
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
    Res = GetVolumeInformation(CbDrvStr, Temp1, Len(Temp1), SerialNum, 0, 0, Temp2, Len(Temp2))
    GetSerial = SerialNum
End Function
Private Function Crypt(Text As String)
    Dim Tmp As String, Itg As Integer
    For Itg = 1 To Len(Text)
    Tmp$ = Tmp$ + Chr$(Asc(Mid(Text, Itg, 1)) Xor 4)
    Next Itg
    Crypt = Tmp
End Function
