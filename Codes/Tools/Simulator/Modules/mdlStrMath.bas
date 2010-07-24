Attribute VB_Name = "mdlStrMath"
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Validate IP] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Function ValidateIP(StrToCheck) As Boolean
    NoneNumChar = "abcdefghijklmnopqrstuvwxyz/,<>;:'""[]{}-=\|+_()*&^%$#@!~`"
    ValidateIP = True
    
    For d = 1 To Len(StrToCheck)
        For i = 1 To Len(NoneNumChar)
            If Mid(LCase(StrToCheck), d, 1) = Mid(LCase(NoneNumChar), i, 1) Then ValidateIP = False: Exit Function
        Next i
    Next d
End Function

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Shortened Format String] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Function GetShortStr(StrToShort As String, Optional IdealLen As Long = 8)
    If Len(StrToShort) > IdealLen Then
        GetShortStr = Left(StrToShort, IdealLen) & ".."
    Else
        GetShortStr = StrToShort
    End If
End Function

Public Function ConvMem2Str(MemLong As Long) As String
    Dim byteBuffer(64) As Byte, lRet As Long
    lRet = lstrcpy(byteBuffer(0), ByVal MemLong)
    ConvMem2Str = StrConv(byteBuffer(), vbUnicode)
    ConvMem2Str = Left$(ConvMem2Str, InStr(ConvMem2Str, vbNullChar) - 1)
End Function

Public Function RemoveNull(StringWithNull As String) As String
    RemoveNull = Left$(StringWithNull, InStr(StringWithNull, vbNullChar) - 1)
End Function
