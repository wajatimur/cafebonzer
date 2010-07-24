Attribute VB_Name = "mdlEnkrip"
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [CryptX] - Nematix Encryption
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Function CryptX(Text As String, Codekey As Integer) As String
Attribute CryptX.VB_Description = "Enkripsi Dengan ""Seal"""
    Dim Tmp As String, Tmp2 As String, Itg As Integer
    Tmp2 = "*nematix*" & Text & "*seal*"
    For Itg = 1 To Len(Tmp2)
        Tmp$ = Tmp$ + Chr$(Asc(Mid(Tmp2, Itg, 1)) Xor Codekey)
    Next Itg
    CryptX = Tmp
End Function

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [CryptX] - Nematix DeCryption
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Function DecryptX(Text As String, Codekey As Integer) As String
Attribute DecryptX.VB_Description = "Dekripsi Dengan ""Seal"""
    Dim Tmp As String, Itg As Integer
    Dim a001 As String, a002 As String
    For Itg = 1 To Len(Text)
        Tmp$ = Tmp$ + Chr$(Asc(Mid(Text, Itg, 1)) Xor Codekey)
    Next Itg
    a001 = Right(Tmp, Len(Tmp) - 9)
    a002 = Left(a001, Len(a001) - 6)
    DecryptX = a002
End Function


