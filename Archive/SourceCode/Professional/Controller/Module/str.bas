Attribute VB_Name = "mStrMath"
'==================================================================
' Aplication codename : CafeBonzer
' Programmer          : Azri Jamil a.k.a wajatimur
' Module Name         : String Control
' Description         :
'==================================================================

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Bandingkan kewujudan huruf dalam ayat yang di berikan
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Function CharCompare(Ayat As String, SetAyat As String) As Boolean
    Pjayat = Len(Ayat)
    pjset = Len(SetAyat)
    
    For g = 1 To Pjayat
        For H = 1 To pjset
            If Mid(Ayat, g, 1) = Mid(SetAyat, H, 1) Then BandingSetHuruf = True: Exit For
        Next H
    Next g
End Function

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Bilang jumlah huruf dalam ayat yang diberikan
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Function CharCount(Ayat As String, Huruf As String) As Integer
    Dim Jhrf As Integer
    Pjayat = Len(Ayat)
    
    If Len(Huruf) > 1 Or Len(Huruf) = 0 Then BilangJumlahHuruf = 0: Exit Function
    
    For H = 1 To Pjayat
        If Huruf = Mid(Ayat, H, 1) Then Jhrf = Jhrf + 1
    Next H
    
    BilangJumlahHuruf = Jhrf
End Function

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Nilai yang di roundup (depend on setting)
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function GetRoundUpVal(dblValue As Double) As Double
    If SetAmbil("roundup") = 1 Then
        GetRoundUpVal = Round(dblValue, 1)
    Else
        GetRoundUpVal = dblValue
    End If
End Function

Public Function GetBulan(MonthNum)
    GetBulan = Choose(MonthNum, "Januari", "Februari", "Mac", "April", "Mei", "Jun", "Julai", "Ogos", "September", "Oktober", "November", "Disember")
End Function

Public Function Text2Num(Text As String) As Double
    On Error Resume Next
    If Text = "" Then
        Text2Num = 0
    Else
        Text2Num = CDbl(Text)
    End If
End Function

Public Function GetDirName(DirPath As String)
    cf = InStrRev(DirPath, "\", -1)
    GetDirName = Mid(DirPath, cf + 1)
End Function

Public Function GetArrNameNum(Nama) As Integer
    j = InStrRev(Nama, "(", Len(Nama) - 1)
    GetArrNameNum = CInt(Mid(Nama, j + 1, Len(Nama) - j - 2))
End Function


