Attribute VB_Name = "mSecurity"
Enum eCbAccessAllow
    [Allow Setting] = 1
    [Allow Statistic] = 2
    [Allow Unlock] = 3
End Enum

Enum eCbAccessTo
    [Configuration] = 1
    [Statistic] = 2
End Enum


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Log Activity Pekerja
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub LogWorker(Activity As String, ParamArray SeqParam())
    If UBound(SeqParam) > -1 Then Activity = LangParam(Activity, SeqParam)
    Activity = LangPrcs(Activity)
    
    If CbLogUser = True Then
        uIDBe.DataSave "pekerja-log", "tarikh", Date, True, False
        uIDBe.DataSave "pekerja-log", "masa", Time, False, False
        uIDBe.DataSave "pekerja-log", "nick", CbUserName, False, False
        uIDBe.DataSave "pekerja-log", "akses", CbUserAccess, False, False
        uIDBe.DataSave "pekerja-log", "perkara", Activity, False, True
    End If
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Cek pasport
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'kena buat encryption untuk ini.. kalau boleh untuk semua setting
Public Function CekPass(NamePass, NeedPass) As Boolean
    CekPass = False
    Cond1 = NamePass = SetAmbil("mu") And NeedPass = SetAmbil("mp")
    If Cond1 Then CekPass = True: CbUserName = "Admin": CbUserAccess = "111": Exit Function
    
    For d = 0 To uSDBe.DataCount("pekerja-list") - 1
        Cond2 = NamePass = uSDBe.DataGet("pekerja-list", "nick", d) And NeedPass = uSDBe.DataGet("pekerja-list", "password", d)
        If Cond2 Then
            CbUserName = uSDBe.DataGet("pekerja-list", "nick", d)
            CbUserAccess = uSDBe.DataGet("pekerja-list", "akses", d)
            CekPass = True
            If SetAmbil("logaktiviti") = True Then CbLogUser = True
            Exit Function
        End If
    Next d
End Function


Public Function CekAkses(AksesPart As eCbAccessAllow) As Boolean
    CekAkses = True
    If Mid(CbUserAccess, AksesPart, 1) = "0" Then
        MsgBox MB(10), vbOKOnly, CbMsgWarn
        CekAkses = False
    End If
End Function

Public Sub Accessing(WhichPart As eCbAccessTo)
    If CekAkses(WhichPart) = False Then
        MsgBox MB(10), vbOKOnly, CbMsgWarn
    Else
        Select Case WhichPart
            Case 1
                FrmSet.Show
            Case 2
                FrmStat.Show
        End Select
    End If
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Tulis log ke list1 dalam frmmain
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub MainLog(Log As String)
    FrmMain.MainLog.AddItem Log
    FrmMain.MainLog.Selected(FrmMain.MainLog.NewIndex) = True
    FrmMain.MainLog.Refresh
End Sub


Function Crypt(Text As String, Codekey As Integer)
    Dim Tmp As String, Itg As Integer
    For Itg = 1 To Len(Text)
    Tmp$ = Tmp$ + Chr$(Asc(Mid(Text, Itg, 1)) Xor Codekey)
    Next Itg
    Crypt = Tmp
End Function

