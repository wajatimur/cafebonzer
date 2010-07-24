Attribute VB_Name = "mCustomer"
'==================================================================
' Aplication codename : CafeBonzer
' Programmer          : Azri Jamil a.k.a wajatimur
' Module Name         : cbUser Module
' Description         :
'==================================================================


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Dapatkan harga bagi PC tersebut
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Function UserHarga(lvIndex As Integer, CusType As String) As String
    Dim pRate As Double, gi01 As Double, gi02 As Double

   'ambil harga mengikut skema
    If LCase(CusType) = LCase(VS(2)) Then
       'default pricing
        pRate = SetAmbil("harga")
    Else
       'pricing for current scheme
        pRate = uSDBe.DataFind("skema", "skema", "harga", CusType)
    End If
   'ambil nilai biasa jika harga skema tidak dapat diterima
    If pRate < 0 Then pRate = SetAmbil("harga")

   'pengiraan harga..
    gi01 = DateDiff("n", AgentSel.CustomerTimeIn, Now)
    gi02 = (gi01 * pRate) + SetAmbil("hargaex")

   'return value.. easy heh..
    gi02 = GetRoundUpVal(gi02)
    UserHarga = Format(gi02, "#0.00")
End Function


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Jumlah masa yang telah di gunakan oleh Pengguna tersebut
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Function UserMasa(Masa, MasaSemasa, Optional GetMinute As Boolean = False) As String
    Dim Jam, Minit, Ai
    Ai = 0
   'cek dari PM ke AM untuk mengelakkan kesilapan dalam pengiraan
    If Right(Masa, 2) = "PM" And Right(MasaSemasa, 2) = "AM" Then Ai = 1440
   'dapatkan jumlah minit
    Minit = DateDiff("n", Masa, MasaSemasa) + Ai
    If GetMinute = True Then UserMasa = Minit: Exit Function
   'dapatkan jumlah jam
    Jam = Minit \ 60
   'dapatkan minit selepas jam
    Minit = Minit - (Jam * 60)
   'gabungkan kesemua
    UserMasa = Jam & " " & VS(7) & ", " & Minit & " " & VS(8)
End Function


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Jumlah masa yang tinggal
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Function UserMasaTinggal(Masa, MasaSemasa)
    Dim Jam, Minit
    
    Minit = DateDiff("n", MasaSemasa, Masa)
    Jam = Minit \ 60
    Minit = Minit - (Jam * 60)
    UserMasaTinggal = Jam & " " & VS(7) & ", " & Minit & " " & VS(8)
End Function
