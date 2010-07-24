Attribute VB_Name = "mStat"

'=-=-====-==-===-===-=-==-=-=--==-=-=-=-===-==-=-==-=-=-==---=--
' Simpan Senarai Penggunaan PC
'=-=-====-==-===-===-=-==-=-=--==-=-=-=-===-==-=-==-=-=-==---=--
Public Sub SavePcUsage(PcName, CusName, TimeIn, TimeOut, JumlahBayar)
    If TimeOut <> VS(1) Then TimeOut = TimeValue(TimeOut)
    
    uIDBe.DataSave "pc-usage", "tahun", Year(Date), True, False
    uIDBe.DataSave "pc-usage", "bulan", Month(Date), False, False
    uIDBe.DataSave "pc-usage", "hari", Day(Date), False, False
    uIDBe.DataSave "pc-usage", "pcname", PcName, False, False
    uIDBe.DataSave "pc-usage", "nama", CusName, False, False
    uIDBe.DataSave "pc-usage", "masuk", TimeIn, False, False
    uIDBe.DataSave "pc-usage", "keluar", TimeOut, False, False
    uIDBe.DataSave "pc-usage", "harga", JumlahBayar, False, True
End Sub


'=-=-====-==-===-===-=-==-=-=--==-=-=-=-===-==-=-==-=-=-==---=--
' Jualan Sebulan Bagi Setiap PC
'=-=-====-==-===-===-=-==-=-=--==-=-=-=-===-==-=-==-=-=-==---=--
Public Sub SavePcBulanan(PcName, JumlahMasa, JumlahBayar)
    Dim mBef As Long, bBef As Double
    Dim RsSS As Recordset
    Set RsSS = uIDB.OpenRecordset("pc-bulanan", dbOpenDynaset)
    
    RsSS.Filter = "tahun = '" & Year(Date) & "' AND bulan = '" & Month(Date) & "'"
    Set Rs = RsSS.OpenRecordset
    
    With Rs
        .FindFirst "namapc = '" & PcName & "'"
        If .NoMatch = False Then
            mBef = !JumlahMasa
            bBef = !JumlahBayar
            .Edit
            !JumlahMasa = JumlahMasa + mBef
            !JumlahBayar = Format$(JumlahBayar + bBef, "#0.00")
            .Update
        Else
            .AddNew
            !Tahun = Year(Date)
            !Bulan = Month(Date)
            !NamaPc = PcName
            !JumlahMasa = JumlahMasa
            !JumlahBayar = JumlahBayar
            .Update
        End If
    End With
End Sub


'=-=-====-==-===-===-=-==-=-=--==-=-=-=-===-==-=-==-=-=-==---=--
' Jualan Harian
'=-=-====-==-===-===-=-==-=-=--==-=-=-=-===-==-=-==-=-=-==---=--
Public Sub SavePcHarian(JumBayar)
    Dim sYear As String, sMonth As String, sDay As String
    Dim vstBef As Integer, jBef As Double
    Set Rs = uIDB.OpenRecordset("pc-harian", dbOpenDynaset)
    

        
    With Rs
        'for current date
        sYear = Year(Date)
        sMonth = Month(Date)
        sDay = Day(Date)
        .FindFirst "tahun = '" & sYear & "' AND bulan = '" & sMonth & "' AND hari = '" & sDay & "'"
        If .NoMatch = False Then
            vstBef = !pelanggan
            jBef = !pungutan
            .Edit
            !pelanggan = vstBef + 1
            !pungutan = Format$(jBef + JumBayar, "#0.00")
            .Update
        Else
            .AddNew
            !Tahun = sYear
            !Bulan = sMonth
            !Hari = sDay
            !pelanggan = 1
            !pungutan = JumBayar
            .Update
        End If
        
        'for session date
        sYear = Year(OpenSessionCur)
        sMonth = Month(OpenSessionCur)
        sDay = Day(OpenSessionCur)
        .FindFirst "tahun = '" & sYear & "' AND bulan = '" & sMonth & "' AND hari = '" & sDay & "'"
        If .NoMatch = False Then
            jBef = !closing
            .Edit
            !closing = Format$(jBef + JumBayar, "#0.00")
            .Update
        End If
    End With
End Sub


'=-=-====-==-===-===-=-==-=-=--==-=-=-=-===-==-=-==-=-=-==---=--
' Jualan Harian
'=-=-====-==-===-===-=-==-=-=--==-=-=-=-===-==-=-==-=-=-==---=--
Public Sub SavePcMingguan(Hari, JumJualan)
    Dim jBef As Double
    Set Rs = uIDB.OpenRecordset("pc-grafminggu", dbOpenDynaset)
    
    With Rs
        .FindFirst "tahun = '" & Year(Date) & "' AND bulan = '" & Month(Date) & "'"
        If .NoMatch = False Then
            jBef = .Fields(Hari).Value
            .Edit
            .Fields(Hari).Value = Format$(jBef + JumJualan, "#0.00")
            .Update
        Else
            .AddNew
            !Tahun = Year(Date)
            !Bulan = Month(Date)
            .Fields(Hari).Value = JumJualan
            .Update
        End If
    End With
End Sub


'=-=-====-==-===-===-=-==-=-=--==-=-=-=-===-==-=-==-=-=-==---=--
' Simpan Senarai Pelanggan
'=-=-====-==-===-===-=-==-=-=--==-=-=-=-===-==-=-==-=-=-==---=--
Public Sub SavePelanggan(CusName, CusID, JumlahMasa, JumlahBayar)
    Dim mBef As Long, bBef As Double, vstBef As Integer
    Set Rs = uSDB.OpenRecordset("pelanggan-list", dbOpenTable)
    
    With Rs
        If CusID <> "" Then
            .Index = "noahli"
            .Seek "=", CusID
        Else
            .Index = "nama"
            .Seek "=", CusName
        End If
    
        If .NoMatch = False Then
            mBef = !JumlahMasa
            bBef = !JumlahBayar
            vstBef = !lawat
            .Edit
            !JumlahMasa = mBef + JumlahMasa
            !JumlahBayar = Format$(bBef + JumlahBayar, "#0.00")
            !lawat = vstBef + 1
            !tarikhakhir = Now
            .Update
        Else
            .AddNew
            !noahli = CusID
            !Nama = CusName
            !JumlahMasa = JumlahMasa
            !JumlahBayar = JumlahBayar
            !lawat = 1
            !tarikhakhir = Now
            .Update
        End If
    End With
End Sub


'=-=-====-==-===-===-=-==-=-=--==-=-=-=-===-==-=-==-=-=-==---=--
' Simpan Transaksi POS
'=-=-====-==-===-===-=-==-=-=--==-=-=-=-===-==-=-==-=-=-==---=--
Public Sub SavePosTrans(Items As ListItems)
    Dim RsF As Recordset
    If Items.Count = 0 Then Exit Sub
    
    Set Rs = uIDB.OpenRecordset("pos-usage", dbOpenDynaset)
    Rs.Filter = "tahun = '" & Year(Date) & "' AND bulan = '" & Month(Date) & "' AND hari = '" & Day(Date) & "'"
    Set RsF = Rs.OpenRecordset
    
    With RsF
        '.MoveLast
        '.MoveFirst
        For g = 1 To Items.Count
            .AddNew
            !Tahun = Year(Date)
            !Bulan = Month(Date)
            !Hari = Day(Date)
            !GroupId = Mid(Items(g).Key, 2, 2)
            !id = Items(g).Key
            !transid = !Tahun & !Bulan & !Hari & GroupId & Format(.RecordCount, "#0000")
            !Item = Items(g).Text
            !qty = Items(g).SubItems(2)
            !Harga = Format$((CDbl(Items(g).SubItems(1)) * CInt(!qty)), "#0.00")
            .Update
        Next g
    End With
End Sub

