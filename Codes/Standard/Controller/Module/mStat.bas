Attribute VB_Name = "MdlStatistic"
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : MdlStatistic
'    Project    : CafeBonzer
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>

'=-=-====-==-===-===-=-==-=-=--==-=-=-=-===-==-=-==-=-=-==---=--
' Simpan Senarai Pelanggan
'=-=-====-==-===-===-=-==-=-=--==-=-=-=-===-==-=-==-=-=-==---=--
Public Sub SaveCustomer(CusName As String, CusId As String, TotalTime As Long, TotalPaid As Double)
    Dim LngTmpTime As Long, DblTmpPaid As Double, LngTmpVisit As Long
    Set CRset = CDataS.OpenRecordset("ListCustomer", dbOpenTable)
    
    With CRset
        If CusId <> "" Then
            .Index = "Id"
            .Seek "=", CusId
        Else
            .Index = "Name"
            .Seek "=", CusName
        End If
    
        If .NoMatch = False Then
            LngTmpTime = !TotalTime
            DblTmpPaid = !TotalPaid
            LngTmpVisit = !CountVisit
            .Edit
            !TotalTime = LngTmpTime + TotalTime
            !TotalPaid = Format$(DblTmpPaid + TotalPaid, "#0.00")
            !CountVisit = LngTmpVisit + 1
            !LastVisit = Now
            .Update
        Else
            .AddNew
            !Id = CusId
            !Name = CusName
            !TotalTime = TotalTime
            !TotalPaid = TotalPaid
            !CountVisit = 1
            !LastVisit = Now
            .Update
        End If
    End With
End Sub


'=-=-====-==-===-===-=-==-=-=--==-=-=-=-===-==-=-==-=-=-==---=--
' Simpan Senarai Penggunaan PC
'=-=-====-==-===-===-=-==-=-=--==-=-=-=-===-==-=-==-=-=-==---=--
Public Sub SaveTransactionPc(Terminal, CusName, TimeIn, TimeOut, Usage)
    Dim StrTransactionId As String
    StrTransactionId = StatGenerateId
    CDataIe.DataSave "LogUsageTerminal", "Year", Year(Date), True, False
    CDataIe.DataSave "LogUsageTerminal", "Month", Month(Date), False, False
    CDataIe.DataSave "LogUsageTerminal", "Day", Day(Date), False, False
    CDataIe.DataSave "LogUsageTerminal", "DaySession", Day(StatGetSessionDate), False, False
    CDataIe.DataSave "LogUsageTerminal", "Operator", CbUserName, False, False
    CDataIe.DataSave "LogUsageTerminal", "TransactionId", StrTransactionId, False, False
    CDataIe.DataSave "LogUsageTerminal", "Terminal", Terminal, False, False
    CDataIe.DataSave "LogUsageTerminal", "Customer", CusName, False, False
    CDataIe.DataSave "LogUsageTerminal", "TimeIn", TimeIn, False, False
    CDataIe.DataSave "LogUsageTerminal", "TimeOut", TimeOut, False, False
    CDataIe.DataSave "LogUsageTerminal", "Price", Usage, False, True
End Sub


'=-=-====-==-===-===-=-==-=-=--==-=-=-=-===-==-=-==-=-=-==---=--
' Simpan Transaksi POS
'=-=-====-==-===-===-=-==-=-=--==-=-=-=-===-==-=-==-=-=-==---=--
Public Sub SaveTransactionPos(Items As ListItems)
    Dim RsF As Recordset, LngIdxA As Long
    If Items.Count = 0 Then Exit Sub
    
    Set CRset = CDataI.OpenRecordset("LogUsageServices", dbOpenDynaset)
    CRset.Filter = "Year = " & Year(Date) & " AND Month = " & Month(Date) & " AND Day = " & Day(Date) '& "'"
    Set RsF = CRset.OpenRecordset
    
    With RsF
        For LngIdxA = 1 To Items.Count
            If Items(LngIdxA).Text <> VS(2, 0) Then
                .AddNew
                !Year = Year(Date)
                !Month = Month(Date)
                !Day = Day(Date)
                !DaySession = Day(StatGetSessionDate)
                !GroupID = Items(LngIdxA).Tag
                !Id = Items(LngIdxA).Key
                !TransactionId = !Year & !Month & !Day & !GroupID & Format(.RecordCount, "#0000")
                !Item = Items(LngIdxA).Text
                !Quantity = Mid(Items(LngIdxA).SubItems(1), InStr(1, Items(LngIdxA).SubItems(1), "=") + 1)
                !Price = Format$(Items(LngIdxA).SubItems(2), "#0.00")
                .Update
            End If
        Next
    End With
End Sub


Public Function StatGenerateId() As String
    Dim CRset As Recordset, Rss As Recordset
    Dim StrSql As String, StrNewId As String
    
    StrSql = "Year = " & Year(Date) & " AND Month = " & Month(Date) & " AND Day = " & Day(Date)
    Set Rss = CDataI.OpenRecordset("LogTransactionDaily", dbOpenSnapshot)
    Set CRset = RsFilter(Rss, StrSql)
    
    With CRset
        If .BOF = True Then
            StrNewId = Year(Date) & Format(Month(Date), "#00") & Format(Day(Date), "#00") & "0000"
        Else
            StrNewId = Year(Date) & Format(Month(Date), "#00") & Format(Day(Date), "#00") & Format(!Customer, "#0000")
        End If
        StatGenerateId = StrNewId
    End With
End Function


Public Function StatGetSalesMonth(Year, Month, TransactionType As Long)
    Dim CDbRs As Recordset, StrSqlQ As String, StrTransactionType As String
    Dim DblSales As Double
    
    StrTransactionType = Choose(TransactionType, "LogUsageTerminal", "LogUsageServices")
    StrSqlQ = "SELECT * FROM " & StrTransactionType & " WHERE Year = " & Year & " AND Month = " & Month
    
    Set CDbRs = CDataI.OpenRecordset(StrSqlQ, dbOpenSnapshot)
    With CDbRs
        If .BOF = False Then
            Do Until .EOF = True
                DblSales = DblSales + !Price
                .MoveNext
            Loop
        End If
    End With
    Set CDbRs = Nothing
    StatGetSalesMonth = DblSales
End Function


Public Function StatGetSalesDay(Year, Month, Day, TransactionType As Long, Optional Session As Boolean)
    Dim CDbRs As Recordset, StrSqlQ As String, StrTransactionType As String
    Dim DblSales As Double
    
    StrTransactionType = Choose(TransactionType, "LogUsageTerminal", "LogUsageServices")
    If Session = False Then
        StrSqlQ = "SELECT * FROM " & StrTransactionType & " WHERE Year = " & Year & " AND Month = " & Month & " AND Day = " & Day
    Else
        StrSqlQ = "SELECT * FROM " & StrTransactionType & " WHERE Year = " & Year & " AND Month = " & Month & " AND DaySession = " & Day
    End If
    
    Set CDbRs = CDataI.OpenRecordset(StrSqlQ, dbOpenSnapshot)
    With CDbRs
        If .BOF = False Then
            Do Until .EOF = True
                DblSales = DblSales + !Price
                .MoveNext
            Loop
        End If
    End With
    Set CDbRs = Nothing
    StatGetSalesDay = DblSales
End Function


Public Function StatGetCustomer(Year, Month, Day, Optional Session As Boolean) As Long
    Dim CDbRs As Recordset, StrSqlQ As String
    If Session = False Then
        StrSqlQ = "SELECT * FROM LogUsageTerminal WHERE Year = " & Year & " AND Month = " & Month & " AND Day = " & Day
    Else
        StrSqlQ = "SELECT * FROM LogUsageTerminal WHERE Year = " & Year & " AND Month = " & Month & " AND DaySession = " & Day
    End If
    
    Set CDbRs = CDataI.OpenRecordset(StrSqlQ, dbOpenSnapshot)
    StatGetCustomer = CDbRs.RecordCount
    Set CDbRs = Nothing
End Function


Public Function StatGetSessionDate() As Date
    Dim DteSessionDate As Date
    Dim LngSessionDay As Long

    LngSessionDay = SetGetDb("FinSessionDay", 0)
    If LngSessionDay = 0 Then
        If Time > CDate("12:00:00 AM") And Time < CDate(OpenSessionCur) Then
            DteSessionDate = CDate(OpenSessionCur) + DateAdd("d", -1, Date)
        Else
            DteSessionDate = Time + Date
        End If
    ElseIf LngSessionDay = 1 Then
        If Time > CDate(OpenSessionCur) Then
            DteSessionDate = Time + DateAdd("d", 1, Date)
        Else
            DteSessionDate = Time + Date
        End If
    End If
    StatGetSessionDate = DteSessionDate
End Function


Public Function StatGetSalesByDay(Year As String, Month As String, DayConstant As VbDayOfWeek) As Double
    Dim CDbRs As Recordset, StrSqlQ As String, DblSalesTerminal As Double, DblSalesServices As Double
    
    StrSqlQ = "SELECT * FROM LogUsageTerminal WHERE Year = " & Year & " AND Month = " & Month
    Set CDbRs = CDataI.OpenRecordset(StrSqlQ, dbOpenSnapshot)
    With CDbRs
        If .BOF = False Then
            Do Until .EOF = True
                If Weekday(DateSerial(Year, Month, !Day)) = DayConstant Then
                    DblSalesTerminal = DblSalesTerminal + !Price
                End If
                .MoveNext
            Loop
        End If
    End With
    
    StrSqlQ = "SELECT * FROM LogUsageServices WHERE Year = " & Year & " AND Month = " & Month
    Set CDbRs = CDataI.OpenRecordset(StrSqlQ, dbOpenSnapshot)
    With CDbRs
        If .BOF = False Then
            Do Until .EOF = True
                If Weekday(DateSerial(Year, Month, !Day)) = DayConstant Then
                    DblSalesServices = DblSalesServices + !Price
                End If
                .MoveNext
            Loop
        End If
    End With
    
    StatGetSalesByDay = DblSalesTerminal + DblSalesServices
End Function


Public Sub LoadDate(CCboxYear As ComboBox, CCboxMonth As ComboBox, CCboxDay As ComboBox, Optional CboxChanged As Long)
    Dim Rss As Recordset, StrTmpDate As String, SqlQ As String
    '   0 = Year Changed
    '   1 = Month Changed

 '{ Load years }'
    If CCboxYear.ListCount = 0 Then
        SqlQ = "SELECT DISTINCT Year FROM LogUsageTerminal"
        Set Rss = CDataI.OpenRecordset(SqlQ, dbOpenSnapshot)
        With Rss
            If .BOF = False Then
                Do Until .EOF = True
                    CbAddEx !Year, CCboxYear
                    .MoveNext
                Loop
            End If
            Call CbSelect(Year(Date), CCboxYear)
        End With
        Set Rss = Nothing
    End If

 '{ Load month }'
    If CboxChanged = 0 Then
        CCboxMonth.Clear
        SqlQ = "SELECT DISTINCT Month FROM LogUsageTerminal WHERE Year = " & CCboxYear
        Set Rss = CDataI.OpenRecordset(SqlQ, dbOpenSnapshot)
        With Rss
            Do Until .EOF = True
                CbAddEx !Month, CCboxMonth
                .MoveNext
            Loop
            Call CbSelect(Month(Date), CCboxMonth)
        End With
    End If

 '{ display to combo todays date }'
    If CboxChanged = 1 Then
        CCboxDay.Clear
        SqlQ = "SELECT DISTINCT Day FROM LogUsageTerminal WHERE Year = " & CCboxYear & " AND Month = " & CCboxMonth
        Set Rss = CDataI.OpenRecordset(SqlQ, dbOpenSnapshot)
        With Rss
            Do Until .EOF = True
                'If !Year = CCboxYear And !Month = CCboxMonth Then
                    StrTmpDate = DateGetSystem(!Day, CInt(CCboxMonth), CInt(CCboxYear))
                    CCboxDay.AddItem StrTmpDate
                'End If
                .MoveNext
            Loop
        End With
    '{ display to combo todays date }'
       If StrYear = Year(Date) And StrMonth = Month(Date) Then
           CCboxDay = Date
       Else
           CCboxDay.ListIndex = 0
       End If
    End If

End Sub

