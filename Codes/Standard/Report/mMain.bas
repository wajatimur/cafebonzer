Attribute VB_Name = "mMain"
Dim CurIDBPath As String

Sub main()
    CurIDBPath = App.Path & "\data\idata.mdb"
    DataEnv1.CnPcSales.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & CurIDBPath & ";Persist Security Info=False;Pwd=nsb2003"

    If Command = "" Then FrmMain.Show
    If Command = "pc-usage" Then RptPcSales.Show
End Sub
