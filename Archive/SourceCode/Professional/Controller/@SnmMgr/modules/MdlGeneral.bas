Attribute VB_Name = "MdlGeneral"
Public uSDB As Database

Sub main()
    Set uSDB = OpenDatabase(App.Path & "\data\sdata.mdb")
    Call LoadLang
    FrmSnmMg.Show
End Sub

Public Sub ErrLog(errType As ErrObject, procName As String, Optional DisplayMsg As Boolean = True)
On Error GoTo ErrInt
    Dim ErrDesc As String
    Dim i_errNum As Integer, s_errDesc As String, s_errSource As String
        
    i_errNum = errType.Number
    s_errSource = errType.Source
    s_errDesc = errType.Description
    
    MsgBox i_errNum & " / " & s_errSource & vbNewLine & s_errDesc, vbExclamation, procName
    ErrDesc = Now & " - " & s_errDesc & " - " & s_errSource & " - " & i_errNum
    
    Open "ErrLog.txt" For Append As #1
    Write #1, ErrDesc
    Close #1
Exit Sub

ErrInt:
    MsgBox "Critical error reach ! Terminate will occur !", vbCritical, "CafeBonzer"
    End
End Sub
