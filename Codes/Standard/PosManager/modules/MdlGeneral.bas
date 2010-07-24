Attribute VB_Name = "MdlGeneral"
Public StrIdHead As String
Public StrIdGrpHead As String
Public uSDB As Database

Sub Main()
    Set uSDB = OpenDatabase(App.Path & "\data\sdata.mdb", False, False, ";pwd=nsb2003")
    Call LoadLang
    
    StrIdHead = "CP"
    StrIdGrpHead = "CPG"
    FrmSnmMg.Show
End Sub

Public Sub ErrLog(ErrType As ErrObject, ProcName As String, Optional DisplayMsg As Boolean = True)
    Dim StrError As String
    Dim IntErrNum As Integer, StrErrDesc As String, StrErrSource As String
        
    IntErrNum = ErrType.Number
    StrErrSource = ErrType.Source
    StrErrDesc = ErrType.Description
    
    If DisplayMsg = True Then
        MsgBox IntErrNum & " / " & StrErrSource & vbNewLine & StrErrDesc, vbExclamation, ProcName
    End If
    
    StrError = Now & " - " & StrErrSource & " - " & StrErrDesc & " - " & IntErrNum
    
    Open "ErrLog.txt" For Append As #1
    Write #1, StrError
    Close #1
End Sub
