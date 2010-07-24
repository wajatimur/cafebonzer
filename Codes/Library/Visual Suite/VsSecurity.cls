VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VsSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function CheckPass(Optional PasswordOnly As Boolean = False, Optional DefaultUser As String = "admin") As Boolean
    Dim Db As Database, Rs As Recordset
    Set Db = OpenDatabase(App.Path & "\data\sdata.mdb", False, False, ";pwd=nsb2003")
    Set Rs = Db.OpenRecordset("ListEmployee", dbOpenTable)
   
    CheckPass = False
    PasswordOnly = BlnVsPasswordOnly
    StrVsCheckPassUser = DefaultUser
    StrVsCheckPassPassword = ""
    
    FrmPass.Show vbModal
    
    Cond1 = StrVsCheckPassUser = SetGetDb("GenAdminName", "admin") And StrVsCheckPassPassword = SetGetDb("GenAdminPass")
    If Cond1 Then
        CheckPass = True
        Exit Function
    End If
    
    With Rs
        .Index = "UserName"
        .Seek "=", StrVsCheckPassUser
        If .NoMatch = False Then
            If !Password = StrVsCheckPassPassword Then CheckPass = True
        End If
    End With
End Function

Public Function GetPass(Optional PasswordOnly As Boolean = False, Optional Seperator As String = ",") As String
    BlnVsPasswordOnly = PasswordOnly
    FrmPass.Show vbModal
    
    If PasswordOnly = False Then
        GetPass = StrVsCheckPassUser + Seperator + StrVsCheckPassPassword
    Else
        GetPass = StrVsCheckPassPassword
    End If
End Function
