VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GCbSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function GetPass(PasswordOnly As Boolean, Optional Seperator As String = ",") As String
    BlnSecPassOnly = PasswordOnly
    StrSecPassUser = ""
    StrSecPassPassword = ""
    FrmPass.Show vbModal
    If PasswordOnly = False Then
        GetPass = StrSecPassUser + Seperator + StrSecPassPassword
    Else
        GetPass = StrSecPassPassword
    End If
End Function
