Attribute VB_Name = "mAplikasi"
'###################################################
'#  Product Name    : CafeBonzer Agent
'#  Copyright       : Nematix Technology
'#  Author          : Azri Jamil
'#  Modul           : mAplikasi
'#  Desc            : Main module for this program
'###################################################

Public lTotalAgent As Long

Public bFirstTime As Boolean
Public bToClose As Boolean
Public bConnected As Boolean



'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Program Entry Points] - Di mana semuanya bermula.....
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Sub Main()
    If App.PrevInstance = True Then End
    FrmHost.Show
End Sub


'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Error handler
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub ErrHand(ErrObj As ErrObject, ProcName As String)
    Dim s_ErrNum As String, s_ErrDesc As String, s_msg As String
    
    s_ErrNum = ErrObj.Number
    s_ErrDesc = ErrObj.Description
    
    Select Case l_errorViewType
    Case 1
        s_msg = "[ " & ProcName & " | "
        s_msg = s_msg & s_ErrNum & " | "
        s_msg = s_msg & s_ErrDesc & " ]"
        MsgBox s_msg, vbCritical + vbOKOnly, App.Title
    End Select
End Sub

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Tutup
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub Tutup()
    Dim Frm As Form
    
  ' Closing Procedure
    'FrmHost.Socket.Cleanup
    Call NetClose

  ' Flushing Object
    For Each Frm In Forms
        Unload Frm
    Next Frm
End Sub
