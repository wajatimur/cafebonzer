Attribute VB_Name = "mdlNet"

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [NetConnect] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub NetConnect()
    FrmHost.Connecter.Enabled = True
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [NetStart] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub NetStart(SckIndex)
On Error GoTo ErrInt
    FrmHost.Pinger = True
    FrmHost.Monitor = True
    
   'send station information
    NetSend SckIndex, "/info.net:" & SubBuild("mac", GetMACAddress)
    NetSend SckIndex, "/info.printers:" & GetPrinters
    
   'jika komputer = lock, hantar status
    If bConLock = 1 Then NetSend SckIndex, "/info.me:lock"
    
    bConnected = True
    FrmHost.mLv.ListItems(SckIndex).SubItems(1) = "Connected"
Exit Sub

ErrInt:
    ErrHand Err, "mdlNet | NetStart"
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [NetClose] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub NetClose()
On Error GoTo ErrInt
    bConnected = False
    FrmHost.Pinger = False
    FrmHost.Monitor = False
    
    FrmHost.mLv.ListItems.Clear
    lTotalAgent = 0
    For a = 1 To FrmHost.Txt(0)
        Unload FrmHost.Socket(a)
    Next a
Exit Sub

ErrInt:
    ErrHand Err, "NetClose | mdlNet"
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [NetPing] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub NetPing(SckIndex)
    Call NetSend(SckIndex, "/hey")
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [NetSend] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub NetSend(SckIndex, Data)
On Error GoTo ErrInt
    If FrmHost.Socket(SckIndex).IsWritable = True Then
        FrmHost.Socket(SckIndex).SendLen = Len(Data)
        FrmHost.Socket(SckIndex).SendData = Data
    End If
Exit Sub

ErrInt:
    If Err.Number = 24054 Then
        NetClose
        NetConnect
    Else
        ErrHand Err, "mdlNet | NetSend"
    End If
End Sub
