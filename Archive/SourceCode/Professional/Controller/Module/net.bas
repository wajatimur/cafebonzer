Attribute VB_Name = "mNetwork"
'==================================================================
' Aplication codename : CafeBonzer
' Programmer          : Azri Jamil a.k.a wajatimur
' Module Name         : Network
' Description         : Network Engine
'==================================================================
Public StackNetData As New Collection
Public Type dNetData
    SockIndex As Long
    Data As String
End Type

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' hidupkan server dan tunggu setiap sambungan
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub NetUp()
On Error GoTo ErrTrap
    FrmHost.Socket(0).LocalPort = SetAmbil("porttempatan")
    FrmHost.Socket(0).Listen
    FrmHost.NetTimer = True
Exit Sub

ErrTrap:
    ErrLog Err, "mNet | NetUp"
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' terminate all network
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub NetClose()
On Error GoTo ErrTrap
    FrmHost.NetTimer = False
    For w% = 0 To FrmHost.Socket.Count - 1
        Set Sck = FrmHost.Socket(w%)
        Sck.Cleanup
        Set Sck = Nothing
    Next w%
Exit Sub

ErrTrap:
    ErrLog Err, "mNet | NetClose"
    Resume Next
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Ping client
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub Ping(SockIndex As Long)
    Send SockIndex, "//hey"
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Hantar data (masukkan kedalam "query list")
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub Send(SockIndex As Long, Data As String)
    Dim TmpDst As New clsDataStore
    
    TmpDst.Name = SockIndex
    TmpDst.Add SockIndex, "sockindex"
    TmpDst.Add Data, "data"
    StackNetData.Add TmpDst
    Set TmpDst = Nothing
End Sub
