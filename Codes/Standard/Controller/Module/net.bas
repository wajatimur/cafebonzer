Attribute VB_Name = "MdlNetworking"
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : MdlNetworking
'    Project    : CafeBonzer
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
'==================================================================
' Aplication codename : CafeBonzer
' Programmer          : Azri Jamil a.k.a wajatimur
' Module Name         : Network
' Description         : Network Engine
'==================================================================
Public StackNetData As New Collection

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Network Up - Server Listen
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub NetUp()
On Error GoTo ErrTrap
    FrmSysHost.Socket(0).LocalPort = SetGetDb("NetPortLocal", 8180)
    FrmSysHost.Socket(0).Listen
    FrmSysHost.TmrNet = True
Exit Sub

ErrTrap:
    AppErrorLog Err, "mNet | NetUp"
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' terminate all network
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub NetClose()
On Error GoTo ErrTrap
    Dim LngIdx As Long
    FrmSysHost.TmrNet = False
    For LngIdx = 0 To FrmSysHost.Socket.Count - 1
        Set Sck = FrmSysHost.Socket(LngIdx)
        Sck.Cleanup
        Set Sck = Nothing
    Next LngIdx
Exit Sub

ErrTrap:
    AppErrorLog Err, "mNet | NetClose"
    Resume Next
End Sub


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Hantar data (masukkan kedalam "query list")
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub Send(SockIndex As Long, Data As String)
    Dim TmpDst As New ClsDataStore
    
    TmpDst.Name = SockIndex
    TmpDst.Add SockIndex, "sockindex"
    TmpDst.Add Data, "data"
    StackNetData.Add TmpDst

    Set TmpDst = Nothing
End Sub
