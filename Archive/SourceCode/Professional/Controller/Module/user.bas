Attribute VB_Name = "mUser"
'==================================================================
' Aplication codename : CafeBonzer
' Programmer          : Azri Jamil a.k.a wajatimur
' Module Name         : cbUser Module
' Description         :
'==================================================================


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Henti guna PC tersebut
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub UserHenti()
On Error GoTo ErrInt
    Dim s_cOutTime As String, s_cInTime As String, hO As Double
    'Dim uA As clsAgent
    
    'Set uA = uAgents(SelTag)
    Load FrmHarga
    'FrmHost.Timer1.Enabled = False
    
    hO = SelAgn.CusGetPrice
    mHo = SelAgn.CusGetTimeUse
    mIno = SelAgn.CusGetTimeUse(True)
    
    s_cInTime = TimeValue(SelAgn.CustomerTimeIn)
    s_cOutTime = Time
    'If s_cOutTime <> VS(1) Then s_cOutTime = TimeValue(Time)
    
   'masukkan kedalam variable FrmHarga
    FrmHarga.PcName = SelAgn.AgentName
    FrmHarga.pcCusName = SelAgn.CustomerName
    FrmHarga.pcInTime = s_cInTime
    FrmHarga.pcOutTime = s_cOutTime
    FrmHarga.pcTotalTime = mIno
    FrmHarga.pcPaid = hO

   'prepare for output
    FrmHarga.Harga = Format(hO, "#0.00")
    FrmHarga.Masa.Caption = mHo
    FrmHarga.Jumlah = Crnc & " " & Format(hO, "#0.00")
        
   'display output
    FrmHarga.Show vbModal
Exit Sub

ErrInt:
    ErrLog Err, "UserHenti"
End Sub


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Henti guna PC tersebut - Prepaid & Fixed Time
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub UserHenti2()
On Error GoTo ErrInt
    Dim tJam As Integer, tMinit As Integer, uMin As Integer
    Dim mHo As String, hO As Double, mIno As Double
    Load FrmHarga
    
   'pengiraan harga dan masa
    If Left(SelAgn.CustomerFlag, 1) = "p" Then
        hO = Mid(SelAgn.CustomerFlag, 2)
        mIno = hO / CDbl(SetAmbil("harga"))
    ElseIf Left(SelAgn.CustomerFlag, 1) = "f" Then
        uMin = CDbl(Mid(SelAgn.CustomerFlag, 2))
        hO = uMin * CDbl(SetAmbil("harga"))
        mIno = uMin
    End If
    tJam = mIno \ 60
    tMinit = mIno Mod 60
    mHo = tJam & " " & VS(7) & ", " & tMinit & " " & VS(8)
    
   'masukkan kedalam variable FrmHarga
    FrmHarga.PcName = SelAgn.AgentName
    FrmHarga.pcCusName = SelAgn.CustomerName
    FrmHarga.pcInTime = TimeValue(SelAgn.CustomerTimeIn)
    FrmHarga.pcOutTime = TimeValue(SelAgn.CustomerTimeOut)
    FrmHarga.pcTotalTime = mIno
    FrmHarga.pcPaid = hO
    
   'prepare for output
    FrmHarga.Harga = Format(hO, "#0.00")
    FrmHarga.Masa.Caption = mHo
    FrmHarga.Jumlah = Crnc & " " & Format(hO, "#0.00")
    
   'display output
    FrmHarga.Show vbModal
Exit Sub

ErrInt:
    ErrLog Err, "UserHenti2"
End Sub
