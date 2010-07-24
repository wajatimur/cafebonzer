Attribute VB_Name = "mAplikasi"
'==================================================================
' Aplication codename : CafeBonzer
' Programmer          : Azri Jamil a.k.a wajatimur
' Module Name         : Main
' Description         : Main Module
'==================================================================

'Public declaration
Public Const EcKey1 = 8
Public Const EcKey2 = 6

Public CbAppVersion As String
Public CbAppBuild As String
Public CbAppLatestAgn As String

Public CbPathDatRecv As String
Public CurSDBPath As String
Public CurIDBPath As String

Public Rs As Recordset
Public uSDB As Database
Public uIDB As Database
Public uSDBe As New clsData
Public uIDBe As New clsData

Public CbUserName As String
Public CbUserAccess As String
Public CbDemoMode As Boolean
Public CbDrvStr As String
Public CbLogUser As Boolean
Public CbViewMode As Long
Public CbMsgRcv As Boolean
Public CbConsole As Boolean

Public OpenSessionCur As String
Public OpenSessionLast As String

'//cbViewMode Constant//
' 0 = normal mode
' 1 = map mode

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Program Entry Point
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Sub Main()
  '[ Prevent previous instance ]'
    If App.PrevInstance = True Then Exit Sub

  '[ Splash screen ]'
    FrmSplash.Show
    DoEvents: Sleep 500

  '[ Load language ]'
    Call LangLoad

  '[ Demo mode check ]'
    If SetAmbil("demo", False) = True Then
            Dim l_Day As Long
            l_Day = SetAmbil("demoday", 0)
            If SetAmbil("demodate", Tarikh) <> Tarikh Then SetSimpan "demoday", (l_Day + 1)
          ' simpan tarikh demo terakhir dibuka
            SetSimpan "demodate", Tarikh
            FrmDemo.Show vbModal
          ' set variable global cbDemo = True
            CbDemoMode = True
    End If
    SetSimpan "demo", False
    CbDemoMode = False
    
  '[ jalankan rutin bagi buka program ]'
  '[ cek pertama kali..... ]'
    Call PreSetting
  '[ load semua setting ]'
    Call SettingUp
  '[ setup mainform ]'
    Call SettingFrm
  '[ hidupkan network ]'
    Call NetUp
  '[ terminate form frmSplash ]'
    Unload FrmSplash

  '[ minta password ]'
    'FrmPass.Show
    FrmMaster.Show: CbUserName = "admin": CbUserAccess = "111"
End Sub

Public Function RsFilter(RsTmp As Recordset, FilterStr As String)
    If FilterStr = "" Then Exit Function
    RsTmp.Filter = FilterStr
    Set RsFilter = RsTmp.OpenRecordset
End Function


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Keluar dari Program
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub Keluar(Optional Ask As Boolean = True)
    Dim Frm As Form
    
    If CbDemoMode = True Then FrmDemo.Show vbModal
    
    'save form position and size
    Call CbFrmMetricSave(FrmMain)
    
    'check routine
    If Ask = True And FrmMain.Lv1.ListItems.Count = 0 Then msgret = MsgBox(MB(11), vbOKCancel, CbMsgApp)
    If Ask = True And FrmMain.Lv1.ListItems.Count <> 0 Then msgret = MsgBox(MB(12), vbOKCancel, CbMsgApp)
    
    'user decision
    If msgret = vbCancel Then
        Exit Sub
    Else:
        Call UnloadIconic
        Set oSpc = Nothing
        Call NetClose: FrmMain.Hide
        For Each Frm In Forms
            Unload Frm
        Next Frm
    End If
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Tarikh dalam format US
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function Tarikh() As Date
    Hari = Day(Date)
    Bulan = Month(Date)
    Tahun = Year(Date)
    Tarikh = Bulan & "/" & Hari & "/" & Tahun
End Function


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Pengiraan beza objek antara objek kecil dan besar
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function KiraBezaSaizX(ObjekBesar As Object, ObjekKecil As Object) As Long
    KiraBezaSaizX = ObjekBesar.Width - ObjekKecil.Width
End Function
Function KiraBezaSaizY(ObjekBesar As Object, ObjekKecil As Object) As Long
    KiraBezaSaizY = ObjekBesar.Height - ObjekKecil.Height
End Function

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Rounding Nombor
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Function RoundNum(numVal As Double, numDigits As Integer) As Double
    RoundNum = Int(numVal * (10 ^ numDigits) + 0.5) / (10 ^ numDigits)
End Function

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Untuk mendapatkan jumlah hari yang telah digunakan
' dalam demomode
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Function GetDayUse()
    Dim buf As String
    Open App.Path & "\" & App.EXEName For Binary As 1
    Get #1, LOF(1) - 1, buf
    GetDayUse = buf
End Function

Function GetSystemDate(tDay As Integer, tMonth As Integer, tYear As Integer) As String
    Dim St As SYSTEMTIME
    Dim tBuffer As String
    
    St.wDay = tDay
    St.wMonth = tMonth
    St.wYear = tYear
    tBuffer = String(255, 0)
        
    GetDateFormat ByVal 0&, 0, St, vbNullString, tBuffer, Len(tBuffer)
    GetSystemDate = Left(tBuffer, InStr(1, tBuffer, Chr$(0)) - 1)
End Function
