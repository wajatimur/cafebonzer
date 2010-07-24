Attribute VB_Name = "MdlInterface"
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : MdlInterface
'    Project    : CafeBonzer
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Public CFrmConst As New ClsFormConstraint
Public CIconMenu As New VsIconMenu
Public CVsInput As New VsInput


Public Sub MenuIconAttach()
    CIconMenu.Attach FrmMain.Hwnd
    With CIconMenu
        .HighlightStyle = ECPHighlightStyleGradient
        .ImageList = FrmMain.ImgList16
        .IconIndex(FrmMain.MnuMainConfig.Caption) = FrmMain.ImgList16.ListImages("penalaan").Index - 1
        .IconIndex(FrmMain.MnuMainLogout.Caption) = FrmMain.ImgList16.ListImages("logoff").Index - 1
        .IconIndex(FrmMain.MnuMainClose.Caption) = FrmMain.ImgList16.ListImages("power").Index - 1
        .IconIndex(FrmMain.MnuAgentBroad.Caption) = FrmMain.ImgList16.ListImages("broad").Index - 1
        .IconIndex(FrmMain.MnuAgentBroadSub(0).Caption) = FrmMain.ImgList16.ListImages("mesej").Index - 1
        .IconIndex(FrmMain.MnuAgentCtlLock(0).Caption) = FrmMain.ImgList16.ListImages("TerminalLock").Index - 1
        .IconIndex(FrmMain.MnuAgentCtlLock(1).Caption) = FrmMain.ImgList16.ListImages("TerminalLock").Index - 1
        .IconIndex(FrmMain.MnuAgentCtlLock(2).Caption) = FrmMain.ImgList16.ListImages("kuncibuka").Index - 1
        .IconIndex(FrmMain.MnuAgentCtlWinExit(0).Caption) = FrmMain.ImgList16.ListImages("off").Index - 1
        .IconIndex(FrmMain.MnuAgentCtlWinExit(2).Caption) = FrmMain.ImgList16.ListImages("boot").Index - 1
        .IconIndex(FrmMain.MnuInfoHelp.Caption) = FrmMain.ImgList16.ListImages("help").Index - 1
        .IconIndex(FrmMain.MnuInfoAbout.Caption) = FrmMain.ImgList16.ListImages("info").Index - 1
        
        .IconIndex(FrmMain.PopMnu1Flog.Caption) = FrmMain.ImgList16.ListImages("jalan1").Index - 1
        .IconIndex(FrmMain.PopMnu1Cancel.Caption) = FrmMain.ImgList16.ListImages("no").Index - 1
        .IconIndex(FrmMain.PopMnu1Trans.Caption) = FrmMain.ImgList16.ListImages("transfer").Index - 1
        .IconIndex(FrmMain.PopMnu1Cln.Caption) = FrmMain.ImgList16.ListImages("TerminalClean").Index - 1
        .IconIndex(FrmMain.PopMnu1Ctl.Caption) = FrmMain.ImgList16.ListImages("cpu").Index - 1
        .IconIndex(FrmMain.PopMnu1CtlSub(0).Caption) = FrmMain.ImgList16.ListImages("TerminalLock").Index - 1
        .IconIndex(FrmMain.PopMnu1CtlSub(1).Caption) = FrmMain.ImgList16.ListImages("kuncibuka").Index - 1
        .IconIndex(FrmMain.PopMnu1CtlSub(2).Caption) = FrmMain.ImgList16.ListImages("boot").Index - 1
        .IconIndex(FrmMain.PopMnu1CtlSub(3).Caption) = FrmMain.ImgList16.ListImages("off").Index - 1
    End With
End Sub

Public Sub MenuIconDetach()
    CIconMenu.Detach
End Sub

Sub StatText(Optional Panel As Integer = 0, Optional Text As String = "")
    FrmMain.MainSbar.Panels(Panel).Text = Text
End Sub


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Tutup form dan bebaskan resource
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub FormClose(AnyFrm As Form)
    AnyFrm.Hide
    Unload AnyFrm
    Set AnyFrm = Nothing
End Sub


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Pengiraan beza objek antara objek kecil dan besar
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function MetricDiffrenceX(ObjekBesar As Object, ObjekKecil As Object) As Long
    MetricDiffrenceX = ObjekBesar.Width - ObjekKecil.Width
End Function
Function MetricDiffrenceY(ObjekBesar As Object, ObjekKecil As Object) As Long
    MetricDiffrenceY = ObjekBesar.Height - ObjekKecil.Height
End Function

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Simpan info metric bagi form
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub MetricFrmSave(FormName As Form)
    If FormName.WindowState = 0 Then
        SettingSave FormName.Name & "SaizX", FormName.Width
        SettingSave FormName.Name & "SaizY", FormName.Height
        SettingSave FormName.Name & "PosX", FormName.Left
        SettingSave FormName.Name & "PosY", FormName.Top
    End If
    SettingSave FormName.Name & "State", FormName.WindowState
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Load metric info
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub MetricFrmLoad(FormName As Form)
    FormName.Width = SettingGet(FormName.Name & "SaizX", 11460)
    FormName.Height = SettingGet(FormName.Name & "SaizY", 8640)
    FormName.Top = SettingGet(FormName.Name & "PosX", 0)
    FormName.Left = SettingGet(FormName.Name & "PosY", 0)
    FormName.WindowState = SettingGet(FormName.Name & "State", 0)
End Sub

Public Sub FormMove(Hwnd As Long)
    ReleaseCapture
    SendMessage Hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Public Sub FormOnTop(Hwnd As Long)
    SetWindowPos Hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_WNDFLAGS
End Sub
