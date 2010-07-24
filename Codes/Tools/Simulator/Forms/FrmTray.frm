VERSION 5.00
Begin VB.Form FrmTray 
   BorderStyle     =   0  'None
   ClientHeight    =   450
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   3015
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmTray.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   450
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Menu pmenu1 
      Caption         =   "<popmenu1>"
      Visible         =   0   'False
      Begin VB.Menu pmenu1_about 
         Caption         =   "About"
      End
      Begin VB.Menu pmenu1sep2 
         Caption         =   "-"
      End
      Begin VB.Menu pmenu1_setting 
         Caption         =   "Penalaan"
      End
      Begin VB.Menu pmenu1_kunci 
         Caption         =   "Kunci PC"
      End
      Begin VB.Menu pmenu1sep1 
         Caption         =   "-"
      End
      Begin VB.Menu pmenu1_shutdown 
         Caption         =   "Shutdown PC"
      End
      Begin VB.Menu pmenu1_close 
         Caption         =   "Tutup Agent"
      End
   End
End
Attribute VB_Name = "FrmTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const ID01 = "Please enter your password"
Private Const ID02 = "Lock This PC ?"
Private Const FM01 = "Close Agent"
Private Const FM02 = "Lock Station"
Private Const FM03 = "Configuration"
Private Const FM04 = "Shutdown PC"


Private Sub Form_Load()
    'caption for menu
    pmenu1_close.Caption = FM01
    pmenu1_kunci.Caption = FM02
    pmenu1_setting.Caption = FM03
    pmenu1_shutdown.Caption = FM04
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Pnt As POINTAPI
    GetCursorPos Pnt
    x = x / Screen.TwipsPerPixelX
    If x = WM_RBUTTONUP Then Me.PopupMenu pmenu1
End Sub


Private Sub pmenu1_about_Click()
    FrmAbout.Show
End Sub

'menu - tutup
Private Sub pmenu1_close_Click()
    Dim cp As String
    
    cp = GetInput(ID01)
    If cp = "" Then Exit Sub
    If CekPass(cp) = "ok" Then b_ToClose = True: Tutup
End Sub
'menu - lock
Private Sub pmenu1_kunci_Click()
    Dim lRet
    lRet = MsgBox(ID02, vbOKCancel)
    If lRet = vbOK Then FncKunci 1
End Sub
'menu - setting
Private Sub pmenu1_setting_Click()
    Dim cp As String
    
    cp = GetInput(ID01)
    If cp = "" Then Exit Sub
    If CekPass(cp) = "ok" Then
        'TickerHide
        'frmTicker.Hide
        FrmMain.Show
    End If
End Sub
'menu - shutdown
Private Sub pmenu1_shutdown_Click()
    Dim cp As String
    
    cp = GetInput(ID01)
    If cp = "" Then Exit Sub
    If CekPass(cp) = "ok" Then
        FncShutdown 1
    End If
End Sub

