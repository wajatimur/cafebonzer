VERSION 5.00
Begin VB.Form FrmTray 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1395
   ClientLeft      =   240
   ClientTop       =   1395
   ClientWidth     =   2295
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmTray.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "FrmTray"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   2295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Menu PopMenu1 
      Caption         =   "PopMenu1"
      Visible         =   0   'False
      Begin VB.Menu Pmenu1About 
         Caption         =   "About"
      End
      Begin VB.Menu Pmenu1Debug 
         Caption         =   "Debug"
      End
      Begin VB.Menu Pmenu1Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu Pmenu1Setting 
         Caption         =   "Configuration"
      End
      Begin VB.Menu Pmenu1Kunci 
         Caption         =   "Lock Terminal"
      End
      Begin VB.Menu Pmenu1Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu Pmenu1Shutdown 
         Caption         =   "Shutdown PC"
      End
      Begin VB.Menu Pmenu1Close 
         Caption         =   "Close Agent"
      End
   End
End
Attribute VB_Name = "FrmTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : FrmTray
'    Project    : CafeBonzerAG
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    x = x / Screen.TwipsPerPixelX
    If x = WM_RBUTTONUP Then Me.PopupMenu PopMenu1
End Sub

Private Sub Pmenu1About_Click()
    FrmAbout.Show
End Sub

Private Sub Pmenu1Close_Click()
    Dim StrPassword As String
    Dim ObjSecurity As New GCbSecurity
    
    StrPassword = ObjSecurity.GetPass(True)
    If StrPassword = "" Then Exit Sub
    If SecCheckPassword(StrPassword) = 1 Then AppExit
End Sub

Private Sub Pmenu1Debug_Click()
    FrmHost.Show
End Sub

Private Sub Pmenu1Kunci_Click()
    Dim lRet As Long
    lRet = MsgBox("Lock this Terminal?", vbOKCancel)
    If lRet = vbOK Then SysShellLock 1
End Sub

Private Sub Pmenu1Setting_Click()
    Dim StrPassword As String
    Dim ObjSecurity As New GCbSecurity
    
    StrPassword = ObjSecurity.GetPass(True)
    If SecCheckPassword(StrPassword) = 1 Then
        FrmMain.Show
    End If
End Sub

Private Sub Pmenu1Shutdown_Click()
    Dim StrPassword As String
    Dim ObjSecurity As New GCbSecurity
    
    StrPassword = ObjSecurity.GetPass(True)
    If StrPassword = "" Then Exit Sub
    If SecCheckPassword(StrPassword) = 1 Then
        SysWindowsExit 1
    End If
End Sub

