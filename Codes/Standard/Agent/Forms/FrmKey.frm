VERSION 5.00
Object = "{B280D12A-792E-4DF1-AA2A-E84D836A12CC}#3.0#0"; "VISUAL~1.OCX"
Begin VB.Form FrmKey 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2190
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   4650
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   Icon            =   "FrmKey.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   146
   ScaleMode       =   0  'User
   ScaleWidth      =   310
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox ImgBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   4650
      TabIndex        =   11
      Top             =   0
      Width           =   4650
      Begin VB.Label LblBuild 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Build 1.7.42"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   180
         Left            =   3690
         TabIndex        =   13
         Top             =   0
         Width           =   930
      End
      Begin VB.Image ImgBackBrand 
         Height          =   450
         Left            =   75
         Picture         =   "FrmKey.frx":000C
         Top             =   180
         Width           =   2700
      End
      Begin VB.Label LblCopy 
         BackStyle       =   0  'Transparent
         Caption         =   "Nematix Technology© 1996-2004"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   105
         TabIndex        =   12
         Top             =   600
         Width           =   2730
      End
      Begin VB.Image ImgBackLogo 
         Height          =   960
         Left            =   3780
         Picture         =   "FrmKey.frx":04AE
         Top             =   165
         Width           =   960
      End
   End
   Begin VisualSuiteX.VsGuiLine VsLine 
      Height          =   45
      Left            =   15
      TabIndex        =   10
      Top             =   1770
      Width           =   4620
      _ExtentX        =   8149
      _ExtentY        =   79
      horizon         =   -1  'True
   End
   Begin VB.Timer TmrGeneral 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   90
      Top             =   2460
   End
   Begin VB.PictureBox Pages 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   885
      Index           =   0
      Left            =   0
      ScaleHeight     =   885
      ScaleWidth      =   4665
      TabIndex        =   0
      Top             =   855
      Width           =   4665
      Begin VisualSuiteX.VsGuiButton MainBtn 
         Height          =   615
         Index           =   0
         Left            =   405
         TabIndex        =   1
         ToolTipText     =   "Unlock | CTRL+U"
         Top             =   135
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmKey.frx":0EAC
         PICN            =   "FrmKey.frx":0EC8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VisualSuiteX.VsGuiButton mainBut 
         Height          =   615
         Index           =   1
         Left            =   1200
         TabIndex        =   2
         ToolTipText     =   "Shutdown | CTRL+S"
         Top             =   135
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmKey.frx":2F4A
         PICN            =   "FrmKey.frx":2F66
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VisualSuiteX.VsGuiButton mainBut 
         Height          =   615
         Index           =   2
         Left            =   2010
         TabIndex        =   3
         ToolTipText     =   "Restart | CTRL+R"
         Top             =   135
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmKey.frx":4FE8
         PICN            =   "FrmKey.frx":5004
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VisualSuiteX.VsGuiButton mainBut 
         Height          =   615
         Index           =   3
         Left            =   2805
         TabIndex        =   4
         ToolTipText     =   "Configurations | CTRL+C"
         Top             =   135
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmKey.frx":7086
         PICN            =   "FrmKey.frx":70A2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VisualSuiteX.VsGuiButton mainBut 
         Height          =   615
         Index           =   4
         Left            =   3615
         TabIndex        =   5
         ToolTipText     =   "Exit | CTRL+X"
         Top             =   135
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmKey.frx":9324
         PICN            =   "FrmKey.frx":9340
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.PictureBox Pages 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   885
      Index           =   1
      Left            =   0
      ScaleHeight     =   885
      ScaleWidth      =   4665
      TabIndex        =   6
      Top             =   855
      Width           =   4665
      Begin VB.TextBox TxtPass 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   11.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   2025
         PasswordChar    =   "l"
         TabIndex        =   8
         Top             =   285
         Width           =   2175
      End
      Begin VB.Label LblPass 
         BackStyle       =   0  'Transparent
         Caption         =   "Password :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   855
         TabIndex        =   7
         Top             =   330
         Width           =   1080
      End
      Begin VB.Image ImgPass 
         Height          =   480
         Left            =   285
         Picture         =   "FrmKey.frx":BAF2
         Top             =   225
         Width           =   480
      End
   End
   Begin VB.PictureBox PicTicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   4650
      TabIndex        =   9
      Top             =   855
      Width           =   4650
   End
   Begin VB.Image ImgStatus 
      Height          =   240
      Left            =   4380
      Picture         =   "FrmKey.frx":DB64
      ToolTipText     =   "Not Connected"
      Top             =   1890
      Width           =   240
   End
End
Attribute VB_Name = "FrmKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : FrmKey
'    Project    : CafeBonzerAG
'
'    Description: [type_description_here]
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
' height with ticker = 1260
' height withour ticker = 2280
Private BlnIdle As Boolean
Private BlnAccessGranted As Boolean

Public Enum EnStatusIcon
    Connected = 1
    Discconnet = 2
    Cleaning = 3
End Enum


Private Sub Form_Load()
    lblBuild = "Build " & App.Major & "." & StrAppBuild
    If BlnConnected = True Then Call StatIcon(Connected)
    Call IdleMode(SettingGet("AppLockMode", 0) Xor 1)
    Call BoxPos
End Sub


Private Sub Form_Unload(Cancel As Integer)
    FormTrap FrmKey, False
End Sub


Private Sub Form_Activate()
    PutOnTop Me.hWnd
    FormTrap FrmKey, True
    TxtPass.SetFocus
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim VarResult As Variant
    
    If Shift = 7 And KeyCode = vbKeyInsert Then
        Call IdleMode(BlnIdle Xor True)
        Call BoxPos
        Call FormTrap(FrmKey, True)
        Exit Sub
    End If
    If (Shift And vbCtrlMask) > 0 And BlnIdle = False And BlnAccessGranted = True Then
        VarResult = Switch(KeyCode = vbKeyU, 0, KeyCode = vbKeyS, 1, KeyCode = vbKeyR, 2, KeyCode = vbKeyC, 3, KeyCode = vbKeyX, 4)
        If CStr(VarResult & "") = vbNullString Then Exit Sub
        Call MainBtn_Click(CLng(VarResult))
    End If
End Sub


Private Sub MainBtn_Click(Index As Integer)
    Select Case Index
        Case 0: Call SysShellLock(0): Unload Me
        Case 1: Call SysWindowsExit(2)
        Case 2: Call SysWindowsExit(3)
        Case 3: Call SysShellLock(0): FrmMain.Show: Unload Me
        Case 4: Call SysShellLock(0): Call AppExit
    End Select
End Sub


Private Sub TmrGeneral_Timer()
    Dim StrTickMsg As String, LngTickMsgLen As Long
    Static LngTickMsgPos As Long

  ' + TICKER FRAME TIMING ------------------------------------------
    If StrTickMsg = "" Then StrTickMsg = StrTickMsgWelcome
    LngTickMsgLen = PicTicker.TextWidth(StrTickMsg)
    
    LngTickMsgPos = LngTickMsgPos + 15
    PicTicker.Cls
    PicTicker.CurrentY = 60
    PicTicker.CurrentX = PicTicker.ScaleWidth - LngTickMsgPos
    PicTicker.Print StrTickMsg
    If -(LngTickMsgPos) < -(LngTickMsgLen + PicTicker.ScaleWidth) Then LngTickMsgPos = 0

  ' + ICON STATUS TIMER ---------------------------------------------
    If BlnConnected = 1 Then FrmKey.StatIcon (Connected)
End Sub


Private Sub TxtPass_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrInt
    If BlnIdle = True Then Exit Sub
    If KeyCode = vbKeyReturn Then
        If SecCheckPassword(TxtPass) = 1 Then
            Pages(0).ZOrder 0
            TxtPass = ""
            BlnAccessGranted = True
        End If
    End If
Exit Sub

ErrInt:
    Dim LngRet As Long
    LngRet = MsgBox("Error occur ! Restart computer ?", vbOKCancel + vbCritical, appmsg)
    If LngRet = vbOK Then SysWindowsExit 3
End Sub


Public Sub IdleMode(Optional IdleTrue As Boolean = True)
    If IdleTrue Then
        Me.Height = 1260
        TmrGeneral.Enabled = True
        Pages(1).ZOrder 0
        PicTicker.ZOrder 0
        BlnIdle = True
    Else
        Me.Height = 2280
        TmrGeneral.Enabled = False
        Pages(1).ZOrder 0
        PicTicker.ZOrder 1
        BlnIdle = False
    End If
End Sub


Public Sub BoxPos()
    Dim lPos As Long, lMidX As Long, lMidY As Long

    lPos = SettingGet("AppLockPos", 0)
    Select Case lPos
        Case 0
            lMidX = (Screen.Width \ 2) - (Me.Width \ 2)
            lMidY = (Screen.Height \ 2) - (Me.Height \ 2)
            Me.Left = lMidX
            Me.Top = lMidY
        Case 1
            Me.Top = 0
            Me.Left = 0
        Case 2
            Me.Top = 0
            Me.Left = Screen.Width - Me.Width
        Case 3
            Me.Top = Screen.Height - Me.Height
            Me.Left = 0
        Case 4
            Me.Top = Screen.Height - Me.Height
            Me.Left = Screen.Width - Me.Width
    End Select
End Sub


Public Sub StatIcon(StatType As EnStatusIcon)
    Select Case StatType
        Case 1
            Set ImgStatus.Picture = FrmHost.IconStatOn.Picture
            ImgStatus.ToolTipText = "Connected"
        Case 2
            Set ImgStatus.Picture = FrmHost.IconStatOff.Picture
            ImgStatus.ToolTipText = "Disconnected"
        Case 3
            Set ImgStatus.Picture = FrmHost.IconStatClean.Picture
            ImgStatus.ToolTipText = "Cleaning.."
    End Select
End Sub
