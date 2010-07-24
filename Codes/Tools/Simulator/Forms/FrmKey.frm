VERSION 5.00
Begin VB.Form FrmKey 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
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
   Begin VB.Timer timerClock 
      Interval        =   900
      Left            =   570
      Top             =   2445
   End
   Begin CafeBonzerAGSim.Label3D lblClock 
      Height          =   240
      Left            =   105
      TabIndex        =   13
      Top             =   1890
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer timerTicker 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   90
      Top             =   2460
   End
   Begin CafeBonzerAGSim.uLine3D UcLine3D 
      Height          =   45
      Left            =   15
      TabIndex        =   1
      Top             =   1740
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   79
      horizon         =   -1  'True
   End
   Begin VB.PictureBox Pages 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   885
      Index           =   1
      Left            =   0
      ScaleHeight     =   885
      ScaleWidth      =   4665
      TabIndex        =   7
      Top             =   855
      Width           =   4665
      Begin VB.TextBox txtPass 
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
      Begin VB.Label lblPass 
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
         TabIndex        =   9
         Top             =   330
         Width           =   1080
      End
      Begin VB.Image imgPass 
         Height          =   480
         Left            =   285
         Picture         =   "FrmKey.frx":000C
         Top             =   225
         Width           =   480
      End
   End
   Begin VB.PictureBox Pages 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   885
      Index           =   0
      Left            =   0
      ScaleHeight     =   885
      ScaleWidth      =   4665
      TabIndex        =   2
      Top             =   855
      Width           =   4665
      Begin CafeBonzerAGSim.chameleonButton mainBut 
         Height          =   615
         Index           =   0
         Left            =   405
         TabIndex        =   3
         ToolTipText     =   "Unlock | CTRL+U"
         Top             =   135
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmKey.frx":207E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CafeBonzerAGSim.chameleonButton mainBut 
         Height          =   615
         Index           =   1
         Left            =   1200
         TabIndex        =   4
         ToolTipText     =   "Shutdown | CTRL+S"
         Top             =   135
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmKey.frx":209A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CafeBonzerAGSim.chameleonButton mainBut 
         Height          =   615
         Index           =   2
         Left            =   2010
         TabIndex        =   5
         ToolTipText     =   "Restart | CTRL+R"
         Top             =   135
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmKey.frx":20B6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CafeBonzerAGSim.chameleonButton mainBut 
         Height          =   615
         Index           =   3
         Left            =   2805
         TabIndex        =   6
         ToolTipText     =   "Configurations | CTRL+C"
         Top             =   135
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmKey.frx":20D2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CafeBonzerAGSim.chameleonButton mainBut 
         Height          =   615
         Index           =   4
         Left            =   3615
         TabIndex        =   10
         ToolTipText     =   "Exit | CTRL+X"
         Top             =   135
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmKey.frx":20EE
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
   Begin VB.PictureBox picTicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   4650
      TabIndex        =   12
      Top             =   855
      Width           =   4650
   End
   Begin VB.Image imgIdle 
      Height          =   240
      Index           =   1
      Left            =   1140
      Picture         =   "FrmKey.frx":210A
      Top             =   2565
      Width           =   240
   End
   Begin VB.Image imgStatClean 
      Height          =   240
      Index           =   0
      Left            =   1425
      Picture         =   "FrmKey.frx":2694
      ToolTipText     =   "Not Connected"
      Top             =   2580
      Width           =   240
   End
   Begin VB.Image imgStatNet 
      Height          =   240
      Index           =   1
      Left            =   1710
      Picture         =   "FrmKey.frx":2C1E
      ToolTipText     =   "Not Connected"
      Top             =   2580
      Width           =   240
   End
   Begin VB.Image imgStatNet 
      Height          =   240
      Index           =   0
      Left            =   1965
      Picture         =   "FrmKey.frx":31A8
      ToolTipText     =   "Not Connected"
      Top             =   2580
      Width           =   240
   End
   Begin VB.Image imgIdle 
      Height          =   240
      Index           =   0
      Left            =   4395
      Picture         =   "FrmKey.frx":3732
      Top             =   600
      Width           =   240
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nematix Technology© 1996-2002"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   195
      TabIndex        =   11
      Top             =   570
      Width           =   2985
   End
   Begin VB.Image imgStatus 
      Height          =   240
      Left            =   4395
      Picture         =   "FrmKey.frx":3CBC
      ToolTipText     =   "Not Connected"
      Top             =   1875
      Width           =   240
   End
   Begin VB.Label lblBuild 
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
      Height          =   195
      Left            =   3645
      TabIndex        =   0
      Top             =   30
      Width           =   945
   End
   Begin VB.Image imgLogo 
      Height          =   420
      Left            =   705
      Picture         =   "FrmKey.frx":4246
      Top             =   135
      Width           =   3000
   End
   Begin VB.Image imgGuard 
      Height          =   480
      Left            =   120
      Picture         =   "FrmKey.frx":54A1
      ToolTipText     =   "I'm guarding you cafe !"
      Top             =   105
      Width           =   480
   End
   Begin VB.Image imgBg 
      Height          =   900
      Left            =   0
      Picture         =   "FrmKey.frx":5D6B
      Top             =   -45
      Width           =   4650
   End
End
Attribute VB_Name = "FrmKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private obj_PicTmp As StdPicture
Private l_tPos As Long
Private b_IdleMode As Boolean
Public s_TickText As String

Public Enum ue_StatType
    Connected = 1
    Discconnet = 2
    Cleaning = 3
End Enum

Private Sub Form_Activate()
    PutOnTop Me.hwnd
    FormTrap FrmKey, True
    txtPass.SetFocus
End Sub

Private Sub Form_Initialize()
 ' height with ticker = 1260
 ' height withour ticker = 2280
    Me.Width = 4740
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim l_retSw As Integer
    
    If (Shift And vbCtrlMask) > 0 And b_IdleMode = False Then
        l_retSw = Switch(KeyCode = vbKeyU, 0, KeyCode = vbKeyS, 1, KeyCode = vbKeyR, 2, KeyCode = vbKeyC, 3, KeyCode = vbKeyX, 4)
        Call mainBut_Click(l_retSw)
    End If
End Sub

Private Sub Form_Load()
    Call timerClock_Timer
    lblBuild = "Build " & App.Major & "." & s_appBuild
    Set obj_PicTmp = imgIdle(0).Picture
    If b_Connected = True Then Call StatIcon(Connected)
    
    Call IdleMode(SetGet("lock.boxstate", 0) Xor 1)
    Call BoxPos
End Sub

Private Sub imgIdle_Click(Index As Integer)
    If b_IdleMode = True Then
        Call IdleMode(False)
    Else
        Call IdleMode(True)
    End If
    Call BoxPos
    Call FormTrap(FrmKey, True)
End Sub


Private Sub mainBut_Click(Index As Integer)
    Select Case Index
        Case 0: Call FncKunci(0): Unload Me
        Case 1: Call FncShutdown(2)
        Case 2: Call FncShutdown(3)
        Case 3: Call FncKunci(0): FrmMain.Show: Unload Me
        Case 4: Call FncKunci(0): Call Tutup
    End Select
End Sub

Private Sub timerClock_Timer()
    lblClock.Caption = Now
    If b_Connected = 1 Then FrmKey.StatIcon (Connected)
End Sub

Private Sub timerTicker_Timer()
On Error GoTo ErrInt
    Dim txtLen As Long
    
        txtLen = picTicker.TextWidth(s_TickText)
        If s_TickText = "" Then s_TickText = s_Welcome
        
        l_tPos = l_tPos + 15
        picTicker.Cls
        picTicker.CurrentY = 60
        picTicker.CurrentX = picTicker.ScaleWidth - l_tPos
        picTicker.Print s_TickText
        If -(l_tPos) < -(txtLen + picTicker.ScaleWidth) Then l_tPos = 0
Exit Sub

ErrInt:
    TickText = Err.Description
End Sub

Private Sub TxtPass_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrInt
    If txtPass = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then
        If CekPass(txtPass) = "ok" Then
            Pages(0).ZOrder 0: txtPass = ""
        End If
    End If
Exit Sub

ErrInt:
    ret = MsgBox("Error occur ! Restart computer ?", vbOKCancel + vbCritical, appmsg)
    If ret = vbOK Then FncShutdown 3
End Sub


Public Sub IdleMode(Optional IdleTrue As Boolean = True)
    If IdleTrue Then
        Me.Height = 1260
        timerTicker.Enabled = True
        Pages(1).ZOrder 0
        picTicker.ZOrder 0
        Set imgIdle(0).Picture = obj_PicTmp
        b_IdleMode = True
    Else
        Me.Height = 2280
        timerTicker.Enabled = False
        Pages(1).ZOrder 0
        picTicker.ZOrder 1
        Set imgIdle(0).Picture = imgIdle(1).Picture
        'txtPass.SetFocus
        b_IdleMode = False
    End If
End Sub

Public Sub BoxPos()
    Dim l_Pos As Long, l_MidX As Long, l_MidY As Long

    l_Pos = SetGet("lock.boxpos", 0)
    Select Case l_Pos
        Case 0
            l_MidX = (Screen.Width \ 2) - (Me.Width \ 2)
            l_MidY = (Screen.Height \ 2) - (Me.Height \ 2)
            Me.Left = l_MidX
            Me.Top = l_MidY
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

Public Sub StatIcon(StatType As ue_StatType)
    Select Case StatType
        Case 1
            Set imgStatus.Picture = imgStatNet(1).Picture
            imgStatus.ToolTipText = "Connected"
        Case 2
            Set imgStatus.Picture = imgStatNet(0).Picture
            imgStatus.ToolTipText = "Not Connected"
        Case 3
            Set imgStatus.Picture = imgStatClean(0).Picture
            imgStatus.ToolTipText = "Cleaning.."
    End Select
End Sub
