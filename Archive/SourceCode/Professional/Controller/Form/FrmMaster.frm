VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm FrmMaster 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "CafeBonzer"
   ClientHeight    =   7155
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11610
   LinkTopic       =   "CafeHost"
   NegotiateToolbars=   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin CafeBonzer.PageDock MainVbar 
      Align           =   4  'Align Right
      Height          =   6780
      Left            =   10440
      TabIndex        =   0
      Top             =   0
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   11959
      HldrBtnPos      =   1
      HldrLne         =   0   'False
      PageState       =   0
      PageWidth       =   1170
      Begin VB.PictureBox MainVbarCtn 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6780
         Left            =   345
         ScaleHeight     =   6720
         ScaleWidth      =   765
         TabIndex        =   1
         Tag             =   "subcontainer"
         Top             =   0
         Width           =   825
         Begin CafeBonzer.XpButton DbMnu 
            Height          =   555
            Index           =   0
            Left            =   105
            TabIndex        =   2
            ToolTipText     =   "Configuration"
            Top             =   75
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   979
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
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   16777215
            MPTR            =   1
            MICON           =   "FrmMaster.frx":0000
            PICN            =   "FrmMaster.frx":001C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin CafeBonzer.XpButton DbMnu 
            Height          =   555
            Index           =   1
            Left            =   105
            TabIndex        =   3
            ToolTipText     =   "Statistic"
            Top             =   675
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   979
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
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   16777215
            MPTR            =   1
            MICON           =   "FrmMaster.frx":229E
            PICN            =   "FrmMaster.frx":22BA
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin CafeBonzer.XpButton DbMnu 
            Height          =   555
            Index           =   2
            Left            =   105
            TabIndex        =   4
            ToolTipText     =   "Monitoring : Printer"
            Top             =   1275
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   979
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
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   16777215
            MPTR            =   1
            MICON           =   "FrmMaster.frx":433C
            PICN            =   "FrmMaster.frx":4358
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin CafeBonzer.XpButton DbMnu 
            Height          =   555
            Index           =   3
            Left            =   105
            TabIndex        =   5
            ToolTipText     =   "Monitoring : Resources"
            Top             =   1875
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   979
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
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   16777215
            MPTR            =   1
            MICON           =   "FrmMaster.frx":63DA
            PICN            =   "FrmMaster.frx":63F6
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin CafeBonzer.XpButton DbMnu 
            Height          =   555
            Index           =   4
            Left            =   105
            TabIndex        =   6
            ToolTipText     =   "Monitoring : Process"
            Top             =   2475
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   979
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
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   16777215
            MPTR            =   1
            MICON           =   "FrmMaster.frx":8478
            PICN            =   "FrmMaster.frx":8494
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin CafeBonzer.XpButton DbMnu 
            Height          =   555
            Index           =   5
            Left            =   105
            TabIndex        =   7
            ToolTipText     =   "Terminal"
            Top             =   3075
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   979
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
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   16777215
            MPTR            =   1
            MICON           =   "FrmMaster.frx":A516
            PICN            =   "FrmMaster.frx":A532
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin CafeBonzer.XpButton DbMnu 
            Height          =   555
            Index           =   6
            Left            =   105
            TabIndex        =   8
            ToolTipText     =   "Shutdown"
            Top             =   3675
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   979
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
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   16777215
            MPTR            =   1
            MICON           =   "FrmMaster.frx":C5B4
            PICN            =   "FrmMaster.frx":C5D0
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
   End
   Begin MSComctlLib.StatusBar MainStat 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   6780
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2469
            MinWidth        =   2469
            TextSave        =   "4:58 AM"
            Key             =   "stat1"
            Object.ToolTipText     =   "CafeBonzer"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
            Key             =   "stat2"
            Object.ToolTipText     =   "Panel Informasi"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10389
            MinWidth        =   4939
         EndProperty
      EndProperty
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
   Begin VB.Menu MnuFile 
      Caption         =   "File"
      Begin VB.Menu MnuFileSub 
         Caption         =   "Configuration"
         Index           =   0
      End
      Begin VB.Menu MnuFileSub 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu MnuFileSub 
         Caption         =   "LogOut"
         Index           =   2
      End
      Begin VB.Menu MnuFileSub 
         Caption         =   "Close"
         Index           =   3
      End
   End
   Begin VB.Menu MnuStation 
      Caption         =   "Station"
      Begin VB.Menu MnuStationBcast 
         Caption         =   "Broadcast"
         Begin VB.Menu MnuStationBcastSub 
            Caption         =   "Message"
            Index           =   0
         End
         Begin VB.Menu MnuStationBcastSub 
            Caption         =   "Ticker"
            Index           =   1
         End
      End
      Begin VB.Menu MnuStationCtl 
         Caption         =   "Control"
         Begin VB.Menu MnuStationCtlSub 
            Caption         =   "Lock All"
            Index           =   0
         End
         Begin VB.Menu MnuStationCtlSub 
            Caption         =   "Lock Unused"
            Index           =   1
         End
         Begin VB.Menu MnuStationCtlSub 
            Caption         =   "Unlock All"
            Index           =   2
         End
         Begin VB.Menu MnuStationCtlSub 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu MnuStationCtlSub 
            Caption         =   "Shutdown All"
            Index           =   4
         End
         Begin VB.Menu MnuStationCtlSub 
            Caption         =   "Shutdown Unused"
            Index           =   5
         End
         Begin VB.Menu MnuStationCtlSub 
            Caption         =   "Reboot All"
            Index           =   6
         End
         Begin VB.Menu MnuStationCtlSub 
            Caption         =   "Reboot Unused"
            Index           =   7
         End
      End
      Begin VB.Menu MnuStationCln 
         Caption         =   "Cleaning"
         Begin VB.Menu MnuStationClnSub 
            Caption         =   "Clean All"
            Index           =   0
         End
         Begin VB.Menu MnuStationClnSub 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu MnuStationClnSub 
            Caption         =   "Temp Folder"
            Index           =   2
         End
         Begin VB.Menu MnuStationClnSub 
            Caption         =   "Recycle Bin"
            Index           =   3
         End
         Begin VB.Menu MnuStationClnSub 
            Caption         =   "Internet History"
            Index           =   4
         End
         Begin VB.Menu MnuStationClnSub 
            Caption         =   "Recent Docs"
            Index           =   5
         End
      End
   End
   Begin VB.Menu MnuTools 
      Caption         =   "Tools"
      Begin VB.Menu MnuToolsSub 
         Caption         =   "Service && Merchandise"
         Index           =   0
      End
      Begin VB.Menu MnuToolsSub 
         Caption         =   "Statistic System"
         Index           =   1
      End
      Begin VB.Menu MnuToolsSub 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu MnuToolsSub 
         Caption         =   "Resources Monitoring"
         Index           =   3
      End
   End
   Begin VB.Menu MnuHelp 
      Caption         =   "Help"
      Begin VB.Menu MnuHelpSub 
         Caption         =   "Contents"
         Index           =   0
      End
      Begin VB.Menu MnuHelpSub 
         Caption         =   "About"
         Index           =   1
      End
   End
   Begin VB.Menu PopMenu 
      Caption         =   "<PopMenu>"
      Visible         =   0   'False
      Begin VB.Menu PopMnuFlog 
         Caption         =   "Fast Login"
      End
      Begin VB.Menu PopMnuFlout 
         Caption         =   "Fast Logout"
      End
      Begin VB.Menu PopMnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu PopMnuCancel 
         Caption         =   "Cancel User"
      End
      Begin VB.Menu PopMnuTrans 
         Caption         =   "Transfer PC"
      End
      Begin VB.Menu PopMnuConsole 
         Caption         =   "Terminal"
      End
      Begin VB.Menu PopMnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu PopMnuCln 
         Caption         =   "Cleaning"
         Begin VB.Menu PopMnuClnSub 
            Caption         =   "All"
            Index           =   0
         End
         Begin VB.Menu PopMnuClnSub 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu PopMnuClnSub 
            Caption         =   "Temp Folder"
            Index           =   2
         End
         Begin VB.Menu PopMnuClnSub 
            Caption         =   "Recycle Bin"
            Index           =   3
         End
         Begin VB.Menu PopMnuClnSub 
            Caption         =   "Internet History"
            Index           =   4
         End
         Begin VB.Menu PopMnuClnSub 
            Caption         =   "Recent Docs"
            Index           =   5
         End
      End
      Begin VB.Menu PopMnuCtl 
         Caption         =   "Control"
         Begin VB.Menu PopMnuCtlSub 
            Caption         =   "Lock Computer"
            Index           =   0
         End
         Begin VB.Menu PopMnuCtlSub 
            Caption         =   "Unlock Computer"
            Index           =   1
         End
         Begin VB.Menu PopMnuCtlSub 
            Caption         =   "Reboot Computer"
            Index           =   2
         End
         Begin VB.Menu PopMnuCtlSub 
            Caption         =   "Shutdown Computer"
            Index           =   3
         End
      End
   End
End
Attribute VB_Name = "FrmMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public LngWorkSpaceX As Long
Public LngWorkSpaceY As Long

Private Sub MDIForm_Load()
    FrmMainTool.Show
    'FrmMain.Show
End Sub

Private Sub MDIForm_Resize()
    LngWorkSpaceX = FrmMaster.Width - (MainVbar.Width + 200)
    LngWorkSpaceY = FrmMaster.Height - (MainStat.Height + 750)
    
    FrmMainTool.Top = LngWorkSpaceY - FrmMainTool.Height
    FrmMainTool.Width = LngWorkSpaceX
End Sub

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' DockBar | Menu
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub DbMnu_Click(Index As Integer)
    Select Case Index
        Case 0
            Call Accessing(Configuration)
        Case 1
            Call Accessing(Statistic)
        Case 2
            Call ConstructPage(1)
        Case 3
            Call ConstructPage(2)
        Case 4
            Call ConstructPage(0)
        Case 5
            FrmTerminal.Show
        Case 6
            Call Keluar
    End Select
End Sub

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Page Dock
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub MainPdock_PageFliped(ByVal Flipped As Boolean)
    'If Flipped = True Then
    '    NegateXtmp = NegateX
    '    Pages(0).Width = MainPdock.Left
    '    NegateX = FrmMain.Width - Pages(0).Width
    'Else
    '    NegateX = NegateXtmp
    '    Pages(0).Width = FrmMain.Width - NegateX
    'End If
    'Menu4EnvSub(1).Checked = Flipped Xor True
    'Call LayOutSize
    'SetSimpan "dockbar", FrmMain.Menu4EnvSub(1).Checked
End Sub




Private Sub MnuFileSub_Click(Index As Integer)
    Select Case Index
    '[ Penalaan ]'
    Case 0
        If Mid(CbUserAccess, 1, 1) = "0" Then MsgBox MB(10), vbOKOnly, CbMsgWarn: Exit Sub
        LogWorker SL(5) '((security log))
        FrmSet.Show vbModal
    '[ Logout ]'
    Case 2
        LogWorker SL(2) '((security log))
        FrmMain.Hide
        FrmPass.Show
    '[ Keluar ]'
    Case 3
        Call Keluar
    End Select
End Sub

Private Sub MnuHelpSub_Click(Index As Integer)
    Select Case Index
    '[ Help ]'
    Case 0
        Call LoadModule(Help)
    '[ About ]'
    Case 1
        FrmAbout.Show vbModal
    End Select
End Sub

Private Sub MnuStationBcastSub_Click(Index As Integer)
    Dim StrMesej As String
    Select Case Index
    '[ Pengumuman : Mesej ]'
    Case 0
        StrMesej = MgoInpt.GetInput("Please enter your announcement", BtnClose)
        If Trim(StrMesej) <> "" Then UniAgents.SendCommand "mesej:Server:" & StrMesej
    '[ Pengumuman : Tiker ]'
    Case 1
        StrMesej = MgoInpt.GetInput("Please enter your message ticker", BtnClose)
        If Trim(StrMesej) <> "" Then UniAgents.SendCommand "tiker:" & StrMesej
    End Select
End Sub

Private Sub MnuStationCtlSub_Click(Index As Integer)
    Dim uA As clsAgent, j As Long
    
    For j% = 1 To AgentCount
        Select Case Index
            Case 0
                UniAgents(j%).NetSend "//kunci:1"
            Case 1
                If UniAgents(j%).AgentStatus = VS(4) Then UniAgents(j%).NetSend "//kunci:1"
            Case 2
                UniAgents(j%).NetSend "//kunci:0"
        End Select
    Next j%
    
    For j = 1 To AgentCount
        Set uA = UniAgents(j)
        Select Case Index
            Case 0
                uA.NetSend "//sdown:2"
            Case 1
                If uA.AgentStatus = VS(4) Then uA.NetSend "//sdown:2"
            Case 2
                uA.NetSend "//sdown:3"
            Case 3
                If uA.AgentStatus = VS(4) Then uA.NetSend "//sdown:3"
        End Select
    Next j
    Select Case Index
        Case 0: LogWorker SL(11) '((security log))
        Case 2: LogWorker SL(10) '((security log))
    End Select
End Sub

Private Sub MnuToolsSub_Click(Index As Integer)
    Select Case Index
    Case 0
        Call LoadModule(CafeSnmMgr)
    Case 1
        Call Accessing(Statistic)
    End Select
End Sub
