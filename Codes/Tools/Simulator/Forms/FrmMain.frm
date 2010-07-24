VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{1FB2138E-EE1B-11D6-9361-B0DA59D02E57}#1.0#0"; "NAXCLRPICK.OCX"
Begin VB.Form FrmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuration"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8220
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
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   8220
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin CafeBonzerAGSim.chameleonButton CbtMainPages 
      Height          =   390
      Index           =   2
      Left            =   1125
      TabIndex        =   20
      ToolTipText     =   "Security Settings 1"
      Top             =   6390
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   688
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
      MICON           =   "FrmMain.frx":08CA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CafeBonzerAGSim.chameleonButton CbtMainPages 
      Height          =   390
      Index           =   1
      Left            =   615
      TabIndex        =   19
      ToolTipText     =   "Appearance Settings"
      Top             =   6390
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   688
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
      MICON           =   "FrmMain.frx":08E6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CafeBonzerAGSim.chameleonButton CbtMainPages 
      Height          =   390
      Index           =   0
      Left            =   105
      TabIndex        =   18
      ToolTipText     =   "General Settings"
      Top             =   6390
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   688
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
      MICON           =   "FrmMain.frx":0902
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CafeBonzerAGSim.chameleonButton CbtMain 
      Height          =   375
      Index           =   0
      Left            =   7260
      TabIndex        =   15
      Top             =   6405
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   661
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
      MICON           =   "FrmMain.frx":091E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox mainTop2 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   -15
      ScaleHeight     =   285
      ScaleWidth      =   8220
      TabIndex        =   13
      Top             =   750
      Width           =   8250
      Begin VB.Label mainTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "General Settings"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Left            =   75
         TabIndex        =   14
         Top             =   15
         Width           =   4950
      End
   End
   Begin VB.PictureBox mainTop1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   -15
      ScaleHeight     =   780
      ScaleWidth      =   8250
      TabIndex        =   5
      Top             =   -15
      Width           =   8250
      Begin CafeBonzerAGSim.Label3D mainTopVer 
         Height          =   255
         Left            =   5490
         TabIndex        =   6
         Top             =   45
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   450
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
      Begin CafeBonzerAGSim.Label3D mainTopCopy 
         Height          =   270
         Left            =   5145
         TabIndex        =   7
         Top             =   300
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   476
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
      Begin VB.Image mainTopFace 
         Height          =   480
         Left            =   75
         Picture         =   "FrmMain.frx":093A
         Top             =   90
         Width           =   480
      End
      Begin VB.Image mainTopLogo 
         Height          =   630
         Left            =   480
         Picture         =   "FrmMain.frx":1204
         Top             =   45
         Width           =   4500
      End
   End
   Begin CafeBonzerAGSim.chameleonButton CbtMain 
      Height          =   375
      Index           =   1
      Left            =   6015
      TabIndex        =   16
      Top             =   6405
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   661
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
      MICON           =   "FrmMain.frx":283E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CafeBonzerAGSim.chameleonButton CbtMain 
      Height          =   375
      Index           =   2
      Left            =   5100
      TabIndex        =   17
      Top             =   6405
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   661
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
      MICON           =   "FrmMain.frx":285A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CafeBonzerAGSim.chameleonButton CbtMainPages 
      Height          =   390
      Index           =   3
      Left            =   1635
      TabIndex        =   65
      ToolTipText     =   "Security Settings 2"
      Top             =   6390
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   688
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
      MICON           =   "FrmMain.frx":2876
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox Page 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5220
      Index           =   0
      Left            =   -15
      ScaleHeight     =   5190
      ScaleWidth      =   8220
      TabIndex        =   8
      Top             =   1050
      Width           =   8250
      Begin VB.CheckBox GenOpt1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Retrive default password on start."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   1
         Left            =   450
         TabIndex        =   55
         ToolTipText     =   "Retrive default password from server when windows start."
         Top             =   2415
         Width           =   2850
      End
      Begin VB.CheckBox GenOpt2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Monitor network traffic."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   3
         Left            =   4560
         TabIndex        =   51
         ToolTipText     =   "Monitor network traffic."
         Top             =   3450
         Width           =   2850
      End
      Begin VB.CheckBox GenOpt2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Monitor applications."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   2
         Left            =   4560
         TabIndex        =   50
         ToolTipText     =   "Monitor process & applications."
         Top             =   3105
         Width           =   2850
      End
      Begin VB.CheckBox GenOpt2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Monitor system resource."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   1
         Left            =   4560
         TabIndex        =   49
         ToolTipText     =   "Monitor print activity."
         Top             =   2745
         Width           =   2850
      End
      Begin VB.TextBox GenWelcome 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   4755
         TabIndex        =   48
         Text            =   ":: CafeBonzer Agent R1 ::"
         Top             =   1260
         Width           =   3150
      End
      Begin VB.CheckBox GenOpt2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Monitor print activity."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   0
         Left            =   4560
         TabIndex        =   47
         ToolTipText     =   "Monitor print activity."
         Top             =   2415
         Width           =   2850
      End
      Begin VB.CheckBox GenOpt1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Autostart on windows begin."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   0
         Left            =   4560
         TabIndex        =   46
         ToolTipText     =   "Automatic start CafeBonzer"
         Top             =   555
         Width           =   2850
      End
      Begin VB.TextBox GenNetPort 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1905
         TabIndex        =   1
         Text            =   "56266"
         Top             =   1005
         Width           =   1650
      End
      Begin VB.TextBox GenNetIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1905
         TabIndex        =   2
         Top             =   1410
         Width           =   1650
      End
      Begin VB.TextBox GenNetName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1905
         TabIndex        =   0
         Text            =   "Cake"
         Top             =   615
         Width           =   1650
      End
      Begin VB.TextBox GenPass1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1860
         PasswordChar    =   "l"
         TabIndex        =   3
         Top             =   2820
         Width           =   1700
      End
      Begin VB.TextBox GenPass2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1860
         PasswordChar    =   "l"
         TabIndex        =   4
         Top             =   3225
         Width           =   1700
      End
      Begin VB.Image GenHdrIco 
         Height          =   240
         Index           =   3
         Left            =   4155
         Picture         =   "FrmMain.frx":2892
         Top             =   1950
         Width           =   240
      End
      Begin VB.Image GenHdrIco 
         Height          =   240
         Index           =   2
         Left            =   4170
         Picture         =   "FrmMain.frx":2E1C
         Top             =   105
         Width           =   240
      End
      Begin VB.Image GenHdrIco 
         Height          =   240
         Index           =   1
         Left            =   120
         Picture         =   "FrmMain.frx":33A6
         Top             =   1935
         Width           =   240
      End
      Begin VB.Image GenHdrIco 
         Height          =   240
         Index           =   0
         Left            =   120
         Picture         =   "FrmMain.frx":3930
         Top             =   135
         Width           =   240
      End
      Begin VB.Label GenMiscLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome Message :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   4575
         TabIndex        =   54
         Top             =   975
         Width           =   1455
      End
      Begin VB.Label GenHdr 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " PC Monitoring"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   3
         Left            =   4350
         TabIndex        =   53
         Top             =   2025
         Width           =   3720
      End
      Begin VB.Label GenHdr 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Miscellaneous"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   2
         Left            =   4350
         TabIndex        =   52
         Top             =   195
         Width           =   3720
      End
      Begin VB.Label GenHdr 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Agent Password"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   1
         Left            =   300
         TabIndex        =   41
         Top             =   2025
         Width           =   3510
      End
      Begin VB.Label GenNetLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server Port :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   495
         TabIndex        =   31
         Top             =   1035
         Width           =   915
      End
      Begin VB.Label GenNetLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server IP :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   480
         TabIndex        =   30
         Top             =   1440
         Width           =   750
      End
      Begin VB.Label GenNetLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Computer Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   480
         TabIndex        =   29
         Top             =   645
         Width           =   1230
      End
      Begin VB.Label GenHdr 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Network Configuration"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   0
         Left            =   315
         TabIndex        =   28
         Top             =   195
         Width           =   3510
      End
      Begin VB.Label GenPassLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Password :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   210
         Index           =   0
         Left            =   450
         TabIndex        =   27
         Top             =   2865
         Width           =   855
      End
      Begin VB.Label GenPassLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Retype :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   210
         Index           =   1
         Left            =   450
         TabIndex        =   26
         Top             =   3270
         Width           =   855
      End
   End
   Begin VB.PictureBox Page 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5220
      Index           =   1
      Left            =   -15
      ScaleHeight     =   5190
      ScaleWidth      =   8220
      TabIndex        =   10
      Top             =   1050
      Width           =   8250
      Begin VB.ComboBox AprLckCb 
         Height          =   315
         ItemData        =   "FrmMain.frx":3EBA
         Left            =   6195
         List            =   "FrmMain.frx":3EC4
         Style           =   2  'Dropdown List
         TabIndex        =   34
         ToolTipText     =   "Choose initial state for Lock Interface"
         Top             =   570
         Width           =   1425
      End
      Begin VB.PictureBox AprLckCon 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   930
         Left            =   6195
         ScaleHeight     =   900
         ScaleWidth      =   1380
         TabIndex        =   57
         Top             =   1020
         Width           =   1410
         Begin VB.PictureBox AprLckPos 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   0
            Left            =   525
            ScaleHeight     =   210
            ScaleWidth      =   300
            TabIndex        =   62
            Top             =   330
            Width           =   330
         End
         Begin VB.PictureBox AprLckPos 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   4
            Left            =   1065
            ScaleHeight     =   210
            ScaleWidth      =   300
            TabIndex        =   61
            Top             =   675
            Width           =   330
         End
         Begin VB.PictureBox AprLckPos 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   3
            Left            =   -15
            ScaleHeight     =   210
            ScaleWidth      =   300
            TabIndex        =   60
            Top             =   675
            Width           =   330
         End
         Begin VB.PictureBox AprLckPos 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   2
            Left            =   1065
            ScaleHeight     =   210
            ScaleWidth      =   300
            TabIndex        =   59
            Top             =   -15
            Width           =   330
         End
         Begin VB.PictureBox AprLckPos 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   1
            Left            =   -15
            ScaleHeight     =   210
            ScaleWidth      =   300
            TabIndex        =   58
            Top             =   -15
            Width           =   330
         End
      End
      Begin ClrPckr.ColorPicker AprClrPick 
         Height          =   300
         Index           =   0
         Left            =   2145
         TabIndex        =   38
         Tag             =   "grpTick"
         Top             =   1620
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   529
         Value           =   0
         Appearance      =   0
         BackColor       =   14737632
      End
      Begin VB.TextBox AprTickSize 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2700
         TabIndex        =   37
         Tag             =   "grpTick"
         Text            =   "7"
         Top             =   795
         Width           =   555
      End
      Begin VB.CheckBox AprOpt1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Disable Tray Ticker."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   0
         Left            =   285
         TabIndex        =   12
         ToolTipText     =   "Enable/Disable the cool tray ticker."
         Top             =   495
         Width           =   2850
      End
      Begin ClrPckr.ColorPicker AprClrPick 
         Height          =   300
         Index           =   1
         Left            =   2145
         TabIndex        =   66
         Tag             =   "grpTick"
         Top             =   2040
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   529
         Value           =   14737632
         Appearance      =   0
         BackColor       =   14737632
      End
      Begin CafeBonzerAGSim.chameleonButton AprFntChoose 
         Height          =   315
         Left            =   1620
         TabIndex        =   67
         Tag             =   "grpTick"
         Top             =   1200
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   556
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
         MICON           =   "FrmMain.frx":3ED7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label AprLckLbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Initial Interface State  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   4455
         TabIndex        =   40
         Top             =   600
         Width           =   1590
      End
      Begin VB.Label AprLckLbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Interface Position  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   4455
         TabIndex        =   63
         Top             =   1005
         Width           =   1380
      End
      Begin VB.Label AprHdr 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Lock Interface Appearance"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   1
         Left            =   4260
         TabIndex        =   56
         Top             =   150
         Width           =   3570
      End
      Begin VB.Label AprTraLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ticker Back Colour  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   540
         TabIndex        =   39
         Top             =   2070
         Width           =   1485
      End
      Begin VB.Label AprTraLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ticker Size  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   555
         TabIndex        =   36
         Top             =   870
         Width           =   930
      End
      Begin VB.Label AprTraLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ticker Font Colour  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   540
         TabIndex        =   35
         Top             =   1680
         Width           =   1440
      End
      Begin VB.Label AprTraLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ticker Font  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   540
         TabIndex        =   33
         Top             =   1260
         Width           =   930
      End
      Begin VB.Label AprHdr 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tray Ticker®"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   0
         Left            =   150
         TabIndex        =   32
         Top             =   150
         Width           =   3525
      End
   End
   Begin VB.PictureBox Page 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5220
      Index           =   3
      Left            =   -15
      ScaleHeight     =   5190
      ScaleWidth      =   8220
      TabIndex        =   11
      Top             =   1050
      Width           =   8250
      Begin VB.CheckBox Sec2Opt1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Protect Desktop Icons."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   1
         Left            =   330
         TabIndex        =   42
         ToolTipText     =   "Protect desktop from change."
         Top             =   870
         Width           =   2355
      End
      Begin VB.CheckBox Sec2Opt1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Protect Desktop Wallpaper."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   0
         Left            =   330
         TabIndex        =   44
         ToolTipText     =   "Wallpaper will be restore when Logout or Lock."
         Top             =   510
         Width           =   2355
      End
      Begin VB.Label Sec2Hdr 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Persistents Setting"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   0
         Left            =   150
         TabIndex        =   45
         Top             =   150
         Width           =   3510
      End
   End
   Begin VB.PictureBox Page 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5220
      Index           =   2
      Left            =   -15
      ScaleHeight     =   5190
      ScaleWidth      =   8220
      TabIndex        =   9
      Top             =   1050
      Width           =   8250
      Begin CafeBonzerAGSim.XTreeOpt SecXtree1 
         Height          =   4995
         Left            =   90
         TabIndex        =   64
         Top             =   105
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   8811
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Indentation     =   256.251983642578
      End
      Begin MSComctlLib.ListView SecDpLv 
         Height          =   2505
         Left            =   4995
         TabIndex        =   43
         Top             =   2115
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   4419
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Application Name"
            Object.Width           =   5115
         EndProperty
      End
      Begin VB.CheckBox SecOpt1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Disable CTL+ALT+DEL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   5070
         TabIndex        =   24
         ToolTipText     =   "Disable the user from using CTL+ALT+DEL"
         Top             =   810
         Width           =   2085
      End
      Begin VB.CheckBox SecOpt1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Disable Registry Editing Tool"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   2
         Left            =   5070
         TabIndex        =   23
         ToolTipText     =   "Disable the famous Regedit"
         Top             =   1110
         Width           =   2355
      End
      Begin VB.CheckBox SecOpt1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Autolock on Windows start."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   0
         Left            =   5070
         TabIndex        =   22
         ToolTipText     =   "Enable/Disable autolock function."
         Top             =   435
         Width           =   2355
      End
      Begin VB.Label SecHrd 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Allow Only Below Applications"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   4965
         TabIndex        =   25
         Top             =   1695
         Width           =   3165
      End
      Begin VB.Label SecHdr 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " System Configuration"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   0
         Left            =   4965
         TabIndex        =   21
         Top             =   120
         Width           =   3165
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const l_DrvMax = 14
Private WithEvents objPolDP As clsPolicies
Attribute objPolDP.VB_VarHelpID = -1


Private Sub AprClrPick_Click(Index As Integer)
    If Index = 0 Then
        SetSave "ticker.forecolor", AprClrPick(0).Color
        FrmTicker.picTicker.ForeColor = AprClrPick(0).Color
    Else
        SetSave "ticker.backcolor", AprClrPick(1).Color
        FrmTicker.picTicker.BackColor = AprClrPick(1).Color
    End If
End Sub

Private Sub AprFntChoose_Click()
    FrmPickFont.PopFontPicker AprFntChoose, FrmTicker.picTicker.FontName, FrmTicker.picTicker.FontSize, FrmTicker.picTicker.ForeColor
End Sub


Private Sub AprLckPos_Click(Index As Integer)
    For Each PictureBox In AprLckPos
        PictureBox.BackColor = vbWhite
    Next
    AprLckPos(Index).BackColor = &H808080
End Sub

Private Sub Form_Initialize()
    Me.Width = 8310
    Me.Height = 7230
End Sub

Private Sub Form_Load()
On Error GoTo ErrInt
    bConTick = False
    
 '{ Make-up }'
    FrmMain.Caption = "Configuration - " & MyName
    SaveString HKEY_CURRENT_USER, "Control Panel\Desktop", "FontSmoothing", "2"
    If Command = "/setup" Then CbtMain(2).Enabled = False
    If b_FirstTime = True Then CbtMain(1).Enabled = False
    
 ' LOAD ALL SETTINGS '*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*
    GenNetName = MyName
    GenNetPort = SetGet("porttempatan", 8180)
    GenNetIP = SetGet("nomborip")
    GenPass1 = SetGet("noid")
    GenPass2 = GenPass1
    GenWelcome = SetGet("misc.welcome", s_VerWelcome)
    GenOpt1(0) = SetGet("autoload", 1)
    GenOpt2(0) = SetGet("mon.printer", 1)
    GenOpt2(1) = SetGet("mon.resource", 1)
    GenOpt2(2) = SetGet("mon.app", 1)
    GenOpt2(3) = SetGet("mon.traffic", 1)
    
    Call LoadSecurityItems
    objPol.DisProgEnum
    
    SecOpt1(0).Value = SetGet("autolock", 1)
    SecOpt1(1).Value = SetGet("disCAD", 1)
    SecOpt1(2).Value = objPol.Misc_DisableRegedit
    Sec2Opt1(0).Value = SetGet("persist.wpaper", 1)
    Sec2Opt1(1).Value = SetGet("persist.deskicons", 0)
    
    AprOpt1(0).Value = SetGet("noticker", 0)
    AprTickSize = SetGet("ticker.size", 7)
    AprFntChoose.Caption = GetShortStr(SetGet("ticker.fontface", "Verdana"))
    AprClrPick(0).Color = SetGet("ticker.fontcolor", vbBlack)
    AprClrPick(1).Color = SetGet("ticker.backcolor", &HE0E0E0)
    
    For Each PictureBox In AprLckPos
        PictureBox.BackColor = vbWhite
    Next
    AprLckCb.ListIndex = SetGet("lock.boxstate", 0)
    AprLckPos(SetGet("lock.boxpos", 0)).BackColor = &H808080
    
Exit Sub

ErrInt:
    ErrHand Err, "FrmMain | Form_Load"
End Sub


'easter egg - really ?
Private Sub mainTopFace_Click()
    Dim Tip2u(5) As String
    Static idx As Long
    Tip2u(1) = "Hai apa kabar !"
    Tip2u(2) = "Boleh berkenalan"
    Tip2u(3) = "Lelaki ke perempuan"
    Tip2u(4) = "Dinosaur sudah lama pupus"
    Tip2u(5) = "Boringnyer hari ini"
    idx = idx + 1
    mainTopLogo.ToolTipText = Tip2u(idx)
    If idx = 5 Then idx = 0
End Sub


Private Sub CbtMain_Click(Index As Integer)
On Error GoTo ErrInt
    Select Case Index
    Case 0
        Call SaveAllSetting
    Case 1
        If SetGet("mula1") <> "tidak" Then End
        Unload FrmMain
    Case 2
        Unload FrmMain
        b_ToClose = True
        Call Tutup
    End Select
Exit Sub

ErrInt:
    ErrHand Err, "FrmMain | cbtMain_Click"
End Sub


Private Sub CbtMainPages_Click(Index As Integer)
    Dim s_PageTitle(3) As String
    s_PageTitle(0) = "General Settings"
    s_PageTitle(1) = "Appearance Settings"
    s_PageTitle(2) = "Security Settings 1"
    s_PageTitle(3) = "Security Settings 2"
    
    mainTitle = s_PageTitle(Index)
    Page(Index).ZOrder 0
End Sub


Private Sub objPolDP_EnumerateDisProg(sKeyPath As String, l_ID As Long, sProgName As String, sEnable As Boolean)
On Error GoTo ErrInt
    Dim nItem As ListItem
    Set nItem = SecDpLv.ListItems.Add(, , sProgName)
    nItem.Checked = sEnable
    nItem.Tag = l_ID
Exit Sub

ErrInt:
    ErrHand Err, "objPolDP_EnumerateDisProg | FrmMain"
End Sub


Private Sub AprOpt1_Click(Index As Integer)
    Call EnableGroup("grpTick", (True Xor b_SelOpt))
End Sub


Private Sub SecDpCbt_Click(Index As Integer)
On Error GoTo ErrInt
    Dim sRet As String
    
    If Index = 0 Then
        sRet = DlgFileOpen("Select program to disable.", "c:\", Me.hwnd, "Executable (*.exe)" + Chr$(0) + "*.exe", , True)
        If Trim(sRet) <> "" Then
            objPol.DisProgAdd sRet
            SecDpLv.ListItems.Clear
            objPol.DisProgEnum
        End If
    ElseIf Index = 1 Then
        objPol.DisProgRemove SecDpLv.SelectedItem.Tag
        SecDpLv.ListItems.Clear
        objPol.DisProgEnum
    End If
Exit Sub

ErrInt:
    ErrHand Err, "SecDpCbt_Click | FrmMain"
End Sub

Private Sub SecDpLv_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ErrInt
    If Item.Checked = True Then
        objPol.DisProgAdd Item.Text, Item.Tag, True
        SecDpLv.ListItems.Clear
        objPol.DisProgEnum
    Else
        objPol.DisProgRemove Item.Tag, True
        SecDpLv.ListItems.Clear
        objPol.DisProgEnum
    End If
Exit Sub

ErrInt:
    ErrHand Err, "SecDpLv_ItemCheck | FrmMain"
End Sub


Public Sub SaveAllSetting()
On Error GoTo ErrInt
        
     ' { Check ip routine.. untuk mengelakkan no to route error }'
        If GenNetPort = "" Then GenNetPort = 8180
        If GenNetIP = "" Then MsgBox "Please enter master computer IP number !!", vbOKOnly, "IP Number": Exit Sub
        If Len(GenNetIP) < 9 Then MsgBox "Please the right IP number ! !": Exit Sub
        If ValidateIP(GenNetIP) = False Then MsgBox "Please enter your IP Number in numerical form !": Exit Sub
        If GenPass1 = "" Or GenPass2 = "" Then MsgBox "Please enter your password !": Exit Sub
        If GenPass1 <> GenPass2 Then MsgBox "Please retype your password !": Exit Sub
        If IsNumeric(AprTickSize) = False Then AprTickSize = 7
        If GenWelcome = "" Then GenWelcome = s_Welcome
        
     ' + SAVING ALL SETTINGS -----------------------------------------------------
     ' { GENERAL SETTINGS } '
        MyNameSet Trim(GenNetName)
        SetSave "porttempatan", Trim(GenNetPort)      'local port
        SetSave "nomborip", Trim(GenNetIP)            'server ip
        SetSave "noid", Trim(GenPass2)                'password
        SetSave "misc.welcome", GenWelcome
        SetSave "autoload", GenOpt1(0)                'autoload
        
        SetSave "mon.printer", GenOpt2(0)
        SetSave "mon.resource", GenOpt2(1)
        SetSave "mon.app", GenOpt2(2)
        SetSave "mon.traffic", GenOpt2(3)
        
     ' { SECURITY SETTING }
        Call SaveSecurityItems
        
        SetSave "autolock", SecOpt1(0).Value
        SetSave "disCAD", SecOpt1(1)
        objPol.Misc_DisableRegedit = SecOpt1(2).Value
        SetSave "persist.wpaper", Sec2Opt1(0).Value
        SetSave "persist.deskicons", Sec2Opt1(1).Value

       '- APPEARANCE
     ' { No Tray Ticker } '
        SetSave "noticker", AprOpt1(0).Value
        SetSave "ticker.size", AprTickSize
        SetSave "ticker.fontface", FrmTicker.picTicker.FontName
        SetSave "ticker.fontsize", FrmTicker.picTicker.FontSize
        SetSave "ticker.fontcolor", AprClrPick(0).Color
        SetSave "ticker.backcolor", AprClrPick(1).Color
        
        SetSave "lock.boxstate", AprLckCb.ListIndex
        For Each PictureBox In AprLckPos
            If PictureBox.BackColor = &H808080 Then
                SetSave "lock.boxpos", PictureBox.Index
            End If
        Next
        
    ' { First Time Flag, malam pertama ?? } '
        SetSave "mula1", "tidak"
        
    ' { Activate Settings } '
        If Command <> "/setup" Then
            If b_FirstTime = False Then Call ActivateSetting
        End If
        
    ' { Unload FrmMain, Sebab ia jarang digunakan } '
        Unload FrmMain
Exit Sub

ErrInt:
    ErrHand Err, "FrmMain | SaveAllSetting"
End Sub

Public Sub LoadSecurityItems()
On Error GoTo ErrInt
    Dim nodExp As Object, nodNet As Object, nodSys As Object, nodDos As Object, nodDrv As Object
    Dim e As Long, f As Long, ret As Long, l_Pcount As Long, o_TopNod As Object
    Dim s_DrvName As String, l_DrvVal As Long
    Dim s_CatKey As String
    
    Set objPolDP = objPol
    Set nodExp = SecXtree1.AddFolder(, , "Explorer", OptionTreeFolder)
    Set nodNet = SecXtree1.AddFolder(, , "Network", OptionTreeFolder)
    Set nodSys = SecXtree1.AddFolder(, , "System", OptionTreeFolder)
    Set nodDos = SecXtree1.AddFolder(, , "Dos", OptionTreeFolder)
    Set nodDrv = SecXtree1.AddFolder(, , "Hide Drives", OptionTreeFolder)
        
 ' { LOAD POLICIES SETTING }
    With SecXtree1
        For f = 1 To 4
            s_CatKey = Choose(f, "exp", "net", "sys", "dos")
            l_Pcount = objPol.GetPoliciesCount(f)  'Choose(f, 14, 6, 13, 2)
            Set o_TopNod = Choose(f, nodExp, nodNet, nodSys, nodDos)
            For e = 1 To l_Pcount
                ret = objPol.GetPolicies2(f, e)
                .AddCheck s_CatKey & e, o_TopNod, objPol.GetPoliciesDesc(f, e), ret
            Next e
        Next f
        
        .ExpandAll
    End With


 ' { LOAD HIDE DRIVES }
 '   Info !
 '   ------------------------
 '      4096 = drive m
 '      m is 13th drive
 '
    With SecXtree1
        ret = objPol.Misc_DriveHide
        
        For f = 1 To l_DrvMax
            s_DrvName = "Drive " & Mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ", f, 1)
            .AddCheck "drv" & f, nodDrv, s_DrvName, OptionTreeCheckNone
        Next f
        
        For e = l_DrvMax To 1 Step -1
            l_DrvVal = 2 ^ (e - 1)
            If ret >= l_DrvVal Then
                .Value("drv" & e) = OptionTreeCheckFull
                ret = ret - l_DrvVal
            End If
        Next e
    End With
Exit Sub

ErrInt:
    ErrHand Err, "FrmMain | LoadSecurityItems"
End Sub


Public Sub SaveSecurityItems()
On Error GoTo ErrInt
    Dim s_CatKey As String, l_Pcount As Long, l_DrvVal As Long
    Dim f As Long, e As Long
    
 ' { SAVE POLICIES SETTING }
    With SecXtree1
        For f = 1 To 4
            s_CatKey = Choose(f, "exp", "net", "sys", "dos")
            l_Pcount = objPol.GetPoliciesCount(f)
            For e = 1 To l_Pcount
                If .Value(s_CatKey & e) = OptionTreeCheckFull Then
                    objPol.SetPolicies f, e, True
                Else
                    objPol.SetPolicies f, e, False
                End If
            Next e
        Next f
    End With

 ' { SAVE HIDE DRIVES }
    With SecXtree1
        For f = 1 To l_DrvMax
            If .Value("drv" & f) = OptionTreeCheckFull Then
                l_DrvVal = l_DrvVal + (2 ^ f)
            End If
        Next f
        objPol.Misc_DriveHide = l_DrvVal
    End With
Exit Sub

ErrInt:
    ErrHand Err, "SaveSecurityItems | FrmMain"
End Sub

Public Sub ActivateSetting()
On Error GoTo ErrInt
    s_Welcome = GenWelcome
    
    Call DisableCtlAltDel(SecOpt1(1))
    Call DeskWallProtect
    Call DeskIconProtect
    
    If GenOpt1(0).Value = 1 Then
        SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", _
        "Tahoma", "cbag.exe"
    Else
        DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", _
        "Tahoma"
    End If
    
    If AprOpt1(0).Value = 1 Then
        FrmTicker.tmrCheck.Enabled = False
        FrmTicker.tmrCheck.Enabled = False
        TickerHide False
        TrayStart
    Else
        TrayRemove
        TickerStart
        If IconCount <> AprTickSize Then
            IconCount = AprTickSize
            TickerHide False
            TickerResize
            TickerShow
        End If
    End If

Exit Sub

ErrInt:
    ErrHand Err, "ActivateSetting | FrmMain"
End Sub

