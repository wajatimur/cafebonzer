VERSION 5.00
Object = "{B280D12A-792E-4DF1-AA2A-E84D836A12CC}#3.0#0"; "VISUAL~1.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CafeBonzer Agent Configuration"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VisualSuiteX.VsGuiLine CfgLine 
      Height          =   45
      Index           =   0
      Left            =   15
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   6570
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   79
      horizon         =   -1  'True
   End
   Begin VisualSuiteX.VsGuiButton CbtMain 
      Height          =   375
      Index           =   0
      Left            =   6735
      TabIndex        =   38
      Top             =   6705
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   661
      TX              =   "OK"
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
      MICON           =   "FrmMain.frx":6852
      PICN            =   "FrmMain.frx":686E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox CfgHeader 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   7695
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   855
      Width           =   7695
      Begin VisualSuiteX.VsGuiButton BtnPage 
         Height          =   405
         Index           =   0
         Left            =   45
         TabIndex        =   0
         Tag             =   "General"
         ToolTipText     =   "General Configuration"
         Top             =   60
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
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
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmMain.frx":6E08
         PICN            =   "FrmMain.frx":6E24
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   -1  'True
         VALUE           =   -1  'True
      End
      Begin VisualSuiteX.VsGuiButton BtnPage 
         Height          =   405
         Index           =   1
         Left            =   570
         TabIndex        =   1
         Tag             =   "Appearence"
         ToolTipText     =   "Appearence"
         Top             =   60
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
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
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmMain.frx":7736
         PICN            =   "FrmMain.frx":7752
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   -1  'True
         VALUE           =   0   'False
      End
      Begin VisualSuiteX.VsGuiButton BtnPage 
         Height          =   405
         Index           =   2
         Left            =   1095
         TabIndex        =   2
         Tag             =   "Security"
         ToolTipText     =   "Security Configuration"
         Top             =   60
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
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
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmMain.frx":7CEC
         PICN            =   "FrmMain.frx":7D08
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   -1  'True
         VALUE           =   0   'False
      End
   End
   Begin VB.PictureBox CfgBanner 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   7710
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   0
      Width           =   7710
      Begin VisualSuiteX.VsGuiLabel CfgHeaderInfo 
         Height          =   210
         Index           =   0
         Left            =   6045
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   360
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   370
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor1      =   14737632
         ForeColor2      =   8421504
         Caption         =   "Cafebonzer Agent"
         BackColor       =   16777215
      End
      Begin VisualSuiteX.VsGuiLabel CfgHeaderInfo 
         Height          =   225
         Index           =   1
         Left            =   3990
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   585
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   397
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor1      =   14737632
         ForeColor2      =   8421504
         Caption         =   "Copyright 1996-2004 Nematix Technology"
         BackColor       =   16777215
      End
      Begin VB.Label CfgBannerLbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "General"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6615
         TabIndex        =   48
         Top             =   -30
         Width           =   1050
      End
      Begin VB.Image CfgBannerImg 
         Height          =   960
         Left            =   30
         Picture         =   "FrmMain.frx":82A2
         Top             =   30
         Width           =   960
      End
   End
   Begin VisualSuiteX.VsGuiButton CbtMain 
      Height          =   375
      Index           =   1
      Left            =   5580
      TabIndex        =   39
      Top             =   6705
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   661
      TX              =   "Cancel"
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
      MICON           =   "FrmMain.frx":8CA0
      PICN            =   "FrmMain.frx":8CBC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VisualSuiteX.VsGuiButton CbtMain 
      Height          =   375
      Index           =   2
      Left            =   4680
      TabIndex        =   40
      Top             =   6705
      Visible         =   0   'False
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   661
      TX              =   "Exit"
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
      MICON           =   "FrmMain.frx":9256
      PICN            =   "FrmMain.frx":9272
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VisualSuiteX.VsGuiLine CfgLine 
      Height          =   45
      Index           =   1
      Left            =   0
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   1380
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   79
      horizon         =   -1  'True
   End
   Begin VB.PictureBox Page 
      Appearance      =   0  'Flat
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
      Height          =   5000
      Index           =   0
      Left            =   30
      ScaleHeight     =   4995
      ScaleWidth      =   7635
      TabIndex        =   3
      Top             =   1545
      Width           =   7635
      Begin VisualSuiteX.VsGuiLine CfgGenLine 
         Height          =   45
         Index           =   1
         Left            =   1830
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   2070
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   79
         horizon         =   -1  'True
      End
      Begin VisualSuiteX.VsGuiLine CfgGenLine 
         Height          =   45
         Index           =   0
         Left            =   2625
         TabIndex        =   4
         Top             =   150
         Width           =   5010
         _ExtentX        =   8837
         _ExtentY        =   79
         horizon         =   -1  'True
      End
      Begin VB.CheckBox GenOpt1 
         Caption         =   "Syncronize server password."
         Height          =   390
         Index           =   0
         Left            =   3960
         TabIndex        =   8
         ToolTipText     =   "Retrive default password on connect."
         Top             =   420
         Width           =   3630
      End
      Begin VB.TextBox GenWelcome 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1965
         TabIndex        =   11
         Top             =   2415
         Width           =   5490
      End
      Begin VB.CheckBox GenOpt2 
         Caption         =   "Autostart on windows begin."
         Height          =   315
         Index           =   0
         Left            =   150
         TabIndex        =   12
         ToolTipText     =   "Automatic start CafeBonzer on Windows Start."
         Top             =   2820
         Width           =   2850
      End
      Begin VB.TextBox GenNetPort 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1845
         TabIndex        =   6
         Top             =   855
         Width           =   1650
      End
      Begin VB.TextBox GenNetIP 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1845
         TabIndex        =   7
         Top             =   1260
         Width           =   1650
      End
      Begin VB.TextBox GenNetName 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1845
         TabIndex        =   5
         Top             =   465
         Width           =   1650
      End
      Begin VB.TextBox GenPass1 
         Alignment       =   2  'Center
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
         Left            =   5370
         PasswordChar    =   "l"
         TabIndex        =   9
         Tag             =   "GRPPASSWORD"
         Top             =   885
         Width           =   1700
      End
      Begin VB.TextBox GenPass2 
         Alignment       =   2  'Center
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
         Left            =   5370
         PasswordChar    =   "l"
         TabIndex        =   10
         Tag             =   "GRPPASSWORD"
         Top             =   1290
         Width           =   1700
      End
      Begin VB.Image GenHdrIco 
         Height          =   240
         Index           =   2
         Left            =   45
         Picture         =   "FrmMain.frx":960C
         Top             =   1935
         Width           =   240
      End
      Begin VB.Image GenHdrIco 
         Height          =   240
         Index           =   0
         Left            =   45
         Picture         =   "FrmMain.frx":9B96
         Top             =   45
         Width           =   240
      End
      Begin VB.Label GenMiscLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome Message :"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   56
         Top             =   2445
         Width           =   1710
      End
      Begin VB.Label GenHdr 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   330
         TabIndex        =   55
         Top             =   1965
         Width           =   1425
      End
      Begin VB.Label GenNetLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server Port :"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   51
         Top             =   885
         Width           =   1125
      End
      Begin VB.Label GenNetLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server IP :"
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   53
         Top             =   1290
         Width           =   960
      End
      Begin VB.Label GenNetLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Computer Name :"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   50
         Top             =   495
         Width           =   1545
      End
      Begin VB.Label GenHdr 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   49
         Top             =   60
         Width           =   2250
      End
      Begin VB.Label GenPassLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password :"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   0
         Left            =   3960
         TabIndex        =   52
         Top             =   930
         Width           =   945
      End
      Begin VB.Label GenPassLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Retype :"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   1
         Left            =   3960
         TabIndex        =   54
         Top             =   1335
         Width           =   735
      End
   End
   Begin VB.PictureBox Page 
      Appearance      =   0  'Flat
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
      Height          =   5000
      Index           =   1
      Left            =   30
      ScaleHeight     =   4995
      ScaleWidth      =   7635
      TabIndex        =   13
      Top             =   1545
      Width           =   7635
      Begin VisualSuiteX.VsGfxPicker AprClrPick 
         Height          =   300
         Index           =   0
         Left            =   2055
         TabIndex        =   17
         Tag             =   "GRPTICKER"
         Top             =   1245
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         Value           =   0
         Appearance      =   0
         BackColor       =   14737632
      End
      Begin VisualSuiteX.VsGuiLine CfgAprLine 
         Height          =   45
         Index           =   0
         Left            =   1410
         TabIndex        =   14
         Top             =   135
         Width           =   6240
         _ExtentX        =   11007
         _ExtentY        =   79
         horizon         =   -1  'True
      End
      Begin VisualSuiteX.VsGuiLine CfgAprLine 
         Height          =   45
         Index           =   1
         Left            =   2790
         TabIndex        =   19
         Top             =   2775
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   79
         horizon         =   -1  'True
      End
      Begin VB.ComboBox AprLckCb 
         Height          =   315
         ItemData        =   "FrmMain.frx":A120
         Left            =   1890
         List            =   "FrmMain.frx":A12A
         Style           =   2  'Dropdown List
         TabIndex        =   20
         ToolTipText     =   "Choose initial state for Lock Interface"
         Top             =   3105
         Width           =   1425
      End
      Begin VB.PictureBox AprLckCon 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   930
         Left            =   1890
         ScaleHeight     =   900
         ScaleWidth      =   1380
         TabIndex        =   21
         Top             =   3555
         Width           =   1410
         Begin VB.PictureBox AprLckPos 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   0
            Left            =   525
            ScaleHeight     =   210
            ScaleWidth      =   300
            TabIndex        =   22
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
            TabIndex        =   26
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
            TabIndex        =   25
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
            TabIndex        =   24
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
            TabIndex        =   23
            Top             =   -15
            Width           =   330
         End
      End
      Begin VB.TextBox AprTickSize 
         Alignment       =   2  'Center
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
         TabIndex        =   16
         Tag             =   "GRPTICKER"
         Text            =   "7"
         Top             =   780
         Width           =   555
      End
      Begin VB.CheckBox AprOpt1 
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
         Height          =   315
         Index           =   0
         Left            =   195
         TabIndex        =   15
         ToolTipText     =   "Enable/Disable the cool tray ticker."
         Top             =   450
         Width           =   2850
      End
      Begin VisualSuiteX.VsGfxPicker AprClrPick 
         Height          =   300
         Index           =   1
         Left            =   2055
         TabIndex        =   18
         Tag             =   "GRPTICKER"
         Top             =   1665
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         Value           =   14737632
         Appearance      =   0
         BackColor       =   14737632
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
         Left            =   150
         TabIndex        =   63
         Top             =   3135
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
         Left            =   150
         TabIndex        =   62
         Top             =   3540
         Width           =   1380
      End
      Begin VB.Label AprHdr 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   15
         TabIndex        =   61
         Top             =   2670
         Width           =   2715
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
         TabIndex        =   60
         Top             =   1695
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
         TabIndex        =   58
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
         TabIndex        =   59
         Top             =   1305
         Width           =   1440
      End
      Begin VB.Label AprHdr 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   15
         TabIndex        =   57
         Top             =   45
         Width           =   1350
      End
   End
   Begin VB.PictureBox Page 
      Appearance      =   0  'Flat
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
      Height          =   5000
      Index           =   2
      Left            =   30
      ScaleHeight     =   4995
      ScaleWidth      =   7635
      TabIndex        =   27
      Top             =   1545
      Width           =   7635
      Begin VB.CheckBox SecOpt1 
         Caption         =   "Autolock on Windows start."
         Height          =   315
         Index           =   0
         Left            =   270
         TabIndex        =   28
         ToolTipText     =   "Enable/Disable autolock function."
         Top             =   345
         Width           =   2790
      End
      Begin VisualSuiteX.VsGuiXTree SecTree 
         Height          =   4890
         Left            =   3195
         TabIndex        =   37
         Top             =   45
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   8625
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         Indentation     =   256.251983642578
         Style           =   3
      End
      Begin VB.CheckBox SecOpt2 
         Caption         =   "Monitor network traffic."
         Enabled         =   0   'False
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
         Index           =   3
         Left            =   270
         TabIndex        =   36
         ToolTipText     =   "Monitor network traffic."
         Top             =   3705
         Width           =   2355
      End
      Begin VB.CheckBox SecOpt2 
         Caption         =   "Monitor applications."
         Enabled         =   0   'False
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
         Index           =   2
         Left            =   270
         TabIndex        =   35
         ToolTipText     =   "Monitor process & applications."
         Top             =   3360
         Width           =   2355
      End
      Begin VB.CheckBox SecOpt2 
         Caption         =   "Monitor system resource."
         Enabled         =   0   'False
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
         Index           =   1
         Left            =   270
         TabIndex        =   34
         ToolTipText     =   "Monitor print activity."
         Top             =   3015
         Width           =   2355
      End
      Begin VB.CheckBox SecOpt2 
         Caption         =   "Monitor print activity."
         Enabled         =   0   'False
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
         Index           =   0
         Left            =   270
         TabIndex        =   33
         ToolTipText     =   "Monitor print activity."
         Top             =   2670
         Width           =   2355
      End
      Begin VB.CheckBox Sec2Opt1 
         Caption         =   "Protect Desktop Wallpaper."
         Enabled         =   0   'False
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
         Index           =   0
         Left            =   270
         TabIndex        =   31
         ToolTipText     =   "Wallpaper will be restore when Logout or Lock."
         Top             =   1440
         Width           =   2355
      End
      Begin VB.CheckBox Sec2Opt1 
         Caption         =   "Protect Desktop Icons."
         Enabled         =   0   'False
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
         Index           =   1
         Left            =   270
         TabIndex        =   32
         ToolTipText     =   "Protect desktop from change."
         Top             =   1770
         Width           =   2355
      End
      Begin VB.CheckBox SecOpt1 
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
         Height          =   210
         Index           =   1
         Left            =   270
         TabIndex        =   29
         ToolTipText     =   "Disable the user from using CTL+ALT+DEL"
         Top             =   727
         Width           =   2085
      End
      Begin VB.CheckBox SecOpt1 
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
         Height          =   315
         Index           =   2
         Left            =   270
         TabIndex        =   30
         ToolTipText     =   "Disable the famous Regedit"
         Top             =   1005
         Width           =   2355
      End
      Begin VB.Image GenHdrIco 
         Height          =   240
         Index           =   3
         Left            =   135
         Picture         =   "FrmMain.frx":A13D
         Top             =   2325
         Width           =   240
      End
      Begin VB.Label SecHdr 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "PC Monitoring"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   450
         TabIndex        =   65
         Top             =   2355
         Width           =   1335
      End
      Begin VB.Label SecHdr 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Security Configuration"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   64
         Top             =   75
         Width           =   2190
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : FrmMain
'    Project    : CafeBonzerAG
'
'    Description: Configuration Form
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Public ObjPol As New ClsPolicies
Private Const LngDriveMax = 14

Private Sub Form_Load()
'On Error GoTo ErrInt
        
 '[ Preliminaries ]'
    FrmMain.Caption = "Configuration - " & SysInfoGetName
    If BlnAppFirstTime = True Then CbtMain(1).Enabled = False
       
 '[ Load All Configuration ]'
    GenNetName = SysInfoGetName
    GenNetPort = SettingGet("NetServerPort", 8180)
    GenNetIP = SettingGet("NetServerIp", "192.168.0.1")
    GenOpt1(0) = SettingGet("GenAdminPassSyncServer", 0)
    GenPass1 = SettingGet("GenAdminPass")
    GenPass2 = GenPass1
    GenWelcome = SettingGet("TickMsgWelcome", CStrTickMsgWelcome)
    GenOpt2(0) = SettingGet("AppAutoStart", 1)
    SecOpt2(0) = SettingGet("SysMonPrinter", 1)
    SecOpt2(1) = SettingGet("SysMonResource", 1)
    SecOpt2(2) = SettingGet("SysMonApp", 1)
    SecOpt2(3) = SettingGet("SysMonTraffic", 1)
    
    Call LoadSecurityItems
    ObjPol.DisProgEnum
    
    SecOpt1(0).Value = SettingGet("SysAutoLock", 1)
    SecOpt1(1).Value = SettingGet("SysDisCad", 1)
    SecOpt1(2).Value = ObjPol.Misc_DisableRegedit
    'Sec2Opt1(0).Value = SettingGet("SysSecWallpaper", 1)
    'Sec2Opt1(1).Value = SettingGet("SysSecDesktop", 0)
    
    AprOpt1(0).Value = SettingGet("TickGuiDisable", 0)
    AprTickSize = SettingGet("TickGuiSize", 5)
    'AprFntChoose.Caption = GetShortStr(SettingGet("TickGuiFont", "Verdana"))
    AprClrPick(0).Color = SettingGet("TickGuiForeColor", &H80000012)
    AprClrPick(1).Color = SettingGet("TickGuiBackColor", &H8000000F)
    
    AprLckCb.ListIndex = SettingGet("AppLockMode", 0)
    AprLckPos(SettingGet("AppLockPos", 0)).BackColor = &H808080
Exit Sub

ErrInt:
    AppErrorLog Err, "FrmMain | Form_Load"
End Sub

Private Sub BtnPage_Click(Index As Integer)
    For Each VsGuiButton In BtnPage
        If VsGuiButton.Index <> Index Then VsGuiButton.Value = False
    Next
    Page(Index).ZOrder 0
    CfgBannerLbl.Caption = BtnPage(Index).Tag
End Sub

Private Sub AprLckPos_Click(Index As Integer)
    For Each PictureBox In AprLckPos
        PictureBox.BackColor = vbWhite
    Next
    AprLckPos(Index).BackColor = &H808080
End Sub

Private Sub CbtMain_Click(Index As Integer)
    Select Case Index
    Case 0
        Call SaveAllSetting
    Case 1
        Unload Me
    End Select
End Sub

Public Sub SaveAllSetting()
'On Error GoTo ErrInt
        
     ' { Check Input Error } '
        If GenNetPort = "" Then GenNetPort = 8180
        If GenNetIP = "" Then MsgBox "Please enter server IP Address !", vbOKOnly, "IP Number": Exit Sub
        If Len(GenNetIP) < 9 Then MsgBox "Please use the right IP number !": Exit Sub
        If ValidateIP(GenNetIP) = False Then MsgBox "Please enter your IP number in numerical form !": Exit Sub
        If GenPass1 = "" Or GenPass2 = "" Then MsgBox "Please enter your password !": Exit Sub
        If GenPass1 <> GenPass2 Then MsgBox "Please retype your password !": Exit Sub
        If IsNumeric(AprTickSize) = False Then AprTickSize = 5
        If GenWelcome = "" Then GenWelcome = StrTickMsgWelcome
        
     ' + SAVING ALL SETTINGS -----------------------------------------------------
     ' { GENERAL SETTINGS } '
        SysInfoSetName Trim$(GenNetName)
        SettingSave "NetServerPort", Trim$(GenNetPort)
        SettingSave "NetServerIp", Trim$(GenNetIP)
        SettingSave "GenAdminPassSyncServer", GenOpt1(0)
        SettingSave "GenAdminPass", Trim$(GenPass2)
        SettingSave "TickMsgWelcome", GenWelcome
        SettingSave "AppAutoStart", GenOpt2(0)
        
        SettingSave "SysMonPrinter", SecOpt2(0)
        SettingSave "SysMonResource", SecOpt2(1)
        SettingSave "SysMonApp", SecOpt2(2)
        SettingSave "SysMonTraffic", SecOpt2(3)
        
     ' { SECURITY SETTING }
        Call SaveSecurityItems
        
        SettingSave "SysAutoLock", SecOpt1(0).Value
        SettingSave "SysDisCad", SecOpt1(1)
        ObjPol.Misc_DisableRegedit = SecOpt1(2).Value
        'SettingSave "SysSecWallpaper", Sec2Opt1(0).Value
        'SettingSave "SysSecDesktop", Sec2Opt1(1).Value

       '- APPEARANCE
     ' { No Tray Ticker } '
        SettingSave "TickGuiDisable", AprOpt1(0).Value
        SettingSave "TickGuiSize", AprTickSize
        'SettingSave "TickGuiFont", FrmTicker.picTicker.FontName
        'SettingSave "TickGuiFontSize", FrmTicker.picTicker.FontSize
        SettingSave "TickGuiForeColor", AprClrPick(0).Color
        SettingSave "TickGuiBackColor", AprClrPick(1).Color
        
        SettingSave "AppLockMode", AprLckCb.ListIndex
        For Each PictureBox In AprLckPos
            If PictureBox.BackColor = &H808080 Then
                SettingSave "AppLockPos", PictureBox.Index
            End If
        Next
            
    ' { First First Time Flag } '
        SettingSave "AppFirstTime", "APPNOFIRST"
        
        If BlnAppFirstTime = False Then
            Call ActivateSetting
        Else
            MsgBox "Please restart your computer!"
        End If
        Unload FrmMain
Exit Sub

ErrInt:
    AppErrorLog Err, "FrmMain | SaveAllSetting"
End Sub

Public Sub LoadSecurityItems()
On Error GoTo ErrInt
    Dim NodExp As Object, NodNet As Object, NodSys As Object
    Dim NodDos As Object, NodDrv As Object, ObjTopNode As Object
    Dim LngIdxA As Long, LngIdxB As Long, LngRet As Long, LngPolCount As Long
    Dim StrDrvName As String, LngDrvValue As Long, StrCatKey As String
    
    Set ObjPolicies = ObjPol
    Set NodExp = SecTree.AddFolder(, , "Explorer", OptionTreeFolder)
    Set NodNet = SecTree.AddFolder(, , "Network", OptionTreeFolder)
    Set NodSys = SecTree.AddFolder(, , "System", OptionTreeFolder)
    Set NodDos = SecTree.AddFolder(, , "Dos", OptionTreeFolder)
    Set NodDrv = SecTree.AddFolder(, , "Hide Drives", OptionTreeFolder)
        
 ' { LOAD POLICIES SETTING }
    With SecTree
        For LngIdxA = 1 To 4
            StrCatKey = Choose(LngIdxA, "EXP", "NET", "SYS", "DOS")
            LngPolCount = ObjPol.GetPoliciesCount(LngIdxA)  'Choose(LngIdxA, 14, 6, 13, 2)
            Set ObjTopNode = Choose(LngIdxA, NodExp, NodNet, NodSys, NodDos)
            For LngIdxB = 1 To LngPolCount
                LngRet = ObjPol.GetPolicies2(LngIdxA, LngIdxB)
                .AddCheck StrCatKey & LngIdxB, ObjTopNode, ObjPol.GetPoliciesDesc(LngIdxA, LngIdxB), LngRet
            Next
        Next
        .ExpandAll
    End With


 ' { LOAD HIDE DRIVES }
 '   Info !
 '   ------------------------
 '      4096 = drive m
 '      m is 13th drive
    With SecTree
        LngRet = ObjPol.Misc_DriveHide
        
        For LngIdxA = 1 To LngDriveMax
            StrDrvName = "Drive " & Mid$("ABCDEFGHIJKLMNOPQRSTUVWXYZ", LngIdxA, 1)
            .AddCheck "DRV" & LngIdxA, NodDrv, StrDrvName, OptionTreeCheckNone
        Next
        
        For LngIdxB = LngDriveMax To 1 Step -1
            LngDrvValue = 2 ^ (LngIdxB - 1)
            If LngRet >= LngDrvValue Then
                .Value("DRV" & LngIdxB) = OptionTreeCheckFull
                LngRet = LngRet - LngDrvValue
            End If
        Next
        .ExpandAll
    End With
Exit Sub

ErrInt:
    AppErrorLog Err, "FrmMain | LoadSecurityItems"
End Sub

Public Sub ActivateSetting()
On Error GoTo ErrInt

    If GenOpt2(0) = 1 Then
        SaveString LngEnvRegistryRoot, CStrAutoStartPath, "Component", "CbAg.exe"
    Else
        DeleteValue LngEnvRegistryRoot, CStrAutoStartPath, "Component"
    End If

    FrmTicker.ForeColor = AprClrPick(0).Color
    FrmTicker.BackColor = AprClrPick(1).Color
    If AprOpt1(0).Value = 1 Then
        FrmTicker.TmrCheck.Enabled = False
        Call TickerStop
        Call TrayStart
    Else
        Call TrayRemove
        Call TickerStart(FrmTicker)
        If MdlTicker.LngIconCount <> AprTickSize Then
            MdlTicker.LngIconCount = AprTickSize
            Call TickerNormal
        End If
    End If

Exit Sub

ErrInt:
    AppErrorLog Err, "ActivateSetting | FrmMain"
End Sub

Public Sub SaveSecurityItems()
'On Error GoTo ErrInt
    Dim StrCatKey As String, LngPolCount As Long, LngDrvValue As Long
    Dim LngIdxA As Long, LngIdxB As Long
    
 ' { SAVE POLICIES SETTING }
    With SecTree
        For LngIdxA = 1 To 4
            StrCatKey = Choose(LngIdxA, "EXP", "NET", "SYS", "DOS")
            LngPolCount = ObjPol.GetPoliciesCount(LngIdxA)
            For LngIdxB = 1 To LngPolCount
                If .Value(StrCatKey & LngIdxB) = OptionTreeCheckFull Then
                    ObjPol.SetPolicies LngIdxA, LngIdxB, True
                Else
                    ObjPol.SetPolicies LngIdxA, LngIdxB, False
                End If
            Next
        Next
    End With

 ' { SAVE HIDE DRIVES }
    With SecTree
        For LngIdxA = 1 To LngDriveMax
            If .Value("DRV" & LngIdxA) = OptionTreeCheckFull Then
                LngDrvValue = LngDrvValue + (2 ^ LngIdxA)
            End If
        Next
        ObjPol.Misc_DriveHide = LngDrvValue
    End With
Exit Sub

ErrInt:
    AppErrorLog Err, "SaveSecurityItems | FrmMain"
End Sub
