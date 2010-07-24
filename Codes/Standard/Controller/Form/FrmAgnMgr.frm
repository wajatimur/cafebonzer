VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAgnMgr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CafeBonzer - Agent Manager"
   ClientHeight    =   6975
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   10305
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmAgnMgr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   10305
   StartUpPosition =   3  'Windows Default
   Begin CafeBonzer.PageHolder MainPhld 
      Align           =   2  'Align Bottom
      Height          =   945
      Left            =   0
      TabIndex        =   36
      Top             =   6030
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   1667
      HldrStyle       =   2
      HldrTxt         =   "Control Option"
      HldrTxtClr      =   4210752
      HldrLne         =   0   'False
      PageHeight      =   945
      Begin CafeBonzer.PageDock MainPdck 
         Height          =   615
         Left            =   10020
         TabIndex        =   42
         Top             =   345
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   1085
         HldrBtnPos      =   0
         HldrLne         =   -1  'True
         PageState       =   1
         PageWidth       =   10305
         Begin CafeBonzer.XpButton AgsMnu 
            Height          =   420
            Index           =   3
            Left            =   1815
            TabIndex        =   47
            Top             =   90
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
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
            MICON           =   "FrmAgnMgr.frx":058A
            PICN            =   "FrmAgnMgr.frx":05A6
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin CafeBonzer.XpButton AgsMnu 
            Height          =   420
            Index           =   0
            Left            =   465
            TabIndex        =   46
            Top             =   90
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
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
            MICON           =   "FrmAgnMgr.frx":0B40
            PICN            =   "FrmAgnMgr.frx":0B5C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin CafeBonzer.XpButton AgsMnu 
            Height          =   420
            Index           =   1
            Left            =   915
            TabIndex        =   45
            Top             =   90
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
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
            MICON           =   "FrmAgnMgr.frx":10F6
            PICN            =   "FrmAgnMgr.frx":1112
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin CafeBonzer.XpButton AgsMnu 
            Height          =   420
            Index           =   2
            Left            =   1365
            TabIndex        =   44
            Top             =   90
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
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
            MICON           =   "FrmAgnMgr.frx":16AC
            PICN            =   "FrmAgnMgr.frx":16C8
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin CafeBonzer.XpButton AgsCmd 
            Height          =   420
            Index           =   0
            Left            =   9270
            TabIndex        =   43
            Top             =   90
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   741
            TX              =   "Send"
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
            MICON           =   "FrmAgnMgr.frx":1C62
            PICN            =   "FrmAgnMgr.frx":1C7E
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
      Begin CafeBonzer.XpButton CptMnu 
         Height          =   480
         Index           =   0
         Left            =   90
         TabIndex        =   41
         ToolTipText     =   "UnLock"
         Top             =   405
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   847
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
         MICON           =   "FrmAgnMgr.frx":2218
         PICN            =   "FrmAgnMgr.frx":2234
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CafeBonzer.XpButton CptMnu 
         Height          =   480
         Index           =   1
         Left            =   615
         TabIndex        =   40
         ToolTipText     =   "Lock"
         Top             =   405
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   847
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
         MICON           =   "FrmAgnMgr.frx":42B6
         PICN            =   "FrmAgnMgr.frx":42D2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CafeBonzer.XpButton CptMnu 
         Height          =   480
         Index           =   2
         Left            =   1125
         TabIndex        =   39
         ToolTipText     =   "Shutdown"
         Top             =   405
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   847
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
         MICON           =   "FrmAgnMgr.frx":6554
         PICN            =   "FrmAgnMgr.frx":6570
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CafeBonzer.XpButton CptMnu 
         Height          =   480
         Index           =   3
         Left            =   1650
         TabIndex        =   38
         ToolTipText     =   "Restart"
         Top             =   405
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   847
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
         MICON           =   "FrmAgnMgr.frx":85F2
         PICN            =   "FrmAgnMgr.frx":860E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CafeBonzer.XpButton CptMnu 
         Height          =   480
         Index           =   4
         Left            =   2175
         TabIndex        =   37
         ToolTipText     =   "Close"
         Top             =   405
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   847
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
         MICON           =   "FrmAgnMgr.frx":A690
         PICN            =   "FrmAgnMgr.frx":A6AC
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
   Begin VB.PictureBox MainBnr 
      BackColor       =   &H00FFFFFF&
      Height          =   810
      Left            =   3045
      Picture         =   "FrmAgnMgr.frx":CE5E
      ScaleHeight     =   750
      ScaleWidth      =   7170
      TabIndex        =   26
      Top             =   75
      Width           =   7230
      Begin VB.Label MainBnrCap 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Send general command to agent."
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   840
         TabIndex        =   28
         Top             =   360
         Width           =   2865
      End
      Begin VB.Label MainBnrLbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "General Control"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   840
         TabIndex        =   27
         Top             =   60
         Width           =   2070
      End
   End
   Begin CafeBonzer.Line3D MainLne 
      Height          =   45
      Left            =   0
      TabIndex        =   1
      Top             =   -15
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   79
      horizon         =   -1  'True
   End
   Begin MSComctlLib.ListView LstVw1 
      Height          =   5925
      Left            =   30
      TabIndex        =   2
      Top             =   75
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   10451
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Station"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Status"
         Object.Width           =   1940
      EndProperty
   End
   Begin VB.Frame Pages 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5040
      Index           =   0
      Left            =   3060
      TabIndex        =   0
      Top             =   930
      Width           =   7230
      Begin VB.ComboBox GcnCmdCB 
         Height          =   315
         Left            =   285
         TabIndex        =   35
         Text            =   "block:1"
         Top             =   3495
         Width           =   6495
      End
      Begin VB.ListBox GcnList 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         ForeColor       =   &H00404040&
         Height          =   2400
         ItemData        =   "FrmAgnMgr.frx":D73D
         Left            =   300
         List            =   "FrmAgnMgr.frx":D73F
         TabIndex        =   32
         Top             =   495
         Width           =   6525
      End
      Begin VB.Image GcnBtnClr 
         Height          =   240
         Left            =   6885
         Picture         =   "FrmAgnMgr.frx":D741
         ToolTipText     =   "Send Command"
         Top             =   495
         Width           =   240
      End
      Begin VB.Image GcnCmdBtn 
         Height          =   240
         Left            =   6840
         Picture         =   "FrmAgnMgr.frx":DCCB
         ToolTipText     =   "Send Command"
         Top             =   3525
         Width           =   240
      End
      Begin VB.Label GcnHdr 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Custom Command"
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
         Left            =   75
         TabIndex        =   34
         Top             =   3090
         Width           =   7080
      End
      Begin VB.Label GcnHdr 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Summary"
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
         Left            =   60
         TabIndex        =   33
         Top             =   105
         Width           =   7080
      End
   End
   Begin VB.Frame Pages 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5040
      Index           =   1
      Left            =   3060
      TabIndex        =   3
      Top             =   930
      Width           =   7230
      Begin VB.CheckBox GenOpt2 
         Appearance      =   0  'Flat
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
         Left            =   3945
         TabIndex        =   23
         ToolTipText     =   "Monitor print activity."
         Top             =   435
         Width           =   2850
      End
      Begin VB.CheckBox GenOpt2 
         Appearance      =   0  'Flat
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
         Left            =   3945
         TabIndex        =   22
         ToolTipText     =   "Monitor print activity."
         Top             =   780
         Width           =   2850
      End
      Begin VB.CheckBox GenOpt2 
         Appearance      =   0  'Flat
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
         Left            =   3945
         TabIndex        =   21
         ToolTipText     =   "Monitor process & applications."
         Top             =   1125
         Width           =   2850
      End
      Begin VB.CheckBox GenOpt2 
         Appearance      =   0  'Flat
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
         Left            =   3945
         TabIndex        =   20
         ToolTipText     =   "Monitor network traffic."
         Top             =   1470
         Width           =   2850
      End
      Begin VB.CheckBox GenOpt1 
         Appearance      =   0  'Flat
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
         Left            =   3960
         TabIndex        =   18
         ToolTipText     =   "Automatic start CafeBonzer"
         Top             =   2445
         Width           =   2850
      End
      Begin VB.TextBox GenWelcome 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   4155
         TabIndex        =   17
         Text            =   ":: CafeBonzer Agent R1 ::"
         Top             =   3150
         Width           =   2805
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
         Left            =   1605
         PasswordChar    =   "l"
         TabIndex        =   13
         Top             =   3135
         Width           =   1700
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
         Left            =   1605
         PasswordChar    =   "l"
         TabIndex        =   12
         Top             =   2730
         Width           =   1700
      End
      Begin VB.CheckBox GenOpt1 
         Appearance      =   0  'Flat
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
         Left            =   345
         TabIndex        =   11
         ToolTipText     =   "Retrive default password from server when windows start."
         Top             =   2325
         Width           =   2850
      End
      Begin VB.TextBox GenNetName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1635
         TabIndex        =   6
         Text            =   "Cake"
         Top             =   495
         Width           =   1650
      End
      Begin VB.TextBox GenNetIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1635
         TabIndex        =   5
         Top             =   1290
         Width           =   1650
      End
      Begin VB.TextBox GenNetPort 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1635
         TabIndex        =   4
         Text            =   "56266"
         Top             =   885
         Width           =   1650
      End
      Begin VB.Label GenHdr 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Miscelaneous"
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
         Left            =   3795
         TabIndex        =   25
         Top             =   2040
         Width           =   3195
      End
      Begin VB.Label GenHdr 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Pc monitoring"
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
         Left            =   3795
         TabIndex        =   24
         Top             =   105
         Width           =   3195
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
         Left            =   3975
         TabIndex        =   19
         Top             =   2865
         Width           =   1455
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
         Left            =   240
         TabIndex        =   16
         Top             =   1905
         Width           =   3195
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
         Left            =   345
         TabIndex        =   15
         Top             =   3180
         Width           =   855
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
         Left            =   345
         TabIndex        =   14
         Top             =   2775
         Width           =   855
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
         Left            =   150
         TabIndex        =   10
         Top             =   105
         Width           =   3195
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
         Left            =   315
         TabIndex        =   9
         Top             =   525
         Width           =   1230
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
         Left            =   315
         TabIndex        =   8
         Top             =   1320
         Width           =   750
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
         Left            =   330
         TabIndex        =   7
         Top             =   915
         Width           =   915
      End
   End
   Begin VB.Frame Pages 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5040
      Index           =   2
      Left            =   3060
      TabIndex        =   29
      Top             =   930
      Width           =   7230
   End
   Begin VB.Frame Pages 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5040
      Index           =   4
      Left            =   3060
      TabIndex        =   31
      Top             =   930
      Width           =   7230
   End
   Begin VB.Frame Pages 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5040
      Index           =   3
      Left            =   3060
      TabIndex        =   30
      Top             =   930
      Width           =   7230
   End
   Begin VB.Menu Mnu1 
      Caption         =   "Menu"
      Begin VB.Menu Mnu1Rfsh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu Mnu1Close 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu Mnu2 
      Caption         =   "Select"
      Begin VB.Menu Mnu2Sel 
         Caption         =   "Select All"
         Index           =   0
      End
      Begin VB.Menu Mnu2Sel 
         Caption         =   "DeSelect All"
         Index           =   1
      End
      Begin VB.Menu Mnu2Sel 
         Caption         =   "Select Unused"
         Index           =   2
      End
      Begin VB.Menu Mnu2Sel 
         Caption         =   "Select Used"
         Index           =   3
      End
      Begin VB.Menu Mnu2Sel 
         Caption         =   "Select Unlock"
         Index           =   4
      End
   End
End
Attribute VB_Name = "FrmAgnMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : FrmAgnMgr
'    Project    : CafeBonzer
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Private sBnrLabel(1 To 4) As String


''*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
'' FUNCTION
''
''
''*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Private Sub Form_Load()
    sBnrLabel(1) = "General Settings | Main settings for agent."
    sBnrLabel(2) = "Appearance Settings | Set appearance for agent."
    sBnrLabel(3) = "Security Settings 1 | Protect your system."
    sBnrLabel(4) = "Security Settings 2 | More power to protect."
    
    Call LoadAgents
End Sub

Private Sub MainPdck_PageFliped(ByVal Flipped As Boolean)
    If Flipped = False Then
        MainBnrLbl = "Agent Configuration"
        MainBnrCap = "General Settings | Main settings for agent"
        Pages(1).ZOrder 0
    Else
        MainBnrLbl = "General Control"
        MainBnrCap = "Send general command to agent."
        Pages(0).ZOrder 0
    End If
End Sub

Private Sub MainPhld_FrameExpand(ByVal Expanded As Boolean)
    LstVw1.Height = MainPhld.Top - 100
End Sub

Private Sub AgsMnu_Click(Index As Integer)
    MainBnrCap = sBnrLabel(Index + 1)
    Pages(Index + 1).ZOrder 0
End Sub

Private Sub GcnCmdBtn_Click()
    Dim s_Cmd2Send As String
    GcnCmdCB = Trim$(GcnCmdCB)
    If GcnCmdCB <> "" Then
        If Left$(GcnCmdCB, 2) <> "//" Then
            s_Cmd2Send = GcnCmdCB
        Else
            '!!!!!! HELP
            s_Cmd2Send = "" & GcnCmdCB
        End If
        Call SendSel(s_Cmd2Send)
    End If
End Sub

Private Sub GcnCmdCB_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call GcnCmdBtn_Click
    End If
End Sub



''*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
'' MENU
''
''
''*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Private Sub Mnu1Close_Click()
    Unload Me
End Sub



''*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
'' FUNCTION
''
''
''*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Private Function LoadAgents() As Long
    Dim mLv As ListView, CTmpAgent As ClsAgent
    Dim nItm As ListItem, LngIdxA As Long
    Set mLv = FrmMain.ListView
    
    If mLv.ListItems.Count = 0 Then
        Call GcnSmr("No agent loaded !")
        Exit Function
    End If
    
    For LngIdxA = 1 To UniAgents.Count
        Set CTmpAgent = UniAgents.Agents(LngIdxA)
        Set nItm = LstVw1.ListItems.Add(, CTmpAgent.AgentName, CTmpAgent.AgentName)
    Next
    
    Call GcnSmr(LngIdxA & " agent loaded !")
End Function

Private Function SendSel(sCommand) As Long
    Dim l_AgentSel As Long, LngIdxA As Long
    Dim LngItemCnt As Long
    
    If LstVw1.ListItems.Count = 0 Then
        SendSel = -1
        Exit Function
    End If
    
    LngItemCnt = LstVw1.ListItems.Count
    For LngIdxA = 1 To LngItemCnt
        If LstVw1.ListItems(LngIdxA).Selected = True Then
            l_AgentSel = l_AgentSel + 1
            UniAgents(LstVw1.ListItems(LngIdxA).Text).Commands (sCommand)
        End If
    Next
    
    If l_AgentSel = 0 Then SendSel = -1
End Function

Private Sub GcnSmr(Text)
    If Trim$(Text) = "" Then Exit Sub
    GcnList.AddItem ">> " & Text
    GcnList.ListIndex = GcnList.NewIndex
End Sub

