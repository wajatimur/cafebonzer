VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmStat 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CafeBonzer - Statistic"
   ClientHeight    =   7680
   ClientLeft      =   255
   ClientTop       =   1695
   ClientWidth     =   9780
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmStat.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   9780
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame StatFme 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Option"
      ForeColor       =   &H00FF0000&
      Height          =   1815
      Left            =   60
      TabIndex        =   2
      Top             =   5475
      Width           =   9675
      Begin CafeBonzer.Line3D StatLine 
         Height          =   1620
         Index           =   1
         Left            =   2475
         TabIndex        =   56
         Top             =   150
         Width           =   45
         _ExtentX        =   79
         _ExtentY        =   2858
         horizon         =   0   'False
      End
      Begin VB.ComboBox cbHari 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   960
         TabIndex        =   49
         Text            =   "cbHari"
         Top             =   1260
         Width           =   1380
      End
      Begin VB.ComboBox cbBulan 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   810
         Width           =   1380
      End
      Begin VB.ComboBox cbTahun 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   345
         Width           =   1380
      End
      Begin CafeBonzer.XpButton StatBtn 
         Height          =   435
         Left            =   9045
         TabIndex        =   57
         ToolTipText     =   "Close statistic."
         Top             =   1275
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   767
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmStat.frx":058A
         PICN            =   "FrmStat.frx":05A6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CafeBonzer.XpButton StatSesBtn 
         Height          =   360
         Left            =   4590
         TabIndex        =   58
         Top             =   570
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   635
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
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmStat.frx":0B40
         PICN            =   "FrmStat.frx":0B5C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label StatLbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Date :"
         Height          =   195
         Index           =   2
         Left            =   315
         TabIndex        =   50
         Top             =   1275
         Width           =   570
      End
      Begin VB.Label StatLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Current Session Date :"
         Height          =   240
         Index           =   3
         Left            =   2685
         TabIndex        =   47
         Top             =   270
         Width           =   2055
         WordWrap        =   -1  'True
      End
      Begin VB.Label CurSession 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12/12/2000"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2880
         TabIndex        =   46
         Top             =   600
         Width           =   1635
      End
      Begin VB.Label StatLbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Month :"
         Height          =   240
         Index           =   1
         Left            =   225
         TabIndex        =   6
         Top             =   855
         Width           =   660
      End
      Begin VB.Label StatLbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Year :"
         Height          =   240
         Index           =   0
         Left            =   225
         TabIndex        =   4
         Top             =   390
         Width           =   660
      End
   End
   Begin MSComctlLib.StatusBar StatBar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   7335
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14631
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab StatTab 
      Height          =   5445
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   9604
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   12632256
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Monthly\Daily Overview"
      TabPicture(0)   =   "FrmStat.frx":10F6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "GenHdr(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "GenLbl(4)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "GenLbl(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblJualanPos"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblJualanPc"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "GenLbl(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "GenLbl(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblUntung"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblJualan"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblModal"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "GenLbl(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "GenHdr(1)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "GenLbl(6)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblPelanggan"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblPungut"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "GenLbl(5)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "GenHdr(2)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Graf1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "StatLine(0)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "Stations"
      TabPicture(1)   =   "FrmStat.frx":1112
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Lv1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Customers"
      TabPicture(2)   =   "FrmStat.frx":112E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Lv2"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Sales Record"
      TabPicture(3)   =   "FrmStat.frx":114A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "SlsLbl(0)"
      Tab(3).Control(1)=   "SlsBtn"
      Tab(3).Control(2)=   "Lv3"
      Tab(3).Control(3)=   "cbHari2"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Services Record"
      TabPicture(4)   =   "FrmStat.frx":1166
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "SrvLbl(1)"
      Tab(4).Control(1)=   "SrvLbl(0)"
      Tab(4).Control(2)=   "SrvBtn"
      Tab(4).Control(3)=   "PosLV1"
      Tab(4).Control(4)=   "PosCmbItems"
      Tab(4).Control(5)=   "PosCmbDate"
      Tab(4).ControlCount=   6
      Begin CafeBonzer.Line3D StatLine 
         Height          =   5025
         Index           =   0
         Left            =   4545
         TabIndex        =   48
         Top             =   360
         Width           =   45
         _ExtentX        =   79
         _ExtentY        =   8864
         horizon         =   0   'False
      End
      Begin VB.PictureBox Graf1 
         BackColor       =   &H00808080&
         Height          =   2940
         Left            =   5055
         ScaleHeight     =   2880
         ScaleWidth      =   4170
         TabIndex        =   28
         Top             =   1005
         Width           =   4230
         Begin VB.PictureBox GrafDock 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   -45
            ScaleHeight     =   420
            ScaleWidth      =   4245
            TabIndex        =   29
            Top             =   2565
            Width           =   4245
            Begin VB.Label GrafDay 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "A"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0C0&
               Height          =   330
               Index           =   0
               Left            =   435
               TabIndex        =   36
               Top             =   -30
               Width           =   330
            End
            Begin VB.Label GrafDay 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "I"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0C0&
               Height          =   330
               Index           =   1
               Left            =   960
               TabIndex        =   35
               Top             =   -30
               Width           =   330
            End
            Begin VB.Label GrafDay 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "S"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0C0&
               Height          =   330
               Index           =   2
               Left            =   1455
               TabIndex        =   34
               Top             =   -30
               Width           =   330
            End
            Begin VB.Label GrafDay 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "R"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0C0&
               Height          =   330
               Index           =   3
               Left            =   1980
               TabIndex        =   33
               Top             =   -30
               Width           =   330
            End
            Begin VB.Label GrafDay 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "K"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0C0&
               Height          =   330
               Index           =   4
               Left            =   2520
               TabIndex        =   32
               Top             =   -15
               Width           =   330
            End
            Begin VB.Label GrafDay 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "J"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0C0&
               Height          =   330
               Index           =   5
               Left            =   2985
               TabIndex        =   31
               Top             =   -15
               Width           =   330
            End
            Begin VB.Label GrafDay 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "S"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0C0&
               Height          =   330
               Index           =   6
               Left            =   3495
               TabIndex        =   30
               Top             =   -15
               Width           =   330
            End
         End
         Begin MSComctlLib.ProgressBar Bar1 
            Height          =   2460
            Index           =   0
            Left            =   360
            TabIndex        =   37
            Top             =   255
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   4339
            _Version        =   393216
            Appearance      =   0
            Orientation     =   1
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar Bar1 
            Height          =   2460
            Index           =   1
            Left            =   870
            TabIndex        =   38
            Top             =   255
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   4339
            _Version        =   393216
            Appearance      =   0
            Orientation     =   1
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar Bar1 
            Height          =   2460
            Index           =   2
            Left            =   1395
            TabIndex        =   39
            Top             =   255
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   4339
            _Version        =   393216
            Appearance      =   0
            Orientation     =   1
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar Bar1 
            Height          =   2460
            Index           =   3
            Left            =   1920
            TabIndex        =   40
            Top             =   255
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   4339
            _Version        =   393216
            Appearance      =   0
            Orientation     =   1
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar Bar1 
            Height          =   2460
            Index           =   4
            Left            =   2415
            TabIndex        =   41
            Top             =   255
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   4339
            _Version        =   393216
            Appearance      =   0
            Orientation     =   1
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar Bar1 
            Height          =   2460
            Index           =   5
            Left            =   2925
            TabIndex        =   42
            Top             =   255
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   4339
            _Version        =   393216
            Appearance      =   0
            Orientation     =   1
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar Bar1 
            Height          =   2460
            Index           =   6
            Left            =   3420
            TabIndex        =   43
            Top             =   255
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   4339
            _Version        =   393216
            Appearance      =   0
            Orientation     =   1
            Scrolling       =   1
         End
         Begin VB.Label GrafHigh 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   60
            TabIndex        =   44
            Top             =   45
            Width           =   765
         End
      End
      Begin VB.ComboBox cbHari2 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   -74205
         TabIndex        =   12
         Top             =   525
         Width           =   1620
      End
      Begin VB.ComboBox PosCmbDate 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   -74205
         TabIndex        =   9
         Top             =   525
         Width           =   1620
      End
      Begin VB.ComboBox PosCmbItems 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   -70755
         TabIndex        =   8
         Top             =   510
         Width           =   1620
      End
      Begin MSComctlLib.ListView PosLV1 
         Height          =   4305
         Left            =   -74835
         TabIndex        =   7
         Top             =   975
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   7594
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648384
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Group"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Transaction ID"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Item"
            Object.Width           =   5115
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Qty"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Total"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView Lv3 
         Height          =   4305
         Left            =   -74835
         TabIndex        =   13
         Top             =   975
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   7594
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648384
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Station"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Customer"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "In"
            Object.Width           =   2258
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Out"
            Object.Width           =   2258
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Paid"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView Lv2 
         Height          =   4725
         Left            =   -74835
         TabIndex        =   15
         Top             =   540
         Width           =   9345
         _ExtentX        =   16484
         _ExtentY        =   8334
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648384
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Customer"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Visits"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Last Visit"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Total Time"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Paid"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComctlLib.ListView Lv1 
         Height          =   4725
         Left            =   -74835
         TabIndex        =   16
         Top             =   540
         Width           =   9345
         _ExtentX        =   16484
         _ExtentY        =   8334
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648384
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Station"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Total Time"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Profit"
            Object.Width           =   2646
         EndProperty
      End
      Begin CafeBonzer.XpButton SrvBtn 
         Height          =   360
         Left            =   -69060
         TabIndex        =   59
         Top             =   495
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   635
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
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmStat.frx":1182
         PICN            =   "FrmStat.frx":119E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CafeBonzer.XpButton SlsBtn 
         Height          =   360
         Left            =   -65910
         TabIndex        =   60
         Top             =   465
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   635
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
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmStat.frx":1738
         PICN            =   "FrmStat.frx":1754
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label GenHdr 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Today's Collection"
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
         Height          =   255
         Index           =   2
         Left            =   210
         TabIndex        =   55
         Top             =   3660
         Width           =   4140
      End
      Begin VB.Label GenLbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Collection :"
         Height          =   195
         Index           =   5
         Left            =   585
         TabIndex        =   54
         Top             =   4125
         Width           =   1845
      End
      Begin VB.Label lblPungut 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RM 0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2550
         TabIndex        =   53
         Top             =   4095
         Width           =   1470
      End
      Begin VB.Label lblPelanggan 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RM 0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2550
         TabIndex        =   52
         Top             =   4590
         Width           =   1470
      End
      Begin VB.Label GenLbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Customer :"
         Height          =   165
         Index           =   6
         Left            =   585
         TabIndex        =   51
         Top             =   4620
         Width           =   1860
      End
      Begin VB.Label GenHdr 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Daily Average Graft"
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
         Height          =   255
         Index           =   1
         Left            =   4830
         TabIndex        =   45
         Top             =   600
         Width           =   4635
      End
      Begin VB.Label GenLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Monthly Overhead :"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   27
         Top             =   1050
         Width           =   1815
      End
      Begin VB.Label lblModal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RM 0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2295
         TabIndex        =   26
         Top             =   1020
         Width           =   1725
      End
      Begin VB.Label lblJualan 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RM 0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2295
         TabIndex        =   25
         Top             =   1470
         Width           =   1725
      End
      Begin VB.Label lblUntung 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RM 0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2295
         TabIndex        =   24
         Top             =   1950
         Width           =   1725
      End
      Begin VB.Label GenLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Monthly Sales :"
         Height          =   195
         Index           =   1
         Left            =   345
         TabIndex        =   23
         Top             =   1500
         Width           =   1815
      End
      Begin VB.Label GenLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Monthly Profit :"
         Height          =   195
         Index           =   2
         Left            =   345
         TabIndex        =   22
         Top             =   1995
         Width           =   1815
      End
      Begin VB.Label lblJualanPc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RM 0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2535
         TabIndex        =   21
         Top             =   2490
         Width           =   1470
      End
      Begin VB.Label lblJualanPos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RM 0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2535
         TabIndex        =   20
         Top             =   2925
         Width           =   1470
      End
      Begin VB.Label GenLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "PC Rent Sales :"
         Height          =   195
         Index           =   3
         Left            =   1050
         TabIndex        =   19
         Top             =   2535
         Width           =   1365
      End
      Begin VB.Label GenLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Service && Maintenance :"
         Height          =   195
         Index           =   4
         Left            =   330
         TabIndex        =   18
         Top             =   2970
         Width           =   2100
      End
      Begin VB.Label GenHdr 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Monthly Statistic"
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
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   17
         Top             =   600
         Width           =   4140
      End
      Begin VB.Label SlsLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Date :"
         Height          =   195
         Index           =   0
         Left            =   -74835
         TabIndex        =   14
         Top             =   555
         Width           =   570
      End
      Begin VB.Label SrvLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Date :"
         Height          =   195
         Index           =   0
         Left            =   -74835
         TabIndex        =   11
         Top             =   555
         Width           =   570
      End
      Begin VB.Label SrvLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Filter By Items :"
         Height          =   195
         Index           =   1
         Left            =   -72225
         TabIndex        =   10
         Top             =   555
         Width           =   1410
      End
   End
End
Attribute VB_Name = "FrmStat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Rs As Recordset
Private DBloaded As Boolean

Private Modal As Double
Private Jualan As Double
Private Untung As Double
Private sTahun As String
Private sBulan  As String


Private Sub Form_Load()
 '{ Reset all date container }'
    cbTahun.Clear
    cbBulan.Clear
    cbHari.Clear
    cbHari2.Clear
    PosCmbDate.Clear
    
 '{ load all year }'
    Call LoadYear(cbTahun)
    
 '{ select current year }'
    For a = 0 To cbTahun.ListCount - 1
        If cbTahun.List(a) = Year(Date) Then cbTahun.ListIndex = a
    Next a
    
 '{ display currents session }'
    CurSession = OpenSessionCur
End Sub

Private Sub cbTahun_Click()
 '{ Reset date container }'
    cbBulan.Clear
    cbHari.Clear
    cbHari2.Clear
    
 '{ load all month }'
    sTahun = cbTahun
    Call LoadMonth(cbBulan)
    
 '{ select current month }'
    For a = 0 To cbBulan.ListCount - 1
        If cbBulan.List(a) = Month(Date) Then cbBulan.ListIndex = a
    Next a
End Sub

Private Sub cbBulan_Click()
 '{ Reset date container }'
    Lv1.ListItems.Clear
    Lv2.ListItems.Clear
    Lv3.ListItems.Clear
    PosLV1.ListItems.Clear
    
    sBulan = cbBulan
    Call LoadDate
    
    Call StatKewangan
    Call StatTerminal
    Call StatPelanggan
    Call StatHarian
    Call StatPOS
    GenHdr(0) = " Monthly Statistic - " & GetBulan(sBulan) & " \ " & sTahun
End Sub

Private Sub cbHari_Click()
    If cbHari = "" Then Exit Sub
    If IsDate(cbHari) = False Then Exit Sub
    Lv3.ListItems.Clear
    PosLV1.ListItems.Clear
    
    cbHari2 = cbHari
    Call StatHarian(cbHari, False, False)
    Call StatHarian(cbHari2, True)
End Sub

Private Sub cbHari2_click()
    If cbHari2 = "" Then Exit Sub
    If IsDate(cbHari) = False Then Exit Sub
    Lv3.ListItems.Clear
    
    cbHari = cbHari2
    Call StatHarian(cbHari2, True)
    Call StatHarian(cbHari, False, False)
End Sub


Private Sub PosCmbDate_Click()
    If PosCmbDate = "" Then Exit Sub
    If Not IsDate(PosCmbDate) Then Exit Sub
    
    PosLV1.ListItems.Clear
    Call StatPOS(PosCmbDate, PosCmbItems)
End Sub

Private Sub PosCmbItems_Click()
    If PosCmbItems = "" Then Exit Sub
    
    PosLV1.ListItems.Clear
    Call StatPOS(PosCmbDate, PosCmbItems)
End Sub


Public Sub LoadYear(Cbox As ComboBox)
    Dim Rss As Recordset, Rs As Recordset
    Dim SqlQ As String
   
 '{ add list month to pc sales combo }'
    Set Rss = uIDB.OpenRecordset("pc-harian", dbOpenSnapshot)
    With Rss
        Do Until .EOF = True
            CbAddEx !Tahun, Cbox
            .MoveNext
        Loop
    End With
    
 '{ add list of previous date to pos combo }'
    Set Rs = uIDB.OpenRecordset("pos-usage", dbOpenSnapshot)
    Do Until Rs.BOF = True
        With Rs
            .MoveFirst
            CbAddEx !Tahun, Cbox
            SqlQ = "tahun = '" & !Tahun & "' AND bulan = '" & !Bulan & "' AND hari <> '" & !Hari & "'"
        End With
        Set Rs = RsFilter(Rs, SqlQ)
    Loop
End Sub

Public Sub LoadMonth(Cbox As ComboBox)
    Dim Rss As Recordset, Rs As Recordset
    Dim SqlQ As String
    
 '{ add list month to pc sales combo }'
    Set Rss = uIDB.OpenRecordset("pc-harian", dbOpenSnapshot)
    With Rss
        Do Until .EOF = True
            If !Tahun = sTahun Then CbAddEx !Bulan, Cbox
            .MoveNext
        Loop
    End With
    
 '{ add list of previous date to pos combo }'
    Set Rs = uIDB.OpenRecordset("pos-usage", dbOpenSnapshot)
    Do Until Rs.BOF = True
        With Rs
            .MoveFirst
            If !Tahun = sTahun Then CbAddEx !Bulan, Cbox
            SqlQ = "tahun = '" & !Tahun & "' AND bulan = '" & !Bulan & "' AND hari <> '" & !Hari & "'"
        End With
        Set Rs = RsFilter(Rs, SqlQ)
    Loop
End Sub

Public Sub LoadDate()
    Dim Rss As Recordset, Rs As Recordset, tDate As String
    Dim SqlQ As String
    
    cbHari.Clear
    cbHari2.Clear
    
 '{ add list of previous date to pc sales combo }'
    Set Rss = uIDB.OpenRecordset("pc-harian", dbOpenSnapshot)
    With Rss
        Do Until .EOF = True
            If !Tahun = sTahun And !Bulan = sBulan Then
                tDate = GetSystemDate(!Hari, !Bulan, !Tahun)
                cbHari.AddItem tDate
                cbHari2.AddItem tDate
            End If
            .MoveNext
        Loop
    End With
    
 '{ add list of previous date to pos combo }'
    Set Rs = uIDB.OpenRecordset("pos-usage", dbOpenSnapshot)
    Do Until Rs.BOF = True
        With Rs
            .MoveFirst
            If !Tahun = sTahun And !Bulan = sBulan Then
                tDate = GetSystemDate(!Hari, !Bulan, !Tahun)
                PosCmbDate.AddItem tDate
            End If
            SqlQ = "tahun = '" & !Tahun & "' AND bulan = '" & !Bulan & "' AND hari <> '" & !Hari & "'"
        End With
        Set Rs = RsFilter(Rs, SqlQ)
    Loop
    
 '{ display to combo todays date }'
    If sTahun = Year(Date) And sBulan = Month(Date) Then
        cbHari = Date
    Else
        cbHari = cbHari.List(0)
    End If
    cbHari2 = cbHari
    PosCmbDate = PosCmbDate.List(0)
End Sub




Public Sub StatKewangan()
    Dim Rss As Recordset
    Dim SqlQ As String, JualanPc As Double, JualanPos As Double
    Dim Gaji, Sewa, Bil, BilLain, Biggest, tmpVal As Double
    Dim cTahun, cBulan
    
    Modal = 0: Jualan = 0: Untung = 0: Biggest = 0
    cTahun = Year(sTahun)
    cBulan = Month(sBulan)
    'cHari = Day(OpenSessionCur)
    
    'ambil jumlah kesemua modal
    For K = 0 To uSDBe.DataCount("pekerja-list") - 1
        Gaji = Gaji + CDbl(uSDBe.DataGet("pekerja-list", "gaji", K))
    Next K
    Sewa = CDbl(SetAmbil("sewa"))
    Bil = CDbl(SetAmbil("bil"))
    BilLain = CDbl(SetAmbil("billain"))
    Modal = Gaji + Sewa + Bil + BilLain
    lblModal = Crnc & " " & Format(Modal, "#0.00")
    
    
    'Query-Code untuk filter bulan semasa
    SqlQ = "tahun = '" & sTahun & "' AND bulan = '" & sBulan & "'"
    
    'ambil jumlah jualan - PC
    Set Rss = uIDB.OpenRecordset("pc-harian", dbOpenSnapshot)
    Rss.Filter = SqlQ
    Set Rs = Rss.OpenRecordset
    With Rs
        Do Until .EOF = True
            JualanPc = JualanPc + !pungutan
            .MoveNext
        Loop
    End With
    lblJualanPc = Crnc & " " & Format(JualanPc, "#0.00")
    'ambil jumlah jualan - POS
    Set Rss = uIDB.OpenRecordset("pos-usage", dbOpenSnapshot)
    Rss.Filter = SqlQ
    Set Rs = Rss.OpenRecordset
    With Rs
        Do Until .EOF = True
            JualanPos = JualanPos + !Harga
            .MoveNext
        Loop
    End With
    lblJualanPos = Crnc & " " & Format(JualanPos, "#0.00")
    
    'pemer jumlah harga
    Jualan = JualanPc + JualanPos
    lblJualan = Crnc & " " & Format(Jualan, "#0.00")
    
    'ambil data untuk bar1
    Set Rs = uIDB.OpenRecordset("pc-grafminggu", dbOpenSnapshot)
    If Rs.BOF = True Then Exit Sub
    Rs.FindFirst SqlQ
    If Rs.NoMatch = False Then
        For H = 1 To 7
            dh = Choose(H, "Ahad", "Isnin", "Selasa", "Rabu", "Khamis", "Jumaat", "Sabtu")
            tmpVal = Rs.Fields(dh).Value
            If tmpVal > Biggest Then
                Biggest = tmpVal
                For p = 0 To 6
                    Bar1(p).Max = Biggest
                Next p
            End If
            Bar1(H - 1).Value = tmpVal '(tmpVal * Biggest) / 100
            GrafHigh = Crnc & " " & Format(Biggest, "#0.00")
        Next H
    End If
    
    'kira keuntungan dan juga filter... pasti enakkk
    Untung = Jualan - Modal
    lblUntung = Crnc & " " & Format(Untung, "#0.00")
    If Jualan < Modal Then
        lblJualan.ForeColor = &H8080FF
        lblUntung.ForeColor = &H8080FF
    End If
    
    Set Rss = Nothing
End Sub

Public Sub StatTerminal()
    Dim Rss As Recordset
    Dim tItm As ListItem
    Dim SqlQ As String
    
    'Query-Code untuk filter bulan semasa
    SqlQ = "tahun = '" & sTahun & "' AND bulan = '" & sBulan & "'"
    
    Set Rss = uIDB.OpenRecordset("pc-bulanan", dbOpenSnapshot)
    Rss.Filter = SqlQ
    Set Rs = Rss.OpenRecordset
    If Rs.BOF = True Then Exit Sub
    With Rs
        .MoveLast
        .MoveFirst
        For g = 1 To .RecordCount
            Set tItm = Lv1.ListItems.Add(, , !NamaPc)
            tItm.SubItems(1) = !JumlahMasa & " Minit"
            tItm.SubItems(2) = Crnc & " " & Format(!JumlahBayar, "#0.00")
            .MoveNext
        Next g
    End With
    
    Set Rss = Nothing
    Set tItm = Nothing
End Sub

Public Sub StatPelanggan()
    Dim tItm As ListItem
    
    Set Rs = uSDB.OpenRecordset("pelanggan-list", dbOpenSnapshot)
        
    If Rs.BOF = True Then Exit Sub
    With Rs
        .MoveFirst
        Do Until .EOF = True
            Set tItm = Lv2.ListItems.Add(, , !Nama)
            tItm.SubItems(1) = !lawat & " kali"
            tItm.SubItems(2) = !tarikhakhir
            tItm.SubItems(3) = !JumlahMasa
            tItm.SubItems(4) = Crnc & " " & Format(!JumlahBayar, "#0.00")
            .MoveNext
       Loop
    End With
End Sub


Public Sub StatHarian(Optional TarikhStr As String, Optional LoadListOnly As Boolean = False, Optional LoadDetail As Boolean = True)
    Dim Rss As Recordset
    Dim TrkhTmp As String, SqlQ As String
    Dim nItm As ListItem
    
    If TarikhStr = "" Then
        TrkhTmp = cbHari
    Else
        TrkhTmp = TarikhStr
    End If
    
    SqlQ = "tahun =  '" & Year(TrkhTmp) & "' AND bulan = '" & Month(TrkhTmp) & "' AND hari = '" & Day(TrkhTmp) & "'"
    
    If LoadListOnly = True Then GoTo ListOnly
    Set Rs = uIDB.OpenRecordset("pc-harian", dbOpenSnapshot)
    If Rs.BOF = True Then Exit Sub
    With Rs
        .FindFirst SqlQ
        If .NoMatch = False Then
            lblPungut = Crnc & " " & Format(!pungutan, "#0.00")
            lblPelanggan = !pelanggan
        End If
    End With
    If LoadDetail = False Then Exit Sub
    
ListOnly:
    Set Rss = uIDB.OpenRecordset("pc-usage", dbOpenSnapshot)
    Rss.Filter = SqlQ
    Set Rs = Rss.OpenRecordset
    If Rs.BOF = True Then Exit Sub
    Do While Rs.EOF <> True
        With Rs
            Set nItm = Lv3.ListItems.Add(, , TrkhTmp)
            nItm.SubItems(1) = !PcName
            nItm.SubItems(2) = !Nama
            nItm.SubItems(3) = !masuk
            nItm.SubItems(4) = !Keluar
            nItm.SubItems(5) = Crnc & " " & !Harga
        End With
        Rs.MoveNext
    Loop
    
    Set Rss = Nothing
    Set nItm = Nothing
End Sub

Public Sub StatPOS(Optional FilterDate As String, Optional FilterItem As String)
    Dim tItm As ListItem, SqlQ As String
    
    If FilterDate = "" Then
        SqlQ = "tahun = '" & sTahun & "' AND bulan = '" & sBulan & "'"
    Else
        SqlQ = "tahun = '" & Year(FilterDate) & "' AND bulan = '" & Month(FilterDate) & "' AND hari = '" & Day(FilterDate) & "'"
    End If
    If FilterItem <> "" Then SqlQ = SqlQ & " AND item = '" & FilterItem & "'"
    
    Set Rs = uIDB.OpenRecordset("pos-usage", dbOpenSnapshot)
    Rs.Filter = SqlQ
    Set Rs = Rs.OpenRecordset

    If Rs.BOF = True Then Exit Sub
    With Rs
        .MoveFirst
        Do Until .EOF = True
            Set tItm = PosLV1.ListItems.Add(, , GetSystemDate(!Hari, !Bulan, !Tahun))
            tItm.SubItems(1) = !GroupId
            tItm.SubItems(2) = !transid
            tItm.SubItems(3) = !Item
            tItm.SubItems(4) = !qty
            tItm.SubItems(5) = !Harga
            .MoveNext
       Loop
    End With
    
    Set tItm = Nothing
End Sub

Private Sub SlsBtn_Click()
    'ret = ShellExecute(Me.hwnd, "open", App.Path & "\CafeReport.exe", "pc-usage", vbNullString, SW_NORMAL)
    'If ret <= 32 Then MsgBox MB(20), vbCritical, CbMsgWarn
    Call LoadModule(CafeReport)
End Sub

Private Sub SrvBtn_Click()
    Set Rs = uSDB.OpenRecordset("pos-items", dbOpenSnapshot)
    PosCmbItems.Clear
    
    If Rs.BOF = True Then Exit Sub
    With Rs
        Do Until .EOF = True
            PosCmbItems.AddItem !Nama
            .MoveNext
        Loop
    End With
End Sub

Private Sub StatBtn_Click()
    DBloaded = False
    Unload Me
End Sub

Private Sub StatSesBtn_Click()
    mSj = "Close current session ?"
    ret = MsgBox(mSj, vbOKCancel, CbMsgApp)
    If ret = vbOK Then
        uSDBe.DbSaveSetting "lastsession", OpenSessionCur
        OpenSessionCur = Date
        uSDBe.DbSaveSetting "opensession", OpenSessionCur
    End If
End Sub
