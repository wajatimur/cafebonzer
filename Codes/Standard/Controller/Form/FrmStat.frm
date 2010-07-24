VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmStat 
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
   Begin CafeBonzer.Line3D Line3D1 
      Height          =   45
      Left            =   15
      TabIndex        =   55
      Top             =   735
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   79
      horizon         =   -1  'True
   End
   Begin VB.PictureBox LgBanner 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   9765
      TabIndex        =   53
      Top             =   0
      Width           =   9765
      Begin CafeBonzer.Label3D LgBannerTerminal 
         Height          =   390
         Left            =   6495
         TabIndex        =   54
         ToolTipText     =   "Nama Pc"
         Top             =   45
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor1      =   12632256
         ForeColor2      =   0
         Caption         =   "Statistic && Analysis"
         BackColor       =   16777215
      End
      Begin VB.Image LgBannerImg 
         Height          =   960
         Left            =   60
         Picture         =   "FrmStat.frx":23D2
         Top             =   15
         Width           =   960
      End
   End
   Begin VB.Frame StatFme 
      Caption         =   "Option"
      ForeColor       =   &H00FF0000&
      Height          =   1530
      Left            =   60
      TabIndex        =   2
      Top             =   5760
      Width           =   9675
      Begin CafeBonzer.Line3D StatLine 
         Height          =   1320
         Index           =   1
         Left            =   2475
         TabIndex        =   50
         Top             =   150
         Width           =   45
         _ExtentX        =   79
         _ExtentY        =   2328
         horizon         =   0   'False
      End
      Begin VB.ComboBox StatCbDay 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   1065
         Width           =   1380
      End
      Begin VB.ComboBox StatCbMonth 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   675
         Width           =   1380
      End
      Begin VB.ComboBox StatCbYear 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   285
         Width           =   1380
      End
      Begin CafeBonzer.XpButton StatBtn 
         Height          =   435
         Left            =   9060
         TabIndex        =   51
         ToolTipText     =   "Save settings and exit."
         Top             =   975
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   767
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
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmStat.frx":2E40
         PICN            =   "FrmStat.frx":2E5C
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
         TabIndex        =   44
         Top             =   1080
         Width           =   570
      End
      Begin VB.Label StatLbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Month :"
         Height          =   240
         Index           =   1
         Left            =   225
         TabIndex        =   6
         Top             =   720
         Width           =   660
      End
      Begin VB.Label StatLbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Year :"
         Height          =   240
         Index           =   0
         Left            =   225
         TabIndex        =   4
         Top             =   330
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
      Height          =   4905
      Left            =   45
      TabIndex        =   0
      Top             =   825
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   8652
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
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
      TabPicture(0)   =   "FrmStat.frx":33F6
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
      Tab(0).Control(13)=   "lblServis"
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
      TabPicture(1)   =   "FrmStat.frx":3412
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Lv1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Customers"
      TabPicture(2)   =   "FrmStat.frx":342E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Lv2"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Sales Record"
      TabPicture(3)   =   "FrmStat.frx":344A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Lv3"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Services Record"
      TabPicture(4)   =   "FrmStat.frx":3466
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "PosCmbItems"
      Tab(4).Control(1)=   "PosLV1"
      Tab(4).Control(2)=   "SrvBtn"
      Tab(4).Control(3)=   "SrvLbl(1)"
      Tab(4).ControlCount=   4
      Begin CafeBonzer.Line3D StatLine 
         Height          =   4500
         Index           =   0
         Left            =   4545
         TabIndex        =   42
         Top             =   360
         Width           =   45
         _ExtentX        =   79
         _ExtentY        =   7938
         horizon         =   0   'False
      End
      Begin VB.PictureBox Graf1 
         BackColor       =   &H00808080&
         Height          =   2940
         Left            =   5055
         ScaleHeight     =   2880
         ScaleWidth      =   4170
         TabIndex        =   24
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
            TabIndex        =   25
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
               TabIndex        =   32
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
               TabIndex        =   31
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
               TabIndex        =   30
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
               TabIndex        =   29
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
               TabIndex        =   28
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
               TabIndex        =   27
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
               TabIndex        =   26
               Top             =   -15
               Width           =   330
            End
         End
         Begin MSComctlLib.ProgressBar Bar1 
            Height          =   2460
            Index           =   0
            Left            =   360
            TabIndex        =   33
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
            TabIndex        =   34
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
            TabIndex        =   35
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
            TabIndex        =   36
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
            Index           =   5
            Left            =   2925
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
            Index           =   6
            Left            =   3420
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
            TabIndex        =   40
            Top             =   45
            Width           =   765
         End
      End
      Begin VB.ComboBox PosCmbItems 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   -73380
         TabIndex        =   8
         Top             =   510
         Width           =   1620
      End
      Begin MSComctlLib.ListView PosLV1 
         Height          =   3750
         Left            =   -74835
         TabIndex        =   7
         Top             =   975
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   6615
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
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Transaction ID"
            Object.Width           =   3175
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
         Height          =   4260
         Left            =   -74835
         TabIndex        =   10
         Top             =   465
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   7514
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
            Object.Width           =   2117
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
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Out"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Paid"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView Lv2 
         Height          =   4260
         Left            =   -74835
         TabIndex        =   11
         Top             =   465
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   7514
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
            Object.Width           =   2469
         EndProperty
      End
      Begin MSComctlLib.ListView Lv1 
         Height          =   4245
         Left            =   -74835
         TabIndex        =   12
         Top             =   465
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   7488
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
         Left            =   -71700
         TabIndex        =   52
         ToolTipText     =   "Delete selected employee."
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
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmStat.frx":3482
         PICN            =   "FrmStat.frx":349E
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
         TabIndex        =   49
         Top             =   3450
         Width           =   4140
      End
      Begin VB.Label GenLbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PC Collection :"
         Height          =   195
         Index           =   5
         Left            =   1155
         TabIndex        =   48
         Top             =   3915
         Width           =   1275
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
         TabIndex        =   47
         Top             =   3885
         Width           =   1470
      End
      Begin VB.Label lblServis 
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
         TabIndex        =   46
         Top             =   4380
         Width           =   1470
      End
      Begin VB.Label GenLbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Service && Merchandise :"
         Height          =   195
         Index           =   6
         Left            =   345
         TabIndex        =   45
         Top             =   4410
         Width           =   2085
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
         TabIndex        =   41
         Top             =   600
         Width           =   4635
      End
      Begin VB.Label GenLbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monthly Overhead :"
         Height          =   195
         Index           =   0
         Left            =   465
         TabIndex        =   23
         Top             =   1050
         Width           =   1695
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
         Top             =   1950
         Width           =   1725
      End
      Begin VB.Label GenLbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monthly Sales :"
         Height          =   195
         Index           =   1
         Left            =   840
         TabIndex        =   19
         Top             =   1500
         Width           =   1320
      End
      Begin VB.Label GenLbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monthly Profit :"
         Height          =   195
         Index           =   2
         Left            =   855
         TabIndex        =   18
         Top             =   1995
         Width           =   1305
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
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   2925
         Width           =   1470
      End
      Begin VB.Label GenLbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PC Rent Sales :"
         Height          =   195
         Index           =   3
         Left            =   1065
         TabIndex        =   15
         Top             =   2535
         Width           =   1350
      End
      Begin VB.Label GenLbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Service && Merchandise :"
         Height          =   195
         Index           =   4
         Left            =   345
         TabIndex        =   14
         Top             =   2970
         Width           =   2085
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
         TabIndex        =   13
         Top             =   600
         Width           =   4140
      End
      Begin VB.Label SrvLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Filter By Items :"
         Height          =   195
         Index           =   1
         Left            =   -74850
         TabIndex        =   9
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

'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : FrmStat
'    Project    : CafeBonzer
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Private Sub Form_Load()
 '{ Reset all date container }'
    StatCbYear.Clear
    StatCbMonth.Clear
    StatCbDay.Clear
    
 '{ Load all year }'
    Call LoadDate(StatCbYear, StatCbMonth, StatCbDay, -1)
End Sub

Private Sub StatCbYear_Click()
 '{ select current month }'
    Call LoadDate(StatCbYear, StatCbMonth, StatCbDay, 0)
End Sub

Private Sub StatCbMonth_Click()
 '{ Reset date container }'
    Lv1.ListItems.Clear
    Lv2.ListItems.Clear
    Lv3.ListItems.Clear
    PosLV1.ListItems.Clear

    Call LoadDate(StatCbYear, StatCbMonth, StatCbDay, 1)
    Call StatLoadFinancial
    Call StatLoadCustomer
    Call StatLoadUsagePc(StatCbDay)
    Call StatLoadUsagePos(StatCbDay)
    GenHdr(0) = " Monthly Statistic - " & GetMonthString(StatCbMonth) & " \ " & StatCbYear
End Sub

Private Sub StatCbDay_Click()
    If StatCbDay = "" Then Exit Sub
    If IsDate(StatCbDay) = False Then Exit Sub
    Lv3.ListItems.Clear
    PosLV1.ListItems.Clear
    
    Call StatLoadUsagePc(StatCbDay)
    Call StatLoadUsagePos(StatCbDay)
End Sub

Private Sub SrvBtn_Click()
    Dim CRset As Recordset
    Set CRset = CDataS.OpenRecordset("ServiceItems", dbOpenSnapshot)
    PosCmbItems.Clear
    
    If CRset.BOF = True Then Exit Sub
    With CRset
        Do Until .EOF = True
            PosCmbItems.AddItem !Name
            .MoveNext
        Loop
    End With
    Set CRset = Nothing
End Sub

Private Sub StatBtn_Click()
    Unload FrmStat
End Sub


Public Sub StatLoadFinancial()
    Dim CDbRs As Recordset, LngIdxA As Long, LngIdxB As Long
    Dim DblProfit As Double, DblSalesTerminal As Double, DblSalesServices As Double
    Dim LngOverHead As Double, LngSalary As Double
    Dim LngGraftHigh, LngGraftValue As Double
    
    For LngIdxA = 0 To CDataSe.DataCount("ListEmployee") - 1
        LngSalary = LngSalary + CDbl(CDataSe.DataGet("ListEmployee", "Salary", LngIdxA))
    Next
    For LngIdxA = 0 To CDataSe.DataCount("FinanceOverhead") - 1
        LngOverHead = LngOverHead + CDbl(CDataSe.DataGet("FinanceOverhead", "Value", LngIdxA))
    Next
    LngOverHead = LngOverHead + LngSalary
    
    DblSalesTerminal = StatGetSalesMonth(StatCbYear, StatCbMonth, 1)
    DblSalesServices = StatGetSalesMonth(StatCbYear, StatCbMonth, 2)
    DblProfit = DblSalesTerminal + DblSalesServices

    For LngIdxA = 1 To 7
        LngGraftValue = StatGetSalesByDay(StatCbYear, StatCbMonth, LngIdxA)
        If LngGraftValue > LngGraftHigh Then LngGraftHigh = LngGraftValue
        If LngGraftHigh > 0 Then
            Bar1(LngIdxA - 1).Max = LngGraftHigh
            For LngIdxB = 0 To 6
                Bar1(LngIdxB).Max = LngGraftHigh
            Next
        End If
        Bar1(LngIdxA - 1).Value = CStr(LngGraftValue)
    Next
    
    GrafHigh = Crnc & " " & Format(LngGraftHigh, "#0.00")
    lblJualanPc = Crnc & " " & Format(DblSalesTerminal, "#0.00")
    lblJualanPos = Crnc & " " & Format(DblSalesServices, "#0.00")
    lblModal = Crnc & " " & Format(LngOverHead, "#0.00")
    lblJualan = Crnc & " " & Format((DblSalesTerminal + DblSalesServices), "#0.00")
    lblUntung = Crnc & " " & Format((DblProfit - LngOverHead), "#0.00")
    
    Set CDbRs = Nothing
End Sub


Public Sub StatLoadCustomer()
    Dim CListItem As ListItem
    
    Set CRset = CDataS.OpenRecordset("ListCustomer", dbOpenSnapshot)
        
    If CRset.BOF = True Then Exit Sub
    With CRset
        .MoveFirst
        Do Until .EOF = True
            Set CListItem = Lv2.ListItems.Add(, , !Name)
            CListItem.SubItems(1) = !CountVisit & " Times"
            CListItem.SubItems(2) = !LastVisit
            CListItem.SubItems(3) = !TotalTime
            CListItem.SubItems(4) = Crnc & " " & Format(!TotalPaid, "#0.00")
            .MoveNext
       Loop
    End With
End Sub


Public Sub StatLoadUsagePc(DteStatDate As Date)
    Dim CRset As Recordset, StrSqlQ As String, CListItem As ListItem

    lblPungut = Crnc & " " & StatGetSalesDay(Year(DteStatDate), Month(DteStatDate), Day(DteStatDate), 1, False)
    
    StrSqlQ = "SELECT * FROM LogUsageTerminal WHERE Year = " & Year(DteStatDate) & " AND Month = " & Month(DteStatDate) & " AND Day = " & Day(DteStatDate)
    StrSqlQ = StrSqlQ & " ORDER BY TimeIn"
    Set CRset = CDataI.OpenRecordset(StrSqlQ, dbOpenSnapshot)
    If CRset.BOF = True Then Exit Sub
    Do While CRset.EOF <> True
        With CRset
            Set CListItem = Lv3.ListItems.Add(, , DateGetSystem(!Day, !Month, !Year))
            CListItem.SubItems(1) = !Terminal
            CListItem.SubItems(2) = !Customer
            CListItem.SubItems(3) = !TimeIn
            CListItem.SubItems(4) = !TimeOut
            CListItem.SubItems(5) = Crnc & " " & !Price
        End With
        CRset.MoveNext
    Loop

    Set CRset = Nothing
    Set CListItem = Nothing
End Sub

Public Sub StatLoadUsagePos(DteStatDate As Date, Optional FilterItem As String)
    Dim CRset As Recordset, StrSqlQ As String, CListItem As ListItem
    
    lblServis = Crnc & " " & StatGetSalesDay(Year(DteStatDate), Month(DteStatDate), Day(DteStatDate), 2, True)

    StrSqlQ = "SELECT * FROM LogUsageServices WHERE Year = " & Year(DteStatDate) & " AND Month = " & Month(DteStatDate) & " AND Day = " & Day(DteStatDate)
    If FilterItem <> "" Then StrSqlQ = StrSqlQ & " AND Item = " & FilterItem & " ORDER BY Item"
    
    Set CRset = CDataI.OpenRecordset(StrSqlQ, dbOpenSnapshot)
    With CRset
        If .BOF = True Then Exit Sub
        .MoveFirst
        Do Until .EOF = True
            Set CListItem = PosLV1.ListItems.Add(, , DateGetSystem(!Day, !Month, !Year))
            CListItem.SubItems(1) = !GroupID
            CListItem.SubItems(2) = !TransactionId
            CListItem.SubItems(3) = !Item
            CListItem.SubItems(4) = !Quantity
            CListItem.SubItems(5) = Crnc & " " & !Price
            .MoveNext
       Loop
    End With
    
    Set CRset = Nothing
    Set CListItem = Nothing
End Sub
