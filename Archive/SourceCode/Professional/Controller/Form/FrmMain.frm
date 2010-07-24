VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Station Manager"
   ClientHeight    =   4575
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   10170
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
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   10170
   Begin MSComctlLib.ImageList Iml 
      Left            =   9570
      Top             =   1185
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":0B26
            Key             =   "user"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":157A
            Key             =   "akses"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1B16
            Key             =   "item"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":20B2
            Key             =   "services"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":264E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":2766
            Key             =   "foods"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":2B02
            Key             =   "beverages"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":2E9E
            Key             =   "magazines"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":343A
            Key             =   "others"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":39D6
            Key             =   "none"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglist2 
      Left            =   9570
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":3F72
            Key             =   "aktif"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":4C4E
            Key             =   "jalan"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":592A
            Key             =   "tamat"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":6606
            Key             =   "lock"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":72E2
            Key             =   "aktif1"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglist 
      Left            =   9570
      Top             =   615
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   37
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":7FBE
            Key             =   "aktif1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":855A
            Key             =   "aktif2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":88F6
            Key             =   "logoff"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":8C9E
            Key             =   "boot"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":923A
            Key             =   "jalan1"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":97D6
            Key             =   "jalan"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":9D72
            Key             =   "tamat"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":A30E
            Key             =   "lock"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":A8AA
            Key             =   "mouse"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":AE46
            Key             =   "info"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":B3E2
            Key             =   "query"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":B97E
            Key             =   "dump"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":BF22
            Key             =   "mesej"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":C4BE
            Key             =   "kuncibuka"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":CA5A
            Key             =   "off"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":CFF6
            Key             =   "penalaan1"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":D592
            Key             =   "penalaan"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":DEA6
            Key             =   "hoi"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":E442
            Key             =   "power"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":E7DE
            Key             =   "net"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":ED7A
            Key             =   "gift"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":F31E
            Key             =   "graft"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":F8C2
            Key             =   "no"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":FE5E
            Key             =   "transfer"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":103FA
            Key             =   "term"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":10996
            Key             =   "key"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":10F32
            Key             =   "cleaning"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":114CE
            Key             =   "broad"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1186A
            Key             =   "flagin"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":11E06
            Key             =   "flagout"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":123A2
            Key             =   "meter"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1293E
            Key             =   "wait"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":12EDA
            Key             =   "people"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":13476
            Key             =   "printer"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":13A12
            Key             =   "help"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":13FAE
            Key             =   "cpu"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":14556
            Key             =   "paper"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Pages 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Height          =   4605
      Index           =   0
      Left            =   15
      ScaleHeight     =   4605
      ScaleWidth      =   10170
      TabIndex        =   0
      Top             =   15
      Width           =   10170
      Begin MSComctlLib.ListView Lv1 
         Height          =   4590
         Left            =   -15
         TabIndex        =   1
         Top             =   -15
         Width           =   10185
         _ExtentX        =   17965
         _ExtentY        =   8096
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imglist2"
         SmallIcons      =   "imglist"
         ColHdrIcons     =   "imglist"
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Station"
            Object.Width           =   2822
            ImageKey        =   "dump"
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Status"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Username"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "User Type"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Time In"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Time Out"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Current"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Printed"
            Object.Width           =   2117
         EndProperty
      End
   End
   Begin VB.PictureBox Pages 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Height          =   4605
      Index           =   1
      Left            =   15
      ScaleHeight     =   4605
      ScaleWidth      =   10170
      TabIndex        =   2
      Top             =   15
      Width           =   10170
      Begin MSComctlLib.ProgressBar DynaPbar 
         Height          =   240
         Index           =   0
         Left            =   75
         TabIndex        =   6
         Top             =   405
         Visible         =   0   'False
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.ComboBox DynaCombo 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "FrmMain.frx":14AF2
         Left            =   45
         List            =   "FrmMain.frx":14AF9
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   60
         Visible         =   0   'False
         Width           =   1545
      End
      Begin MSComctlLib.ListView DynaLv 
         Height          =   4530
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   15
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   7990
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imglist2"
         SmallIcons      =   "imglist"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView DynaLv 
         Height          =   4530
         Index           =   1
         Left            =   3705
         TabIndex        =   4
         Top             =   15
         Width           =   6465
         _ExtentX        =   11404
         _ExtentY        =   7990
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imglist2"
         SmallIcons      =   "imglist"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin CafeBonzer.PageHolder MainPhold 
      Height          =   2970
      Left            =   0
      TabIndex        =   7
      Top             =   4620
      Width           =   10170
      _ExtentX        =   17939
      _ExtentY        =   5239
      HldrTxt         =   "Toolbox"
      HldrTxtClr      =   16777215
      HldrLne         =   -1  'True
      PageHeight      =   2970
      Begin CafeBonzer.Line3D MainLine 
         Height          =   2595
         Left            =   465
         TabIndex        =   31
         Top             =   345
         Width           =   45
         _ExtentX        =   79
         _ExtentY        =   4577
         horizon         =   0   'False
      End
      Begin CafeBonzer.XpButton SubPagesMnu 
         Height          =   405
         Index           =   3
         Left            =   30
         TabIndex        =   30
         ToolTipText     =   "Log"
         Top             =   1635
         Width           =   405
         _ExtentX        =   714
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmMain.frx":14B0A
         PICN            =   "FrmMain.frx":14B26
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CafeBonzer.XpButton SubPagesMnu 
         Height          =   405
         Index           =   2
         Left            =   30
         TabIndex        =   29
         ToolTipText     =   "Note"
         Top             =   1215
         Width           =   405
         _ExtentX        =   714
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmMain.frx":150C0
         PICN            =   "FrmMain.frx":150DC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CafeBonzer.XpButton SubPagesMnu 
         Height          =   405
         Index           =   1
         Left            =   30
         TabIndex        =   28
         ToolTipText     =   "Service & Merchandise"
         Top             =   795
         Width           =   405
         _ExtentX        =   714
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmMain.frx":15676
         PICN            =   "FrmMain.frx":15692
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CafeBonzer.XpButton SubPagesMnu 
         Height          =   405
         Index           =   0
         Left            =   30
         TabIndex        =   27
         ToolTipText     =   "Information"
         Top             =   375
         Width           =   405
         _ExtentX        =   714
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmMain.frx":15C2C
         PICN            =   "FrmMain.frx":15C48
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.PictureBox SubPages 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2565
         Index           =   1
         Left            =   570
         ScaleHeight     =   2565
         ScaleWidth      =   9660
         TabIndex        =   8
         Top             =   360
         Width           =   9660
         Begin VB.TextBox SerTxtJumlah 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BeginProperty Font 
               Name            =   "Endless Showroom"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Left            =   7965
            Locked          =   -1  'True
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   120
            Width           =   1365
         End
         Begin VB.TextBox SerTxtQty 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0C0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1275
            TabIndex        =   17
            Text            =   "1"
            Top             =   1125
            Width           =   615
         End
         Begin VB.TextBox SerTxtBaki 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BeginProperty Font 
               Name            =   "Endless Showroom"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Left            =   7965
            Locked          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   660
            Width           =   1365
         End
         Begin VB.TextBox SerTxtBayar 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BeginProperty Font 
               Name            =   "Endless Showroom"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Left            =   7965
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   1215
            Width           =   1365
         End
         Begin VB.VScrollBar SerScroll1 
            Height          =   330
            Left            =   1920
            Max             =   999
            Min             =   1
            TabIndex        =   14
            Top             =   1125
            Value           =   999
            Width           =   165
         End
         Begin VB.TextBox SerTxtTotalItm 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFC0&
            Height          =   315
            Left            =   1275
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   2145
            Width           =   1740
         End
         Begin VB.TextBox SerTxtPriItm 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1275
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   1635
            Width           =   1740
         End
         Begin MSComctlLib.ImageCombo SerImgCb2 
            Height          =   330
            Left            =   1275
            TabIndex        =   11
            Top             =   600
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Locked          =   -1  'True
            Text            =   "None"
         End
         Begin MSComctlLib.ImageCombo SerImgCb1 
            Height          =   330
            Left            =   1275
            TabIndex        =   12
            Top             =   75
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Locked          =   -1  'True
            Text            =   "None"
            ImageList       =   "Iml"
         End
         Begin MSComctlLib.ListView SerLv1 
            Height          =   2415
            Left            =   3150
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   60
            Width           =   3270
            _ExtentX        =   5768
            _ExtentY        =   4260
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
            Appearance      =   1
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Item"
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Price"
               Object.Width           =   1235
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Qty."
               Object.Width           =   1058
            EndProperty
         End
         Begin CafeBonzer.XpButton SerBtn 
            Height          =   480
            Index           =   0
            Left            =   8895
            TabIndex        =   50
            ToolTipText     =   "Accept"
            Top             =   1995
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
            MICON           =   "FrmMain.frx":161E2
            PICN            =   "FrmMain.frx":161FE
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin CafeBonzer.XpButton SerBtnItm 
            Height          =   360
            Index           =   0
            Left            =   2130
            TabIndex        =   51
            Top             =   1095
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
            MICON           =   "FrmMain.frx":18280
            PICN            =   "FrmMain.frx":1829C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin CafeBonzer.XpButton SerBtnItm 
            Height          =   360
            Index           =   1
            Left            =   2565
            TabIndex        =   52
            Top             =   1095
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   635
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
            MICON           =   "FrmMain.frx":18836
            PICN            =   "FrmMain.frx":18852
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label SerLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   240
            Index           =   3
            Left            =   6585
            TabIndex        =   26
            Top             =   180
            Width           =   675
         End
         Begin VB.Label SerLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quantity :"
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   2
            Left            =   75
            TabIndex        =   25
            Top             =   1125
            Width           =   855
         End
         Begin VB.Label SerLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Items :"
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   1
            Left            =   75
            TabIndex        =   24
            Top             =   615
            Width           =   630
         End
         Begin VB.Label SerLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Category :"
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   0
            Left            =   75
            TabIndex        =   23
            Top             =   105
            Width           =   930
         End
         Begin VB.Label SerLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Balanced :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   240
            Index           =   4
            Left            =   6585
            TabIndex        =   22
            Top             =   735
            Width           =   1140
         End
         Begin VB.Label SerLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Received :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   240
            Index           =   5
            Left            =   6585
            TabIndex        =   21
            Top             =   1275
            Width           =   1125
         End
         Begin VB.Label SerLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Items Price :"
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   6
            Left            =   75
            TabIndex        =   20
            Top             =   2160
            Width           =   1110
         End
         Begin VB.Label SerLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Price :"
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   7
            Left            =   75
            TabIndex        =   19
            Top             =   1650
            Width           =   555
         End
      End
      Begin VB.PictureBox SubPages 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2565
         Index           =   3
         Left            =   570
         ScaleHeight     =   2565
         ScaleWidth      =   9660
         TabIndex        =   35
         Top             =   360
         Width           =   9660
         Begin VB.ListBox MainLog 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   2370
            Left            =   60
            TabIndex        =   36
            Top             =   75
            Width           =   7740
         End
      End
      Begin VB.PictureBox SubPages 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2565
         Index           =   2
         Left            =   570
         ScaleHeight     =   2565
         ScaleWidth      =   9660
         TabIndex        =   33
         Top             =   360
         Width           =   9660
         Begin VB.TextBox MainNote 
            Height          =   2430
            Left            =   45
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   34
            Top             =   60
            Width           =   7185
         End
         Begin CafeBonzer.XpButton MainNoteBtn 
            Height          =   345
            Index           =   0
            Left            =   7290
            TabIndex        =   53
            Top             =   75
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   609
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
            MICON           =   "FrmMain.frx":18DEC
            PICN            =   "FrmMain.frx":18E08
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin CafeBonzer.XpButton MainNoteBtn 
            Height          =   345
            Index           =   1
            Left            =   7290
            TabIndex        =   54
            Top             =   435
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   609
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
            MICON           =   "FrmMain.frx":193A2
            PICN            =   "FrmMain.frx":193BE
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
      Begin VB.PictureBox SubPages 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2565
         Index           =   0
         Left            =   570
         ScaleHeight     =   2565
         ScaleWidth      =   9660
         TabIndex        =   32
         Top             =   360
         Width           =   9660
         Begin CafeBonzer.Line3D SpgInfoLine 
            Height          =   2610
            Left            =   4170
            TabIndex        =   49
            Top             =   -15
            Width           =   45
            _ExtentX        =   79
            _ExtentY        =   4604
            horizon         =   0   'False
         End
         Begin VB.Label SpgInfoLblD 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00404040&
            Height          =   285
            Index           =   3
            Left            =   1860
            TabIndex        =   48
            Top             =   1845
            Width           =   2115
         End
         Begin VB.Label SpgInfoLblD 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00404040&
            Height          =   285
            Index           =   2
            Left            =   1860
            TabIndex        =   47
            Top             =   1410
            Width           =   2115
         End
         Begin VB.Label SpgInfoLblD 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00404040&
            Height          =   285
            Index           =   1
            Left            =   1860
            TabIndex        =   46
            Top             =   960
            Width           =   2115
         End
         Begin VB.Label SpgInfoLblD 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00404040&
            Height          =   285
            Index           =   0
            Left            =   1860
            TabIndex        =   45
            Top             =   525
            Width           =   2115
         End
         Begin VB.Label SpgInfoLblC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current Used :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   3
            Left            =   300
            TabIndex        =   44
            Top             =   1860
            Width           =   1395
         End
         Begin VB.Label SpgInfoLblC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MAC Address :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   2
            Left            =   315
            TabIndex        =   43
            Top             =   1425
            Width           =   1380
         End
         Begin VB.Label SpgInfoLblC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IP Address :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   1
            Left            =   510
            TabIndex        =   42
            Top             =   975
            Width           =   1185
         End
         Begin VB.Label SpgInfoLblC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Connected At :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   0
            Left            =   285
            TabIndex        =   41
            Top             =   540
            Width           =   1410
         End
         Begin VB.Image SpgInfoHdr 
            Height          =   300
            Index           =   1
            Left            =   90
            Picture         =   "FrmMain.frx":19958
            Top             =   60
            Width           =   3225
         End
         Begin VB.Image SpgInfoHdr 
            Height          =   270
            Index           =   0
            Left            =   4365
            Picture         =   "FrmMain.frx":1A879
            Top             =   60
            Width           =   2400
         End
         Begin VB.Label SpgInfoLblB 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00404040&
            Height          =   285
            Index           =   1
            Left            =   6585
            TabIndex        =   40
            Top             =   915
            Width           =   885
         End
         Begin VB.Label SpgInfoLblA 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unused Station :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   1
            Left            =   4830
            TabIndex        =   39
            Top             =   930
            Width           =   1590
         End
         Begin VB.Label SpgInfoLblB 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00404040&
            Height          =   285
            Index           =   0
            Left            =   6585
            TabIndex        =   38
            Top             =   480
            Width           =   885
         End
         Begin VB.Label SpgInfoLblA 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Connected Agent :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   0
            Left            =   4650
            TabIndex        =   37
            Top             =   495
            Width           =   1770
         End
      End
   End
   Begin VB.Menu popmenu1 
      Caption         =   "<popmenu1>"
      Visible         =   0   'False
      Begin VB.Menu pmenu1flog 
         Caption         =   "Fast Login"
      End
      Begin VB.Menu pmenu1flout 
         Caption         =   "Fast Logout"
      End
      Begin VB.Menu psep2 
         Caption         =   "-"
      End
      Begin VB.Menu pmenu1cancel 
         Caption         =   "Cancel User"
      End
      Begin VB.Menu pmenu1trans 
         Caption         =   "Transfer PC"
      End
      Begin VB.Menu pmenu1terminal 
         Caption         =   "Terminal"
      End
      Begin VB.Menu psep1 
         Caption         =   "-"
      End
      Begin VB.Menu pmenu1cln 
         Caption         =   "Cleaning"
         Begin VB.Menu pmenu1clnsub 
            Caption         =   "All"
            Index           =   0
         End
         Begin VB.Menu pmenu1clnsub 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu pmenu1clnsub 
            Caption         =   "Temp Folder"
            Index           =   2
         End
         Begin VB.Menu pmenu1clnsub 
            Caption         =   "Recycle Bin"
            Index           =   3
         End
         Begin VB.Menu pmenu1clnsub 
            Caption         =   "Internet History"
            Index           =   4
         End
         Begin VB.Menu pmenu1clnsub 
            Caption         =   "Recent Docs"
            Index           =   5
         End
      End
      Begin VB.Menu pmenu1ctl 
         Caption         =   "Control"
         Begin VB.Menu pmenu1ctlsub 
            Caption         =   "Lock Computer"
            Index           =   0
         End
         Begin VB.Menu pmenu1ctlsub 
            Caption         =   "Unlock Computer"
            Index           =   1
         End
         Begin VB.Menu pmenu1ctlsub 
            Caption         =   "Reboot Computer"
            Index           =   2
         End
         Begin VB.Menu pmenu1ctlsub 
            Caption         =   "Shutdown Computer"
            Index           =   3
         End
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NegateX As Long      'beza x saiz untuk Lv
Public NegateY As Long      'beza y saiz untuk Lv
Public NegateXtmp As Long
Public NegateYtmp As Long

Private SerTotal As Double


'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Form Initialize
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub Form_Initialize()
    'FrmMain.Height = 8640
    'Call LayOutMeasure
    'Call CbFrmMetricLoad(FrmMain)
End Sub


'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Form Resizing
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub Form_Resize()
    On Error Resume Next
    
    If FrmMain.Height < 7000 Then FrmMain.Height = 8000
    If FrmMain.Width < 11400 Then FrmMain.Width = 11400
    'Call LayOutSize
End Sub


'###############################################################################################
'# MAIN GUI
'#
'#
'###############################################################################################
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Page Holder
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub MainPhold_PageFlip(ByVal Collapse As Boolean)
    If Collapse = True Then
        NegateYtmp = NegateY
        Pages(0).Height = MainPhold.Top
        NegateY = FrmMain.Height - Pages(0).Height
    Else
        NegateY = NegateYtmp
        Pages(0).Height = FrmMain.Height - NegateY
    End If
    Menu4EnvSub(0).Checked = Collapse Xor True
    Call LayOutSize
    SetSimpan "tooltab", FrmMain.Menu4EnvSub(0).Checked
End Sub


'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Main Page - Note
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub MainNoteBtn_Click(Index As Integer)
    Select Case Index
    Case 0
        SetSimpan "mainnote", MainNote
    Case 1
        MainNote = ""
        SetSimpan "mainnote", " "
    End Select
End Sub




'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' SubPages | Menu
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub SubPagesMnu_Click(Index As Integer)
    SubPages(Index).ZOrder 0
End Sub

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ListView | DoubleClick
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub Lv1_DblClick()
    If AgentCount = 0 Then Exit Sub
        
    If SelKey = "" Then Exit Sub
    If SelSubItm(1) = VS(4) And SelTag = "dump" Then
        FrmUserMdump1.Show vbModal
        Exit Sub
    End If
    If SelSubItm(1) = VS(3) And SelTag = "dump" Then
        FrmUserMdump2.Show vbModal
        Exit Sub
    End If
    
    If SelSubItm(1) = VS(4) Then FrmUserMenu1.Show vbModal: Exit Sub
    If SelSubItm(1) = VS(3) Then FrmUserMenu2.Show vbModal: Exit Sub
    If SelSubItm(1) = VS(5) Then FrmUserMenu3.Show vbModal: Exit Sub
End Sub
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ListView | ItemClick
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub Lv1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call UpdatePanel(Item.Text)
    Call UpdateStat(Item)
End Sub
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ListView | When the mouse goes up
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub Lv1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If AgentCount = 0 Then Exit Sub
    If Button = 2 Then FrmMain.PopupMenu popmenu1, , X, Y: Exit Sub
End Sub



'###############################################################################################
'# DYNAMIC PAGES
'#
'#
'###############################################################################################
Private Sub DynaCombo_Click()
    DynaLv(0).SelectedItem.SubItems(1) = DynaCombo.Text
    DynaCombo.Visible = False
End Sub

Private Sub DynaLv_DblClick(Index As Integer)
    Select Case Index
        Case 0
            If MglPageLast = 1 Then MgoSlv1.MatrixExpand DynaLv(0).SelectedItem
    End Select
End Sub

Private Sub DynaLv_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    Select Case Index
        Case 0
            DynaCombo.Visible = False
            If MglPageLast = 1 Then InfoJobsEnum (Item.Key), Item.SubItems(1)
        Case 1
            If MglPageLast = 2 Then UpdatePanel Item.Text
    End Select
End Sub



'###############################################################################################
'# SERVICES & MERCHANDISE
'#
'#
'###############################################################################################
Private Sub SerBtn_Click(Index As Integer)
    Dim sItm As ListItem
    
    If SerTotal = 0 Then
        MsgBox MB(14), vbInformation, CbMsgWarn
        Exit Sub
    End If
    
    msg = "Confirm transaction of " & Crnc & " " & Format(SerTotal, "#0.00")
    msg = msg & vbCrLf & "for the below items :-" & vbCrLf & vbCrLf
    For d = 1 To SerLv1.ListItems.Count
        Set sItm = SerLv1.ListItems(d)
        msg = msg & "  " & d & ". " & sItm.Text & vbTab & " =   " & sItm.SubItems(2) & vbCrLf
    Next d
    ret = MsgBox(msg, vbOKCancel, CbMsgApp)
    If ret = vbOK Then
        SavePosTrans SerLv1.ListItems
        Call SerCtlReset
    End If
End Sub

Private Sub SerBtnItm_Click(Index As Integer)
    Dim SCbItm As ComboItem, lvItm As ListItem, fItm As ListItem
     
    Select Case Index
    Case 0
        Set SCbItm = SerImgCb2.SelectedItem
        If SerTxtQty > 0 Then
            Set fItm = SerLv1.FindItem(SCbItm.Text)
            
            If fItm Is Nothing Then
                Set lvItm = SerLv1.ListItems.Add(, SCbItm.Key, SCbItm.Text)
                lvItm.SubItems(1) = SCbItm.Tag
                lvItm.SubItems(2) = SerTxtQty
            Else
                fItm.SubItems(2) = SerTxtQty
            End If
        Else
            MsgBox "Please enter quantity !", vbOKOnly, CbMsgWarn
            SerTxtQty.SelStart = 1
            SerTxtQty.SelLength = Len(SerTxtQty.Text)
            Exit Sub
        End If
    Case 1
        If SerLv1.ListItems.Count = 0 Then Exit Sub
        SerLv1.ListItems.Remove (SerLv1.SelectedItem.Index)
    End Select
    
    'recalculate total
    SerTotal = 0
    For g = 1 To SerLv1.ListItems.Count
        SerTotal = SerTotal + (CDbl(SerLv1.ListItems(g).SubItems(1)) * CInt(SerLv1.ListItems(g).SubItems(2)))
    Next g
    SerTxtJumlah = Crnc & " " & Format(SerTotal, "#0.00")
End Sub
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Services Bayar - Change
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub SerTxtBayar_Change()
    If Trim(SerTxtBayar.Text) <> "" And IsNumeric(SerTxtBayar.Text) = True Then
        SerTxtBaki = Crnc & Format$((SerTxtBayar - SerTotal), "#0.00")
    Else
        SerTxtBaki = ""
    End If
End Sub
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Services Price PerItem - KeyUp
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub SerTxtPriItm_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If SerTxtPriItm = "" Then Exit Sub
        If IsNumeric(SerTxtPriItm) = False Then Exit Sub
        SerLv1.SelectedItem.SubItems(1) = Format(SerTxtPriItm, "#0.00")
    End If
End Sub
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Services ImageCombo1 - Change
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub SerImgCb1_Change()
    If SerImgCb1.Text <> VS(1) Then
        Call SerCtlEnable
        Call LoadPosItmCB(SerImgCb2, Mid(SerImgCb1.SelectedItem.Key, 2))
    Else
        Call SerCtlEnable(False)
    End If
End Sub
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Services ImageCombo1 - Click
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub SerImgCb1_Click()
    Call SerImgCb1_Change
End Sub
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Services Lv1 - Total Item Price
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub SerLv1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim ItemTotal As Double
    
    ItemTotal = CDbl(Item.SubItems(1)) * CInt(Item.SubItems(2))
    SerTxtJumlah = Crnc & " " & Format(SerTotal, "#0.00")
    SerTxtTotalItm = Crnc & " " & Format(ItemTotal, "#0.00")
    SerTxtPriItm = Item.SubItems(1)
End Sub
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Quantity scroller
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub SerScroll1_Change()
    SerTxtQty = 1000 - SerScroll1.Value
End Sub



'###############################################################################################
'# MENU & SUBMENU
'#
'#
'###############################################################################################



'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' POPUPMENU
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'[ popmenu fast login ]'
Private Sub pmenu1flog_Click()
    If AgentCount = 0 Then Exit Sub
    ' Get station status
    If SelSubItm(1) = VS(3) Or SelSubItm(1) = VS(5) Then
        MsgBox MB(21), vbInformation, CbMsgWarn
        Exit Sub
    End If
    FrmGuna.FastLogin
    Unload FrmGuna
End Sub
'[ popmenu fast logout ]'
Private Sub pmenu1flout_Click()
    If AgentCount = 0 Then Exit Sub
    If SelSubItm(1) = VS(3) Then Call UserHenti
    If SelSubItm(1) = VS(5) Then Call UserHenti2
End Sub
'[ popmenu transfer ]'
Private Sub pmenu1trans_Click()
    If SelAgn.AgentStatus = VS(4) Then Exit Sub
    If SelAgn.AgentStatus = VS(5) Then Exit Sub
    FrmTrans.Show
End Sub
'[ popmenu terminal ]'
Private Sub pmenu1terminal_Click()
    If AgentCount = 0 Then Exit Sub
    FrmTerminal.Show
    FrmTerminal.SetFocus
End Sub
'[ popmenu cancel ]'
Private Sub pmenu1cancel_Click()
    Dim l_UsedTime As Long
    l_UsedTime = SelAgn.CusGetTimeUseEx
    
    If AgentCount = 0 Then Exit Sub
    If SelAgn.CustomerName = "" Then Exit Sub
    If l_UsedTime > 10 Then MsgBox MB(23), vbCritical, CbMsgApp: Exit Sub

    ret = MsgBox(MB(22), vbOKCancel, CbMsgApp)
    LogWorker SL(7) '((security log))
    
    If ret = vbOK Then
        SelAgn.CusStop
        If SelTag <> "dump" Then
            SelAgn.NetSend "//kunci:1"
        End If
    End If
End Sub
'[ popmenu cleaning ]'
Private Sub pmenu1clnsub_Click(Index As Integer)
    If AgentCount = 0 Then Exit Sub
    If Index = 0 Then
        SelAgn.NetSend "//cleand:0"
    Else
        SelAgn.NetSend "//cleand:" & (Index - 1)
    End If
    AgentIcon SelTag, "cleaning", True
End Sub
'[ popmenu controlling ]'
Private Sub pmenu1ctlsub_Click(Index As Integer)
    If AgentCount = 0 Then Exit Sub
    Select Case Index
        Case 0
            SelAgn.NetSend "//kunci:1"
        Case 1
            If SelAgn.AgentStatus = VS(3) Then
                SelAgn.NetSend "//kunci:0"
                LogWorker SL(4) '((security log))
            ElseIf SelAgn.AgentStatus = VS(4) Then
                If Mid(CbUserAccess, 3, 1) = 0 Then
                    MsgBox MB(10), vbOKOnly, CbMsgWarn
                Else
                    SelAgn.NetSend "//kunci:0"
                    LogWorker SL(4) '((security log))
                End If
            End If
        Case 2
            SelAgn.NetSend "//sdown:3"
            LogWorker SL(8) '((security log))
        Case 3
            SelAgn.NetSend "//sdown:2"
            LogWorker SL(9) '((security log))
    End Select
End Sub



'###############################################################################################
'# FUNCTIONS
'#
'#
'###############################################################################################
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Layout Measuring
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub LayOutMeasure()
    NegateX = KiraBezaSaizX(FrmMain, Pages(0))
    NegateY = KiraBezaSaizY(FrmMain, Pages(0))
End Sub
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Layout Resizing
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub LayOutSize()
    For Each PictureBox In Pages
        PictureBox.Width = FrmMain.Width - NegateX
        PictureBox.Height = FrmMain.Height - NegateY
    Next
    
    MainPhold.Top = FrmMain.Height - NegateY
    MainPhold.Width = FrmMain.Width - NegateX
    
    Lv1.Width = Pages(0).Width - 15
    Lv1.Height = Pages(0).Height - 15
    DynaLv(1).Width = Pages(1).Width - DynaLv(1).Left - 15
    DynaLv(0).Height = Lv1.Height
    DynaLv(1).Height = Lv1.Height
End Sub

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Service & Merchandise Reset
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub SerCtlReset(Optional EnableCtl As Boolean = False)
    SerImgCb1.ComboItems.Item(VS(1)).Selected = True
    SerImgCb2.Text = VS(1)
    
    SerImgCb2.Enabled = EnableCtl
    SerTxtQty.Enabled = EnableCtl
    SerBtnItm(0).Enabled = EnableCtl
    SerBtnItm(0).Enabled = EnableCtl
    SerLv1.Enabled = EnableCtl
    SerScroll1.Enabled = EnableCtl
    
    SerLv1.ListItems.Clear
    SerScroll1.Value = 999
    SerTxtQty = 1
    SerTotal = 0
    SerTxtJumlah = ""
    SerTxtBaki = ""
    SerTxtBayar = ""
    SerTxtTotalItm = ""
    SerTxtPriItm = ""
End Sub
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Service & Merchandise Control Enable\Disable
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub SerCtlEnable(Optional Enabled As Boolean = True)
    SerImgCb2.Enabled = Enabled
    SerTxtQty.Enabled = Enabled
    SerBtn(0).Enabled = Enabled
    SerBtnItm(0).Enabled = Enabled
    SerBtnItm(1).Enabled = Enabled
    SerLv1.Enabled = Enabled
    SerScroll1.Enabled = Enabled
    If Enabled = False Then
        SerLv1.ListItems.Clear
        SerScroll1.Value = 999
        SerTxtQty = 1
        SerTotal = 0
        SerTxtJumlah = ""
        SerTxtBaki = ""
        SerTxtBayar = ""
        SerTxtTotalItm = ""
        SerTxtPriItm = ""
    End If
End Sub
