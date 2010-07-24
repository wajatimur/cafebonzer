VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   Caption         =   "CafeBonzer v2.0 beta"
   ClientHeight    =   7950
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11340
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
   ScaleHeight     =   7950
   ScaleWidth      =   11340
   Begin CafeBonzer.PageHolder Toolbox 
      Height          =   2460
      Left            =   3075
      TabIndex        =   5
      Top             =   5115
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   4339
      HldrStyle       =   2
      HldrTxt         =   "Toolbox"
      HldrTxtClr      =   4210752
      HldrLne         =   0   'False
      PageHeight      =   2460
      Begin CafeBonzer.XpButton MainNoteBtn 
         Height          =   315
         Index           =   1
         Left            =   45
         TabIndex        =   8
         Top             =   2130
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   556
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
      Begin CafeBonzer.XpButton MainNoteBtn 
         Height          =   315
         Index           =   0
         Left            =   435
         TabIndex        =   7
         Top             =   2130
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   556
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
         MICON           =   "FrmMain.frx":6E08
         PICN            =   "FrmMain.frx":6E24
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox MainNote 
         BackColor       =   &H00C0FFFF&
         Height          =   1710
         Left            =   15
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   375
         Width           =   4005
      End
   End
   Begin VB.PictureBox SideBar 
      Align           =   3  'Align Left
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   0
      Picture         =   "FrmMain.frx":73BE
      ScaleHeight     =   6915
      ScaleWidth      =   3000
      TabIndex        =   2
      Tag             =   "subcontainer"
      Top             =   600
      Width           =   3060
      Begin MSComctlLib.ListView InfoListView1 
         Height          =   870
         Left            =   60
         TabIndex        =   3
         Top             =   1080
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   1535
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImgList16"
         SmallIcons      =   "ImgList16"
         ColHdrIcons     =   "ImgList16"
         ForeColor       =   4210752
         BackColor       =   16777215
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Items"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   2399
         EndProperty
      End
      Begin MSComctlLib.ListView InfoListView 
         Height          =   2025
         Left            =   60
         TabIndex        =   4
         Top             =   2475
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   3572
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImgList16"
         SmallIcons      =   "ImgList16"
         ColHdrIcons     =   "ImgList16"
         ForeColor       =   4210752
         BackColor       =   16777215
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Items"
            Object.Width           =   2434
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   2611
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImgList32 
      Left            =   9570
      Top             =   630
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":8A49
            Key             =   "POS"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":9095
            Key             =   "STAT"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgList16 
      Left            =   10155
      Top             =   630
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   45
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":96DA
            Key             =   "TerminalOnline"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":9A82
            Key             =   "TerminalOffline"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":9E1E
            Key             =   "UserOnline"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":A1C6
            Key             =   "UserOffline"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":A56E
            Key             =   "UserEnded"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":AB0A
            Key             =   "TerminalLock"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":B0A6
            Key             =   "logoff"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":B44E
            Key             =   "boot"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":B9EA
            Key             =   "jalan1"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":BF86
            Key             =   "tamat"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":C522
            Key             =   "mouse"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":CABE
            Key             =   "info"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":D05A
            Key             =   "query"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":D5F6
            Key             =   "dump"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":DB9A
            Key             =   "mesej"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":E136
            Key             =   "kuncibuka"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":E6D2
            Key             =   "off"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":EC6E
            Key             =   "penalaan1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":F20A
            Key             =   "penalaan"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":FB1E
            Key             =   "hoi"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":100BA
            Key             =   "power"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":10456
            Key             =   "net"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":109F2
            Key             =   "gift"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":10F96
            Key             =   "graft"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1153A
            Key             =   "no"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":11AD6
            Key             =   "transfer"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":12072
            Key             =   "term"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1260E
            Key             =   "key"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":12BAA
            Key             =   "TerminalClean"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":13146
            Key             =   "broad"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":134E2
            Key             =   "flagin"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":13A7E
            Key             =   "flagout"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1401A
            Key             =   "meter"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":145B6
            Key             =   "wait"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":14B52
            Key             =   "people"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":150EE
            Key             =   "printer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1568A
            Key             =   "help"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":15C26
            Key             =   "cpu"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":161CE
            Key             =   "paper"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1676A
            Key             =   "TIME"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":16D04
            Key             =   "ITEM"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1709E
            Key             =   "CHAIN"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":17638
            Key             =   "NET"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":17BD2
            Key             =   "PCI"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1816C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgListSnm 
      Left            =   10740
      Top             =   630
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   48
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":18579
            Key             =   "FOLDER"
            Object.Tag             =   "GRP"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":18B15
            Key             =   "USER"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":19569
            Key             =   "ACCESS"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":19B05
            Key             =   "ITEM"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":19EA1
            Key             =   "SERVICES"
            Object.Tag             =   "GRP"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1A43D
            Key             =   "MAN"
            Object.Tag             =   "GRP"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1A555
            Key             =   "FOODS"
            Object.Tag             =   "GRP"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1A8F1
            Key             =   "BEVERAGES"
            Object.Tag             =   "GRP"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1AC8D
            Key             =   "MERCHANDISE"
            Object.Tag             =   "GRP"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1B227
            Key             =   "MAGAZINES"
            Object.Tag             =   "GRP"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1B7C3
            Key             =   "OTHERS"
            Object.Tag             =   "GRP"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1BD5F
            Key             =   "NONE"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1C2FB
            Key             =   "OBJECT"
            Object.Tag             =   "GRP"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1C895
            Key             =   "CHIP"
            Object.Tag             =   "GRP"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1CE2F
            Key             =   "DRIVE"
            Object.Tag             =   "GRP"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1D3C9
            Key             =   "PHONE"
            Object.Tag             =   "GRP"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1D963
            Key             =   "PEN"
            Object.Tag             =   "GRP"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1DEFD
            Key             =   "BRICK"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1E497
            Key             =   "PAPER"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1EA31
            Key             =   "CLIP"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1EFCB
            Key             =   "CRAYON"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1F565
            Key             =   "GEAR"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1FAFF
            Key             =   "FILM"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":20099
            Key             =   "FLOPPY"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":20633
            Key             =   "FLOPPYDRIVE"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":20BCD
            Key             =   "HARDDRIVE"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":21167
            Key             =   "DRAFT"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":21701
            Key             =   "OTO"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":21C9B
            Key             =   "WINDOWS"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":22235
            Key             =   "MAC"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":227CF
            Key             =   "TENT"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":22D69
            Key             =   "DROPS"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":23303
            Key             =   "DICE"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":2389D
            Key             =   "GLUE"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":23E37
            Key             =   "LADYBUG"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":241D1
            Key             =   "OFFICE"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":2476B
            Key             =   "STOP"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":24D05
            Key             =   "SKULL"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":2529F
            Key             =   "SMILE1"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":25839
            Key             =   "SMILE2"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":25DD3
            Key             =   "SMILE3"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":2616D
            Key             =   "SMILE4"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":26507
            Key             =   "SMILE5"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":26AA1
            Key             =   "SMILE6"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":2703B
            Key             =   "BALL"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":275D5
            Key             =   "STAR"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":27B6F
            Key             =   "TOXIC"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":28109
            Key             =   "FRAME"
            Object.Tag             =   "ITEM"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImgList32"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "POS"
            Description     =   "Services & Merchandises"
            Object.ToolTipText     =   "Services & Merchandises"
            Object.Tag             =   "ServicesMerchandise"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "STATISTIC"
            Description     =   "Statistic"
            Object.ToolTipText     =   "Statistic"
            Object.Tag             =   "Statistic"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar MainSbar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   7575
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2469
            MinWidth        =   2469
            TextSave        =   "21:31"
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
            Object.Width           =   9913
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
   Begin MSComctlLib.ListView ListView 
      Height          =   4515
      Left            =   3075
      TabIndex        =   1
      Top             =   585
      Width           =   8280
      _ExtentX        =   14605
      _ExtentY        =   7964
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imglist2"
      SmallIcons      =   "ImgList16"
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Terminal"
         Object.Width           =   2822
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
   End
   Begin VB.Menu MnuMain 
      Caption         =   "Menu"
      Begin VB.Menu MnuMainConfig 
         Caption         =   "Configuration"
      End
      Begin VB.Menu MnuMainSep1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuMainLogout 
         Caption         =   "Logout"
      End
      Begin VB.Menu MnuMainClose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu MnuAgent 
      Caption         =   "Station"
      Begin VB.Menu MnuAgentCtl 
         Caption         =   "Control"
         Begin VB.Menu MnuAgentCtlLock 
            Caption         =   "Lock All"
            Index           =   0
         End
         Begin VB.Menu MnuAgentCtlLock 
            Caption         =   "Lock Unused"
            Index           =   1
         End
         Begin VB.Menu MnuAgentCtlLock 
            Caption         =   "Unlock All"
            Index           =   2
         End
         Begin VB.Menu MnuAgentSep2 
            Caption         =   "-"
         End
         Begin VB.Menu MnuAgentCtlWinExit 
            Caption         =   "Shutdown All"
            Index           =   0
         End
         Begin VB.Menu MnuAgentCtlWinExit 
            Caption         =   "Shutdown Unused"
            Index           =   1
         End
         Begin VB.Menu MnuAgentCtlWinExit 
            Caption         =   "Reboot All"
            Index           =   2
         End
         Begin VB.Menu MnuAgentCtlWinExit 
            Caption         =   "Reboot Unused"
            Index           =   3
         End
      End
      Begin VB.Menu MnuAgentClean 
         Caption         =   "Cleaning"
         Begin VB.Menu MnuAgentCleanSub 
            Caption         =   "All"
            Index           =   0
         End
         Begin VB.Menu MnuAgentCleanSub 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu MnuAgentCleanSub 
            Caption         =   "Temp Folder"
            Index           =   2
         End
         Begin VB.Menu MnuAgentCleanSub 
            Caption         =   "Recycle Bin"
            Index           =   3
         End
         Begin VB.Menu MnuAgentCleanSub 
            Caption         =   "Internet History"
            Index           =   4
         End
         Begin VB.Menu MnuAgentCleanSub 
            Caption         =   "Recent Docs"
            Index           =   5
         End
      End
      Begin VB.Menu MnuAgentBroad 
         Caption         =   "Broadcast"
         Begin VB.Menu MnuAgentBroadSub 
            Caption         =   "Message"
            Index           =   0
         End
      End
      Begin VB.Menu MnuAgentSep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuAgentManager 
         Caption         =   "Agent Manager"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu MnuView 
      Caption         =   "View"
      Visible         =   0   'False
      Begin VB.Menu MnuViewMonPrint 
         Caption         =   "Printing Monitoring"
         Index           =   0
      End
      Begin VB.Menu MnuViewMonRes 
         Caption         =   "Resource Monitoring"
         Index           =   1
      End
      Begin VB.Menu MnuViewMonApp 
         Caption         =   "Application Monitoring"
         Index           =   2
      End
      Begin VB.Menu MnuViewMonTraf 
         Caption         =   "Traffic Monitoring"
         Index           =   3
      End
   End
   Begin VB.Menu MnuTools 
      Caption         =   "Tools"
      Begin VB.Menu MnuToolsBuiltIn 
         Caption         =   "Statistic System"
         Index           =   0
      End
      Begin VB.Menu MnuToolsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuToolsModules 
         Caption         =   "Modules"
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu MnuInfo 
      Caption         =   "Info"
      Begin VB.Menu MnuInfoHelp 
         Caption         =   "Help"
      End
      Begin VB.Menu MenuInfoLiscense 
         Caption         =   "Liscense"
         Begin VB.Menu MenuInfoLiscenseSub 
            Caption         =   "Activate"
            Index           =   0
         End
         Begin VB.Menu MenuInfoLiscenseSub 
            Caption         =   "Transfer"
            Index           =   1
         End
      End
      Begin VB.Menu MnuInfoSep1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuInfoAbout 
         Caption         =   "About.."
      End
   End
   Begin VB.Menu PopMnu1 
      Caption         =   "<popmenu1>"
      Visible         =   0   'False
      Begin VB.Menu PopMnu1Flog 
         Caption         =   "Fast Login"
      End
      Begin VB.Menu PopMnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu PopMnu1Cancel 
         Caption         =   "Cancel User"
      End
      Begin VB.Menu PopMnu1Trans 
         Caption         =   "Transfer PC"
      End
      Begin VB.Menu PopMnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu PopMnu1Cln 
         Caption         =   "Cleaning"
         Begin VB.Menu PopMnu1ClnSub 
            Caption         =   "All"
            Index           =   0
         End
         Begin VB.Menu PopMnu1ClnSub 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu PopMnu1ClnSub 
            Caption         =   "Temp Folder"
            Index           =   2
         End
         Begin VB.Menu PopMnu1ClnSub 
            Caption         =   "Recycle Bin"
            Index           =   3
         End
         Begin VB.Menu PopMnu1ClnSub 
            Caption         =   "Internet History"
            Index           =   4
         End
         Begin VB.Menu PopMnu1ClnSub 
            Caption         =   "Recent Docs"
            Index           =   5
         End
      End
      Begin VB.Menu PopMnu1Ctl 
         Caption         =   "Control"
         Begin VB.Menu PopMnu1CtlSub 
            Caption         =   "Lock Computer"
            Index           =   0
         End
         Begin VB.Menu PopMnu1CtlSub 
            Caption         =   "Unlock Computer"
            Index           =   1
         End
         Begin VB.Menu PopMnu1CtlSub 
            Caption         =   "Reboot Computer"
            Index           =   2
         End
         Begin VB.Menu PopMnu1CtlSub 
            Caption         =   "Shutdown Computer"
            Index           =   3
         End
      End
      Begin VB.Menu PopMnu1Cmc 
         Caption         =   "Communicate"
         Begin VB.Menu PopMnu1CmcSub 
            Caption         =   "Message"
            Index           =   0
         End
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
'    Project    : CafeBonzer
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Private CAsListView As New ClsAutoSize
Private CAsToolBox As New ClsAutoSize

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Form Initialize
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub Form_Initialize()
    FrmMain.Height = 8640
    
    CAsListView.InitControl ListView
    CAsToolBox.InitControl Toolbox
    CAsToolBox.AnchorTop = False
    CFrmConst.InitWindow FrmMain.Hwnd
    CFrmConst.TrackMin True, 500, 800

    LngGap = FrmMain.Toolbox.Top - (FrmMain.ListView.Top + FrmMain.ListView.Height)
    
    FrmMain.Toolbox.PageCollapse = SetGetDb("UiToolBox", True)
    If FrmMain.Toolbox.PageCollapse = True Then
        FrmMain.ListView.Height = (FrmMain.Toolbox.Top + LngGap) - FrmMain.ListView.Top
        CAsListView.NewSize
        CAsToolBox.NewSize
    End If
End Sub

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Form Query Unload
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Cancel = AppExit
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CFrmConst.InitEnd
End Sub

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Form Resizing
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub Form_Resize()
    On Error Resume Next
    
    CAsListView.ResizeControl
    CAsToolBox.ResizeControl
End Sub

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Toolbar
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case Is = "POS"
        FrmPos.Show vbModal
    Case Is = "STATISTIC"
        SecOpenModules ModStatistic
    End Select
End Sub


'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' MAIN LAYOUT
'
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'

Private Sub Toolbox_HolderButtonClick(ByVal Collapse As Boolean)
    If Collapse = True Then
        ListView.Height = (Toolbox.Top + LngGap) - ListView.Top
        CAsListView.NewSize
        CAsToolBox.NewSize
    Else
        CAsListView.NewSizeUndo
        CAsToolBox.NewSizeUndo
        CAsListView.ResizeControl
    End If
    Call SetSaveDb("UiToolBox", CStr(Collapse))
End Sub

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Main Page - Note
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub MainNoteBtn_Click(Index As Integer)
    Select Case Index
    Case 0
        SetSaveDb "AppMainNote", MainNote
    Case 1
        MainNote = ""
        SetSaveDb "AppMainNote", " "
    End Select
End Sub
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ListView | DoubleClick
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub ListView_DblClick()
    If UniAgents.Count = 0 Then Exit Sub
    
    If SelSubItm(1) = VS(1, 2) Then
        FrmLogin.Show vbModal
    Else
        FrmLogout.Show vbModal
    End If
End Sub
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ListView | ItemClick
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub ListView_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call UpdatePanel(Item.Text)
    Call UpdateStat(Item)
End Sub
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ListView | When the mouse goes up
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub ListView_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If UniAgents.Count = 0 Then Exit Sub
    If Button = 2 Then FrmMain.PopupMenu PopMnu1, , X + SideBar.Width, Y + Toolbar.Height
End Sub



'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' Section untuk Menu - SubMenu Object
'
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
'[ Configuration ]
Private Sub MnuMainConfig_Click()
    Call SecOpenModules(ModConfiguration)
End Sub
'[ Exit ]'
Private Sub MnuMainClose_Click()
    Call AppExit
End Sub
'[ Help ]'
Private Sub MnuInfoHelp_Click()
    Call AppLoadInfo(Help)
End Sub
'[ About ]
Private Sub MnuInfoAbout_Click()
    FrmAppAbout.Show vbModal
End Sub
'[ LogOut ]
Private Sub MnuMainLogOut_Click()
    SecUserLog SL(2)
    FrmMain.Hide
    FrmAppPass.Show
End Sub
'[ Lock ]'
Private Sub MnuAgentCtlLock_Click(Index As Integer)
    Select Case Index
        Case 0
            UniAgents.AgentControl LockAll
        Case 1
            UniAgents.AgentControl LockAllUnused
        Case 2
            UniAgents.AgentControl UnlockAll
    End Select
End Sub
'[ Mass Reboot/Shutdown ]'
Private Sub MnuAgentCtlWinExit_Click(Index As Integer)
    Select Case Index
        Case 0
            UniAgents.AgentControl ShutdownAll
        Case 1
            UniAgents.AgentControl ShutdownAllUnused
        Case 2
            UniAgents.AgentControl RebootAll
        Case 3
            UniAgents.AgentControl RebootUnused
    End Select
End Sub
'[ Broadcast ]'
Private Sub MnuAgentBroadSub_Click(Index As Integer)
    Select Case Index
        Case 0
            FrmAgnMsg.Show
            FrmAgnMsg.SetToAll
    End Select
End Sub
'[ Agent Manager ]'
Private Sub MnuAgentManager_Click()
    Call SecOpenModules(ModAgentManager)
End Sub
'[ BuiltIn Tools ]'
Private Sub MnuToolsBuiltIn_Click(Index As Integer)
    Call SecOpenModules((Index + 1) * 4)
End Sub
'[ Modules ]'
Private Sub MnuToolsModules_Click(Index As Integer)
    Dim StrMdlPath As String
    Select Case Index
    Case 1
        StrMdlPath = App.Path & "\CafeSmMgr.exe"
    End Select
    
    Call SecOpenModules(ModExternal, StrMdlPath)
End Sub
'[ Liscense ]'
Private Sub MenuInfoLiscenseSub_Click(Index As Integer)
    Select Case Index
        Case 0
            FrmSysDemo.Show
        Case 1
            Call SecLiscenseTransfer
    End Select
End Sub

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' POPUPMENU
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'[ Fast Login ]'
Private Sub PopMnu1Flog_Click()
    If UniAgents.Count = 0 Then Exit Sub
    ' Get station status
    If SelSubItm(1) = VS(1, 1) Or SelSubItm(1) = VS(1, 3) Then
        MsgBox ST(2, 0), vbInformation, CbMsgWarn
        Exit Sub
    End If
    FrmLogin.FastLogin
    Unload FrmLogin
End Sub
'[ Transfer ]'
Private Sub PopMnu1Trans_Click()
    If AgentSel.AgentStatus = VS(1, 2) Then Exit Sub
    If AgentSel.AgentStatus = VS(1, 3) Then Exit Sub
    FrmAgnTrans.Show
End Sub
'[ Cancel ]'
Private Sub PopMnu1Cancel_Click()
    Dim LngTimeUsed As Long, LngRet As Long
    LngTimeUsed = AgentSel.CusGetTimeUseEx
    
    If UniAgents.Count = 0 Then Exit Sub
    If AgentSel.CustomerName = "" Then Exit Sub
    If LngTimeUsed > 10 Then MsgBox ST(2, 2), vbCritical, CbMsgApp: Exit Sub

    LngRet = MsgBox(ST(2, 1), vbOKCancel, CbMsgApp)
    SecUserLog SL(7)
    
    If LngRet = vbOK Then AgentSel.CusStop
End Sub
'[ Cleaning ]'
Private Sub PopMnu1ClnSub_Click(Index As Integer)
    If UniAgents.Count = 0 Then Exit Sub
    If Index = 0 Then
        AgentSel.Commands.ConCleaning (Clean)
    Else
        AgentSel.Commands.ConCleaning (Index - 1)
    End If
    AgentSel.AgentSmallIcon = "TerminalClean"
End Sub
'[ Controlling ]'
Private Sub PopMnu1CtlSub_Click(Index As Integer)
    ' 0 = Lock
    ' 1 = Unlock
    ' 2 = Reboot
    ' 3 = Shutdown
    
    If UniAgents.Count = 0 Then Exit Sub
    Select Case Index
        Case 0
            AgentSel.Commands.ConLock (TerminalLock)
        Case 1
            If AgentSel.AgentStatus = VS(1, 1) Then
                AgentSel.Commands.ConLock (TerminalUnlock)
            ElseIf AgentSel.AgentStatus = VS(1, 2) Then
                If SecAccessRequest(TaskUnlock) = False Then
                    MsgBox ST(1, 1), vbOKOnly, CbMsgWarn
                Else
                    AgentSel.Commands.ConLock (TerminalUnlock)
                End If
            End If
        Case 2
            AgentSel.Commands.ConExitWin (Reboot)
        Case 3
            AgentSel.Commands.ConExitWin (Shutdown)
    End Select
End Sub
'[ Messagging ]'
Private Sub PopMnu1CmcSub_Click(Index As Integer)
    Select Case Index
        Case 0
            FrmAgnMsg.Show
    End Select
End Sub



'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' FUNCTION
'
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'

