VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmSnmMg 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Services & Merchandise Manager"
   ClientHeight    =   4800
   ClientLeft      =   255
   ClientTop       =   1980
   ClientWidth     =   9735
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
   Icon            =   "FrmSnm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   9735
   StartUpPosition =   2  'CenterScreen
   Begin CbSnmMgr.Line3D Line3D1 
      Height          =   45
      Left            =   4530
      TabIndex        =   20
      Top             =   525
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   79
      horizon         =   -1  'True
   End
   Begin VB.TextBox ItemInfoTxt 
      BackColor       =   &H00C0FFC0&
      Height          =   330
      Index           =   4
      Left            =   6270
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   3435
      Width           =   3165
   End
   Begin VB.TextBox ItemInfoTxt 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Height          =   330
      Index           =   3
      Left            =   6270
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   2925
      Width           =   915
   End
   Begin VB.TextBox ItemInfoTxt 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Height          =   330
      Index           =   2
      Left            =   8115
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2925
      Width           =   840
   End
   Begin VB.TextBox ItemInfoTxt 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   1
      Left            =   6270
      TabIndex        =   9
      Top             =   2370
      Width           =   915
   End
   Begin VB.TextBox ItemInfoTxt 
      Height          =   330
      Index           =   0
      Left            =   6270
      TabIndex        =   7
      Top             =   750
      Width           =   3210
   End
   Begin MSComctlLib.ImageCombo IcGroup 
      Height          =   330
      Left            =   60
      TabIndex        =   6
      Top             =   120
      Width           =   4410
      _ExtentX        =   7779
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
      ImageList       =   "Iml"
   End
   Begin CbSnmMgr.Line3D GuiLine 
      Height          =   45
      Left            =   15
      TabIndex        =   5
      Top             =   0
      Width           =   9690
      _ExtentX        =   17092
      _ExtentY        =   79
      horizon         =   -1  'True
   End
   Begin CbSnmMgr.XpButton BtnMenu 
      Height          =   450
      Index           =   0
      Left            =   8235
      TabIndex        =   1
      ToolTipText     =   "Save Changes"
      Top             =   3960
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   794
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmSnm.frx":23D2
      PICN            =   "FrmSnm.frx":23EE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.StatusBar GuiStat 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   4485
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12674
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
   Begin MSComctlLib.ImageList Iml 
      Left            =   105
      Top             =   585
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
            Picture         =   "FrmSnm.frx":2988
            Key             =   "FOLDER"
            Object.Tag             =   "GRP"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":2F24
            Key             =   "USER"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":3978
            Key             =   "ACCESS"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":3F14
            Key             =   "ITEM"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":42B0
            Key             =   "SERVICES"
            Object.Tag             =   "GRP"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":484C
            Key             =   "MAN"
            Object.Tag             =   "GRP"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":4964
            Key             =   "FOODS"
            Object.Tag             =   "GRP"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":4D00
            Key             =   "BEVERAGES"
            Object.Tag             =   "GRP"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":509C
            Key             =   "MERCHANDISE"
            Object.Tag             =   "GRP"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":5636
            Key             =   "MAGAZINES"
            Object.Tag             =   "GRP"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":5BD2
            Key             =   "OTHERS"
            Object.Tag             =   "GRP"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":616E
            Key             =   "NONE"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":670A
            Key             =   "OBJECT"
            Object.Tag             =   "GRP"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":6CA4
            Key             =   "CHIP"
            Object.Tag             =   "GRP"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":723E
            Key             =   "DRIVE"
            Object.Tag             =   "GRP"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":77D8
            Key             =   "PHONE"
            Object.Tag             =   "GRP"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":7D72
            Key             =   "PEN"
            Object.Tag             =   "GRP"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":830C
            Key             =   "BRICK"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":88A6
            Key             =   "PAPER"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":8E40
            Key             =   "CLIP"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":93DA
            Key             =   "CRAYON"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":9974
            Key             =   "GEAR"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":9F0E
            Key             =   "FILM"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":A4A8
            Key             =   "FLOPPY"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":AA42
            Key             =   "FLOPPYDRIVE"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":AFDC
            Key             =   "HARDDRIVE"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":B576
            Key             =   "DRAFT"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":BB10
            Key             =   "OTO"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":C0AA
            Key             =   "WINDOWS"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":C644
            Key             =   "MAC"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":CBDE
            Key             =   "TENT"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":D178
            Key             =   "DROPS"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":D712
            Key             =   "DICE"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":DCAC
            Key             =   "GLUE"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":E246
            Key             =   "LADYBUG"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":E5E0
            Key             =   "OFFICE"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":EB7A
            Key             =   "STOP"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":F114
            Key             =   "SKULL"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":F6AE
            Key             =   "SMILE1"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":FC48
            Key             =   "SMILE2"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":101E2
            Key             =   "SMILE3"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":1057C
            Key             =   "SMILE4"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":10916
            Key             =   "SMILE5"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":10EB0
            Key             =   "SMILE6"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":1144A
            Key             =   "BALL"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":119E4
            Key             =   "STAR"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":11F7E
            Key             =   "TOXIC"
            Object.Tag             =   "ITEM"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSnm.frx":12518
            Key             =   "FRAME"
            Object.Tag             =   "ITEM"
         EndProperty
      EndProperty
   End
   Begin CbSnmMgr.XpButton BtnMenu 
      Height          =   450
      Index           =   1
      Left            =   8715
      TabIndex        =   2
      ToolTipText     =   "Add New Item"
      Top             =   3960
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   794
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmSnm.frx":12AB2
      PICN            =   "FrmSnm.frx":12ACE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CbSnmMgr.XpButton BtnMenu 
      Height          =   450
      Index           =   2
      Left            =   9195
      TabIndex        =   3
      ToolTipText     =   "Delete Selected Item"
      Top             =   3960
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   794
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmSnm.frx":13068
      PICN            =   "FrmSnm.frx":13084
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView LvItem 
      Height          =   3900
      Left            =   60
      TabIndex        =   4
      Top             =   540
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   6879
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "Iml"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item ID"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Price"
         Object.Width           =   1764
      EndProperty
   End
   Begin CbSnmMgr.XpButton BtnMenu 
      Height          =   375
      Index           =   3
      Left            =   9030
      TabIndex        =   13
      ToolTipText     =   "Stock Management"
      Top             =   2895
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   661
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmSnm.frx":1341E
      PICN            =   "FrmSnm.frx":1343A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView ItemSymLv 
      Height          =   945
      Left            =   6270
      TabIndex        =   14
      Top             =   1245
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   1667
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "Iml"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
   Begin VB.Image HdrItemImg 
      Height          =   315
      Left            =   4650
      Top             =   120
      Width           =   315
   End
   Begin VB.Label HdrItemId 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CP01001"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   5055
      TabIndex        =   21
      Top             =   135
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Purchased :"
      Height          =   195
      Index           =   5
      Left            =   4710
      TabIndex        =   19
      Top             =   3480
      Width           =   1425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cosumed :"
      Height          =   195
      Index           =   4
      Left            =   5190
      TabIndex        =   17
      Top             =   2970
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Symbol :"
      Height          =   195
      Index           =   3
      Left            =   5355
      TabIndex        =   15
      Top             =   1260
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock :"
      Height          =   195
      Index           =   2
      Left            =   7395
      TabIndex        =   12
      Top             =   2970
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Price :"
      Height          =   195
      Index           =   0
      Left            =   5580
      TabIndex        =   10
      Top             =   2415
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name :"
      Height          =   195
      Index           =   1
      Left            =   5040
      TabIndex        =   8
      Top             =   795
      Width           =   1095
   End
   Begin VB.Menu menu1 
      Caption         =   "Menu"
      Begin VB.Menu menu1close 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu menu2 
      Caption         =   "Tools"
      Begin VB.Menu menu2itemfunc 
         Caption         =   "Group Add"
         Index           =   0
      End
      Begin VB.Menu menu2itemfunc 
         Caption         =   "Group Delete"
         Index           =   1
      End
      Begin VB.Menu menu2itemfunc 
         Caption         =   "Item Add"
         Index           =   2
      End
      Begin VB.Menu menu2itemfunc 
         Caption         =   "Item Delete"
         Index           =   3
      End
   End
End
Attribute VB_Name = "FrmSnmMg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Form_Load] - Form Entry Points
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Private Sub Form_Load()
    If LoadGroups > 0 Then
        If LoadItems > 0 Then
            Call LoadSymbol(ItemSymLv)
            Call ItemGet(LvItem.SelectedItem)
        End If
    Else
        MsgBox Var(0), vbInformation, CbMsgWarn
        Call CtlDis(1)
    End If
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Form_Resize] - Form Resize Event
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Private Sub Form_Resize()
    Call AutoColumn(LvItem, 2)
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [IcGroup] - Image Combo .Click
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Private Sub IcGroup_Click()
    LvItem.ListItems.Clear
    If LoadItems > 0 Then
        Call ItemGet(LvItem.SelectedItem)
    End If
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [BtnMenu] - Button menu
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Private Sub BtnMenu_Click(Index As Integer)
On Error GoTo ErrInt
    Select Case Index
    Case 0  'Save
        Call ItemSave(LvItem.SelectedItem.Key, ItemInfoTxt(0), ItemInfoTxt(1), ItemSymLv.SelectedItem.Key)
    Case 1  'Add Item
        FrmItemAdd.Show vbModal
    Case 2  'Delete Item
        Call ItemDel
    Case 3  'Stock Management
        FrmStockMg.Init LvItem.SelectedItem.Key, ItemInfoTxt(2)
    End Select
Exit Sub

ErrInt:
    ErrLog Err, "Snm Manager | BtnMenu_Click"
End Sub


Private Sub LvItem_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call ItemGet(Item.Key)
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [menu1close] - Close Module
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Private Sub menu1close_Click()
    Set FrmPosMg = Nothing
    Unload Me
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [menu2itemfunc] - Item and Group Function
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Private Sub menu2itemfunc_Click(Index As Integer)
    Select Case Index
        Case 0
            FrmGrpAdd.Show vbModal
        Case 1
            Call GroupDel
        Case 2
            Call BtnMenu_Click(2)
        Case 3
            Call BtnMenu_Click(3)
    End Select
End Sub
